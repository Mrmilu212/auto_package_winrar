import os
import time
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk
from pathlib import Path
from typing import Iterable
import subprocess
from auto_package.core.compress import find_rar_exe, run_rar_archive, run_double_compress, run_triple_compress
from auto_package.core.extract import find_winrar_exe, run_auto_extract
from auto_package.core.utils import decode_drop_path, _commit_outputs_atomic
from auto_package.config.settings import _load_window_geometry, _save_window_geometry, _load_upload_paths, _save_upload_paths
from auto_package.gui.pages.home import Home
from auto_package.gui.pages.transfer import Transfer
from auto_package.gui.pages.upload_path import UploadPath
from auto_package.utils.logging_config import setup_logging, get_logger

try:
    import windnd
except ImportError:
    windnd = None

class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        
        # 初始化日志
        self.logger = setup_logging()
        self.logger.info('程序启动')
        
        self.title("拖放文件夹 — WinRAR 打包")
        self.minsize(480, 320)
        
        # 加载上次的窗口大小
        saved_geometry = _load_window_geometry()
        if saved_geometry:
            width, height = saved_geometry
            self.geometry(f"{width}x{height}")
        else:
            self.geometry("560x400")

        self._rar = find_rar_exe()
        self._winrar = find_winrar_exe()
        self._busy = False
        self._var_double = tk.BooleanVar(value=True)
        self._var_triple = tk.BooleanVar(value=True)
        self._status_var = tk.StringVar(value="")
        self._phase_text = "就绪"
        self._job_start_t: float | None = None
        self._job_tick_id: str | None = None
        self._cancel_ev: threading.Event | None = None
        self._current_proc: subprocess.Popen[str] | None = None
        self._close_after_cancel = False
        # 解压相关变量
        self._extract_busy = False
        self._extract_cancel_ev: threading.Event | None = None
        self._extract_current_proc: subprocess.Popen[str] | None = None
        
        # 上传线程计数
        self._upload_thread_count = 0
        self._max_upload_threads = 5
        
        # 百度网盘上传选项
        self._var_upload = tk.BooleanVar(value=False)
        
        # 上传路径配置
        self._upload_path, self._upload_path_history = _load_upload_paths()

        # 标签页管理
        self._notebook = ttk.Notebook(self)
        self._notebook.pack(fill=tk.BOTH, expand=True)
        
        # 主页标签
        self._home_frame = tk.Frame(self._notebook)
        self._notebook.add(self._home_frame, text="主页")
        
        # 传输标签
        self._transfer_frame = tk.Frame(self._notebook)
        self._notebook.add(self._transfer_frame, text="传输")
        
        # 上传路径管理标签
        self._upload_path_frame = tk.Frame(self._notebook)
        self._notebook.add(self._upload_path_frame, text="上传路径")
        
        # 初始化页面
        self._home_page = Home(self._home_frame, self)
        self._transfer_page = Transfer(self._transfer_frame, self)
        self._upload_path_page = UploadPath(self._upload_path_frame, self)
        
        self._wire_mode_vars()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _wire_mode_vars(self) -> None:
        def on_triple_changed(*_args) -> None:
            if self._var_triple.get():
                self._var_double.set(True)

        def on_double_changed(*_args) -> None:
            if not self._var_double.get() and self._var_triple.get():
                self._var_triple.set(False)

        self._var_triple.trace_add("write", on_triple_changed)
        self._var_double.trace_add("write", on_double_changed)



    def _fmt_time(self, seconds: float) -> str:
        seconds = max(0.0, float(seconds))
        m, s = divmod(int(seconds + 0.5), 60)
        h, m = divmod(m, 60)
        if h:
            return f"{h:02d}:{m:02d}:{s:02d}"
        return f"{m:02d}:{s:02d}"

    def _set_phase(self, step: int | None, total: int) -> None:
        if self._extract_busy:
            if total <= 1:
                self._phase_text = "正在解压"
                return
            if step is None:
                self._phase_text = "正在解压"
                return
            self._phase_text = f"正在进行第{step}次解压"
        else:
            if total <= 1:
                self._phase_text = "正在压缩" if self._busy else "就绪"
                return
            if step is None:
                self._phase_text = "正在压缩" if self._busy else "就绪"
                return
            self._phase_text = f"正在进行第{step}次压缩"

    def _set_status(self, elapsed: float) -> None:
        self._status_var.set(f"{self._phase_text}    已用: {self._fmt_time(elapsed)}")

    def _start_job_timer(self) -> None:
        self._job_start_t = time.monotonic()
        if self._job_tick_id is not None:
            try:
                self.after_cancel(self._job_tick_id)
            except Exception:
                pass
            self._job_tick_id = None

        def tick() -> None:
            if not self._busy and not self._extract_busy or self._job_start_t is None:
                self._job_tick_id = None
                return
            elapsed = time.monotonic() - self._job_start_t
            self._set_status(elapsed)
            self._job_tick_id = self.after(200, tick)

        tick()

    def _stop_job_timer(self) -> None:
        self._job_start_t = None
        if self._job_tick_id is not None:
            try:
                self.after_cancel(self._job_tick_id)
            except Exception:
                pass
            self._job_tick_id = None
        self._phase_text = "就绪"
        self._set_status(0.0)

    def _set_current_proc(self, p: subprocess.Popen[str] | None) -> None:
        self._current_proc = p

    def _set_extract_current_proc(self, p: subprocess.Popen[str] | None) -> None:
        self._extract_current_proc = p

    def _cancel_current(self) -> None:
        if self._cancel_ev is not None:
            self._cancel_ev.set()
        # 尝试立刻终止当前进程（_rar_run 里也会检测并终止）
        p = self._current_proc
        if p is not None:
            try:
                p.terminate()
            except Exception:
                pass

    def _cancel_extract(self) -> None:
        """取消解压操作"""
        if self._extract_cancel_ev is not None:
            self._extract_cancel_ev.set()
        # 尝试立刻终止当前进程
        p = self._extract_current_proc
        if p is not None:
            try:
                p.terminate()
            except Exception:
                pass

    def _on_close(self) -> None:
        if self._busy:
            self._close_after_cancel = True
            self._cancel_current()
            # 等任务结束后在 done() 里关闭
            self._log_line("正在取消并退出…")
            self.logger.info('正在取消压缩操作并退出')
            return
        if self._extract_busy:
            self._close_after_cancel = True
            self._cancel_extract()
            # 等任务结束后在 extract_done() 里关闭
            self._log_line("正在取消解压并退出…")
            self.logger.info('正在取消解压操作并退出')
            return
        # 保存窗口大小
        try:
            width = self.winfo_width()
            height = self.winfo_height()
            if width > 0 and height > 0:
                _save_window_geometry(width, height)
                self.logger.debug('保存窗口大小: %dx%d', width, height)
        except Exception as e:
            self.logger.debug('保存窗口大小失败: %s', e)
            pass
        self.logger.info('程序退出')
        self.destroy()

    def _hook_drop(self) -> None:
        if windnd is None:
            self._log_line("未安装 windnd，拖放不可用。请运行: pip install -r requirements.txt")
            self.logger.warning('未安装 windnd，拖放不可用')
            return
        try:
            # 压缩拖放区域
            kw = {"func": self._on_drop_files, "force_unicode": True}
            windnd.hook_dropfiles(self._home_page.drop_frame, **kw)
            windnd.hook_dropfiles(self._home_page.lbl_drop, **kw)
            # 解压拖放区域
            kw_extract = {"func": self._on_drop_extract, "force_unicode": True}
            windnd.hook_dropfiles(self._home_page.extract_drop_frame, **kw_extract)
            windnd.hook_dropfiles(self._home_page.lbl_extract_drop, **kw_extract)
            self.logger.debug('拖放注册成功')
        except Exception as e:
            self._log_line(f"拖放注册失败: {e}")
            self.logger.warning('拖放注册失败: %s', e)

    def _clear_log(self) -> None:
        self._home_page.log.configure(state=tk.NORMAL)
        self._home_page.log.delete("1.0", tk.END)
        self._home_page.log.configure(state=tk.DISABLED)
    
    def _open_log_folder(self) -> None:
        """打开日志文件夹"""
        import os
        project_root = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
        log_dir = os.path.join(project_root, 'logs')
        os.makedirs(log_dir, exist_ok=True)
        try:
            os.startfile(log_dir)
            self.logger.debug('打开日志文件夹: %s', log_dir)
        except Exception as e:
            self._log_line(f"打开日志文件夹失败: {e}")
            self.logger.error('打开日志文件夹失败: %s', e)

    def _log_line(self, msg: str) -> None:
        if hasattr(self, '_home_page') and hasattr(self._home_page, 'log'):
            self._home_page.log.configure(state=tk.NORMAL)
            self._home_page.log.insert(tk.END, msg + "\n")
            self._home_page.log.see(tk.END)
            self._home_page.log.configure(state=tk.DISABLED)

    def _on_drop_files(self, paths: Iterable[bytes]) -> None:
        if self._busy:
            self._log_line("正在打包，请稍候…")
            self.logger.debug('用户尝试拖放文件，但当前正在打包')
            return
        if not self._rar:
            messagebox.showerror("WinRAR", "未找到 Rar.exe，无法打包。")
            self.logger.error('未找到 Rar.exe，无法打包')
            return

        decoded = [
            decode_drop_path(p) if isinstance(p, bytes) else str(p) for p in paths
        ]
        items = [Path(p) for p in decoded if Path(p).exists()]
        if not items:
            messagebox.showinfo("提示", "请拖放文件或文件夹。")
            self.logger.debug('用户拖放了不存在的文件')
            return
        file_paths = [str(item) for item in items]
        self.logger.info('用户拖放文件: %s', ', '.join(file_paths))
        self._start_pack(items)

    def _pick_folder(self) -> None:
        d = filedialog.askdirectory(title="选择要打包的文件夹")
        if not d:
            self.logger.debug('用户取消选择文件夹')
            return
        self.logger.info('用户选择文件夹: %s', d)
        self._on_drop_files([os.fsencode(d)])

    def _start_pack(self, items: list[Path]) -> None:
        self._busy = True
        self._home_page.lbl_drop.configure(text="正在打包…")
        if len(items) == 1:
            self._log_line(f"开始: {items[0]}")
            self.logger.info('开始压缩: %s', items[0])
        else:
            self._log_line(f"开始: {items[0]}（共 {len(items)} 个输入）")
            self.logger.info('开始压缩: %s（共 %d 个输入）', items[0], len(items))
        
        # 记录压缩参数
        compress_mode = '三次压缩' if self._var_triple.get() else '二次压缩' if self._var_double.get() else '一次压缩'
        self.logger.debug('压缩模式: %s', compress_mode)
        
        # 状态栏阶段文案（不显示百分比/剩余时间）
        if self._var_triple.get():
            self._set_phase(1, 3)
        elif self._var_double.get():
            self._set_phase(1, 2)
        else:
            self._set_phase(None, 1)
        self._start_job_timer()
        self._home_page.btn_cancel.configure(state=tk.NORMAL)
        self._close_after_cancel = False
        self._cancel_ev = threading.Event()
        self._set_current_proc(None)

        def work() -> None:
            assert self._rar is not None
            target_dir = items[0].parent
            temp_dir = target_dir / f".apwr_tmp_{os.urandom(8).hex()}"
            try:
                temp_dir.mkdir(parents=False, exist_ok=False)
            except Exception as e:
                return self.after(
                    0, self._pack_done, False, f"创建临时目录失败: {e}", None
                )

            def progress_cb(_pct: int | None, _elapsed: float) -> None:
                # 已用时间由界面计时器实时刷新；这里不再处理进度/剩余时间
                return

            def phase_cb(step: int, total: int) -> None:
                self.after(0, self._set_phase, step, total)

            def proc_cb(p: subprocess.Popen[str] | None) -> None:
                self._set_current_proc(p)

            ok: bool = False
            msg: str = "初始化失败"
            extra: str | None = None
            is_volumes = False
            temp_outputs: list[Path] = []

            try:
                extra = None
                if self._var_triple.get():
                    # 三次：最终为分卷
                    ok_t, pattern_or_err, extra = run_triple_compress(
                        self._rar,
                        items,
                        output_dir=temp_dir,
                        progress_cb=progress_cb,
                        phase_cb=phase_cb,
                        cancel_ev=self._cancel_ev,
                        proc_cb=proc_cb,
                    )
                    if not ok_t:
                        ok, msg = False, pattern_or_err
                    else:
                        # 在 temp_dir 收集 *.part*.rar
                        # pattern_or_err 形如 "<stem>.part*.rar"
                        stem = pattern_or_err.split(".part*")[0]
                        # stem 可能包含路径，也可能是文件名；用目录扫描更可靠
                        # 直接收集 temp_dir 下所有 part*.rar
                        temp_outputs = sorted(temp_dir.glob("*.part*.rar"))
                        if not temp_outputs:
                            ok, msg = False, "未找到分卷产物（内部错误）"
                        else:
                            ok, msg = True, "OK"
                            is_volumes = True
                elif self._var_double.get():
                    ok_d, out_path_or_err, extra = run_double_compress(
                        self._rar,
                        items,
                        output_dir=temp_dir,
                        progress_cb=progress_cb,
                        phase_cb=phase_cb,
                        cancel_ev=self._cancel_ev,
                        proc_cb=proc_cb,
                    )
                    if not ok_d:
                        ok, msg = False, out_path_or_err
                    else:
                        temp_outputs = [Path(out_path_or_err)]
                        ok, msg = True, "OK"
                else:
                    phase_cb(1, 1)
                    ok_o, out_or_err, extra = run_rar_archive(
                        self._rar,
                        items,
                        output_dir=temp_dir,
                        progress_cb=progress_cb,
                        cancel_ev=self._cancel_ev,
                        proc_cb=proc_cb,
                    )
                    if not ok_o:
                        ok, msg = False, str(out_or_err)
                    else:
                        assert isinstance(out_or_err, Path)
                        temp_outputs = [out_or_err]
                        ok, msg = True, "OK"

                if self._cancel_ev is not None and self._cancel_ev.is_set():
                    ok, msg = False, "已取消"

                if ok:
                    if not temp_outputs:
                        ok, msg = False, "未找到压缩产物（内部错误）"
                    else:
                        okc, outmsg, finals = _commit_outputs_atomic(
                            temp_outputs, target_dir, is_volumes=is_volumes
                        )
                        if not okc:
                            ok, msg = False, outmsg
                        else:
                            msg = outmsg
            finally:
                try:
                    import shutil
                    shutil.rmtree(temp_dir, ignore_errors=True)
                except Exception:
                    pass

            # 确保无论如何都会调用 _pack_done
            self.after(0, self._pack_done, ok, msg, extra)

        threading.Thread(target=work, daemon=True).start()

    def _pack_done(self, ok: bool, msg: str, extra: str | None) -> None:
        self._busy = False
        self._home_page.lbl_drop.configure(text="拖放文件/文件夹")
        self._home_page.btn_cancel.configure(state=tk.DISABLED)
        self._cancel_ev = None
        self._set_current_proc(None)
        self._stop_job_timer()

        if ok:
            self._log_line(f"成功: {msg}")
            self.logger.info('压缩成功: %s', msg)
            messagebox.showinfo("成功", msg)
            
            # 检查是否需要上传到百度网盘
            if self._var_upload.get():
                # 检查上传线程数限制
                if self._upload_thread_count >= self._max_upload_threads:
                    self._log_line(f"上传线程数达到上限 ({self._max_upload_threads})，请稍后再试")
                    self.logger.debug('上传线程数达到上限: %d/%d', self._upload_thread_count, self._max_upload_threads)
                else:
                    self._log_line("开始上传到百度网盘...")
                    self.logger.info('开始上传到百度网盘')
                    # 创建后台线程执行上传，不阻塞GUI
                    def upload_task():
                        # 创建取消事件
                        cancel_ev = threading.Event()
                        # 保存进程引用
                        current_proc = None
                        
                        def proc_cb(p):
                            nonlocal current_proc
                            current_proc = p
                        
                        try:
                            # 增加上传线程计数
                            self._upload_thread_count += 1
                            self.after(0, self._log_line, f"当前上传线程数: {self._upload_thread_count}/{self._max_upload_threads}")
                            self.logger.debug('上传线程启动: %d/%d', self._upload_thread_count, self._max_upload_threads)
                            
                            # 获取压缩产物路径
                            output_path = Path(msg)
                            
                            # 取消回调函数
                            def cancel_callback():
                                if not cancel_ev.is_set():
                                    cancel_ev.set()
                                    self.after(0, self._log_line, f"取消上传任务: {output_path.name}")
                                    self.logger.info('取消上传任务: %s', output_path.name)
                            
                            # 添加传输任务，传递取消回调
                            self.after(0, lambda: self._transfer_page.add_transfer_task(output_path.name, cancel_callback))
                            
                            # 检查路径是否存在
                            if not output_path.exists():
                                self.after(0, self._log_line, f"上传失败: 产物路径不存在: {output_path}")
                                self.logger.error('上传失败: 产物路径不存在: %s', output_path)
                                return
                            
                            # 调用上传模块
                            from auto_package.core.upload import upload_to_baidu_pan
                            
                            # 确保日志回调在GUI线程中执行
                            def log_callback(message):
                                self.after(0, self._log_line, message)
                            
                            ok, result_msg = upload_to_baidu_pan(
                                output_path=output_path,
                                upload_base=self._upload_path,
                                log_cb=log_callback,
                                cancel_ev=cancel_ev,
                                proc_cb=proc_cb,
                            )
                            if not ok:
                                self.after(0, self._log_line, f"上传失败: {result_msg}")
                                self.logger.error('上传失败: %s', result_msg)
                        except Exception as e:
                            self.after(0, self._log_line, f"上传异常: {e}")
                            self.logger.error('上传异常: %s', e, exc_info=True)
                        finally:
                            # 减少上传线程计数
                            self._upload_thread_count -= 1
                            self.after(0, self._log_line, f"上传线程结束，当前上传线程数: {self._upload_thread_count}/{self._max_upload_threads}")
                            self.logger.debug('上传线程结束: %d/%d', self._upload_thread_count, self._max_upload_threads)
                            # 移除传输任务
                            self.after(0, lambda: self._transfer_page.remove_transfer_task(output_path.name))
                    
                    # 启动后台线程
                    import threading
                    threading.Thread(target=upload_task, daemon=True).start()

    def start_upload(self, output_path: Path):
        """启动上传任务（供外部调用）"""
        self.logger.info('外部调用上传: %s', output_path)
        
        def upload_task():
            cancel_ev = threading.Event()
            current_proc = None
            
            def proc_cb(p):
                nonlocal current_proc
                current_proc = p
            
            try:
                self._upload_thread_count += 1
                self.after(0, self._log_line, f"当前上传线程数: {self._upload_thread_count}/{self._max_upload_threads}")
                self.logger.debug('上传线程启动: %d/%d', self._upload_thread_count, self._max_upload_threads)
                
                def cancel_callback():
                    if not cancel_ev.is_set():
                        cancel_ev.set()
                        self.after(0, self._log_line, f"取消上传任务: {output_path.name}")
                        self.logger.info('取消上传任务: %s', output_path.name)
                
                self.after(0, lambda: self._transfer_page.add_transfer_task(output_path.name, cancel_callback))
                
                if not output_path.exists():
                    self.after(0, self._log_line, f"上传失败: 路径不存在: {output_path}")
                    self.logger.error('上传失败: 路径不存在: %s', output_path)
                    return
                
                from auto_package.core.upload import upload_to_baidu_pan
                
                def log_callback(message):
                    self.after(0, self._log_line, message)
                
                ok, result_msg = upload_to_baidu_pan(
                    output_path=output_path,
                    upload_base=self._upload_path,
                    log_cb=log_callback,
                    cancel_ev=cancel_ev,
                    proc_cb=proc_cb,
                )
                if not ok:
                    self.after(0, self._log_line, f"上传失败: {result_msg}")
                    self.logger.error('上传失败: %s', result_msg)
                else:
                    self.logger.info('上传成功: %s', output_path.name)
            except Exception as e:
                self.after(0, self._log_line, f"上传异常: {e}")
                self.logger.error('上传异常: %s', e, exc_info=True)
            finally:
                self._upload_thread_count -= 1
                self.after(0, self._log_line, f"上传线程结束，当前上传线程数: {self._upload_thread_count}/{self._max_upload_threads}")
                self.logger.debug('上传线程结束: %d/%d', self._upload_thread_count, self._max_upload_threads)
                self.after(0, lambda: self._transfer_page.remove_transfer_task(output_path.name))
        
        if self._upload_thread_count >= self._max_upload_threads:
            self._log_line(f"上传线程数达到上限 ({self._max_upload_threads})，请稍后再试")
            self.logger.debug('上传线程数达到上限: %d/%d', self._upload_thread_count, self._max_upload_threads)
            return
        
        self._log_line("开始上传到百度网盘...")
        self.logger.info('开始上传到百度网盘: %s', output_path.name)
        import threading
        threading.Thread(target=upload_task, daemon=True).start()

    def _on_drop_extract(self, paths: Iterable[bytes]) -> None:
        if self._extract_busy:
            self._log_line("正在解压，请稍候…")
            self.logger.debug('用户尝试拖放文件，但当前正在解压')
            return
        if not self._winrar:
            messagebox.showerror("WinRAR", "未找到 WinRAR.exe，无法解压。")
            self.logger.error('未找到 WinRAR.exe，无法解压')
            return

        decoded = [
            decode_drop_path(p) if isinstance(p, bytes) else str(p) for p in paths
        ]
        items = [Path(p) for p in decoded if Path(p).exists()]
        if not items:
            messagebox.showinfo("提示", "请拖放文件。")
            self.logger.debug('用户拖放了不存在的文件')
            return
        if len(items) > 1:
            messagebox.showinfo("提示", "一次只能解压一个文件。")
            self.logger.debug('用户拖放了多个文件，只支持单个文件解压')
            return
        self.logger.info('用户拖放解压文件: %s', items[0])
        self._start_extract(items[0])

    def _start_extract(self, input_path: Path) -> None:
        if not self._winrar:
            messagebox.showerror("WinRAR", "未找到 WinRAR.exe，无法解压。")
            self.logger.error('未找到 WinRAR.exe，无法解压')
            return

        self._extract_busy = True
        self._home_page.lbl_extract_drop.configure(text="正在解压…")
        self._log_line(f"开始解压: {input_path}")
        self.logger.info('开始解压: %s', input_path)
        self._start_job_timer()
        self._home_page.btn_cancel_extract.config(state=tk.NORMAL)
        self._extract_cancel_ev = threading.Event()
        self._extract_current_proc = None
        self._job_start_t = time.monotonic()
        self._phase_text = "正在解压"
        self._set_status(0.0)

        def progress_cb(_percent: int | None, elapsed: float) -> None:
            self._set_status(elapsed)

        def phase_cb(step: int, total: int) -> None:
            self._set_phase(step, total)
            self._set_status(time.monotonic() - self._job_start_t)

        def proc_cb(p: subprocess.Popen[str] | None) -> None:
            self._set_extract_current_proc(p)

        def extract_done():
            ok, msg, extra = run_auto_extract(
                self._winrar,
                input_path,
                progress_cb=progress_cb,
                phase_cb=phase_cb,
                cancel_ev=self._extract_cancel_ev,
                proc_cb=proc_cb,
            )
            self.after(0, self._extract_done, ok, msg, extra)

        threading.Thread(target=extract_done, daemon=True).start()

    def _extract_done(self, ok: bool, msg: str, extra: str | None) -> None:
        """解压完成回调"""
        self._extract_busy = False
        self._home_page.lbl_extract_drop.configure(text="拖放要解压的文件")
        self._home_page.btn_cancel_extract.config(state=tk.DISABLED)
        self._extract_cancel_ev = None
        self._extract_current_proc = None
        self._stop_job_timer()

        if ok:
            self._log_line(f"解压成功: {msg}")
            self.logger.info('解压成功: %s', msg)
        else:
            self._log_line(f"解压失败: {msg}")
            self.logger.error('解压失败: %s', msg)
            if extra:
                self._log_line(f"中间文件: {extra}")
                self.logger.debug('中间文件: %s', extra)

        if self._close_after_cancel:
            self.logger.info('解压取消后关闭程序')
            self.destroy()