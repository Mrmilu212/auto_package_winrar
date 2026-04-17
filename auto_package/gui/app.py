import os
import time
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from pathlib import Path
from typing import Iterable
import subprocess
from auto_package.core.compress import find_rar_exe, run_rar_archive, run_double_compress, run_triple_compress
from auto_package.core.extract import find_winrar_exe, run_auto_extract
from auto_package.core.utils import decode_drop_path, _commit_outputs_atomic
from auto_package.config.settings import _load_window_geometry, _save_window_geometry

try:
    import windnd
except ImportError:
    windnd = None

class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
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
        
        # 上传进程管理
        self._upload_processes = []  # 存储上传进程信息
        self._upload_process_id = 0
        self._transfer_timer_id = None  # 传输窗口定时器ID

        self._build_ui()
        self._hook_drop()
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

    def _build_ui(self) -> None:
        pad = {"padx": 12, "pady": 8}

        # 窗口切换按钮
        self._current_window = "main"
        self._main_frame = None
        self._transfer_frame = None
        
        # 切换按钮框架
        switch_frame = tk.Frame(self)
        switch_frame.pack(fill=tk.X, padx=12, pady=8)
        
        self.btn_main = tk.Button(switch_frame, text="主界面", command=lambda: self._switch_window("main"))
        self.btn_main.pack(side=tk.LEFT, padx=4)
        
        self.btn_transfer = tk.Button(switch_frame, text="传输窗口", command=lambda: self._switch_window("transfer"))
        self.btn_transfer.pack(side=tk.LEFT, padx=4)

        # 主界面框架
        self._main_frame = tk.Frame(self)
        
        if self._rar:
            rar_line = f"已找到: {self._rar}"
        else:
            rar_line = "未找到 WinRAR（Rar.exe）。请安装 WinRAR 或检查安装路径。"

        frm_top = tk.Frame(self._main_frame)
        frm_top.pack(fill=tk.X, **pad)

        self.lbl_hint = tk.Label(
            frm_top,
            text="将文件/文件夹拖放到下方区域，或点击「选择文件夹」",
            justify=tk.LEFT,
            wraplength=520,
        )
        self.lbl_hint.pack(anchor=tk.W)

        self.lbl_rar = tk.Label(frm_top, text=rar_line, fg="#333", justify=tk.LEFT)
        self.lbl_rar.pack(anchor=tk.W, pady=(4, 0))

        self.chk_double = tk.Checkbutton(
            frm_top,
            text="二次伪装后缀压缩",
            variable=self._var_double,
            justify=tk.LEFT,
            wraplength=520,
            anchor=tk.W,
        )
        self.chk_double.pack(anchor=tk.W, pady=(6, 0))

        self.chk_triple = tk.Checkbutton(
            frm_top,
            text="三次分卷压缩",
            variable=self._var_triple,
            justify=tk.LEFT,
            wraplength=520,
            anchor=tk.W,
        )
        self.chk_triple.pack(anchor=tk.W, pady=(2, 0))

        # 百度网盘上传选项
        self._var_upload = tk.BooleanVar(value=False)
        self.chk_upload = tk.Checkbutton(
            frm_top,
            text="上传到百度网盘",
            variable=self._var_upload,
            justify=tk.LEFT,
            wraplength=520,
            anchor=tk.W,
        )
        self.chk_upload.pack(anchor=tk.W, pady=(6, 0))

        # 拖放区域容器
        drop_container = tk.Frame(self._main_frame)
        drop_container.pack(fill=tk.BOTH, expand=True, **pad)

        # 压缩拖放区域
        compress_frame = tk.Frame(drop_container)
        compress_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 6))
        tk.Label(compress_frame, text="压缩拖放区", font=("Segoe UI", 10, "bold"), fg="#666").pack(anchor=tk.W, pady=(0, 4))
        self.drop_frame = tk.Frame(compress_frame, relief=tk.SUNKEN, bd=2, bg="#e8eef5", width=250, height=250)
        self.drop_frame.pack(fill=tk.BOTH, expand=True)
        self.drop_frame.pack_propagate(False)

        self.lbl_drop = tk.Label(
            self.drop_frame,
            text="拖放文件/文件夹",
            bg="#e8eef5",
            fg="#456",
            font=("Segoe UI", 12),
        )
        self.lbl_drop.pack(expand=True)

        # 解压拖放区域
        extract_frame = tk.Frame(drop_container)
        extract_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(6, 0))
        tk.Label(extract_frame, text="解压拖放区", font=("Segoe UI", 10, "bold"), fg="#666").pack(anchor=tk.W, pady=(0, 4))
        self.extract_drop_frame = tk.Frame(extract_frame, relief=tk.SUNKEN, bd=2, bg="#f5e8ee", width=250, height=250)
        self.extract_drop_frame.pack(fill=tk.BOTH, expand=True)
        self.extract_drop_frame.pack_propagate(False)

        self.lbl_extract_drop = tk.Label(
            self.extract_drop_frame,
            text="拖放要解压的文件",
            bg="#f5e8ee",
            fg="#644",
            font=("Segoe UI", 12),
        )
        self.lbl_extract_drop.pack(expand=True)

        btn_row = tk.Frame(self._main_frame)
        btn_row.pack(fill=tk.X, **pad)

        tk.Button(btn_row, text="选择文件夹…", command=self._pick_folder).pack(
            side=tk.LEFT
        )
        tk.Button(btn_row, text="清空日志", command=self._clear_log).pack(
            side=tk.LEFT, padx=(8, 0)
        )
        
        # 按钮容器
        btn_right = tk.Frame(btn_row)
        btn_right.pack(side=tk.RIGHT)
        
        self.btn_cancel = tk.Button(
            btn_right, text="取消压缩", command=self._cancel_current, state=tk.DISABLED
        )
        self.btn_cancel.pack(side=tk.LEFT, padx=(0, 8))
        
        self.btn_cancel_extract = tk.Button(
            btn_right, text="取消解压", command=self._cancel_extract, state=tk.DISABLED
        )
        self.btn_cancel_extract.pack(side=tk.LEFT)

        self.log = scrolledtext.ScrolledText(self._main_frame, height=8, state=tk.DISABLED)
        self.log.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 12))

        self.status = tk.Label(self._main_frame, textvariable=self._status_var, anchor=tk.W)
        self.status.pack(fill=tk.X, padx=12, pady=(0, 10))
        
        # 传输窗口框架
        self._transfer_frame = tk.Frame(self)
        
        # 传输窗口标题
        transfer_title = tk.Label(self._transfer_frame, text="上传进程", font=("Segoe UI", 12, "bold"))
        transfer_title.pack(padx=12, pady=8, anchor=tk.W)
        
        # 传输列表
        self._transfer_list_frame = tk.Frame(self._transfer_frame)
        self._transfer_list_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=4)
        
        # 初始显示
        self._main_frame.pack(fill=tk.BOTH, expand=True)
        self._transfer_frame.pack(fill=tk.BOTH, expand=True)
        self._transfer_frame.pack_forget()  # 初始隐藏传输窗口

        self._log_line(
            "就绪。默认：每进行一次压缩都会生成 5 位随机文件名（英文大小写+数字）的 .rar。"
            " 二次：在第一次产物基础上改伪装后缀再打包，外层 .rar 也会是新的随机名。"
            " 三次（分卷）：单卷目标≈总/2+10MB，最大≤2049MB，至少 2 卷；过小将取消。"
        )
        self._log_line("解压：拖放文件到右侧区域，自动识别并解压，支持去伪装。")
        self._set_status(0.0)

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
    
    def _switch_window(self, window_name: str) -> None:
        """切换窗口"""
        if self._current_window == window_name:
            return
        
        # 隐藏当前窗口
        if self._current_window == "main":
            self._main_frame.pack_forget()
        else:
            self._transfer_frame.pack_forget()
            # 停止定时器
            if self._transfer_timer_id:
                self.after_cancel(self._transfer_timer_id)
                self._transfer_timer_id = None
        
        # 显示目标窗口
        if window_name == "main":
            self._main_frame.pack(fill=tk.BOTH, expand=True)
        else:
            self._transfer_frame.pack(fill=tk.BOTH, expand=True)
            # 刷新传输窗口
            self._update_transfer_window()
            # 启动定时器，每秒刷新一次
            self._start_transfer_timer()
        
        self._current_window = window_name
    
    def _update_transfer_window(self) -> None:
        """更新传输窗口"""
        # 清空传输列表
        for widget in self._transfer_list_frame.winfo_children():
            widget.destroy()
        
        # 添加标题行
        title_frame = tk.Frame(self._transfer_list_frame)
        title_frame.pack(fill=tk.X, pady=2)
        tk.Label(title_frame, text="文件名", width=40, anchor=tk.W, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT)
        tk.Label(title_frame, text="已用时间", width=15, anchor=tk.E, font=("Segoe UI", 10, "bold")).pack(side=tk.RIGHT)
        
        # 添加分割线
        tk.Frame(self._transfer_list_frame, height=2, bg="#ccc").pack(fill=tk.X, pady=2)
        
        # 显示上传进程
        if not self._upload_processes:
            empty_label = tk.Label(self._transfer_list_frame, text="暂无上传进程", fg="#666")
            empty_label.pack(pady=20)
        else:
            for proc in self._upload_processes:
                proc_frame = tk.Frame(self._transfer_list_frame, bg="#f5f5f5")
                proc_frame.pack(fill=tk.X, pady=2, padx=2)
                
                # 文件名
                file_name = proc.get("file_name", "未知文件")
                tk.Label(proc_frame, text=file_name, width=40, anchor=tk.W, bg="#f5f5f5").pack(side=tk.LEFT)
                
                # 已用时间（中间位置）
                elapsed = time.time() - proc.get("start_time", time.time())
                time_str = self._fmt_time(elapsed)
                time_frame = tk.Frame(proc_frame, bg="#f5f5f5")
                time_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
                tk.Label(time_frame, text=time_str, width=15, anchor=tk.CENTER, bg="#f5f5f5").pack(side=tk.LEFT)
                
                # 取消按钮（最右边）
                cancel_btn = tk.Button(proc_frame, text="取消", width=6, command=lambda pid=proc["id"]: self._cancel_upload_process(pid))
                cancel_btn.pack(side=tk.RIGHT, padx=5)
    
    def _add_upload_process(self, file_name: str) -> int:
        """添加上传进程"""
        proc_id = self._upload_process_id
        self._upload_process_id += 1
        
        proc_info = {
            "id": proc_id,
            "file_name": file_name,
            "start_time": time.time(),
            "cancel_ev": threading.Event(),
            "proc": None
        }
        self._upload_processes.append(proc_info)
        return proc_id
    
    def _remove_upload_process(self, proc_id: int) -> None:
        """移除上传进程"""
        self._upload_processes = [p for p in self._upload_processes if p["id"] != proc_id]
        # 刷新传输窗口
        if self._current_window == "transfer":
            self.after(0, self._update_transfer_window)
    
    def _start_transfer_timer(self) -> None:
        """启动传输窗口定时器"""
        def update():
            self._update_transfer_window()
            self._transfer_timer_id = self.after(1000, update)  # 每秒刷新一次
        
        # 先停止之前的定时器
        if self._transfer_timer_id:
            self.after_cancel(self._transfer_timer_id)
        
        # 启动新的定时器
        self._transfer_timer_id = self.after(1000, update)
    
    def _cancel_upload_process(self, proc_id: int) -> None:
        """取消上传进程"""
        # 找到对应的上传进程
        proc_info = None
        for p in self._upload_processes:
            if p["id"] == proc_id:
                proc_info = p
                break
        
        if not proc_info:
            return
        
        # 设置取消事件
        cancel_ev = proc_info.get("cancel_ev")
        if cancel_ev:
            cancel_ev.set()
        
        # 终止进程（如果存在）
        proc = proc_info.get("proc")
        if proc:
            try:
                proc.terminate()
                try:
                    proc.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    proc.kill()
            except Exception as e:
                self._log_line(f"终止上传进程失败: {e}")
        
        # 从上传进程列表中移除
        self._upload_processes = [p for p in self._upload_processes if p["id"] != proc_id]
        
        # 减少上传线程计数
        self._upload_thread_count = max(0, self._upload_thread_count - 1)
        
        # 刷新传输窗口
        if self._current_window == "transfer":
            self._update_transfer_window()
        
        # 记录日志
        self._log_line(f"上传进程已取消: {proc_info.get('file_name', '未知文件')}")

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
            return
        if self._extract_busy:
            self._close_after_cancel = True
            self._cancel_extract()
            # 等任务结束后在 extract_done() 里关闭
            self._log_line("正在取消解压并退出…")
            return
        # 保存窗口大小
        try:
            width = self.winfo_width()
            height = self.winfo_height()
            if width > 0 and height > 0:
                _save_window_geometry(width, height)
        except Exception:
            pass
        self.destroy()

    def _hook_drop(self) -> None:
        if windnd is None:
            self._log_line("未安装 windnd，拖放不可用。请运行: pip install -r requirements.txt")
            return
        try:
            # 压缩拖放区域
            kw = {"func": self._on_drop_files, "force_unicode": True}
            windnd.hook_dropfiles(self.drop_frame, **kw)
            windnd.hook_dropfiles(self.lbl_drop, **kw)
            # 解压拖放区域
            kw_extract = {"func": self._on_drop_extract, "force_unicode": True}
            windnd.hook_dropfiles(self.extract_drop_frame, **kw_extract)
            windnd.hook_dropfiles(self.lbl_extract_drop, **kw_extract)
        except Exception as e:
            self._log_line(f"拖放注册失败: {e}")

    def _clear_log(self) -> None:
        self.log.configure(state=tk.NORMAL)
        self.log.delete("1.0", tk.END)
        self.log.configure(state=tk.DISABLED)

    def _log_line(self, msg: str) -> None:
        self.log.configure(state=tk.NORMAL)
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)
        self.log.configure(state=tk.DISABLED)

    def _on_drop_files(self, paths: Iterable[bytes]) -> None:
        if self._busy:
            self._log_line("正在打包，请稍候…")
            return
        if not self._rar:
            messagebox.showerror("WinRAR", "未找到 Rar.exe，无法打包。")
            return

        decoded = [
            decode_drop_path(p) if isinstance(p, bytes) else str(p) for p in paths
        ]
        items = [Path(p) for p in decoded if Path(p).exists()]
        if not items:
            messagebox.showinfo("提示", "请拖放文件或文件夹。")
            return
        self._start_pack(items)

    def _pick_folder(self) -> None:
        d = filedialog.askdirectory(title="选择要打包的文件夹")
        if not d:
            return
        self._on_drop_files([os.fsencode(d)])

    def _start_pack(self, items: list[Path]) -> None:
        self._busy = True
        self.lbl_drop.configure(text="正在打包…")
        if len(items) == 1:
            self._log_line(f"开始: {items[0]}")
        else:
            self._log_line(f"开始: {items[0]}（共 {len(items)} 个输入）")
        # 状态栏阶段文案（不显示百分比/剩余时间）
        if self._var_triple.get():
            self._set_phase(1, 3)
        elif self._var_double.get():
            self._set_phase(1, 2)
        else:
            self._set_phase(None, 1)
        self._start_job_timer()
        self.btn_cancel.configure(state=tk.NORMAL)
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
        self.lbl_drop.configure(text="拖放文件/文件夹")
        self.btn_cancel.configure(state=tk.DISABLED)
        self._cancel_ev = None
        self._set_current_proc(None)
        self._stop_job_timer()

        if ok:
            self._log_line(f"成功: {msg}")
            messagebox.showinfo("成功", msg)
            
            # 检查是否需要上传到百度网盘
            if self._var_upload.get():
                # 检查上传线程数限制
                if self._upload_thread_count >= self._max_upload_threads:
                    self._log_line(f"上传线程数达到上限 ({self._max_upload_threads})，请稍后再试")
                else:
                    self._log_line("开始上传到百度网盘...")
                    # 创建后台线程执行上传，不阻塞GUI
                    def upload_task():
                        proc_id = None
                        try:
                            # 增加上传线程计数
                            self._upload_thread_count += 1
                            self.after(0, self._log_line, f"当前上传线程数: {self._upload_thread_count}/{self._max_upload_threads}")
                            
                            # 获取压缩产物路径
                            output_path = Path(msg)
                            
                            # 检查路径是否存在
                            if not output_path.exists():
                                self.after(0, self._log_line, f"上传失败: 产物路径不存在: {output_path}")
                                return
                            
                            # 添加上传进程
                            file_name = output_path.name
                            # 直接调用，因为这是在后台线程中，需要通过线程安全的方式
                            def add_proc():
                                nonlocal proc_id
                                proc_id = self._add_upload_process(file_name)
                                # 刷新传输窗口
                                if self._current_window == "transfer":
                                    self._update_transfer_window()
                            self.after(0, add_proc)
                            
                            # 等待进程添加完成
                            import time
                            time.sleep(0.1)
                            
                            # 获取取消事件
                            cancel_ev = None
                            def get_cancel_ev():
                                nonlocal cancel_ev
                                for p in self._upload_processes:
                                    if p["id"] == proc_id:
                                        cancel_ev = p["cancel_ev"]
                                        break
                            self.after(0, get_cancel_ev)
                            
                            # 等待获取取消事件
                            time.sleep(0.1)
                            
                            # 调用上传模块
                            from auto_package.core.upload import upload_to_baidu_pan
                            
                            # 确保日志回调在GUI线程中执行
                            def log_callback(message):
                                self.after(0, self._log_line, message)
                            
                            ok, result_msg, proc = upload_to_baidu_pan(
                                output_path=output_path,
                                upload_base="/测试",
                                log_cb=log_callback,
                                cancel_ev=cancel_ev,
                            )
                            
                            # 存储进程对象
                            def store_proc():
                                for p in self._upload_processes:
                                    if p["id"] == proc_id:
                                        p["proc"] = proc
                                        break
                            self.after(0, store_proc)
                            if not ok:
                                self.after(0, self._log_line, f"上传失败: {result_msg}")
                        except Exception as e:
                            self.after(0, self._log_line, f"上传异常: {e}")
                        finally:
                            # 减少上传线程计数
                            self._upload_thread_count -= 1
                            self.after(0, self._log_line, f"上传线程结束，当前上传线程数: {self._upload_thread_count}/{self._max_upload_threads}")
                            # 移除上传进程
                            if proc_id:
                                def remove_proc():
                                    self._remove_upload_process(proc_id)
                                self.after(0, remove_proc)
                    
                    # 启动后台线程
                    import threading
                    threading.Thread(target=upload_task, daemon=True).start()
        else:
            self._log_line(f"失败: {msg}")
            if extra:
                self._log_line(f"中间文件: {extra}")
            messagebox.showerror("失败", msg)

        if self._close_after_cancel:
            self.destroy()

    def _on_drop_extract(self, paths: Iterable[bytes]) -> None:
        if self._extract_busy:
            self._log_line("正在解压，请稍候…")
            return
        if not self._winrar:
            messagebox.showerror("WinRAR", "未找到 WinRAR.exe，无法解压。")
            return

        decoded = [
            decode_drop_path(p) if isinstance(p, bytes) else str(p) for p in paths
        ]
        items = [Path(p) for p in decoded if Path(p).exists()]
        if not items:
            messagebox.showinfo("提示", "请拖放文件。")
            return
        if len(items) > 1:
            messagebox.showinfo("提示", "一次只能解压一个文件。")
            return
        self._start_extract(items[0])

    def _start_extract(self, input_path: Path) -> None:
        if not self._winrar:
            messagebox.showerror("WinRAR", "未找到 WinRAR.exe，无法解压。")
            return

        self._extract_busy = True
        self.lbl_extract_drop.configure(text="正在解压…")
        self._log_line(f"开始解压: {input_path}")
        self._start_job_timer()
        self.btn_cancel_extract.config(state=tk.NORMAL)
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
        self.lbl_extract_drop.configure(text="拖放要解压的文件")
        self.btn_cancel_extract.config(state=tk.DISABLED)
        self._extract_cancel_ev = None
        self._extract_current_proc = None
        self._stop_job_timer()

        if ok:
            self._log_line(f"解压成功: {msg}")
        else:
            self._log_line(f"解压失败: {msg}")
            if extra:
                self._log_line(f"中间文件: {extra}")

        if self._close_after_cancel:
            self.destroy()