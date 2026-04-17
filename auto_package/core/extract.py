import os
import time
import subprocess
import threading
import shutil
import tempfile
from pathlib import Path
from typing import Callable, Iterable
from auto_package.core.utils import _safe_unlink, _safe_rmtree, _has_exe_files, _has_mixed_content, _is_known_archive, _count_non_archive_files, _random_token
from auto_package.constants import _ARCHIVE_EXTS

def find_winrar_exe() -> Path | None:
    """Locate WinRAR.exe on typical Windows installs (supports rar/7z/zip extraction)."""
    roots = [
        os.environ.get("ProgramFiles", ""),
        os.environ.get("ProgramFiles(x86)", ""),
        os.environ.get("LocalAppData", ""),
    ]
    for root in roots:
        if not root:
            continue
        candidate = Path(root) / "WinRAR" / "WinRAR.exe"
        if candidate.is_file():
            return candidate
    # PATH
    from shutil import which

    w = which("WinRAR.exe")
    if w:
        return Path(w)
    return None

def _winrar_extract(
    winrar_exe: Path,
    archive_path: Path,
    output_dir: Path,
    *, 
    progress_cb: Callable[[int | None, float], None] | None = None,
    cancel_ev: threading.Event | None = None,
    proc_cb: Callable[[subprocess.Popen[str] | None], None] | None = None,
) -> tuple[bool, str]:
    """
    使用 WinRAR.exe 解压文件（支持 rar/7z/zip）
    命令: WinRAR.exe x -o+ archive.rar output_dir/
    """
    cmd = [str(winrar_exe), "x", "-o+", "-idq", str(archive_path), str(output_dir)]
    start = time.monotonic()

    # 避免将大输出管道到 Python（内存峰值）
    # 将 WinRAR 输出写入临时文件；仅在错误时读取小尾部
    tmp_out = subprocess.PIPE
    try:
        p = subprocess.Popen(
            cmd,
            stdout=tmp_out,
            stderr=tmp_out,
            creationflags=subprocess.CREATE_NO_WINDOW,
            cwd=None,
        )
    except OSError as e:
        return False, str(e)
    if proc_cb:
        proc_cb(p)

    # 轮询循环检查取消和耗时
    while True:
        if cancel_ev is not None and cancel_ev.is_set():
            try:
                p.terminate()
                p.wait(timeout=1.5)
            except Exception:
                try:
                    p.kill()
                except Exception:
                    pass
            if proc_cb:
                proc_cb(None)
            return False, "已取消"

        rc = p.poll()
        if rc is not None:
            break
        if progress_cb:
            progress_cb(None, time.monotonic() - start)
        time.sleep(0.1)

    rc = int(rc)
    if proc_cb:
        proc_cb(None)
    if progress_cb:
        progress_cb(None, time.monotonic() - start)

    if rc != 0:
        # 读取最多 16KB 的错误信息
        try:
            out, _ = p.communicate(timeout=1.0)
            try:
                out_str = out.decode("mbcs", errors="replace").strip()
            except Exception:
                out_str = ""
        except Exception:
            out_str = ""
        return False, out_str or f"退出码 {rc}"
    return True, str(output_dir)

def _try_extract_with_formats(
    winrar_exe: Path,
    file_path: Path,
    output_dir: Path,
    *, 
    cancel_ev: threading.Event | None = None,
    proc_cb: Callable[[subprocess.Popen[str] | None], None] | None = None,
) -> tuple[bool, str]:
    """
    尝试用多种格式解压文件：
    1. 如果是 .rar/.7z/.zip，直接解压
    2. 否则，尝试改名后解压
    返回: (是否成功解压, 消息)
    """
    # 检查是否为已知压缩包格式
    suffix = file_path.suffix.lower()
    if suffix in _ARCHIVE_EXTS:
        return _winrar_extract(winrar_exe, file_path, output_dir, 
                               cancel_ev=cancel_ev, proc_cb=proc_cb)
    
    # 尝试改名后解压（只尝试zip格式）
    ext = '.zip'
    temp_path = file_path.with_suffix(ext)
    try:
        shutil.copy2(str(file_path), str(temp_path))
    except Exception:
        return False, "不是压缩包或解压失败"
    
    ok, msg = _winrar_extract(winrar_exe, temp_path, output_dir,
                              cancel_ev=cancel_ev, proc_cb=proc_cb)
    _safe_unlink(temp_path)  # 方案B：删除临时文件
    
    if ok:
        return True, msg
    
    return False, "不是压缩包或解压失败"

def _check_and_extract_recursive(
    winrar_exe: Path,
    current_dir: Path,
    out_dir: Path,
    *, 
    depth: int = 0,
    max_depth: int = 10,
    progress_cb: Callable[[int | None, float], None] | None = None,
    phase_cb: Callable[[int, int], None] | None = None,
    cancel_ev: threading.Event | None = None,
    proc_cb: Callable[[subprocess.Popen[str] | None], None] | None = None,
) -> tuple[bool, str]:
    """
    递归检查并解压：
    1. 只有文件：
       - 包含.exe → 最终文件
       - 有2个及以上非压缩格式文件 → 最终文件
       - 只有1个非压缩格式文件 → 仅尝试zip解压，失败则最终文件
       - 全是压缩格式文件 → 尝试解压每个文件
    2. 同时包含文件和文件夹 → 最终文件
    3. 只有文件夹 → 递归检查每个文件夹
    """
    if cancel_ev and cancel_ev.is_set():
        return False, "已取消"
    
    if depth > max_depth:
        return True, "达到最大递归深度，判定为最终文件"
    
    has_files, has_folders = _has_mixed_content(current_dir)
    
    # 情况 2：同时包含文件和文件夹 → 最终文件
    if has_files and has_folders:
        return True, "混合内容，判定为最终文件"
    
    # 情况 1：只有文件
    if has_files and not has_folders:
        # 检查是否包含 .exe 文件
        if _has_exe_files(current_dir):
            return True, "包含可执行文件，判定为最终文件"
        
        # 统计非压缩格式文件数量
        non_archive_count = _count_non_archive_files(current_dir)
        
        # 规则1：如果有两个及以上非压缩格式文件，则为最终文件
        if non_archive_count >= 2:
            return True, "包含多个非压缩格式文件，判定为最终文件"
        
        # 规则2：如果只有一个非压缩格式文件，仅尝试改为zip解压
        if non_archive_count == 1:
            # 找到那个非压缩格式文件
            non_archive_file = None
            for item in current_dir.iterdir():
                if item.is_file() and not _is_known_archive(item):
                    non_archive_file = item
                    break
            
            if non_archive_file:
                # 仅尝试改为zip解压
                temp_extract = current_dir / f"_extract_{_random_token(4)}"
                try:
                    temp_extract.mkdir(exist_ok=False)
                except Exception:
                    return True, "无法创建临时目录，判定为最终文件"
                
                # 尝试改为.zip解压
                temp_zip = non_archive_file.with_suffix('.zip')
                try:
                    shutil.copy2(str(non_archive_file), str(temp_zip))
                except Exception:
                    _safe_rmtree(temp_extract)
                    return True, "无法复制文件，判定为最终文件"
                
                ok, msg = _winrar_extract(winrar_exe, temp_zip, temp_extract,
                                          cancel_ev=cancel_ev, proc_cb=proc_cb)
                _safe_unlink(temp_zip)
                
                if ok:
                    # 删除原文件
                    _safe_unlink(non_archive_file)
                    # 递归检查解压后的内容
                    ok2, msg2 = _check_and_extract_recursive(
                        winrar_exe, temp_extract, out_dir,
                        depth=depth + 1, max_depth=max_depth,
                        progress_cb=progress_cb, phase_cb=phase_cb,
                        cancel_ev=cancel_ev, proc_cb=proc_cb
                    )
                    if not ok2:
                        return False, msg2
                    return True, "递归解压完成"
                else:
                    # 解压失败，删除临时目录，判定为最终文件
                    _safe_rmtree(temp_extract)
                    return True, "单个非压缩文件解压失败，判定为最终文件"
        
        # 规则3：所有文件都是压缩格式，尝试解压每个文件
        any_success = False
        for file_item in list(current_dir.iterdir()):
            if not file_item.is_file():
                continue
            
            # 创建临时解压目录
            temp_extract = current_dir / f"_extract_{_random_token(4)}"
            try:
                temp_extract.mkdir(exist_ok=False)
            except Exception:
                continue
            
            ok, msg = _try_extract_with_formats(
                winrar_exe, file_item, temp_extract,
                cancel_ev=cancel_ev, proc_cb=proc_cb
            )
            
            if ok:
                any_success = True
                # 删除原压缩包
                _safe_unlink(file_item)
                # 递归检查解压后的内容
                ok2, msg2 = _check_and_extract_recursive(
                    winrar_exe, temp_extract, out_dir,
                    depth=depth + 1, max_depth=max_depth,
                    progress_cb=progress_cb, phase_cb=phase_cb,
                    cancel_ev=cancel_ev, proc_cb=proc_cb
                )
                if not ok2:
                    return False, msg2
            else:
                # 方案B：删除临时目录，保留原文件
                _safe_rmtree(temp_extract)
        
        if not any_success:
            return True, "所有文件解压失败，判定为最终文件"
        return True, "递归解压完成"
    
    # 情况 3：只有文件夹
    if has_folders and not has_files:
        for folder_item in list(current_dir.iterdir()):
            if not folder_item.is_dir():
                continue
            ok, msg = _check_and_extract_recursive(
                winrar_exe, folder_item, out_dir,
                depth=depth + 1, max_depth=max_depth,
                progress_cb=progress_cb, phase_cb=phase_cb,
                cancel_ev=cancel_ev, proc_cb=proc_cb
            )
            if not ok:
                return False, msg
        return True, "文件夹递归检查完成"
    
    return True, "空目录"

def run_auto_extract(
    winrar_exe: Path,
    input_path: Path,
    *, 
    output_dir: Path | None = None,
    progress_cb: Callable[[int | None, float], None] | None = None,
    phase_cb: Callable[[int, int], None] | None = None,
    cancel_ev: threading.Event | None = None,
    proc_cb: Callable[[subprocess.Popen[str] | None], None] | None = None,
) -> tuple[bool, str, str | None]:
    """
    自动去伪装解压：
    1. 尝试用多种格式解压文件（rar/7z/zip）
    2. 递归检查并解压嵌套内容
    3. 移动到目标目录
    """
    if not input_path.exists():
        return False, f"文件不存在: {input_path}", None
    if not input_path.is_file():
        return False, "请拖放文件，不是文件夹", None

    out_dir = output_dir or input_path.parent
    work_dir = input_path.parent / f".apwr_tmp_{_random_token(10)}"
    try:
        work_dir.mkdir(parents=False, exist_ok=False)
    except Exception as e:
        return False, f"创建临时目录失败: {e}", None
    created_files: list[Path] = []

    try:
        # 阶段 1：准备解压
        if phase_cb:
            phase_cb(1, 3)

        # 创建解压目录
        extract_dir = Path(work_dir) / "extract"
        extract_dir.mkdir(exist_ok=True)

        # 尝试解压
        ok, msg = _try_extract_with_formats(
            winrar_exe, input_path, extract_dir,
            cancel_ev=cancel_ev, proc_cb=proc_cb
        )
        if not ok:
            return False, f"解压失败: {msg}", None

        # 阶段 2：递归检查并解压
        if phase_cb:
            phase_cb(2, 3)
        ok, msg = _check_and_extract_recursive(
            winrar_exe, extract_dir, out_dir,
            progress_cb=progress_cb, phase_cb=phase_cb,
            cancel_ev=cancel_ev, proc_cb=proc_cb
        )
        if not ok:
            return False, msg, None

        # 阶段 3：移动到目标目录
        if phase_cb:
            phase_cb(3, 3)

        # 找到最终文件所在的文件夹（递归解压后的最深层目录）
        def _find_final_content_dir(start_dir: Path) -> Path:
            """找到包含最终内容的目录（最深层的非空目录）"""
            items = list(start_dir.iterdir())
            # 如果只有一个子文件夹且没有文件，则递归进入
            dirs = [item for item in items if item.is_dir()]
            files = [item for item in items if item.is_file()]
            if len(dirs) == 1 and not files:
                return _find_final_content_dir(dirs[0])
            return start_dir

        final_content_dir = _find_final_content_dir(extract_dir)

        # 使用最终内容目录的名称作为文件夹名
        folder_name = final_content_dir.name

        # 确保文件夹名不冲突
        result_folder = out_dir / folder_name
        if result_folder.exists():
            result_folder = out_dir / f"{folder_name}_{_random_token(4)}"

        # 移动最终内容目录到目标位置
        if final_content_dir == extract_dir:
            # 没有嵌套目录，直接移动整个 extract_dir
            shutil.move(str(extract_dir), str(result_folder))
        else:
            # 有嵌套目录，移动最深层目录
            shutil.move(str(final_content_dir), str(result_folder))
            # 清理剩余的中间目录
            _safe_rmtree(extract_dir)

        return True, f"解压完成到: {result_folder}", None
    except Exception as e:
        # 回滚：删除所有已创建的文件
        for f in created_files:
            _safe_unlink(f)
        # 清理临时目录
        _safe_rmtree(work_dir)
        return False, f"解压过程出错: {e}", None
    finally:
        # 确保临时目录被删除
        _safe_rmtree(work_dir)