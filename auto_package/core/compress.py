import os
import math
import time
import subprocess
import threading
from pathlib import Path
from typing import Callable, Iterable
from auto_package.core.utils import _pick_fake_ext, _pick_nonexistent_path, _safe_unlink, _commit_outputs_atomic

def find_rar_exe() -> Path | None:
    """Locate WinRAR's command-line Rar.exe on typical Windows installs."""
    roots = [
        os.environ.get("ProgramFiles", ""),
        os.environ.get("ProgramFiles(x86)", ""),
        os.environ.get("LocalAppData", ""),
    ]
    for root in roots:
        if not root:
            continue
        candidate = Path(root) / "WinRAR" / "Rar.exe"
        if candidate.is_file():
            return candidate
    # PATH
    from shutil import which

    w = which("Rar.exe")
    if w:
        return Path(w)
    return None

def _rar_run(
    rar_exe: Path,
    archive_path: Path,
    sources: list[str],
    *, 
    recurse: bool,
    exclude_paths: bool = False,
    volume_spec: str | None = None,
    cwd: Path | None = None,
    progress_cb: Callable[[int | None, float], None] | None = None,
    cancel_ev: threading.Event | None = None,
    proc_cb: Callable[[subprocess.Popen[str] | None], None] | None = None,
) -> tuple[bool, str]:
    cmd = [str(rar_exe), "a"]
    if recurse:
        cmd.append("-r")
    if exclude_paths:
        cmd.append("-ep")
    if volume_spec:
        cmd.append(volume_spec)
    # WinRAR CLI progress isn't reliably parseable via pipes.
    # Use quiet mode to avoid buffered output.
    cmd.extend(["-idq", str(archive_path)])
    cmd.extend(sources)
    start = time.monotonic()

    # Avoid piping large output into Python (memory spike).
    # Write WinRAR output to a temp file; read only a small tail on errors.
    tmp_out = subprocess.PIPE
    try:
        p = subprocess.Popen(
            cmd,
            stdout=tmp_out,
            stderr=tmp_out,
            creationflags=subprocess.CREATE_NO_WINDOW,
            cwd=str(cwd) if cwd else None,
        )
    except OSError as e:
        return False, str(e)
    if proc_cb:
        proc_cb(p)

    # Poll loop for cancel + elapsed
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
        # Read last up to 16KB for error message
        try:
            out, _ = p.communicate(timeout=1.0)
            try:
                out_str = out.decode("mbcs", errors="replace").strip()
            except Exception:
                out_str = ""
        except Exception:
            out_str = ""
        return False, out_str or f"退出码 {rc}"
    return True, str(archive_path)

def run_rar_archive(
    rar_exe: Path,
    items: list[Path],
    output_dir: Path,
    *, 
    progress_cb: Callable[[int | None, float], None] | None = None,
    cancel_ev: threading.Event | None = None,
    proc_cb: Callable[[subprocess.Popen[str] | None], None] | None = None,
) -> tuple[bool, Path | str, str | None]:
    """Run one-pass Rar.exe archive."""
    out_rar = _pick_nonexistent_path(output_dir, ".rar")
    
    # 处理多个文件夹的情况：创建临时目录来保持文件夹结构
    if len(items) > 1 or (len(items) == 1 and items[0].is_dir()):
        import tempfile
        import shutil
        temp_dir = None
        try:
            # 创建临时目录
            temp_dir = tempfile.mkdtemp()
            temp_path = Path(temp_dir)
            
            # 将所有项目复制到临时目录
            for item in items:
                dest = temp_path / item.name
                if item.is_dir():
                    shutil.copytree(str(item), str(dest))
                else:
                    shutil.copy2(str(item), str(dest))
            
            # 压缩临时目录
            sources = [str(temp_path)]
        except Exception as e:
            if temp_dir:
                try:
                    shutil.rmtree(temp_dir)
                except:
                    pass
            _safe_unlink(out_rar)
            return False, f"创建临时目录失败: {e}", None
    else:
        # 单个文件的情况，直接压缩
        sources = [str(p) for p in items]
    
    ok, msg = _rar_run(
        rar_exe,
        out_rar,
        sources,
        recurse=True,
        exclude_paths=False,
        progress_cb=progress_cb,
        cancel_ev=cancel_ev,
        proc_cb=proc_cb,
    )
    
    # 清理临时目录
    if 'temp_dir' in locals() and temp_dir:
        try:
            import shutil
            shutil.rmtree(temp_dir)
        except:
            pass
    
    if not ok:
        _safe_unlink(out_rar)
        return False, msg, None
    return True, out_rar, None

def run_double_compress(
    rar_exe: Path,
    items: list[Path],
    output_dir: Path,
    *, 
    progress_cb: Callable[[int | None, float], None] | None = None,
    phase_cb: Callable[[int, int], None] | None = None,
    cancel_ev: threading.Event | None = None,
    proc_cb: Callable[[subprocess.Popen[str] | None], None] | None = None,
) -> tuple[bool, str, str | None]:
    """Two-pass compression: archive → rename to fake ext → archive again."""
    if phase_cb:
        phase_cb(1, 2)

    ok1, out1, extra1 = run_rar_archive(
        rar_exe,
        items,
        output_dir,
        progress_cb=progress_cb,
        cancel_ev=cancel_ev,
        proc_cb=proc_cb,
    )
    if not ok1:
        assert isinstance(out1, str)
        return False, out1, None
    assert isinstance(out1, Path)

    disguised = out1.with_suffix(_pick_fake_ext())
    try:
        out1.rename(disguised)
    except OSError as e:
        _safe_unlink(out1)
        return False, f"重命名伪装失败: {e}", str(out1)

    if phase_cb:
        phase_cb(2, 2)

    ok2, out2, extra2 = run_rar_archive(
        rar_exe,
        [disguised],
        output_dir,
        progress_cb=progress_cb,
        cancel_ev=cancel_ev,
        proc_cb=proc_cb,
    )

    try:
        disguised.unlink()
    except OSError:
        pass

    if not ok2:
        assert isinstance(out2, str)
        return False, f"二次压缩失败: {out2}", str(disguised)
    assert isinstance(out2, Path)

    return True, str(out2), None

def run_triple_compress(
    rar_exe: Path,
    items: list[Path],
    output_dir: Path,
    *, 
    progress_cb: Callable[[int | None, float], None] | None = None,
    phase_cb: Callable[[int, int], None] | None = None,
    cancel_ev: threading.Event | None = None,
    proc_cb: Callable[[subprocess.Popen[str] | None], None] | None = None,
) -> tuple[bool, str, str | None]:
    """Three-pass compression: double → triple with volume."""
    if phase_cb:
        phase_cb(1, 3)

    ok2, out2, extra2 = run_double_compress(
        rar_exe,
        items,
        output_dir,
        progress_cb=progress_cb,
        phase_cb=phase_cb,
        cancel_ev=cancel_ev,
        proc_cb=proc_cb,
    )
    if not ok2:
        assert isinstance(out2, str)
        return False, out2, extra2
    out2_path = Path(out2)

    if phase_cb:
        phase_cb(2, 3)

    disguised2 = out2_path.with_suffix(_pick_fake_ext())
    try:
        out2_path.rename(disguised2)
    except OSError as e:
        _safe_unlink(out2_path)
        return False, f"重命名伪装失败: {e}", str(out2_path)

    if phase_cb:
        phase_cb(3, 3)

    out_rar_3 = _pick_nonexistent_path(output_dir, ".rar")
    # 估算分卷大小，目标：总大小/2 + 10MB，最小 100MB，最大 2049MB
    total_bytes = disguised2.stat().st_size
    vol_mb = max(100, min(2049, int(math.ceil(total_bytes / 1024 / 1024 / 2) + 10)))
    # 至少 2 卷，否则取消分卷
    if total_bytes < vol_mb * 1024 * 1024:
        # 太小，取消分卷，退化为二次压缩
        try:
            disguised2.rename(out2_path)
        except OSError:
            pass
        return False, f"文件太小（{total_bytes / 1024 / 1024:.1f}MB），取消分卷", str(out2_path)
    vol_spec = f"-v{vol_mb}m"
    est_parts = math.ceil(total_bytes / (vol_mb * 1024 * 1024))

    ok3, msg3 = _rar_run(
        rar_exe,
        out_rar_3,
        [str(disguised2)],
        recurse=False,
        exclude_paths=False,
        volume_spec=vol_spec,
        cwd=disguised2.parent,
        progress_cb=progress_cb,
        cancel_ev=cancel_ev,
        proc_cb=proc_cb,
    )
    if not ok3:
        return (
            False,
            f"三次分卷压缩失败: {msg3}\n中间文件保留为: {disguised2}",
            str(disguised2),
        )

    try:
        disguised2.unlink()
    except OSError:
        pass

    detail = (
        f"三次压缩完成（分卷）。目标单卷≈总/2+10MB，最大≤2049MB。预计至少 {est_parts} 卷，WinRAR 参数 {vol_spec}。"
        f" 输出示例: {out_rar_3.with_suffix('')}.part1.rar"
    )
    return True, f"{out_rar_3.with_suffix('')}.part*.rar", detail