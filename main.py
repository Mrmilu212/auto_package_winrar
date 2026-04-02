"""
Drag a folder onto the window to create a .rar next to it using WinRAR (Rar.exe).
Optional: double compression — rename first .rar to a random non-archive extension, then archive that file again.
"""
from __future__ import annotations

import os
import math
import random
import string
import time
import re
import subprocess
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from pathlib import Path
from typing import Callable, Iterable
import shutil
import tempfile

try:
    import windnd
except ImportError:
    windnd = None


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


def decode_drop_path(raw: bytes) -> str:
    for enc in ("utf-8", "mbcs", "gbk"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


# 伪装用扩展名（排除常见压缩/打包后缀）
_FAKE_EXTS = (
    ".png",
    ".java",
    ".gif",
    ".dll",
    ".jpg",
    ".mp4",
    ".txt",
    ".pdf",
    ".xml",
    ".css",
    ".log",
    ".dat",
    ".bin",
    ".ico",
    ".svg",
    ".webp",
    ".bmp",
    ".wav",
    ".doc",
    ".mp3",
    ".json",
    ".csv",
)


def _pick_fake_ext() -> str:
    return random.choice(_FAKE_EXTS)

_NAME_CHARS = string.ascii_letters + string.digits


def _random_archive_stem(length: int = 5) -> str:
    return "".join(random.choice(_NAME_CHARS) for _ in range(length))

def _random_token(length: int = 8) -> str:
    return "".join(random.choice(_NAME_CHARS) for _ in range(length))


def _pick_nonexistent_path(parent: Path, suffix: str) -> Path:
    for _ in range(200):
        candidate = parent / f"{_random_archive_stem(5)}{suffix}"
        if not candidate.exists():
            return candidate
    raise RuntimeError("无法生成不冲突的随机文件名（尝试次数过多）")


def _safe_unlink(p: Path) -> None:
    try:
        p.unlink()
    except FileNotFoundError:
        return
    except OSError:
        return


def _safe_rmtree(d: Path) -> None:
    try:
        shutil.rmtree(d, ignore_errors=True)
    except Exception:
        pass


def _collect_volume_parts(archive_path: Path) -> list[Path]:
    """
    给定 out.rar，WinRAR 分卷通常生成 out.part1.rar、out.part2.rar...
    这里按前缀扫描同目录并返回排序后的列表。
    """
    parent = archive_path.parent
    stem = archive_path.with_suffix("").name
    parts = sorted(parent.glob(f"{stem}.part*.rar"))
    return parts


def _commit_outputs_atomic(
    temp_outputs: list[Path],
    target_dir: Path,
    *,
    is_volumes: bool,
) -> tuple[bool, str, list[Path]]:
    """
    将临时目录中的产物提交到 target_dir。
    - 单文件：移动为新的随机名 .rar
    - 分卷：生成新的随机名作为前缀，移动为 <prefix>.partN.rar

    返回 (ok, message, final_paths)。
    注意：多文件无法做到“操作系统级原子”，这里保证失败时会尽力回滚到“目标目录无产物”状态。
    """
    moved: list[Path] = []
    try:
        if not is_volumes:
            if len(temp_outputs) != 1:
                return False, "内部错误：单文件提交期望 1 个产物", []
            src = temp_outputs[0]
            dst = _pick_nonexistent_path(target_dir, ".rar")
            shutil.move(str(src), str(dst))
            moved.append(dst)
            return True, str(dst), moved

        # volumes
        if not temp_outputs:
            return False, "内部错误：分卷提交未找到产物", []
        # 生成目标文件名并搬运
        part_re = re.compile(r"\.part(\d+)\.rar$", re.IGNORECASE)
        parts: list[tuple[Path, str]] = []
        for src in temp_outputs:
            m = part_re.search(src.name)
            if not m:
                return False, f"无法识别分卷文件名: {src.name}", moved
            parts.append((src, m.group(1)))

        # 选一个前缀，确保整组目标都不冲突
        for _ in range(200):
            prefix = _random_archive_stem(5)
            dists = [target_dir / f"{prefix}.part{idx}.rar" for _, idx in parts]
            if any(d.exists() for d in dists):
                continue
            # commit
            for (src, idx), dst in zip(parts, dists, strict=True):
                shutil.move(str(src), str(dst))
                moved.append(dst)
            moved.sort()
            return True, f"{prefix}.part*.rar", moved

        return False, "无法为分卷产物生成不冲突的随机前缀", []
    except Exception as e:
        # rollback: delete anything moved
        for p in moved:
            _safe_unlink(p)
        return False, f"提交产物失败: {e}", []


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
    tmp_out = tempfile.TemporaryFile(mode="w+b")
    try:
        p = subprocess.Popen(
            cmd,
            stdout=tmp_out,
            stderr=tmp_out,
            creationflags=subprocess.CREATE_NO_WINDOW,
            cwd=str(cwd) if cwd else None,
        )
    except OSError as e:
        try:
            tmp_out.close()
        except Exception:
            pass
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
            tmp_out.seek(0, 2)
            size = tmp_out.tell()
            tmp_out.seek(max(0, size - 16384), 0)
            tail = tmp_out.read()
            try:
                out = tail.decode("mbcs", errors="replace").strip()
            except Exception:
                out = ""
        except Exception:
            out = ""
        finally:
            try:
                tmp_out.close()
            except Exception:
                pass
        return False, out or f"退出码 {rc}"
    try:
        tmp_out.close()
    except Exception:
        pass
    return True, str(archive_path)


def _compute_out_rar_for_input(input_path: Path) -> Path:
    # 每次压缩都随机命名压缩包（5位英文大小写+数字）
    return _pick_nonexistent_path(input_path.parent, ".rar")


def _compute_base_prefix_for_input(input_path: Path) -> Path:
    """
    用于生成伪装文件名、分卷前缀等。
    - 文件夹：parent/name
    - 文件：parent/stem
    """
    if input_path.is_dir():
        return input_path.parent / input_path.name
    return input_path.parent / input_path.stem


def run_rar_archive(
    rar_exe: Path,
    input_paths: list[Path],
    *,
    output_dir: Path | None = None,
    progress_cb: Callable[[int | None, float], None] | None = None,
    cancel_ev: threading.Event | None = None,
    proc_cb: Callable[[subprocess.Popen[str] | None], None] | None = None,
) -> tuple[bool, Path | str]:
    """
    Pack `input_paths` into a .rar.
    - If only one path:
      - folder: archive the folder itself (top-level folder included)
      - file: archive that file only
    - If multiple paths: pack all inputs into one archive and exclude paths (-ep),
      so the archive contains each input's basename at top level.
    -r recurse; -idq quiet hash; -ibck background priority (optional).
    """
    if not input_paths:
        return False, "输入为空"
    for p in input_paths:
        if not p.exists():
            return False, f"路径不存在: {p}"

    first_parent = input_paths[0].parent
    out_dir = output_dir or first_parent
    out_rar = _pick_nonexistent_path(out_dir, ".rar")

    if len(input_paths) == 1:
        input_path = input_paths[0]
        if input_path.is_dir():
            # 压缩“文件夹本身”，压缩包里顶层会包含该文件夹名
            ok, msg = _rar_run(
                rar_exe,
                out_rar,
                [input_path.name],
                recurse=True,
                cwd=input_path.parent,
                progress_cb=progress_cb,
                cancel_ev=cancel_ev,
                proc_cb=proc_cb,
            )
            return (ok, out_rar) if ok else (False, msg)

        # 单文件
        ok, msg = _rar_run(
            rar_exe,
            out_rar,
            [input_path.name],
            recurse=False,
            cwd=input_path.parent,
            progress_cb=progress_cb,
            cancel_ev=cancel_ev,
            proc_cb=proc_cb,
        )
        return (ok, out_rar) if ok else (False, msg)

    # 多输入：用绝对路径 + -ep（排除路径），保证不同目录来源也能合并到同一压缩包
    recurse = any(p.is_dir() for p in input_paths)
    ok, msg = _rar_run(
        rar_exe,
        out_rar,
        [str(p) for p in input_paths],
        recurse=recurse,
        exclude_paths=True,
        cwd=None,
        progress_cb=progress_cb,
        cancel_ev=cancel_ev,
        proc_cb=proc_cb,
    )
    return (ok, out_rar) if ok else (False, msg)


def run_double_compress(
    rar_exe: Path,
    input_paths: list[Path],
    *,
    output_dir: Path | None = None,
    progress_cb: Callable[[int | None, float], None] | None = None,
    phase_cb: Callable[[int, int], None] | None = None,
    cancel_ev: threading.Event | None = None,
    proc_cb: Callable[[subprocess.Popen[str] | None], None] | None = None,
) -> tuple[bool, str, str | None]:
    """
    第一次：输入(文件/文件夹) -> base.rar；
    改名 base.<随机伪装后缀>；
    第二次：对伪装文件再打包 -> base.rar。
    成功后删除中间伪装文件。返回 (ok, message, detail_for_log)。
    """
    if not input_paths:
        return False, "输入为空", None
    for p in input_paths:
        if not p.exists():
            return False, f"路径不存在: {p}", None

    first_parent = input_paths[0].parent
    out_dir = output_dir or first_parent
    if phase_cb:
        phase_cb(1, 2)
    ok1, out_or_err = run_rar_archive(
        rar_exe,
        input_paths,
        output_dir=out_dir,
        progress_cb=progress_cb,
        cancel_ev=cancel_ev,
        proc_cb=proc_cb,
    )
    if not ok1:
        return False, str(out_or_err), None
    assert isinstance(out_or_err, Path)
    out_rar_1 = out_or_err
    base_1 = out_rar_1.with_suffix("")

    disguised: Path | None = None
    for _ in range(48):
        ext = _pick_fake_ext()
        candidate = base_1.with_suffix(ext)
        if not candidate.exists():
            disguised = candidate
            break
    if disguised is None:
        out_rar_1.unlink(missing_ok=True)
        return False, "无法为伪装文件分配不冲突的文件名", None

    try:
        out_rar_1.rename(disguised)
    except OSError as e:
        out_rar_1.unlink(missing_ok=True)
        return False, f"重命名失败: {e}", None

    # 第二次压缩的输出压缩包也要随机命名
    try:
        out_rar_2 = _pick_nonexistent_path(out_dir, ".rar")
    except RuntimeError as e:
        return False, str(e), str(disguised)

    if phase_cb:
        phase_cb(2, 2)
    ok2, err2 = _rar_run(
        rar_exe,
        out_rar_2,
        [disguised.name],
        recurse=False,
        exclude_paths=True,
        cwd=disguised.parent,
        progress_cb=progress_cb,
        cancel_ev=cancel_ev,
        proc_cb=proc_cb,
    )
    if not ok2:
        return False, f"二次压缩失败: {err2}\n中间文件保留为: {disguised}", str(disguised)

    try:
        disguised.unlink()
    except OSError:
        pass

    detail = (
        f"二次压缩完成。外层: {out_rar_2}（内层为伪装扩展名文件，已删除磁盘上的中间文件）"
    )
    return True, str(out_rar_2), detail


def _compute_volume_spec(input_file_size_bytes: int) -> tuple[str, int, str | None]:
    """
    分卷规则：
    - 至少分 2 卷
    - 单卷目标大小 = 总体积 / 2 + 1MB
    - 单卷最大 = 2048MB + 1MB（即 2049MB）
    - 如果按照上述规则仍不足以形成至少 2 卷，则返回提示并取消

    返回 (WinRAR 的 -v 参数, 预计卷数下限, 取消原因提示或 None)。
    """
    one_mib = 1024 * 1024
    max_per_vol_mb = 2048 + 1  # 2049MB
    max_per_vol_bytes = max_per_vol_mb * one_mib

    if input_file_size_bytes <= 0:
        return "-v1m", 1, "文件大小无效，无法计算分卷"

    # 目标：总体积/2 + 1MB（先做向上取整，避免因为取整导致更少卷）
    target_per_vol_bytes = math.ceil(input_file_size_bytes / 2) + one_mib
    per_vol_bytes = min(target_per_vol_bytes, max_per_vol_bytes)
    per_vol_mb = max(1, math.ceil(per_vol_bytes / one_mib))

    est_parts = max(1, math.ceil(input_file_size_bytes / (per_vol_mb * one_mib)))
    if est_parts < 2:
        # 规则 4：文件过小不足以分卷 -> 给提示并取消
        return f"-v{per_vol_mb}m", est_parts, (
            f"文件过小，无法分卷（目标单卷大小约 {per_vol_mb}MB，但预计仅 {est_parts} 卷）。"
        )

    return f"-v{per_vol_mb}m", est_parts, None


def run_triple_compress(
    rar_exe: Path,
    input_paths: list[Path],
    *,
    output_dir: Path | None = None,
    progress_cb: Callable[[int | None, float], None] | None = None,
    phase_cb: Callable[[int, int], None] | None = None,
    cancel_ev: threading.Event | None = None,
    proc_cb: Callable[[subprocess.Popen[str] | None], None] | None = None,
) -> tuple[bool, str, str | None]:
    """
    三次压缩（分卷）：
    1) folder -> out_rar (一次)
    2) out_rar 改名为伪装后缀文件 disguised (二次输入)，再打包回 out_rar (二次)
    3) out_rar 再改名为伪装后缀文件 disguised2，并将 disguised2 分卷打包为 out_rar.part*.rar (三次)

    成功后删除中间伪装文件；最终产物为分卷：out_rar.part1.rar、out_rar.part2.rar...
    """
    if not input_paths:
        return False, "输入为空", None
    for p in input_paths:
        if not p.exists():
            return False, f"路径不存在: {p}", None

    # 先做二次压缩，得到 out_rar
    first_parent = input_paths[0].parent
    out_dir = output_dir or first_parent
    def phase_2_to_3(step: int, _total: int) -> None:
        if phase_cb:
            phase_cb(step, 3)

    ok2, msg2, extra2 = run_double_compress(
        rar_exe,
        input_paths,
        output_dir=out_dir,
        progress_cb=progress_cb,
        phase_cb=phase_2_to_3 if phase_cb else None,
        cancel_ev=cancel_ev,
        proc_cb=proc_cb,
    )
    if not ok2:
        return False, msg2, extra2
    out_rar_2 = Path(msg2)
    base_2 = out_rar_2.with_suffix("")

    try:
        size = out_rar_2.stat().st_size
    except OSError as e:
        return False, f"读取二次压缩文件大小失败: {e}", None

    vol_spec, est_parts, cancel_hint = _compute_volume_spec(size)
    if cancel_hint is not None:
        return False, cancel_hint, None

    # 第三次：把 out_rar 伪装为普通后缀，再分卷打包回 out_rar（会生成 part*.rar）
    disguised2: Path | None = None
    for _ in range(48):
        ext = _pick_fake_ext()
        candidate = base_2.with_suffix(ext)
        if not candidate.exists():
            disguised2 = candidate
            break
    if disguised2 is None:
        out_rar_2.unlink(missing_ok=True)
        return False, "无法为第三次伪装文件分配不冲突的文件名", None

    try:
        out_rar_2.rename(disguised2)
    except OSError as e:
        out_rar_2.unlink(missing_ok=True)
        return False, f"第三次重命名失败: {e}", None

    # 第三次压缩（分卷）的输出压缩包也要随机命名
    try:
        out_rar_3 = _pick_nonexistent_path(out_dir, ".rar")
    except RuntimeError as e:
        return False, str(e), str(disguised2)

    if phase_cb:
        phase_cb(3, 3)
    ok3, err3 = _rar_run(
        rar_exe,
        out_rar_3,
        [disguised2.name],
        recurse=False,
        exclude_paths=True,
        volume_spec=vol_spec,
        cwd=disguised2.parent,
        progress_cb=progress_cb,
        cancel_ev=cancel_ev,
        proc_cb=proc_cb,
    )
    if not ok3:
        return (
            False,
            f"三次分卷压缩失败: {err3}\n中间文件保留为: {disguised2}",
            str(disguised2),
        )

    try:
        disguised2.unlink()
    except OSError:
        pass

    detail = (
        f"三次压缩完成（分卷）。目标单卷≈总/2+1MB，最大≤2049MB。预计至少 {est_parts} 卷，WinRAR 参数 {vol_spec}。"
        f" 输出示例: {out_rar_3.with_suffix('')}.part1.rar"
    )
    return True, f"{out_rar_3.with_suffix('')}.part*.rar", detail


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("拖放文件夹 — WinRAR 打包")
        self.minsize(480, 320)
        self.geometry("560x400")

        self._rar = find_rar_exe()
        self._busy = False
        self._var_double = tk.BooleanVar(value=False)
        self._var_triple = tk.BooleanVar(value=False)
        self._status_var = tk.StringVar(value="")
        self._phase_text = "就绪"
        self._job_start_t: float | None = None
        self._job_tick_id: str | None = None
        self._cancel_ev: threading.Event | None = None
        self._current_proc: subprocess.Popen[str] | None = None
        self._close_after_cancel = False

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

        if self._rar:
            rar_line = f"已找到: {self._rar}"
        else:
            rar_line = "未找到 WinRAR（Rar.exe）。请安装 WinRAR 或检查安装路径。"

        frm_top = tk.Frame(self)
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

        self.drop_frame = tk.Frame(self, relief=tk.SUNKEN, bd=2, bg="#e8eef5")
        self.drop_frame.pack(fill=tk.BOTH, expand=True, **pad)

        self.lbl_drop = tk.Label(
            self.drop_frame,
            text="拖放区域",
            bg="#e8eef5",
            fg="#456",
            font=("Segoe UI", 14),
        )
        self.lbl_drop.pack(expand=True)

        btn_row = tk.Frame(self)
        btn_row.pack(fill=tk.X, **pad)

        tk.Button(btn_row, text="选择文件夹…", command=self._pick_folder).pack(
            side=tk.LEFT
        )
        tk.Button(btn_row, text="清空日志", command=self._clear_log).pack(
            side=tk.LEFT, padx=(8, 0)
        )
        self.btn_cancel = tk.Button(
            btn_row, text="取消", command=self._cancel_current, state=tk.DISABLED
        )
        self.btn_cancel.pack(side=tk.RIGHT)

        self.log = scrolledtext.ScrolledText(self, height=8, state=tk.DISABLED)
        self.log.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 12))

        self.status = tk.Label(self, textvariable=self._status_var, anchor=tk.W)
        self.status.pack(fill=tk.X, padx=12, pady=(0, 10))

        self._log_line(
            "就绪。默认：每进行一次压缩都会生成 5 位随机文件名（英文大小写+数字）的 .rar。"
            " 二次：在第一次产物基础上改伪装后缀再打包，外层 .rar 也会是新的随机名。"
            " 三次（分卷）：单卷目标≈总/2+1MB，最大≤2049MB，至少 2 卷；过小将取消。"
        )
        self._set_status(0.0)

    def _fmt_time(self, seconds: float) -> str:
        seconds = max(0.0, float(seconds))
        m, s = divmod(int(seconds + 0.5), 60)
        h, m = divmod(m, 60)
        if h:
            return f"{h:02d}:{m:02d}:{s:02d}"
        return f"{m:02d}:{s:02d}"

    def _set_phase(self, step: int | None, total: int) -> None:
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
            if not self._busy or self._job_start_t is None:
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

    def _on_close(self) -> None:
        if self._busy:
            self._close_after_cancel = True
            self._cancel_current()
            # 等任务结束后在 done() 里关闭
            self._log_line("正在取消并退出…")
            return
        self.destroy()

    def _hook_drop(self) -> None:
        if windnd is None:
            self._log_line("未安装 windnd，拖放不可用。请运行: pip install -r requirements.txt")
            return
        try:
            kw = {"func": self._on_drop_files, "force_unicode": True}
            windnd.hook_dropfiles(self.drop_frame, **kw)
            windnd.hook_dropfiles(self.lbl_drop, **kw)
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
            temp_dir = target_dir / f".apwr_tmp_{_random_token(10)}"
            try:
                temp_dir.mkdir(parents=False, exist_ok=False)
            except Exception as e:
                return self.after(
                    0, done_with_result, False, f"创建临时目录失败: {e}", None
                )

            def progress_cb(_pct: int | None, _elapsed: float) -> None:
                # 已用时间由界面计时器实时刷新；这里不再处理进度/剩余时间
                return

            def phase_cb(step: int, total: int) -> None:
                self.after(0, self._set_phase, step, total)

            def proc_cb(p: subprocess.Popen[str] | None) -> None:
                self._set_current_proc(p)

            ok: bool
            msg: str
            extra: str | None
            is_volumes = False
            temp_outputs: list[Path] = []

            try:
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
                    ok_o, out_or_err = run_rar_archive(
                        self._rar,
                        items,
                        output_dir=temp_dir,
                        progress_cb=progress_cb,
                        cancel_ev=self._cancel_ev,
                        proc_cb=proc_cb,
                    )
                    if not ok_o:
                        ok, msg, extra = False, str(out_or_err), None
                    else:
                        assert isinstance(out_or_err, Path)
                        temp_outputs = [out_or_err]
                        ok, msg, extra = True, "OK", None

                if self._cancel_ev is not None and self._cancel_ev.is_set():
                    ok, msg = False, "已取消"

                if ok:
                    okc, outmsg, finals = _commit_outputs_atomic(
                        temp_outputs, target_dir, is_volumes=is_volumes
                    )
                    if not okc:
                        ok, msg = False, outmsg
                    else:
                        ok, msg = True, outmsg
            finally:
                _safe_rmtree(temp_dir)

            return self.after(0, done_with_result, ok, msg, extra)

        def done_with_result(ok: bool, msg: str, extra: str | None) -> None:
            self._busy = False
            self.lbl_drop.configure(text="拖放区域")
            self._stop_job_timer()
            self.btn_cancel.configure(state=tk.DISABLED)
            self._set_current_proc(None)
            if ok:
                self._log_line(f"完成: {msg}")
                if extra:
                    self._log_line(extra)
                messagebox.showinfo("完成", f"已生成:\n{msg}")
            else:
                self._log_line(f"失败: {msg}")
                # 取消时不弹错误框
                if msg != "已取消":
                    messagebox.showerror("打包失败", msg)
                else:
                    self._log_line("已取消。")

            if self._close_after_cancel:
                self.destroy()

        threading.Thread(target=work, daemon=True).start()


if __name__ == "__main__":
    App().mainloop()
