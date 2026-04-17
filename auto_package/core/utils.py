import os
import random
import re
from pathlib import Path
from typing import Callable, Iterable
import shutil
import tempfile
import subprocess
import threading
from auto_package.constants import _FAKE_EXTS, _EXECUTABLE_EXTS, _ARCHIVE_EXTS, _NAME_CHARS

def decode_drop_path(raw: bytes) -> str:
    for enc in ("utf-8", "mbcs", "gbk"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")

def _pick_fake_ext() -> str:
    return random.choice(_FAKE_EXTS)

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

def _is_archive_file(path: Path) -> bool:
    """
    检查文件是否为压缩包文件
    """
    if not path.is_file():
        return False
    # 检查是否为 .rar 文件或 .part*.rar 文件
    name = path.name.lower()
    if name.endswith(".rar"):
        return True
    if ".part" in name and name.endswith(".rar"):
        return True
    return False

def _find_single_rar_in_dir(directory: Path) -> Path | None:
    """
    在目录中查找单个 .rar 文件
    """
    rars = list(directory.glob("*.rar"))
    if len(rars) == 1:
        return rars[0]
    return None

def _has_exe_files(directory: Path) -> bool:
    """
    检查目录中是否包含 .exe 文件
    """
    for item in directory.iterdir():
        if item.is_file() and item.suffix.lower() == '.exe':
            return True
    return False

def _has_mixed_content(directory: Path) -> tuple[bool, bool]:
    """
    检查目录内容类型
    返回: (是否有文件, 是否有文件夹)
    """
    has_files = False
    has_folders = False
    for item in directory.iterdir():
        if item.is_file():
            has_files = True
        elif item.is_dir():
            has_folders = True
    return has_files, has_folders

def _is_known_archive(file_path: Path) -> bool:
    """
    检查文件是否为已知压缩格式（.rar/.7z/.zip）
    """
    return file_path.suffix.lower() in _ARCHIVE_EXTS

def _count_non_archive_files(directory: Path) -> int:
    """
    统计目录中非压缩格式文件的数量
    """
    count = 0
    for item in directory.iterdir():
        if item.is_file() and not _is_known_archive(item):
            count += 1
    return count

def _commit_outputs_atomic(
    temp_outputs: list[Path],
    target_dir: Path,
    *, 
    is_volumes: bool,
) -> tuple[bool, str, list[Path]]:
    """
    将临时目录中的产物提交到 target_dir。
    - 所有压缩模式：生成新的随机名文件夹，将产物放入其中

    返回 (ok, message, final_paths)。
    注意：多文件无法做到"操作系统级原子"，这里保证失败时会尽力回滚到"目标目录无产物"状态。
    """
    import shutil
    moved: list[Path] = []
    try:
        # 生成随机文件夹名
        for _ in range(200):
            folder_name = _random_archive_stem(5)
            output_folder = target_dir / folder_name
            if not output_folder.exists():
                break
        else:
            return False, "无法生成不冲突的文件夹名称", []

        # 创建文件夹
        try:
            output_folder.mkdir(exist_ok=False)
        except Exception as e:
            return False, f"创建文件夹失败: {e}", []

        # 移动所有产物到文件夹
        for src in temp_outputs:
            dst = output_folder / src.name
            shutil.move(str(src), str(dst))
            moved.append(dst)

        moved.sort()
        return True, str(output_folder), moved
    except Exception as e:
        # rollback: delete anything moved
        for p in moved:
            _safe_unlink(p)
        # 尝试删除创建的文件夹
        try:
            if 'output_folder' in locals() and output_folder.exists():
                shutil.rmtree(output_folder, ignore_errors=True)
        except Exception:
            pass
        return False, f"提交产物失败: {e}", []