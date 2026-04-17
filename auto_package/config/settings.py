import json
from pathlib import Path
from auto_package.constants import _CONFIG_FILE

def _load_window_geometry() -> tuple[int, int] | None:
    """加载窗口几何尺寸，返回 (宽, 高) 或 None"""
    try:
        if _CONFIG_FILE.is_file():
            with open(_CONFIG_FILE, "r", encoding="utf-8") as f:
                config = json.load(f)
            width = config.get("window_width")
            height = config.get("window_height")
            if width and height:
                return (width, height)
    except Exception:
        pass
    return None

def _save_window_geometry(width: int, height: int) -> None:
    """保存窗口几何尺寸"""
    try:
        config = {}
        if _CONFIG_FILE.is_file():
            with open(_CONFIG_FILE, "r", encoding="utf-8") as f:
                config = json.load(f)
        config["window_width"] = width
        config["window_height"] = height
        with open(_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False)
    except Exception:
        pass