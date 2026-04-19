import json
from pathlib import Path
from auto_package.constants import _CONFIG_FILE
from auto_package.utils.logging_config import get_logger

logger = get_logger()

def _load_window_geometry() -> tuple[int, int] | None:
    """加载窗口几何尺寸，返回 (宽, 高) 或 None"""
    try:
        if _CONFIG_FILE.is_file():
            with open(_CONFIG_FILE, "r", encoding="utf-8") as f:
                config = json.load(f)
            width = config.get("window_width")
            height = config.get("window_height")
            if width and height:
                logger.debug('加载窗口大小: %dx%d', width, height)
                return (width, height)
    except Exception as e:
        logger.warning('加载窗口大小失败: %s', e)
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
        logger.debug('保存窗口大小: %dx%d', width, height)
    except Exception as e:
        logger.warning('保存窗口大小失败: %s', e)

def _load_upload_paths() -> tuple[str, list[str]]:
    """加载上传路径配置，返回 (当前路径, 历史路径列表)"""
    try:
        if _CONFIG_FILE.is_file():
            with open(_CONFIG_FILE, "r", encoding="utf-8") as f:
                config = json.load(f)
            current = config.get("upload_path", "/测试")
            history = config.get("upload_path_history", ["/测试"])
            if current not in history:
                history.insert(0, current)
            logger.debug('加载上传路径: 当前=%s, 历史=%s', current, history)
            return current, history
    except Exception as e:
        logger.warning('加载上传路径失败: %s', e)
    return "/测试", ["/测试"]

def _save_upload_paths(current: str, history: list[str]) -> None:
    """保存上传路径配置"""
    try:
        config = {}
        if _CONFIG_FILE.is_file():
            with open(_CONFIG_FILE, "r", encoding="utf-8") as f:
                config = json.load(f)
        config["upload_path"] = current
        config["upload_path_history"] = history[:10]
        with open(_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False)
        logger.debug('保存上传路径: 当前=%s, 历史=%s', current, history[:10])
    except Exception as e:
        logger.warning('保存上传路径失败: %s', e)