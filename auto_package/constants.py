from pathlib import Path
import string

# 配置文件路径
_CONFIG_FILE = Path(__file__).parent.parent / ".app_config.json"

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

# 可执行文件后缀（只检查 .exe）
_EXECUTABLE_EXTS = {'.exe'}

# 支持的压缩包后缀（用于识别已知压缩格式，伪装文件只尝试zip格式）
_ARCHIVE_EXTS = ['.rar', '.7z', '.zip']

_NAME_CHARS = string.ascii_letters + string.digits