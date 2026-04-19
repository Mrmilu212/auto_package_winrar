import os
import logging
from logging.handlers import TimedRotatingFileHandler


_logger = None


def setup_logging():
    """配置日志系统"""
    global _logger
    
    project_root = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
    log_dir = os.path.join(project_root, 'logs')
    os.makedirs(log_dir, exist_ok=True)
    
    log_file = os.path.join(log_dir, 'app.log')
    
    _logger = logging.getLogger('auto_package')
    _logger.setLevel(logging.INFO)
    
    for handler in _logger.handlers[:]:
        _logger.removeHandler(handler)
    
    handler = TimedRotatingFileHandler(
        log_file,
        when='D',
        interval=1,
        backupCount=7,
        encoding='utf-8'
    )
    
    formatter = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s'
    )
    handler.setFormatter(formatter)
    
    _logger.addHandler(handler)
    
    return _logger


def get_logger():
    """获取日志记录器"""
    global _logger
    if _logger is None:
        _logger = setup_logging()
    return _logger
