"""
百度网盘上传模块
提供百度网盘上传、分享链接获取等功能
"""

import re
import subprocess
import time
import threading
import psutil
from pathlib import Path
from typing import Callable
from auto_package.utils.logging_config import get_logger

logger = get_logger()


def _run_command(
    cmd: str,
    cancel_ev: threading.Event | None = None,
    proc_cb: Callable[[subprocess.Popen[str] | None], None] | None = None,
    log_cb: Callable[[str], None] | None = None,
) -> tuple[bool, str, str]:
    """
    执行命令并返回结果（非阻塞方式）

    Returns:
        (success, stdout, stderr)
    """
    def log(msg: str) -> None:
        if log_cb:
            log_cb(msg)

    p = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        creationflags=subprocess.CREATE_NO_WINDOW,
        shell=True,
        text=True,
        encoding='utf-8',
        errors='ignore',
    )
    
    # 输出进程ID
    log(f"执行命令: {cmd}")
    log(f"主进程ID: {p.pid}")
    logger.debug('执行命令: %s', cmd)
    logger.debug('主进程ID: %d', p.pid)
    
    # 监控子进程
    child_processes = []
    
    def get_child_processes(parent_pid):
        """获取所有子进程"""
        try:
            parent = psutil.Process(parent_pid)
            return parent.children(recursive=True)
        except:
            return []
    
    if proc_cb:
        proc_cb(p)

    while True:
        # 检查是否有子进程
        current_children = get_child_processes(p.pid)
        for child in current_children:
            if child.pid not in [c.pid for c in child_processes]:
                child_processes.append(child)
                log(f"发现子进程: {child.pid} ({child.name()})")
                logger.debug('发现子进程: %d (%s)', child.pid, child.name())
        
        if cancel_ev is not None and cancel_ev.is_set():
            logger.info('取消命令执行: %s', cmd)
            try:
                log(f"取消主进程: {p.pid}")
                logger.debug('取消主进程: %d', p.pid)
                # 先终止所有子进程
                for child in child_processes:
                    try:
                        log(f"终止子进程: {child.pid} ({child.name()})")
                        logger.debug('终止子进程: %d (%s)', child.pid, child.name())
                        child.terminate()
                        child.wait(timeout=1)
                    except Exception as e:
                        log(f"终止子进程失败: {e}")
                        logger.debug('终止子进程失败: %s', e)
                        try:
                            child.kill()
                        except Exception:
                            pass
                # 再终止主进程
                p.terminate()
                p.wait(timeout=1.5)
                logger.debug('主进程已终止: %d', p.pid)
            except Exception as e:
                log(f"终止进程失败: {e}")
                logger.debug('终止进程失败: %s', e)
                try:
                    p.kill()
                except Exception:
                    pass
            if proc_cb:
                proc_cb(None)
            return False, "", "已取消"

        rc = p.poll()
        if rc is not None:
            break
        time.sleep(0.1)

    stdout, stderr = p.communicate()
    if proc_cb:
        proc_cb(None)
    
    success = (p.returncode == 0)
    if success:
        logger.debug('命令执行成功: %s', cmd)
    else:
        logger.debug('命令执行失败: %s, 退出码: %d', cmd, p.returncode)

    return success, stdout, stderr


def upload_to_baidu_pan(
    output_path: Path,
    upload_base: str = "/测试",
    log_cb: Callable[[str], None] | None = None,
    cancel_ev: threading.Event | None = None,
    proc_cb: Callable[[subprocess.Popen[str] | None], None] | None = None,
) -> tuple[bool, str]:
    """
    上传文件夹到百度网盘并获取分享链接

    Args:
        output_path: 要上传的本地文件夹路径
        upload_base: 百度网盘上传基础目录
        log_cb: 日志回调函数
        cancel_ev: 取消事件
        proc_cb: 进程回调，用于保存进程引用

    Returns:
        (ok, share_link or error_message)
    """
    def log(msg: str) -> None:
        if log_cb:
            log_cb(msg)

    folder_name = output_path.name
    logger.info('开始上传: %s -> %s/%s', output_path, upload_base, folder_name)

    # 1. 创建网盘目录
    mkdir_cmd = f"BaiduPCS-Go mkdir {upload_base}/{folder_name}"
    log(f"执行命令: {mkdir_cmd}")
    logger.debug('创建网盘目录: %s', f"{upload_base}/{folder_name}")
    mkdir_ok, mkdir_out, _ = _run_command(mkdir_cmd, cancel_ev, proc_cb, log_cb)

    # 2. 立即获取分享链接（即使目录为空）
    share_cmd = f"BaiduPCS-Go share set {upload_base}/{folder_name}"
    log(f"执行命令: {share_cmd}")
    logger.debug('获取分享链接: %s', f"{upload_base}/{folder_name}")
    share_ok, share_stdout, share_stderr = _run_command(share_cmd, cancel_ev, proc_cb, log_cb)

    share_output = share_stdout

    if not share_ok:
        log(f"获取分享链接失败: {share_output}")
        logger.error('获取分享链接失败: %s', share_output)
        return False, f"获取分享链接失败"

    # 3. 提取分享链接
    log(f"分享命令输出: {share_output}")
    logger.debug('分享命令输出: %s', share_output)

    # 匹配链接（包含 - 字符、下划线等特殊字符）
    link_pattern = r'https://pan\.baidu\.com/s/[a-zA-Z0-9_-]+'
    # 匹配提取码（支持多种格式）
    pwd_pattern = r'[提取码密码瀵嗙爜][:：]\s*([a-zA-Z0-9]+)'

    links = re.findall(link_pattern, share_output)
    pwds = re.findall(pwd_pattern, share_output)

    if not links:
        log("未找到分享链接")
        logger.error('未找到分享链接')
        return False, "未找到分享链接"

    share_link = links[0]
    # 如果有提取码，添加到链接中
    if pwds:
        pwd = pwds[0]
        if '?' in share_link:
            share_link = f"{share_link}&pwd={pwd}"
        else:
            share_link = f"{share_link}?pwd={pwd}"
    log(f"获取到分享链接: {share_link}")
    logger.info('获取到分享链接: %s', share_link)

    # 4. 保存链接到文本文件（放在最终文件夹同级目录）
    link_file_path = output_path.parent / "链接.txt"
    try:
        with open(link_file_path, 'w', encoding='utf-8') as f:
            f.write(share_link)
        log(f"链接已保存到: {link_file_path}")
        logger.debug('链接已保存到: %s', link_file_path)
    except Exception as e:
        log(f"保存链接失败: {e}")
        logger.debug('保存链接失败: %s', e)

    # 5. 上传文件到创建的目录（上传目录内的内容，不包含目录本身）
    import os
    files = []
    for root, _, filenames in os.walk(str(output_path)):
        for filename in filenames:
            files.append(os.path.join(root, filename))

    if not files:
        log("未找到要上传的文件")
        logger.info('未找到要上传的文件: %s', output_path)
        return True, share_link  # 链接已获取，返回成功

    logger.debug('找到 %d 个文件要上传', len(files))
    # 构建批量上传命令
    upload_cmd = f"BaiduPCS-Go upload {' '.join(files)} {upload_base}/{folder_name}"
    log(f"执行命令: {upload_cmd}")
    logger.debug('上传文件: %s -> %s/%s', output_path, upload_base, folder_name)
    upload_ok, upload_stdout, upload_stderr = _run_command(upload_cmd, cancel_ev, proc_cb, log_cb)

    if not upload_ok:
        log(f"上传文件失败: {upload_stderr}")
        logger.error('上传文件失败: %s', upload_stderr)
        return False, f"上传文件失败: {upload_stderr}"

    log(f"上传成功: {upload_base}/{folder_name}")
    logger.info('上传成功: %s/%s', upload_base, folder_name)
    return True, share_link
