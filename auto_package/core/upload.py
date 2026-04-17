"""
百度网盘上传模块
提供百度网盘上传、分享链接获取等功能
"""

import re
import subprocess
from pathlib import Path
from typing import Callable


def upload_to_baidu_pan(
    output_path: Path,
    upload_base: str = "/测试",
    log_cb: Callable[[str], None] | None = None,
    cancel_ev = None,
) -> tuple[bool, str, subprocess.Popen | None]:
    """
    上传文件夹到百度网盘并获取分享链接

    Args:
        output_path: 要上传的本地文件夹路径
        upload_base: 百度网盘上传基础目录
        log_cb: 日志回调函数

    Returns:
        (ok, share_link or error_message)
    """
    def log(msg: str) -> None:
        if log_cb:
            log_cb(msg)

    folder_name = output_path.name

    # 1. 创建网盘目录
    mkdir_cmd = f"BaiduPCS-Go mkdir {upload_base}/{folder_name}"
    log(f"执行命令: {mkdir_cmd}")
    subprocess.run(mkdir_cmd, shell=True, capture_output=True, text=True)

    # 2. 立即获取分享链接（即使目录为空）
    share_cmd = f"BaiduPCS-Go share set {upload_base}/{folder_name}"
    log(f"执行命令: {share_cmd}")
    share_result = subprocess.run(share_cmd, shell=True, capture_output=True)

    # 尝试多种编码解码
    try:
        share_output = share_result.stdout.decode('utf-8')
    except UnicodeDecodeError:
        try:
            share_output = share_result.stdout.decode('gbk')
        except UnicodeDecodeError:
            share_output = share_result.stdout.decode('latin-1')

    if share_result.returncode != 0:
        log(f"获取分享链接失败: {share_output}")
        return False, f"获取分享链接失败", None

    # 3. 提取分享链接
    log(f"分享命令输出: {share_output}")

    # 匹配链接（包含 - 字符、下划线等特殊字符）
    link_pattern = r'https://pan\.baidu\.com/s/[a-zA-Z0-9_-]+'
    # 匹配提取码（支持多种格式）
    pwd_pattern = r'[提取码密码瀵嗙爜][:：]\s*([a-zA-Z0-9]+)'

    links = re.findall(link_pattern, share_output)
    pwds = re.findall(pwd_pattern, share_output)

    if not links:
        log("未找到分享链接")
        return False, "未找到分享链接", None

    share_link = links[0]
    # 如果有提取码，添加到链接中
    if pwds:
        pwd = pwds[0]
        if '?' in share_link:
            share_link = f"{share_link}&pwd={pwd}"
        else:
            share_link = f"{share_link}?pwd={pwd}"
    log(f"获取到分享链接: {share_link}")

    # 4. 保存链接到文本文件（放在最终文件夹同级目录）
    link_file_path = output_path.parent / "链接.txt"
    try:
        with open(link_file_path, 'w', encoding='utf-8') as f:
            f.write(share_link)
        log(f"链接已保存到: {link_file_path}")
    except Exception as e:
        log(f"保存链接失败: {e}")

    # 5. 上传文件到创建的目录（上传目录内的内容，不包含目录本身）
    import os
    files = []
    for root, _, filenames in os.walk(str(output_path)):
        for filename in filenames:
            files.append(os.path.join(root, filename))

    if not files:
        log("未找到要上传的文件")
        return True, share_link, None  # 链接已获取，返回成功

    # 构建批量上传命令
    upload_cmd = f"BaiduPCS-Go upload {' '.join(files)} {upload_base}/{folder_name}"
    log(f"执行命令: {upload_cmd}")
    
    # 使用 Popen 以便后续可以终止进程
    proc = subprocess.Popen(upload_cmd, shell=True, capture_output=True, text=True)
    
    # 等待进程完成，同时检查取消信号
    while proc.poll() is None:
        if cancel_ev and cancel_ev.is_set():
            log("上传被取消")
            proc.terminate()
            try:
                proc.wait(timeout=5)
            except subprocess.TimeoutExpired:
                proc.kill()
            return False, "上传被取消", None
    
    # 检查进程返回码
    if proc.returncode != 0:
        log(f"上传文件失败: {proc.stderr.read()}")
        return False, f"上传文件失败: {proc.stderr.read()}", None

    log(f"上传成功: {upload_base}/{folder_name}")
    return True, share_link, None
