"""通用工具函数"""

import re
import sys
import platform
from pathlib import Path


def get_exe_dir():
    """获取程序所在目录（兼容 Nuitka / PyInstaller 打包后的路径）"""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent.parent


def get_safe_filename(name, max_length=100):
    """
    生成安全的文件名：
    - 移除文件系统非法字符
    - 限制长度
    - 空值保护
    """
    name = re.sub(r'[\\/*?:"<>|\r\n\t]', "_", str(name).strip())
    # 移除首尾的空格和点号（Windows 不允许以点结尾）
    name = name.strip(". ")
    if len(name) > max_length:
        name = name[:max_length - 3] + "..."
    return name or "未命名"


def get_platform_info():
    """获取平台信息字典"""
    return {
        'system': platform.system(),
        'release': platform.release(),
        'machine': platform.machine(),
        'python': platform.python_version(),
    }


def is_windows():
    return sys.platform == 'win32'


def is_macos():
    return sys.platform == 'darwin'


def is_linux():
    return sys.platform.startswith('linux')
