"""通用工具函数"""

import os
import re
import sys
import platform
from pathlib import Path


def get_exe_dir():
    """获取程序所在目录（兼容 Nuitka / PyInstaller 打包后的路径）"""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent.parent


def get_app_dir():
    """
    获取应用程序主脚本所在目录

    兼容多种运行模式：
    - 源码直接运行: main.py 所在目录
    - Nuitka onefile: 临时解压目录（数据文件在此）
    - Nuitka standalone: dist 文件夹
    - PyInstaller onefile: sys._MEIPASS
    """
    # PyInstaller onefile
    if hasattr(sys, '_MEIPASS'):
        return Path(sys._MEIPASS)

    # Nuitka onefile: __compiled__ 存在时，
    # 数据文件解压到 exe 同级目录或 sys.argv[0] 同级
    if "__compiled__" in dir():
        return Path(sys.argv[0]).resolve().parent

    # Nuitka standalone 或普通 frozen
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent

    # 源码运行: 返回 main.py 所在目录
    return Path(sys.argv[0]).resolve().parent if sys.argv else Path.cwd()


def get_resource_path(relative_path):
    """
    获取资源文件的绝对路径

    :param relative_path: 相对于主脚本目录的路径，如 "images/sales/qr.png"
    :return: Path 对象
    """
    return get_app_dir() / relative_path


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
