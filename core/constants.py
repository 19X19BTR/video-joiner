"""
constants.py - 常量 & 路径配置
"""

import os
import sys


def _get_app_dir():
    """获取应用根目录（打包后为exe所在目录，开发时为脚本所在目录）。"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    # core/constants.py → core/ → 项目根目录
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


APP_DIR = _get_app_dir()
DEFAULT_BGM_DIR = os.path.join(APP_DIR, 'BGM')
DEFAULT_INPUT_ROOT = os.path.join(APP_DIR, 'input')
DEFAULT_OUTPUT_ROOT = os.path.join(APP_DIR, 'output')
CONFIG_PATH = os.path.join(APP_DIR, 'config.json')

VIDEO_EXTS = {'.mp4', '.mov', '.avi', '.mkv', '.flv', '.wmv', '.webm', '.m4v', '.ts'}
AUDIO_EXTS = {'.mp3', '.wav', '.aac', '.flac', '.ogg', '.m4a', '.wma'}
DEFAULT_FPS = 30
DEFAULT_CRF = 23
DEFAULT_AUDIO_BITRATE = '128k'

# 超过此数量时弹出确认
WARN_PERM_COUNT = 500

_probe_cache = {}
