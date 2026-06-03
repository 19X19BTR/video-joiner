"""
config.py - 配置管理
"""

import json
import os

from core.constants import (
    DEFAULT_INPUT_ROOT, DEFAULT_OUTPUT_ROOT, CONFIG_PATH
)


def load_config() -> dict:
    default = {
        'input_root': DEFAULT_INPUT_ROOT,
        'output_root': DEFAULT_OUTPUT_ROOT,
        'last_input': '',
        'last_output': '',
        'resolution': '自动 (以第一个视频为准)',
        'mode': 'ordered',
    }
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                saved = json.load(f)
            default.update(saved)
        except Exception:
            pass
    return default


def save_config(cfg: dict):
    try:
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
