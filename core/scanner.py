"""
scanner.py - 视频扫描与探测
"""

import json
import re
import subprocess
from collections import defaultdict
from pathlib import Path

from core.constants import VIDEO_EXTS, DEFAULT_FPS, _probe_cache


def scan_videos(input_dir: str, rules: dict = None) -> dict:
    """
    扫描文件夹，按规则分组。

    rules 为 None 时：自动按首字母分组 (A1.mp4, B1.mp4...)
    rules 不为 None 时：按关键词匹配 ({'A': '开头', 'B': '痛点', ...})
    """
    groups = defaultdict(list)
    input_path = Path(input_dir)
    if not input_path.exists():
        raise FileNotFoundError(f"输入文件夹不存在: {input_dir}")

    if rules:
        # 关键词匹配模式
        for f in sorted(input_path.iterdir()):
            if f.suffix.lower() not in VIDEO_EXTS:
                continue
            stem = f.stem
            matched_group = None
            for group_name, keyword in rules.items():
                if keyword in stem:
                    matched_group = group_name
                    break
            if not matched_group:
                continue
            # 提取数字：找文件名中的第一个数字序列
            num_match = re.search(r'(\d+)', stem)
            num = int(num_match.group(1)) if num_match else 0
            groups[matched_group].append({
                'file': f.name, 'name': stem, 'group': matched_group,
                'num': num, 'path': str(f.resolve())
            })
    else:
        # 自动按首字母分组 (A1.mp4, B1.mp4...)
        for f in sorted(input_path.iterdir()):
            if f.suffix.lower() not in VIDEO_EXTS:
                continue
            match = re.match(r'^([A-Za-z]+)(\d+)$', f.stem)
            if not match:
                continue
            group = match.group(1).upper()
            num = int(match.group(2))
            groups[group].append({
                'file': f.name, 'name': f.stem, 'group': group,
                'num': num, 'path': str(f.resolve())
            })

    for key in groups:
        groups[key].sort(key=lambda x: x['num'])
    return dict(sorted(groups.items()))


def probe_video(video_path: str) -> dict:
    if video_path in _probe_cache:
        return _probe_cache[video_path]
    cmd = ['ffprobe', '-v', 'quiet', '-print_format', 'json',
           '-show_streams', '-show_format', video_path]
    try:
        result = subprocess.run(cmd, capture_output=True, timeout=30)
    except subprocess.TimeoutExpired:
        raise RuntimeError(f"ffprobe 超时: {video_path}")

    stdout = result.stdout.decode('utf-8', errors='replace') if result.stdout else ''
    if not stdout.strip():
        raise RuntimeError(f"ffprobe 无输出: {video_path}")

    data = json.loads(stdout)
    info = {'width': 1920, 'height': 1080, 'fps': DEFAULT_FPS,
            'has_audio': False, 'duration': 0.0}
    for stream in data.get('streams', []):
        if stream['codec_type'] == 'video':
            info['width'] = int(stream.get('width', 1920))
            info['height'] = int(stream.get('height', 1080))
            r_fps = stream.get('r_frame_rate', '30/1')
            try:
                num, den = map(int, r_fps.split('/'))
                info['fps'] = round(num / den) if den else DEFAULT_FPS
            except (ValueError, ZeroDivisionError):
                info['fps'] = DEFAULT_FPS
        elif stream['codec_type'] == 'audio':
            info['has_audio'] = True
    fmt = data.get('format', {})
    info['duration'] = float(fmt.get('duration', 0))
    _probe_cache[video_path] = info
    return info


def build_video_lookup(groups: dict) -> dict:
    """构建 {video_name: video_info} 查找表。"""
    lookup = {}
    for vids in groups.values():
        for v in vids:
            lookup[v['name']] = v
    return lookup
