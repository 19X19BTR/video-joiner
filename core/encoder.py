"""
encoder.py - FFmpeg 拼接与 BGM 工具
"""

import os
import random
import re
import subprocess
from pathlib import Path

from core.constants import (
    AUDIO_EXTS, DEFAULT_CRF, DEFAULT_AUDIO_BITRATE, HIDDEN_SI
)
from core.scanner import probe_video


def concat_videos(video_list: list, output_path: str,
                  target_w: int = None, target_h: int = None) -> bool:
    if not video_list:
        return False
    paths = [v['path'] for v in video_list]
    n = len(paths)
    infos = [probe_video(p) for p in paths]
    if target_w is None or target_h is None:
        target_w = target_w or infos[0]['width']
        target_h = target_h or infos[0]['height']
    target_fps = infos[0]['fps']  # 以第一个视频的帧率为准
    any_has_audio = any(info['has_audio'] for info in infos)

    filter_parts = []
    for i in range(n):
        filter_parts.append(
            f"[{i}:v]scale={target_w}:{target_h}:"
            f"force_original_aspect_ratio=decrease,"
            f"pad={target_w}:{target_h}:(ow-iw)/2:(oh-ih)/2:black,"
            f"setsar=1,fps={target_fps},format=yuv420p[v{i}]"
        )
        if any_has_audio:
            if infos[i]['has_audio']:
                filter_parts.append(
                    f"[{i}:a]aformat=sample_fmts=fltp:"
                    f"sample_rates=44100:channel_layouts=stereo[a{i}]"
                )
            else:
                dur = max(infos[i]['duration'], 0.1)
                filter_parts.append(
                    f"anullsrc=r=44100:cl=stereo,atrim=0:{dur}[a{i}]"
                )

    if any_has_audio:
        concat_inputs = ''.join(f'[v{i}][a{i}]' for i in range(n))
        filter_parts.append(f"{concat_inputs}concat=n={n}:v=1:a=1[outv][outa]")
        map_args = ['-map', '[outv]', '-map', '[outa]']
    else:
        concat_inputs = ''.join(f'[v{i}]' for i in range(n))
        filter_parts.append(f"{concat_inputs}concat=n={n}:v=1:a=0[outv]")
        map_args = ['-map', '[outv]']

    filter_complex = ';'.join(filter_parts)
    cmd = ['ffmpeg', '-y', '-hide_banner', '-loglevel', 'warning']
    for p in paths:
        cmd.extend(['-i', p])
    cmd.extend(['-filter_complex', filter_complex])
    cmd.extend(map_args)
    cmd.extend(['-c:v', 'libx264', '-preset', 'medium', '-crf', str(DEFAULT_CRF),
                '-movflags', '+faststart'])
    if any_has_audio:
        cmd.extend(['-c:a', 'aac', '-b:a', DEFAULT_AUDIO_BITRATE])
    cmd.append(output_path)

    result = subprocess.run(cmd, capture_output=True, text=True, startupinfo=HIDDEN_SI)
    return result.returncode == 0


# ─── BGM 工具函数 ───

def scan_bgm_files(bgm_dir: str) -> list:
    """扫描BGM目录，返回音频文件列表。"""
    if not os.path.exists(bgm_dir):
        return []
    files = []
    for f in sorted(Path(bgm_dir).iterdir()):
        if f.suffix.lower() in AUDIO_EXTS:
            files.append({'name': f.stem, 'file': f.name, 'path': str(f.resolve())})
    return files


def measure_audio_loudness(audio_path: str, duration_limit: float = None):
    """
    测量音频/视频的响度（使用 ffmpeg volumedetect）。
    返回 {'mean_volume': float, 'max_volume': float}（单位 dB）。
    duration_limit: 只分析前 N 秒（加速预览分析）。
    """
    cmd = ['ffmpeg', '-hide_banner', '-loglevel', 'info']
    if duration_limit:
        cmd.extend(['-t', str(duration_limit)])
    cmd.extend(['-i', audio_path, '-af', 'volumedetect', '-f', 'null', '-'])

    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60, startupinfo=HIDDEN_SI)
    except subprocess.TimeoutExpired:
        return None

    stderr = result.stderr or ''
    info = {}
    for line in stderr.split('\n'):
        if 'mean_volume' in line:
            match = re.search(r'mean_volume:\s*([-\d.]+)\s*dB', line)
            if match:
                info['mean_volume'] = float(match.group(1))
        elif 'max_volume' in line:
            match = re.search(r'max_volume:\s*([-\d.]+)\s*dB', line)
            if match:
                info['max_volume'] = float(match.group(1))

    return info if 'mean_volume' in info else None


def calculate_bgm_volume(video_db: float, bgm_db: float, target_diff_db: float = 10) -> float:
    """
    根据响度测量值计算BGM音量倍数。
    video_db: 原视频平均响度 (dB)
    bgm_db: BGM平均响度 (dB)
    target_diff_db: BGM应比视频低多少dB（默认10dB，不抢人声）
    返回 0.3 ~ 2.0 的音量倍数。
    """
    needed_attenuation = (bgm_db - video_db) + target_diff_db
    volume = 10 ** (-needed_attenuation / 20)
    return max(0.3, min(2.0, round(volume, 2)))


def add_bgm_to_video(video_path: str, bgm_path: str, output_path: str,
                     video_duration: float = None, bgm_volume: float = 0.80) -> bool:
    """
    给视频添加BGM。
    - BGM裁切到视频长度（尾端超出部分截断）
    - BGM循环覆盖（如果BGM比视频短）
    - 保留原视频音频，BGM混音
    - bgm_volume: BGM音量倍数（默认1.0，原始混音）
    - normalize=0: 关闭amix自动衰减，避免音量双重降低
    """
    if video_duration is None:
        info = probe_video(video_path)
        video_duration = info['duration']

    cmd = [
        'ffmpeg', '-y', '-hide_banner', '-loglevel', 'warning',
        '-i', video_path,
        '-stream_loop', '-1',
        '-i', bgm_path,
        '-filter_complex',
        f'[0:a]aformat=sample_fmts=fltp:sample_rates=44100:channel_layouts=stereo[va];'
        f'[1:a]atrim=0:{video_duration},asetpts=PTS-STARTPTS,'
        f'aformat=sample_fmts=fltp:sample_rates=44100:channel_layouts=stereo,'
        f'volume={bgm_volume}[bgm];'
        f'[va][bgm]amerge=inputs=2,pan=stereo|c0=c0+c2|c1=c1+c3[aout]',
        '-map', '0:v', '-map', '[aout]',
        '-c:v', 'copy',
        '-c:a', 'aac', '-b:a', DEFAULT_AUDIO_BITRATE,
        '-t', str(video_duration),
        '-shortest',
        output_path
    ]

    result = subprocess.run(cmd, capture_output=True, text=True, startupinfo=HIDDEN_SI)
    return result.returncode == 0


class BgmManager:
    """BGM分配管理器：确保BGM均匀分布，每首BGM携带各自的音量倍数。"""

    def __init__(self, bgm_files: list, bgm_volumes: dict = None):
        """
        bgm_files: [{'name': ..., 'path': ...}, ...]
        bgm_volumes: {bgm_path: volume} 每首BGM各自的音量倍数
        """
        self.bgm_files = bgm_files
        self.bgm_volumes = bgm_volumes or {}  # {path: volume}
        self.default_volume = 0.80
        self._pool = []
        self._seed = 42

    def _refill_pool(self):
        """重新填充池：复制一份文件列表并打乱。"""
        self._pool = list(range(len(self.bgm_files)))
        random.Random(self._seed).shuffle(self._pool)
        self._seed += 1

    def next(self) -> tuple:
        """获取下一个BGM，返回 (bgm_info, volume) 二元组。"""
        if not self._pool:
            self._refill_pool()
        idx = self._pool.pop()
        bgm = self.bgm_files[idx]
        vol = self.bgm_volumes.get(bgm['path'], self.default_volume)
        return bgm, vol
