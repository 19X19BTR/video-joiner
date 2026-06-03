"""
permuter.py - 排列生成工具
"""

import random
import re
from itertools import product as cartesian_product


def perm_to_key(perm: dict) -> str:
    """排列唯一标识，如 'A1B2C3D1'。"""
    return ''.join(vid['name'] for _, vid in perm['videos'])


def perm_to_filename(perm: dict, date_str: str) -> str:
    """视频文件名，如 '20260601_开头1痛点2.mp4'。"""
    raw_key = perm_to_key(perm)
    # 清理：去掉空格，全角括号转半角，去掉特殊字符
    safe_key = raw_key.replace(' ', '').replace('\uff08', '(').replace('\uff09', ')')
    safe_key = re.sub(r'[^\w\-\(\)\u4e00-\u9fff]', '', safe_key)
    return f"{date_str}_{safe_key}.mp4"


def parse_key_to_names(key: str) -> list:
    """将 'A1B2C3D1' 解析为 ['A1', 'B2', 'C3', 'D1']。"""
    return re.findall(r'[A-Za-z]+\d+', key)


def generate_all_permutations(groups: dict, ordered: bool = True, fixed_groups: list = None) -> list:
    """
    一次性生成所有排列组合。

    有序: 固定 A→B→C→D 顺序
    无序: 每条排列的组别顺序随机（确定性，可复现）
    部分固定: fixed_groups 中的组保持原顺序，其余组随机排列
    """
    group_keys = sorted(groups.keys())
    video_lists = [groups[k][:] for k in group_keys]  # 拷贝，避免污染原始 groups

    fixed_groups = fixed_groups or []
    fixed_set = set(fixed_groups)

    # ── 无序/部分固定时：打乱每组视频的顺序，让视频出现顺序随机 ──
    rng = random.Random(42)
    for vl in video_lists:
        rng.shuffle(vl)
    # 固定组的视频也要打乱（部分固定模式下，组内顺序仍应随机）

    permutations = []
    for i, combo in enumerate(cartesian_product(*video_lists)):
        # combo: (A组视频, B组视频, C组视频, D组视频)
        video_map = dict(zip(group_keys, combo))

        if ordered:
            order = group_keys[:]
        elif fixed_groups:
            # 部分固定模式：固定组保持原位置，其余组随机排列
            unfixed = [k for k in group_keys if k not in fixed_set]
            shuffled = unfixed[:]
            random.Random(i * 31337).shuffle(shuffled)
            order = []
            shuffled_idx = 0
            for k in group_keys:
                if k in fixed_set:
                    order.append(k)
                else:
                    order.append(shuffled[shuffled_idx])
                    shuffled_idx += 1
        else:
            # 完全无序：每个 combo 用固定种子，跨运行可复现
            order = group_keys[:]
            random.Random(i * 31337).shuffle(order)

        ordered_videos = [(k, video_map[k]) for k in order]

        permutations.append({
            'index': i + 1,
            'order': order,
            'videos': ordered_videos,
            'concat_str': ' → '.join(v['name'] for _, v in ordered_videos),
            'status': '⏳ 待拼接',
            'output_file': ''
        })

    # ── 无序/部分固定时：打乱排列顺序，避免规律性 ──
    if not ordered:
        random.Random(99).shuffle(permutations)
        # 重编序号
        for idx, perm in enumerate(permutations):
            perm['index'] = idx + 1

    return permutations
