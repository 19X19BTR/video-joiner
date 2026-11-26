#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
排列生成逻辑的单元测试（纯标准库，不需要 ffmpeg）：

    python -m unittest discover -s tests -t . -v
"""
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from core.permuter import (  # noqa: E402
    generate_all_permutations,
    parse_key_to_names,
    perm_to_filename,
    perm_to_key,
)


def make_groups(prefixes=('A', 'B', 'C'), n=2):
    """构造 {'A': [A1, A2], 'B': [...]} 这样的分组数据"""
    return {p: [{'name': f'{p}{i}', 'path': f'/tmp/{p}{i}.mp4'} for i in range(1, n + 1)]
            for p in prefixes}


class TestKeyHelpers(unittest.TestCase):
    def test_perm_to_key(self):
        perm = {'videos': [('A', {'name': 'A1'}), ('B', {'name': 'B2'})]}
        self.assertEqual(perm_to_key(perm), 'A1B2')

    def test_parse_key_to_names(self):
        self.assertEqual(parse_key_to_names('A1B2C3D1'), ['A1', 'B2', 'C3', 'D1'])

    def test_文件名带日期前缀(self):
        perm = {'videos': [('A', {'name': 'A1'}), ('B', {'name': 'B2'})]}
        self.assertEqual(perm_to_filename(perm, '20260601'), '20260601_A1B2.mp4')

    def test_文件名清理特殊字符(self):
        perm = {'videos': [('A', {'name': '开头 1'}), ('B', {'name': '（痛点）2'})]}
        name = perm_to_filename(perm, '20260601')
        self.assertNotIn(' ', name.replace('20260601_', ''))
        self.assertTrue(name.endswith('.mp4'))


class TestOrderedMode(unittest.TestCase):
    def test_数量等于笛卡尔积(self):
        groups = make_groups(('A', 'B', 'C'), 2)  # 2*2*2
        perms = generate_all_permutations(groups, ordered=True)
        self.assertEqual(len(perms), 8)

    def test_有序模式组顺序固定(self):
        groups = make_groups(('A', 'B', 'C'), 1)
        perms = generate_all_permutations(groups, ordered=True)
        self.assertEqual(perms[0]['order'], ['A', 'B', 'C'])

    def test_每条排列每组只取一个视频(self):
        groups = make_groups(('A', 'B'), 2)
        for perm in generate_all_permutations(groups, ordered=True):
            keys = [k for k, _ in perm['videos']]
            self.assertEqual(sorted(keys), ['A', 'B'])

    def test_排列不重复(self):
        groups = make_groups(('A', 'B', 'C'), 2)
        perms = generate_all_permutations(groups, ordered=True)
        keys = [perm_to_key(p) for p in perms]
        self.assertEqual(len(keys), len(set(keys)))

    def test_序号从1开始连续(self):
        groups = make_groups(('A', 'B'), 2)
        perms = generate_all_permutations(groups, ordered=True)
        self.assertEqual([p['index'] for p in perms], list(range(1, len(perms) + 1)))

    def test_不污染原始分组(self):
        groups = make_groups(('A', 'B'), 3)
        before = {k: [v['name'] for v in vs] for k, vs in groups.items()}
        generate_all_permutations(groups, ordered=True)
        after = {k: [v['name'] for v in vs] for k, vs in groups.items()}
        self.assertEqual(before, after)


class TestUnorderedMode(unittest.TestCase):
    def test_无序模式可复现(self):
        groups = make_groups(('A', 'B', 'C'), 2)
        first = [perm_to_key(p) for p in generate_all_permutations(groups, ordered=False)]
        second = [perm_to_key(p) for p in generate_all_permutations(groups, ordered=False)]
        self.assertEqual(first, second)

    def test_无序模式组顺序会变化(self):
        groups = make_groups(('A', 'B', 'C', 'D'), 2)
        orders = {tuple(p['order']) for p in generate_all_permutations(groups, ordered=False)}
        self.assertGreater(len(orders), 1)

    def test_无序模式数量不变(self):
        groups = make_groups(('A', 'B'), 3)  # 3*3
        perms = generate_all_permutations(groups, ordered=False)
        self.assertEqual(len(perms), 9)
        self.assertEqual(len({perm_to_key(p) for p in perms}), 9)


class TestPartialFixedMode(unittest.TestCase):
    def test_固定组停在原位置(self):
        groups = make_groups(('A', 'B', 'C'), 2)
        perms = generate_all_permutations(groups, ordered=False, fixed_groups=['A', 'C'])
        for perm in perms:
            self.assertEqual(perm['order'][0], 'A')
            self.assertEqual(perm['order'][2], 'C')

    def test_未固定的组仍会换位(self):
        groups = make_groups(('A', 'B', 'C', 'D'), 2)
        perms = generate_all_permutations(groups, ordered=False, fixed_groups=['A'])
        second_slots = {perm['order'][1] for perm in perms}
        self.assertGreater(len(second_slots), 1)


if __name__ == '__main__':
    unittest.main()
