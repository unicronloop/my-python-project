# -*- coding: utf-8 -*-
"""字符串/路径工具单测（unittest，不依赖 pytest）。"""
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from fapiao.adapters.cyj import leaf_category  # noqa: E402
from fapiao.common.strings import (  # noqa: E402
    build_output_name, cn_en, is_blank, normalize_label, num, s, split_dims,
)


class TestLeafCategory(unittest.TestCase):
    def test_takes_last_segment(self):
        self.assertEqual(leaf_category("个人护理->陶瓷陶具"), "陶瓷陶具")
        self.assertEqual(
            leaf_category("交通工具配件及附件->车身件->外饰件->车门拉手->车门拉手盖"),
            "车门拉手盖")

    def test_handles_other_separators(self):
        self.assertEqual(leaf_category("A→B＞C"), "C")
        self.assertEqual(leaf_category("服装服饰"), "服装服饰")

    def test_empty(self):
        self.assertEqual(leaf_category(""), "")
        self.assertEqual(leaf_category(None), "")


class TestStrings(unittest.TestCase):
    def test_split_dims_formats(self):
        self.assertEqual(split_dims("57*44*42"), (57.0, 44.0, 42.0))
        self.assertEqual(split_dims("36×58×32cm"), (36.0, 58.0, 32.0))
        self.assertEqual(split_dims("1.2*0.8*0.5"), (1.2, 0.8, 0.5))
        self.assertEqual(split_dims(None), (None, None, None))
        self.assertEqual(split_dims("abc"), (None, None, None))

    def test_output_name(self):
        self.assertEqual(build_output_name("T2026", "20261103"), "T2026-20261103-已处理.xlsx")
        self.assertEqual(build_output_name("", "20261103"), "未命名-20261103-已处理.xlsx")
        self.assertEqual(build_output_name("a/b", "20261103"), "a_b-20261103-已处理.xlsx")

    def test_basic_helpers(self):
        self.assertTrue(is_blank(None))
        self.assertTrue(is_blank("  "))
        self.assertFalse(is_blank(0))
        self.assertEqual(s(None), "")
        self.assertEqual(num("12.5"), 12.5)
        self.assertIsNone(num(""))
        self.assertEqual(normalize_label(" 总箱数* "), normalize_label("总箱数"))
        self.assertEqual(cn_en("棒球帽"), ("棒球帽", ""))
        self.assertEqual(cn_en("Baseball Cap"), ("", "Baseball Cap"))


if __name__ == "__main__":
    unittest.main()
