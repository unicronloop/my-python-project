# -*- coding: utf-8 -*-
"""真实样例解析断言（数据目录缺失时自动跳过）。"""
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from fapiao.adapters.cyj import parse_chuangyingjun  # noqa: E402
from fapiao.adapters.deppon import parse_deppon  # noqa: E402
from fapiao.config import CYJ_DIR, DEPPON_DIR  # noqa: E402
from fapiao.models.totals import compute_totals  # noqa: E402

CYJ_FILE = CYJ_DIR / "客户发票.xlsx"


@unittest.skipUnless(CYJ_FILE.exists(), "缺少 04创颖峻/客户发票.xlsx")
class TestChuangyingjunParser(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.ships = parse_chuangyingjun(CYJ_FILE)

    def test_split_by_rtd(self):
        self.assertEqual(len(self.ships), 16)
        self.assertEqual(sum(s.total_boxes for s in self.ships), 179)

    def test_head_fields(self):
        ship = next(s for s in self.ships if s.doc_no == "RTD261027129")
        self.assertEqual(ship.ref, "CH2610290192")
        self.assertEqual(ship.warehouse, "MCI4")
        self.assertEqual(ship.channel, "美森特快线")
        self.assertEqual(ship.customs, "不需报关")
        self.assertEqual(ship.battery, "否")

    def test_carton_fields(self):
        ship = next(s for s in self.ships if s.doc_no == "RTD261027129")
        self.assertEqual(len(ship.cartons), 20)
        self.assertTrue(all(c.unit == 5.0 for c in ship.cartons))
        self.assertTrue(all(c.total == 5.0 * c.qty for c in ship.cartons))
        self.assertTrue(all(c.fba_id for c in ship.cartons))

    def test_leaf_category_from_source(self):
        ship = next(s for s in self.ships if s.doc_no == "RTD261026250")
        self.assertTrue(all("->" not in c.name_cn for c in ship.cartons))

    def test_totals_match_huizong(self):
        t = compute_totals(next(s for s in self.ships if s.doc_no == "RTD261027129"))
        self.assertEqual(t.boxes, 20)
        self.assertAlmostEqual(t.qty, 277)
        self.assertAlmostEqual(t.weight, 291.77, places=2)


@unittest.skipUnless(DEPPON_DIR.exists(), "缺少 05德邦 样例")
class TestDepponParser(unittest.TestCase):
    """05德邦 下的真实文件（4 个来源目录、7 种版式）。"""

    @staticmethod
    def _find(prefix: str) -> Path | None:
        hits = [p for p in DEPPON_DIR.rglob(f"{prefix}*.xls*") if not p.name.startswith("~$")]
        return hits[0] if hits else None

    def _parse(self, prefix: str):
        path = self._find(prefix)
        if path is None:
            self.skipTest(f"缺少 {prefix} 样例")
        return parse_deppon(path, self._ref(path))

    @staticmethod
    def _ref(path: Path) -> str:
        from fapiao.adapters.deppon import DPL_RE
        m = DPL_RE.search(path.name)
        return m.group(0).upper() if m else ""

    def test_template_detection(self):
        from fapiao.adapters.deppon import detect_invoice_kind
        from fapiao.common.excel import load_workbook_any

        cases = {
            "DPL112000165026": "发票2",
            "DPL112000134511": "发票2",
            "DPL112000138396": "发票3",
            "DPL112000141187": "发票5",
        }
        for prefix, expect in cases.items():
            path = self._find(prefix)
            if path is None:
                continue
            wb = load_workbook_any(path, data_only=True)
            try:
                self.assertEqual(detect_invoice_kind(wb), expect, prefix)
            finally:
                wb.close()

    def test_invoice1_std(self):
        ship = self._parse("DPL112000165026")
        self.assertEqual(ship.warehouse, "IAH3")
        self.assertEqual(ship.channel, "美森特快线")
        self.assertEqual(ship.total_boxes, 1)
        self.assertTrue(ship.cartons[0].fba_id.startswith("FBA"))

    def test_invoice2_multi_box(self):
        ship = self._parse("DPL112000134511")
        self.assertEqual(ship.warehouse, "POC3")
        self.assertGreaterEqual(len(ship.cartons), 10)
        self.assertTrue(all(c.weight for c in ship.cartons))

    def test_invoice3(self):
        ship = self._parse("DPL112000138396")
        self.assertEqual(ship.warehouse, "TOL3")
        self.assertEqual(ship.total_boxes, 2)

    def test_invoice5_packing_list(self):
        ship = self._parse("DPL112000141187")
        self.assertEqual(ship.warehouse, "RFD2")
        self.assertEqual(ship.total_boxes, 2)
        # B12 预报实重在金标里留空，重量只从明细汇总
        self.assertAlmostEqual(compute_totals(ship).weight or 0, 23.0, places=1)

    def test_xls_supported(self):
        path = self._find("DPL112000165031")
        if path is None or path.suffix != ".xls":
            self.skipTest("缺少 .xls 样例")
        ship = parse_deppon(path, self._ref(path))
        self.assertEqual(ship.customer, "德邦")
        self.assertTrue(ship.cartons)


@unittest.skipUnless((ROOT / "06尚睿" / "客户发票.xlsx").exists(), "缺少 06尚睿 样例")
class TestShangruiParser(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        from fapiao.adapters.shangrui import parse_shangrui
        cls.ship = parse_shangrui(ROOT / "06尚睿" / "客户发票.xlsx")

    def test_head(self):
        s = self.ship
        self.assertEqual(s.ref, "LKL260910-A05")
        self.assertEqual(s.warehouse, "1PL-USC-TX")
        self.assertEqual(s.channel, "美西洛杉矶卡派含税")
        self.assertEqual(s.customs, "合并报关")
        self.assertEqual(s.total_boxes, 250)
        self.assertAlmostEqual(s.total_value or 0, 112900, places=0)

    def test_cartons(self):
        self.assertEqual(len(self.ship.cartons), 250)
        self.assertTrue(all(c.qty for c in self.ship.cartons))
        self.assertTrue(all(c.length and c.width and c.height for c in self.ship.cartons))

    def test_receiver(self):
        rc = self.ship.receiver
        self.assertEqual(rc["state"], "TX")
        self.assertEqual(rc["zip"], "77338")
        self.assertEqual(rc["region"], "美西")
        self.assertTrue(rc["phone"] and rc["email"])


if __name__ == "__main__":
    unittest.main()
