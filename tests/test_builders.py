# -*- coding: utf-8 -*-
"""公司模板构建器：只断言关键单元格，不做全表 diff。"""
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

import openpyxl  # noqa: E402

from fapiao.builders.template import build_workbook  # noqa: E402
from fapiao.config import COMPANY_TEMPLATE  # noqa: E402
from fapiao.models.shipment import Carton, Shipment  # noqa: E402


def sample_shipment() -> Shipment:
    ship = Shipment(customer="德邦", ref="DPL112000167379", doc_no="FBA19V3WLF7WU000001",
                    channel="美森特快线", warehouse="HIA1", battery="否", magnet="否",
                    customs="不需报关", total_boxes=1, total_weight=16.45)
    ship.cartons = [Carton(
        seq="1", boxes=1, fba_id="FBA19V3WLF7WU000001", po="78481RND",
        name_cn="棒球帽", name_en="Baseball Cap", hs_code="6506999000",
        weight=16.45, length=57, width=44, height=42, qty=160, unit=2.0, total=320.0,
        currency="USD", brand="无", model="无",
        mat_cn="", mat_en="20% polyester 80% cotton",
        use_cn="", use_en="wear", image="/")]
    return ship


class TestTemplateBuilder(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.out = Path(self.tmp.name) / "out.xlsx"

    def tearDown(self):
        self.tmp.cleanup()

    @unittest.skipUnless(COMPANY_TEMPLATE.exists(), "缺少 001模板文件/公司模板.xlsx")
    def test_head_and_detail(self):
        build_workbook(sample_shipment(), self.out, COMPANY_TEMPLATE)
        wb = openpyxl.load_workbook(self.out)
        ws = wb["出货发票"]

        self.assertEqual(ws["B4"].value, "DPL112000167379")
        self.assertEqual(ws["B6"].value, "美森特快线")
        self.assertEqual(ws["B7"].value, "HIA1")
        self.assertEqual(ws["B8"].value, "否")
        self.assertEqual(ws["B10"].value, "不需报关")
        self.assertEqual(ws["B11"].value, 1)
        # 预报方数 / 申报总值 保持空白（与手工样例一致）
        self.assertIsNone(ws["B13"].value)
        self.assertIsNone(ws["B14"].value)

        row = 18
        self.assertEqual(ws.cell(row, 5).value, "FBA19V3WLF7WU000001")
        self.assertEqual(ws.cell(row, 6).value, "78481RND")
        self.assertEqual(ws.cell(row, 7).value, "棒球帽")
        self.assertEqual(ws.cell(row, 8).value, "Baseball Cap")
        self.assertEqual(ws.cell(row, 9).value, "6506999000")
        self.assertAlmostEqual(ws.cell(row, 10).value, 16.45)
        self.assertEqual(ws.cell(row, 11).value, 57)
        self.assertEqual(ws.cell(row, 14).value, 160)
        self.assertEqual(ws.cell(row, 15).value, 2)
        self.assertEqual(ws.cell(row, 17).value, 320)
        wb.close()

    @unittest.skipUnless(COMPANY_TEMPLATE.exists(), "缺少 001模板文件/公司模板.xlsx")
    def test_receiver_from_address_book(self):
        ship = sample_shipment()
        build_workbook(ship, self.out, COMPANY_TEMPLATE)
        wb = openpyxl.load_workbook(self.out)
        ws = wb["出货发票"]
        self.assertTrue(str(ws["G6"].value or "").strip())   # 地址一有值
        self.assertEqual(ws["G8"].value, "PA")
        wb.close()


if __name__ == "__main__":
    unittest.main()
