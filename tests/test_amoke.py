"""爱墨科技适配器：清关发票 → Shipment（按箱拆行）。"""

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from src.fapiao.adapters.amoke import parse_clearance

HEAD = [
    "客户参考号", None, None, None, None, "目的国", "美国",
    None, None, None, None, "是否单独报关", "是",
]
DETAIL = [
    "Shipment ID", "Reference ID", "箱号段(起始-结束)", "SKU", "英文品名",
    "中文品名", "海关代码", "Brand品牌", "材质", "用途", "ASIN/销售链接",
    "销售单价", "是否带电", "型号", "总数量", "单位", "每套个数",
    "投保单价(RMB)", "申报单价", "图片", "单箱重量(kg)", "单箱长",
    "单箱宽", "单箱高",
]
ITEM = [
    "FBA19TEST001", "REF1", "FBA19TEST001U00001", "SKU-1", "Pocket machine",
    "迷你标签打印机", "8443396000", "Phomemo", "塑料", "标签打印",
    "https://example.com/dp/B0", 0, "是", "M02", 4, "台", 4,
    2, 5.4, "", 10.35, 51, 33.5, 33.5,
]


def _make_book(path: Path, battery: str = "是") -> None:
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "Commercial Invoice"
    ws["A2"] = "FBA仓库代码*"
    ws["B2"] = "RFD2"
    for i, v in enumerate(HEAD[1:], start=2):
        if v is not None:
            ws.cell(i, 1, v)
    # 明细表头放在第 12 行，与真实清关发票一致
    for c, v in enumerate(DETAIL, start=1):
        ws.cell(12, c, v)
    for r, item in enumerate([ITEM, ITEM, ITEM], start=13):
        for c, v in enumerate(item, start=1):
            ws.cell(r, c, battery if DETAIL[c - 1] == "是否带电" else v)
    wb.save(path)


class TestAmoke(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name) / "清关发票_测试_海运-以星海派-RFD2.xlsx"
        _make_book(self.path)
        self.booking = {
            "FBA19TEST001": {
                "channel": "海运-以星海派",
                "warehouse": "RFD2",
                "tou": "XW",
                "boxes": 1,
                "tax": "是",
            }
        }

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def test_shipment_header(self) -> None:
        ships = parse_clearance(self.path, self.booking)
        self.assertEqual(len(ships), 1)
        ship = ships[0]
        self.assertEqual(ship.warehouse, "RFD2")
        self.assertEqual(ship.channel, "以星")          # 渠道归一（在 config.from 里）
        self.assertEqual(ship.total_boxes, 1)
        self.assertFalse(ship.empty_as_slash)          # 爱墨成品空列留空
        self.assertEqual(ship.customs, "不需报关")      # 包税 → 不需报关

    def test_one_row_per_item(self) -> None:
        ship = parse_clearance(self.path, self.booking)[0]
        self.assertEqual(len(ship.cartons), 3)         # 3 个 SKU 行
        first = ship.cartons[0]
        self.assertEqual(first.fba_id, "FBA19TEST001")
        self.assertEqual(first.name_cn, "迷你标签打印机")
        self.assertEqual(first.hs_code, "8443396000")
        self.assertEqual(first.qty, 4)
        self.assertEqual(first.unit, 5.4)
        self.assertEqual(first.weight, 10.35)
        self.assertEqual((first.length, first.width, first.height), (51, 33.5, 33.5))
        self.assertEqual(first.battery, "是")

    def test_battery_is_file_level(self) -> None:
        """同一清关发票里只要有一行带电，整票就是「是」（人工成品口径）。"""
        _make_book(self.path, battery="否")
        self.booking["FBA19TEST001"]["battery"] = "否"
        ship = parse_clearance(self.path, self.booking)[0]
        self.assertEqual(ship.battery, "否")


if __name__ == "__main__":
    unittest.main()
