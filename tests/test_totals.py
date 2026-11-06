# -*- coding: utf-8 -*-
"""汇总计算单测：Carton.weight 语义（单箱实重）+ 合箱票。"""
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from fapiao.models.shipment import Carton, Shipment  # noqa: E402
from fapiao.models.totals import compute_totals, summarize  # noqa: E402
from fapiao.validate import validate  # noqa: E402


def ship_with(*cartons) -> Shipment:
    return Shipment(customer="测试", ref="R1", channel="美森特快线",
                    warehouse="HIA1", cartons=list(cartons))


class TestTotals(unittest.TestCase):
    def test_weight_is_per_box(self):
        ship = ship_with(
            Carton(seq="1", boxes=1, weight=16.45, length=57, width=44, height=42,
                   qty=160, unit=2.0, fba_id="FBA1"),
            Carton(seq="2", boxes=1, weight=10.0, length=50, width=40, height=30,
                   qty=20, unit=1.0, fba_id="FBA1"),
        )
        t = compute_totals(ship)
        self.assertAlmostEqual(t.weight, 26.45)
        self.assertEqual(t.boxes, 2)
        self.assertAlmostEqual(t.value, 340.0)

    def test_merged_boxes_multiplies_weight(self):
        # 德邦发票2：一行 3 箱，单箱实重 6，总重 18
        ship = ship_with(Carton(seq="1-3", boxes=3, weight=6.0, length=36, width=58,
                                height=32, qty=6, unit=0.2, fba_id="FBA2"))
        t = compute_totals(ship)
        self.assertEqual(t.boxes, 3)
        self.assertAlmostEqual(t.weight, 18.0)
        self.assertAlmostEqual(t.cbm, 36 * 58 * 32 / 1e6 * 3, places=6)

    def test_summarize_matches_compute(self):
        ship = ship_with(Carton(seq="1", boxes=2, weight=3.5, length=10, width=10,
                                height=10, qty=4, unit=5.0, fba_id="FBA3"))
        t = compute_totals(ship)
        st = summarize(ship)
        self.assertEqual(st["rows"], 1)
        self.assertEqual(st["boxes"], 2)
        self.assertAlmostEqual(st["weight"], t.weight)
        self.assertAlmostEqual(st["value"], t.value)

    def test_validate_flags_missing_fields(self):
        ship = ship_with(Carton(seq="1", boxes=1, weight=None, qty=None))
        problems = validate(ship)
        self.assertTrue(any("缺 FBA ID" in p for p in problems))
        self.assertTrue(any("缺单箱实重" in p for p in problems))
        self.assertTrue(any("缺单箱数量" in p for p in problems))

    def test_validate_compares_with_expect(self):
        ship = ship_with(Carton(seq="1", boxes=1, weight=10, qty=2, fba_id="FBA4"))
        ship.expect = {"boxes": 2, "qty": 3}
        problems = validate(ship)
        self.assertTrue(any("箱数与客户汇总不符" in p for p in problems))
        self.assertTrue(any("件数与客户汇总不符" in p for p in problems))


if __name__ == "__main__":
    unittest.main()
