# -*- coding: utf-8 -*-
"""勾稽校验：与客户汇总 / 合计行对账，不做逐格 diff。"""
from __future__ import annotations

from .models.shipment import Shipment
from .models.totals import compute_totals


def validate(ship: Shipment) -> list[str]:
    problems: list[str] = []
    t = compute_totals(ship)
    boxes, qtys, weights, cbm = t.boxes, t.qty, t.weight, t.cbm

    if ship.total_boxes and boxes != ship.total_boxes:
        problems.append(f"箱数不符：明细 {boxes} ≠ 表头 {ship.total_boxes}")

    exp = ship.expect or {}
    if exp.get("boxes") and boxes != exp["boxes"]:
        problems.append(f"箱数与客户汇总不符：{boxes} ≠ {exp['boxes']}")
    if exp.get("qty") and abs(qtys - exp["qty"]) > 0.01:
        problems.append(f"件数与客户汇总不符：{qtys:g} ≠ {exp['qty']:g}")
    if exp.get("weight") and abs(weights - exp["weight"]) > 0.05:
        problems.append(f"实重与客户汇总不符：{weights:.2f} ≠ {exp['weight']:.2f}")
    if exp.get("cbm") and abs(cbm - exp["cbm"]) > 0.05:
        problems.append(f"方数与客户汇总不符：{cbm:.2f} ≠ {exp['cbm']:.2f}")
    if exp.get("value") and abs(t.value - exp["value"]) > 0.05:
        problems.append(f"申报值与客户汇总不符：{t.value:.2f} ≠ {exp['value']:.2f}")

    for c in ship.cartons:
        if not c.fba_id:
            problems.append(f"第 {c.seq} 箱缺 FBA ID")
        if c.weight is None:
            problems.append(f"第 {c.seq} 箱缺单箱实重")
        if c.qty is None:
            problems.append(f"第 {c.seq} 箱缺单箱数量")
    return problems
