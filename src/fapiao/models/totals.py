# -*- coding: utf-8 -*-
"""汇总计算唯一实现：validate / 报告 / builder 都用它，避免各处口径漂移。"""
from __future__ import annotations

from dataclasses import dataclass

from .shipment import Shipment


@dataclass
class Totals:
    rows: int = 0
    boxes: float = 0
    qty: float = 0
    weight: float = 0
    cbm: float = 0
    value: float = 0


def compute_totals(ship: Shipment) -> Totals:
    """按"单箱实重 × 箱数"统一汇总。

    Carton.weight 的语义锁定为单箱实重，因此这里必须乘 boxes；
    否则合箱票（boxes > 1）会少算重量。
    """
    t = Totals(rows=len(ship.cartons))
    box_of_seq: dict[str, float] = {}
    for c in ship.cartons:
        b = c.boxes or 1
        # 同一箱的多条货品共用同一个箱序号，箱数只能算一次
        key = str(c.seq) or f"row{id(c)}"
        box_of_seq.setdefault(key, b)
        t.qty += c.qty or 0
        t.weight += (c.weight or 0) * b
        t.cbm += ((c.length or 0) * (c.width or 0) * (c.height or 0) / 1e6) * b
        t.value += c.total_value
    t.boxes = sum(box_of_seq.values())
    return t


def summarize(ship: Shipment) -> dict:
    """报告用的一行摘要（与 compute_totals 同源，避免两套口径）。"""
    t = compute_totals(ship)
    return {"rows": t.rows, "boxes": t.boxes, "qty": t.qty,
            "weight": round(t.weight, 2), "cbm": round(t.cbm, 4),
            "value": round(t.value, 2)}
