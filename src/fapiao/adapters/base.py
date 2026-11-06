# -*- coding: utf-8 -*-
"""适配器公共落地逻辑：Shipment → 公司模板 + 校验 + 报告条目。

插件协议与数据类型在 `fapiao.plugin`；这里只放需要 openpyxl 的落地代码。
"""
from __future__ import annotations

from pathlib import Path

from ..builders.template import build_workbook
from ..models.shipment import Shipment
from ..models.totals import summarize
from ..plugin import CustomerPlugin, Generated, RunContext
from ..validate import validate

__all__ = ["CustomerPlugin", "Generated", "RunContext", "emit_shipment", "report_line"]


def emit_shipment(ship: Shipment, ctx: RunContext) -> Generated:
    """Canonical Shipment → 公司模板；返回报告条目（失败时 problems 带原因）。"""
    if not ship.doc_no:
        ship.doc_no = ship.cartons[0].fba_id if ship.cartons else ""

    out_path = ctx.out_root / ship.customer / f"{ship.out_stem}.xlsx"
    problems = validate(ship)
    try:
        if not ctx.dry_run:
            build_workbook(ship, out_path)
    except Exception as exc:  # noqa: BLE001  单个文件失败不影响整体
        ctx.log.error("  [失败] 生成 %s: %s", out_path.name, exc)
        return Generated(file=f"{ship.customer}/{out_path.name}",
                         path=out_path, problems=[str(exc)],
                         warnings=list(ship.warnings))

    stat = summarize(ship)
    return Generated(file=f"{ship.customer}/{out_path.name}", path=out_path,
                     stat=stat, problems=problems, warnings=list(ship.warnings))


def report_line(item: Generated) -> str:
    flag = "OK " if item.ok else "!! "
    st = item.stat
    if st:
        return (f"{flag}{item.file}  行={st['rows']} 箱={st['boxes']} 件={st['qty']:g} "
                f"实重={st['weight']} 方数={st['cbm']} 申报值={st['value']}")
    return f"{flag}{item.file}  已生成"
