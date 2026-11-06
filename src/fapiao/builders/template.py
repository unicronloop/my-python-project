# -*- coding: utf-8 -*-
"""公司模板构建器：出货发票单据头 + 收件人 + 装箱单明细。"""
from __future__ import annotations

from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment

from ..common import pacing
from ..common.strings import rnd
from ..config import COMPANY_TEMPLATE, DETAIL_START_ROW, SHEET_INVOICE
from ..models.address import receiver_from_book
from ..models.shipment import Carton, Shipment
from ..models.totals import compute_totals

MAX_DETAIL_COL = 27  # A:AA
# 文本类明细列：源表空的时候统一写 "/"（PO 例外，写 "无"）
TEXT_COLS = (6, 18, 19, 20, 21, 22, 23, 26, 27)


def _clone_row_style(ws, src_row: int, dst_row: int, max_col: int = MAX_DETAIL_COL) -> None:
    for c in range(1, max_col + 1):
        ws.cell(dst_row, c)._style = copy(ws.cell(src_row, c)._style)


def _write_detail(ws, row: int, c: Carton, empty_as_slash: bool = False) -> None:
    # 单箱申报总值：模板惯例是公式 =O*N，只在算不出来时才写死数值
    if c.total is not None:
        total = c.total
    elif c.unit is not None and c.qty is not None:
        total = f"=O{row}*N{row}"
    else:
        total = None
    values = {
        1: int(c.seq) if str(c.seq).isdigit() else c.seq, 2: c.boxes, 3: c.battery, 4: c.magnet, 5: c.fba_id, 6: c.po,
        7: c.name_cn, 8: c.name_en, 9: c.hs_code, 10: c.weight,
        11: c.length, 12: c.width, 13: c.height, 14: c.qty, 15: c.unit,
        16: c.currency, 17: total, 18: c.brand, 19: c.model,
        20: c.mat_cn, 21: c.mat_en, 22: c.use_cn, 23: c.use_en,
        24: c.net_g, 25: c.gross_g, 26: c.image, 27: c.link,
    }
    # 空值约定（与用户确认过）：文本列空 → "/"，PO 空 → "无"
    if empty_as_slash:
        for col in TEXT_COLS:
            if values.get(col) in (None, ""):
                values[col] = "无" if col == 6 else "/"
    for col, val in values.items():
        ws.cell(row, col).value = val if val != "" else None
    for col in (7, 8, 26):
        ws.cell(row, col).alignment = Alignment(vertical="center", wrap_text=True)


def build_workbook(ship: Shipment, out_path: Path,
                   template_path: Path = COMPANY_TEMPLATE) -> Path:
    """把 Shipment 写进公司模板，返回输出路径。"""
    # 头/箱数等用汇总值，保证与明细一致（明细行数 ≠ 箱数时以 boxes 为准）
    t = compute_totals(ship)
    if not ship.total_boxes:
        ship.total_boxes = int(t.boxes)
    # B12 预报实重：金标一律留空，重量只进校验报告
    # 注意：total_cbm / total_value 不用明细反推回填——旧版只在数据源显式给出时
    # 才写 B13/B14（创颖峻取汇总"方数"），德邦等其余客户保持空白，以维持输出一致。

    wb = openpyxl.load_workbook(str(template_path))
    ws = wb[SHEET_INVOICE]
    if ship.invoice_sheet and ship.invoice_sheet != SHEET_INVOICE:
        # 个别客户的上传系统按 sheet 名取数（爱墨科技 = T6系统发票）
        ws.title = ship.invoice_sheet

    # ---- 单据头 ----
    ws.cell(4, 2).value = ship.ref or None
    ws.cell(6, 2).value = ship.channel or None
    ws.cell(7, 2).value = ship.warehouse or None
    ws.cell(8, 2).value = ship.battery
    ws.cell(9, 2).value = ship.magnet
    ws.cell(10, 2).value = ship.customs
    ws.cell(11, 2).value = ship.total_boxes
    ws.cell(12, 2).value = rnd(ship.total_weight) if ship.total_weight else None
    ws.cell(13, 2).value = rnd(ship.total_cbm) if ship.total_cbm else None
    ws.cell(14, 2).value = rnd(ship.total_value) if ship.total_value else None
    ws.cell(15, 2).value = ship.remarks or None
    if ship.delivery_window is not None:
        ws.cell(14, 7).value = ship.delivery_window

    # ---- 收件人（固定值覆盖模板 VLOOKUP，避免依赖外部链接）----
    rc = ship.receiver or receiver_from_book(ship.warehouse) or {}
    ws["G4"] = rc.get("company") or None
    ws["G5"] = rc.get("name") or None
    ws["G6"] = rc.get("addr1") or None
    ws["G7"] = rc.get("addr2") or None
    ws["G8"] = rc.get("state") or None
    ws["G9"] = rc.get("city") or None
    zip_v = rc.get("zip") or None
    if zip_v and str(zip_v).isdigit():
        zip_v = int(zip_v)
    ws["G10"] = zip_v
    ws["G11"] = rc.get("region") or None
    ws["G12"] = rc.get("phone") or None
    ws["G13"] = rc.get("email") or None

    # ---- 装箱单 ----
    for i, c in enumerate(ship.cartons):
        row = DETAIL_START_ROW + i
        if i > 0:
            _clone_row_style(ws, DETAIL_START_ROW, row)
        _write_detail(ws, row, c, empty_as_slash=ship.empty_as_slash)

    # 清掉模板里可能残留的多余数据行（末行号 = last-1）
    last = DETAIL_START_ROW + len(ship.cartons)
    if ws.max_row > last:
        for r in range(last, ws.max_row + 1):
            for col in range(1, MAX_DETAIL_COL):
                ws.cell(r, col).value = None

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out_path))
    wb.close()
    pacing.after_file()  # 每生成一个文件静默停顿（秒数见 config.MIN_SECONDS_PER_FILE）
    return out_path
