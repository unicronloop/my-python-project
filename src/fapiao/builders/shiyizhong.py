# -*- coding: utf-8 -*-
"""时易中构建器：源发票 → 公司模板（Sheet1 明细 + 出货发票公式行 + Z 列嵌图）。"""
from __future__ import annotations

import logging
import re
from copy import copy
from pathlib import Path

import openpyxl

from ..common.excel import force_set, source_row_heights, unmerge_all
from ..common.images import dedupe_media, insert_shiyizhong_images
from ..common.strings import (
    extract_number,
    find_value_by_label,
    is_blank,
    judge_battery_magnetic,
)
from ..common import pacing
from ..config import DETAIL_ROW_HEIGHT, INVOICE_DETAIL_START_ROW

log = logging.getLogger(__name__)


def _fill_invoice_detail_rows(ws, data_count, detail_start, is_battery, is_magnetic, poa_num):
    """填充"出货发票"明细行（第 18 行起）。Sheet1 数据从 Excel 第 13 行开始。"""
    for i in range(data_count):
        invoice_row = INVOICE_DETAIL_START_ROW + i
        # df 第 detail_start 行（0 基）对应 Excel 第 detail_start+1 行
        sheet1_row = detail_start + 1 + i

        o_val = extract_number(ws.parent["Sheet1"].cell(sheet1_row, 9).value)

        force_set(ws, invoice_row, 1, f"=Sheet1!O{sheet1_row}")
        force_set(ws, invoice_row, 2, 1)
        force_set(ws, invoice_row, 3, is_battery)
        force_set(ws, invoice_row, 4, is_magnetic)
        force_set(ws, invoice_row, 5, f"=Sheet1!U{sheet1_row}")
        force_set(ws, invoice_row, 6, poa_num)
        force_set(ws, invoice_row, 7, f"=Sheet1!A{sheet1_row}")
        force_set(ws, invoice_row, 8, f"=Sheet1!B{sheet1_row}")
        force_set(ws, invoice_row, 9, f"=Sheet1!C{sheet1_row}")
        force_set(ws, invoice_row, 10, f"=Sheet1!K{sheet1_row}")
        force_set(ws, invoice_row, 11, f"=Sheet1!L{sheet1_row}")
        force_set(ws, invoice_row, 12, f"=Sheet1!M{sheet1_row}")
        force_set(ws, invoice_row, 13, f"=Sheet1!N{sheet1_row}")
        force_set(ws, invoice_row, 14, f"=Sheet1!H{sheet1_row}")
        force_set(ws, invoice_row, 15, o_val)
        force_set(ws, invoice_row, 16, "USD")
        force_set(ws, invoice_row, 17, f"=O{invoice_row}*N{invoice_row}")
        force_set(ws, invoice_row, 18, f"=Sheet1!D{sheet1_row}")
        force_set(ws, invoice_row, 19, f"=Sheet1!F{sheet1_row}")
        force_set(ws, invoice_row, 20, "/")
        force_set(ws, invoice_row, 21, f"=Sheet1!E{sheet1_row}")
        force_set(ws, invoice_row, 22, "/")
        force_set(ws, invoice_row, 23, f"=Sheet1!G{sheet1_row}")
        force_set(ws, invoice_row, 26, "/")
        force_set(ws, invoice_row, 27, f"=Sheet1!S{sheet1_row}")


def _format_invoice_detail_rows(ws, data_count):
    """把模板第 18 行的格式复制到每一条明细行，避免第 19 行以后没有边框/对齐。"""
    if data_count <= 0:
        return
    last_row = INVOICE_DETAIL_START_ROW + data_count - 1
    for r in range(INVOICE_DETAIL_START_ROW, last_row + 1):
        for c in range(1, 28):  # A:AA
            ws.cell(r, c)._style = copy(ws.cell(INVOICE_DETAIL_START_ROW, c)._style)
        ws.row_dimensions[r].height = DETAIL_ROW_HEIGHT


def channel_from_filename(path: Path, default: str = "") -> str:
    """时易中文件名：FBA号_国家_渠道_退税.xls，去掉 KG/CBM 并把"按方包税"改为"含税"。"""
    parts = Path(path).stem.split("_")
    if len(parts) < 3:
        return default
    channel = parts[2].strip()

    for suffix in ("CBM", "KG"):
        if channel.upper().endswith(suffix):
            channel = channel[:-len(suffix)].strip()

    channel = channel.replace("按方包税", "含税")
    return channel or default


def _total_box_count(df, detail_start: int, fallback: int) -> int:
    """从源发票"箱号"列取最大箱号（混箱时不能用明细行数当箱数）。"""
    max_box = 0
    if 14 >= len(df.columns):
        return fallback
    for i in range(detail_start, len(df)):
        v = df.iloc[i, 14]
        if is_blank(v):
            continue
        nums = re.findall(r"\d+", str(v))
        if not nums:
            continue
        try:
            max_box = max(max_box, max(int(n) for n in nums))
        except ValueError:
            continue
    return max_box or fallback


def build_shiyizhong(input_path: Path, output_path: Path,
                     template_path: Path, config: dict) -> str | bool:
    """生成时易中发票；不是发票明细时返回 False。"""
    import pandas as pd

    try:
        df = pd.read_excel(input_path, sheet_name=0, header=None)
    except ImportError as e:  # 缺 xlrd
        raise RuntimeError(
            f"读取 {input_path.name} 失败：{e}\n"
            f"请确认已在当前 .venv 中执行：pip install \"xlrd>=2.0.1\""
        ) from e

    # 头部字段
    receiver_company = find_value_by_label(df, "COMPANY(收件公司)", prefer_col=8)
    amazon_ref = find_value_by_label(df, "Amazon reference ID", prefer_col=8)
    attr_note = find_value_by_label(df, "属性备注", prefer_col=8)
    remark = find_value_by_label(df, "备注")

    ref_parts = [p.strip() for p in str(amazon_ref).split(",") if p.strip()]
    fba_id = ref_parts[0] if ref_parts else ""
    poa_num = ref_parts[1] if len(ref_parts) > 1 else ""

    is_battery, is_magnetic = judge_battery_magnetic(attr_note, remark)

    # 找数据起始行（"中文品名"下一行）
    detail_start = None
    for i in range(len(df)):
        if not is_blank(df.iloc[i, 0]) and str(df.iloc[i, 0]).strip() == "中文品名":
            detail_start = i + 1
            break
    if detail_start is None:
        return False

    detail = df.iloc[detail_start:].copy()
    detail = detail[detail[0].notna() & (detail[0].astype(str).str.strip() != "")]
    data_count = len(detail)
    total_boxes = _total_box_count(df, detail_start, data_count)

    if not Path(template_path).exists():
        raise FileNotFoundError(f"公司模板不存在：{template_path}")
    wb = openpyxl.load_workbook(str(template_path))

    # ---- 写 Sheet1（破除合并、保留模板公式） ----
    if "Sheet1" in wb.sheetnames:
        sheet1 = wb["Sheet1"]
        unmerge_all(sheet1, fill_merged=False)
        for row in sheet1.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    continue
                cell.value = None
    else:
        sheet1 = wb.create_sheet("Sheet1")

    for i in range(len(df)):
        sheet_row = i + 1
        for j in range(len(df.columns)):
            val = df.iloc[i, j]
            if not is_blank(val):
                force_set(sheet1, sheet_row, j + 1, val)

    # 保留源文件行高，避免 Sheet1 里的长文本被压扁
    for r, h in source_row_heights(input_path).items():
        if h and r <= sheet1.max_row:
            sheet1.row_dimensions[r].height = h

    # ---- 写"出货发票"头部 B4-B11 ----
    ws = wb["出货发票"]
    force_set(ws, 4, 2, fba_id)
    force_set(ws, 5, 2, "")
    # 收货渠道解析为空时把 B6 清空（None），而不是写空字符串
    channel = channel_from_filename(input_path, config.get("收货渠道", ""))
    force_set(ws, 6, 2, channel if channel else None)
    force_set(ws, 7, 2, receiver_company)
    force_set(ws, 8, 2, is_battery)
    force_set(ws, 9, 2, is_magnetic)
    force_set(ws, 10, 2, config.get("报关类型", ""))
    force_set(ws, 11, 2, total_boxes)

    # ---- 写"出货发票"明细行 ----
    _fill_invoice_detail_rows(ws, data_count, detail_start, is_battery, is_magnetic, poa_num)
    _format_invoice_detail_rows(ws, data_count)
    used_images = insert_shiyizhong_images(ws, input_path, detail_start, data_count)
    if used_images < data_count:
        log.info("[时易中 提示] %s：%d 行源文件没有产品图片，Z 列保留 /",
                 input_path.name, data_count - used_images)

    wb.save(str(output_path))
    pacing.after_file()  # 每生成一个文件静默停顿
    if used_images:
        dedupe_media(output_path)
    return fba_id or input_path.stem
