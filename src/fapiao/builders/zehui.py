# -*- coding: utf-8 -*-
"""泽汇构建器：把客户"发票明细"整理成上传系统要求的版式。

流程（与 1.py 完全一致，只是搬进 builder）：
1) 先读收件地址 / 追踪号（合并单元格左上角）
2) 解除合并（保留样式）→ 删除 5 个多余列
3) A 列向下填充、箱号重编号、G 列 = F 列
4) 删除"总数量""是否正式报关"两行，插入"国家*"行
5) 重排收件人 / VAT 区，重建表头与箱规合并单元格
6) 重新嵌入产品图片（S 列）
"""
from __future__ import annotations

import logging
import re
from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor

from ..common.excel import remove_columns_by_header, unmerge_all
from ..common.images import RawImage, dedupe_media, read_drawing_images
from ..common.region import get_us_region_by_zip, parse_zehui_address
from ..common.strings import is_blank, normalize_label
from ..common import pacing
from ..config import (
    ZEHUI_DECLARATION_RULES,
    ZEHUI_DEFAULT_DECLARATION,
    ZEHUI_PICKUP_DIR,
    ZEHUI_UNWANTED_HEADERS,
)

log = logging.getLogger(__name__)

ZEHUI_DECLARATION_LOOKUP = {
    normalize_label(alias): output
    for output, aliases in ZEHUI_DECLARATION_RULES.items()
    for alias in aliases
}


def load_zehui_pickup_map() -> dict:
    """读"泽汇-提货表"：出库单号 → {物流商备注/区域/仓库代码/配送地址/是否正式报关}。"""
    result: dict[str, dict] = {}
    if not ZEHUI_PICKUP_DIR.exists():
        return result

    for path in sorted(ZEHUI_PICKUP_DIR.glob("*.xlsx")):
        if path.name.startswith("~$"):
            continue
        try:
            wb = openpyxl.load_workbook(str(path), data_only=True)
        except Exception as e:  # noqa: BLE001
            log.warning("[泽汇 提货表] 读取失败：%s: %s", path.name, e)
            continue

        try:
            ws = wb.active
            headers = {
                normalize_label(ws.cell(1, c).value): c
                for c in range(1, ws.max_column + 1)
                if ws.cell(1, c).value not in (None, "")
            }

            def col(name):
                return headers.get(normalize_label(name))

            out_col = col("出库单号")
            if not out_col:
                continue

            for r in range(2, ws.max_row + 1):
                out_no = ws.cell(r, out_col).value
                if is_blank(out_no):
                    continue
                record = {}
                for name in ("物流商备注", "区域", "仓库代码", "配送地址", "是否正式报关"):
                    c = col(name)
                    if c:
                        v = ws.cell(r, c).value
                        if not is_blank(v):
                            record[name] = str(v).strip()
                result[str(out_no).strip()] = record
        finally:
            wb.close()
    return result


def build_zehui(input_path: Path, output_path: Path, pickup_map=None) -> str:
    """生成泽汇上传表，返回客户参考号（用于文件名）。"""
    src = Path(input_path)
    wb = openpyxl.load_workbook(str(src))
    ws = wb.active
    # 某些环境装了 Pillow 时 openpyxl 会把源图片读进来，这里统一清空后自行重嵌
    ws._images = []

    header_row = None
    for r in range(1, min(ws.max_row, 80) + 1):
        if normalize_label(ws.cell(r, 1).value) == normalize_label("外部箱号"):
            header_row = r
            break
    if header_row is None:
        raise ValueError(f"[泽汇] 未找到'外部箱号'表头行：{src.name}")

    source_data_start = header_row + 1

    customer_ref = ws.cell(3, 2).value
    pickup = {}
    if pickup_map is not None and not is_blank(customer_ref):
        pickup = pickup_map.get(str(customer_ref).strip(), {})
        if pickup_map and not pickup:
            log.info("[泽汇 提示] %s: 客户参考号1=%s 未在提货表找到，收货渠道保留原值",
                     src.name, customer_ref)

    # ---- 在改动前读取关键值 / 样式 ----
    rec_vals = [ws.cell(r, 9).value for r in range(3, 12)]      # I 列收件人值
    rec_styles = [copy(ws.cell(r, 9)._style) for r in range(3, 12)]
    vat_vals = [ws.cell(r, 16).value for r in range(3, 12)]     # P 列 VAT 值
    vat_styles = [copy(ws.cell(r, 16)._style) for r in range(3, 12)]

    def b_value_by_label(label):
        for r in range(1, header_row + 1):
            if normalize_label(ws.cell(r, 1).value) == normalize_label(label):
                return ws.cell(r, 2).value
        return None

    total_boxes = b_value_by_label("总箱数")

    other_remark_style = None
    for r in range(1, header_row + 1):
        if normalize_label(ws.cell(r, 7).value) == normalize_label("其他备注"):
            other_remark_style = copy(ws.cell(r, 7)._style)
            break

    def value_right_of_label(label):
        target = normalize_label(label)
        for r in range(1, 16):
            for c in range(1, ws.max_column + 1):
                if normalize_label(ws.cell(r, c).value) != target:
                    continue
                end_c = c
                for rng in ws.merged_cells.ranges:
                    if rng.min_row <= r <= rng.max_row and rng.min_col <= c <= rng.max_col:
                        end_c = rng.max_col
                for k in range(end_c + 1, ws.max_column + 1):
                    if not is_blank(ws.cell(r, k).value):
                        return str(ws.cell(r, k).value).strip()
        return None

    addr1 = value_right_of_label("收件人地址1") or pickup.get("配送地址")
    tracking = value_right_of_label("货件追踪号")

    # ---- 解析报关类型（必须在 unmerge / 删行之前读取） ----
    raw_declaration = ""
    v = b_value_by_label("报关类型")           # 标签同一行的 B 列
    if not is_blank(v):
        raw_declaration = str(v).strip()

    decl_key = normalize_label(raw_declaration) if raw_declaration else ""

    # 标签没读到、或读到的取值不认识时，退回全表扫描（前 15 行找能认的写法）
    if decl_key not in ZEHUI_DECLARATION_LOOKUP:
        for r in range(1, min(16, ws.max_row + 1)):
            for c in range(1, ws.max_column + 1):
                cell_val = ws.cell(r, c).value
                if is_blank(cell_val):
                    continue
                cell_key = normalize_label(cell_val)
                matched = next((k for k in ZEHUI_DECLARATION_LOOKUP if k in cell_key), None)
                if matched:
                    decl_key = matched
                    raw_declaration = str(cell_val).strip()
                    break
            if decl_key in ZEHUI_DECLARATION_LOOKUP:
                break

    declaration_type = ZEHUI_DECLARATION_LOOKUP.get(decl_key, ZEHUI_DEFAULT_DECLARATION)
    if raw_declaration and decl_key not in ZEHUI_DECLARATION_LOOKUP:
        log.warning("[泽汇 提示] %s: 未识别的报关类型“%s”，已按默认值“%s”输出",
                    src.name, raw_declaration, ZEHUI_DEFAULT_DECLARATION)

    parsed = parse_zehui_address(addr1)
    wh_code = pickup.get("仓库代码") or parsed["仓码"]
    zip_code = parsed["邮编"]
    region = pickup.get("区域") or get_us_region_by_zip(zip_code)
    channel = pickup.get("物流商备注") or ws.cell(7, 2).value

    # ---- 解除合并（保留原样式） ----
    unmerge_all(ws, fill_merged=False)

    # ---- 删除多余列 ----
    remove_columns_by_header(ws, header_row, ZEHUI_UNWANTED_HEADERS)

    # ---- A 列填充 / 箱号重编号 / G = F ----
    last_box = None
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if not is_blank(v):
            last_box = v
        elif last_box is not None:
            ws.cell(r, 1).value = last_box

    boxes = []
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if not is_blank(v) and v not in boxes:
            boxes.append(v)

    def natural_key(x):
        return [int(t) if t.isdigit() else str(t).lower()
                for t in re.split(r"(\d+)", str(x))]

    box_map = {orig: idx + 1 for idx, orig in enumerate(sorted(boxes, key=natural_key))}
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v in box_map:
            ws.cell(r, 1).value = box_map[v]

    for r in range(header_row + 1, ws.max_row + 1):
        if is_blank(ws.cell(r, 2).value):
            continue
        ws.cell(r, 7).value = ws.cell(r, 6).value

    # ---- 行调整：去掉"总数量""是否正式报关"，补"国家*"行 ----
    def delete_row_by_label(label):
        target = normalize_label(label)
        for r in range(1, min(header_row, 20) + 1):
            if normalize_label(ws.cell(r, 1).value) == target:
                ws.delete_rows(r, 1)
                return True
        return False

    deleted = 0
    for label in ("总数量", "是否正式报关"):
        if delete_row_by_label(label):
            deleted += 1
    header_row -= deleted

    ws.insert_rows(header_row, 1)
    for c in range(1, 19):
        ws.cell(header_row, c)._style = copy(ws.cell(header_row - 1, c)._style)
    ws.cell(header_row, 1).value = "国家*:"
    ws.row_dimensions[header_row].height = 25

    header_row += 1
    data_start = header_row + 1

    # ---- 重排表头 ----
    for i, r in enumerate(range(3, 12)):
        h = ws.cell(r, 8)
        h.value = rec_vals[i]
        h._style = copy(rec_styles[i])
        al = copy(h.alignment)
        al.horizontal = "center"
        h.alignment = al
        ws.cell(r, 9).value = None

        # 货件追踪号单独写到 B6，不放进 VAT 值区
        if normalize_label(ws.cell(r, 12).value) != normalize_label("货件追踪号"):
            o = ws.cell(r, 15)
            o.value = vat_vals[i]
            o._style = copy(vat_styles[i])
            al = copy(o.alignment)
            al.horizontal = "center"
            o.alignment = al
        else:
            ws.cell(r, 15).value = None

    # "总数量"行删掉后，其他备注标签需要补回 G12
    g12 = ws.cell(12, 7)
    g12.value = "其他备注:"
    if other_remark_style is not None:
        g12._style = other_remark_style
    h12 = ws.cell(12, 8)
    h12.value = ""
    h12._style = copy(g12._style)

    ws.cell(3, 8).value = wh_code
    ws.cell(4, 8).value = wh_code
    ws.cell(5, 8).value = addr1 or ""
    ws.cell(6, 8).value = ""
    ws.cell(7, 8).value = parsed["省州"]
    ws.cell(8, 8).value = parsed["城市"]
    if str(zip_code).isdigit() and not str(zip_code).startswith("0"):
        ws.cell(9, 8).value = int(zip_code)
    else:
        ws.cell(9, 8).value = zip_code
    ws.cell(10, 8).value = region
    ws.cell(11, 8).value = ""

    ws.cell(7, 2).value = channel
    ws.cell(6, 2).value = tracking or ""
    ws.cell(8, 2).value = wh_code
    ws.cell(9, 2).value = None
    ws.cell(11, 2).value = declaration_type
    ws.cell(12, 2).value = total_boxes
    ws.cell(13, 2).value = None
    ws.cell(14, 2).value = None
    ws.cell(15, 2).value = region

    # ---- 重建表头合并单元格 ----
    ws.merge_cells("A1:R1")
    ws.merge_cells("A2:F2")
    ws.merge_cells("G2:R2")
    for r in range(3, 13):
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=11)
        ws.merge_cells(start_row=r, start_column=12, end_row=r, end_column=14)
        ws.merge_cells(start_row=r, start_column=15, end_row=r, end_column=18)
    for r in range(13, 16):
        if r != 15:  # 参考文件第 15 行 B 列不合并
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=8)
        ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=11)
        ws.merge_cells(start_row=r, start_column=12, end_row=r, end_column=14)
        ws.merge_cells(start_row=r, start_column=15, end_row=r, end_column=18)

    # ---- 按箱号分组：A 列视觉合并、G 列样式统一、J-N 合并 ----
    last_row = ws.max_row
    groups = []
    r = data_start
    while r <= last_row:
        v = ws.cell(r, 1).value
        end = r
        while end + 1 <= last_row and ws.cell(end + 1, 1).value == v:
            end += 1
        groups.append((r, end))
        r = end + 1

    for start, end in groups:
        src_a = ws.cell(start, 1)
        src_g = ws.cell(start, 7)
        for r in range(start, end + 1):
            ca = ws.cell(r, 1)
            ca.font = copy(src_a.font)
            ca.fill = copy(src_a.fill)
            ca.number_format = src_a.number_format
            al = copy(src_a.alignment)
            al.horizontal = None
            ca.alignment = al

            cg = ws.cell(r, 7)
            cg._style = copy(src_g._style)
            al = copy(cg.alignment)
            al.horizontal = "centerContinuous"
            cg.alignment = al

        if end > start:
            for c in range(10, 15):
                ws.merge_cells(start_row=start, start_column=c, end_row=end, end_column=c)

    # ---- 重新嵌入产品图片 ----
    row_offset = data_start - source_data_start
    for item in read_drawing_images(src):
        trow = item["row"] + row_offset
        if trow < data_start or trow > last_row:
            continue
        img = RawImage(item["data"], item["fmt"],
                       item["width"] or 72, item["height"] or 33)
        img.anchor = TwoCellAnchor(
            _from=AnchorMarker(col=18, row=trow - 1),
            to=AnchorMarker(col=19, row=trow),
        )
        ws.add_image(img)

    for r in range(1, last_row + 1):
        ws.row_dimensions[r].height = 40 if r == 1 else 25

    wb.save(str(output_path))
    pacing.after_file()  # 每生成一个文件静默停顿
    dedupe_media(output_path)

    return str(customer_ref).strip() if not is_blank(customer_ref) else src.stem
