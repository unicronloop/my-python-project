# -*- coding: utf-8 -*-
"""openpyxl / xlrd 读写小工具。"""
from __future__ import annotations

import logging
import re
from contextlib import contextmanager
from pathlib import Path

import openpyxl
from openpyxl.cell.cell import MergedCell

from .strings import normalize_label

log = logging.getLogger(__name__)


def cell_at(ws, row: int, col: int):
    """取单元格（合并区回退到左上角）。"""
    cell = ws.cell(row, col)
    if not isinstance(cell, MergedCell):
        return cell
    return ws.cell(*merged_anchor(ws, row, col))


def merged_anchor(ws, row: int, col: int) -> tuple[int, int]:
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return rng.min_row, rng.min_col
    return row, col


def safe_read(ws, row: int, col: int):
    return cell_at(ws, row, col).value


def force_set(ws, row: int, col: int, value) -> None:
    """写值：合并单元格写到左上角，普通单元格直接写。

    不对 MergedCell 直接赋值（openpyxl 只读），也不手工构造 Cell 塞进 _cells
    （会丢样式、不更新 dimensions）。
    """
    cell = ws.cell(row, col)
    if isinstance(cell, MergedCell):
        r, c = merged_anchor(ws, row, col)
        ws.cell(r, c).value = value
        return
    cell.value = value


def unmerge_all(ws, fill_merged: bool = False) -> None:
    """解除 ws 所有合并单元格（保留每个单元格原有样式）。"""
    for rng in list(ws.merged_cells.ranges):
        min_r, max_r = rng.min_row, rng.max_row
        min_c, max_c = rng.min_col, rng.max_col
        top_left_val = ws.cell(min_r, min_c).value
        ws.unmerge_cells(str(rng))
        if fill_merged and top_left_val is not None:
            for r in range(min_r, max_r + 1):
                for c in range(min_c, max_c + 1):
                    ws.cell(r, c).value = top_left_val


@contextmanager
def open_book(path, data_only: bool = True):
    wb = openpyxl.load_workbook(str(path), data_only=data_only)
    try:
        yield wb
    finally:
        wb.close()


def load_workbook_any(path: Path, data_only: bool = False) -> openpyxl.Workbook:
    """统一读取 xlsx / xlsm / xls（xls 经 xlrd 转成 openpyxl 工作簿）。"""
    suffix = path.suffix.lower()
    if suffix in (".xlsx", ".xlsm"):
        return openpyxl.load_workbook(str(path), data_only=data_only)

    if suffix != ".xls":
        raise ValueError(f"不支持的文件后缀：{path.suffix}（{path.name}）")

    try:
        import xlrd
    except ImportError as e:
        raise RuntimeError(
            f"读取 {path.name} 失败：{e}\n请确认已安装：pip install \"xlrd>=2.0.1\""
        ) from e

    try:
        book = xlrd.open_workbook(str(path), formatting_info=True)
    except Exception:  # noqa: BLE001  有的 xls 不带格式化信息
        book = xlrd.open_workbook(str(path))

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for idx, sheet in enumerate(book.sheets()):
        ws = wb.create_sheet(title=_safe_sheet_title(sheet.name, idx))
        for r in range(sheet.nrows):
            for c in range(sheet.ncols):
                val = sheet.cell(r, c).value
                if val == "" or val is None:
                    continue
                if isinstance(val, float) and val.is_integer():
                    val = int(val)
                ws.cell(r + 1, c + 1).value = val

        try:
            for r1, r2, c1, c2 in sheet.merged_cells:
                ws.merge_cells(start_row=r1 + 1, start_column=c1 + 1,
                               end_row=r2, end_column=c2)
        except Exception:  # noqa: BLE001
            pass

    return wb


def _safe_sheet_title(name: str, idx: int) -> str:
    """xls 的 sheet 名可能带非法字符或超长，转成 openpyxl 能接受的名字。"""
    title = re.sub(r"[\[\]:*?/\\]", "_", str(name or ""))[:31].strip()
    return title or f"Sheet{idx + 1}"


def remove_columns_by_header(ws, header_row: int, headers_to_remove) -> int:
    """按表头名称匹配删除列（从后往前删，避免索引错乱）。"""
    headers_norm = {normalize_label(h) for h in headers_to_remove}
    cols = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if v is not None and normalize_label(v) in headers_norm:
            cols.append(c)
    for c in sorted(cols, reverse=True):
        ws.delete_cols(c, 1)
    return len(cols)


def source_row_heights(path: Path) -> dict[int, float]:
    """读取源文件的行高，尽量保留客户原始版面。"""
    heights: dict[int, float] = {}
    suffix = Path(path).suffix.lower()
    try:
        if suffix == ".xls":
            import xlrd
            book = xlrd.open_workbook(str(path), formatting_info=True)
            sheet = book.sheet_by_index(0)
            for r, info in sheet.rowinfo_map.items():
                if info.height:
                    heights[r + 1] = info.height / 20.0
        elif suffix in (".xlsx", ".xlsm"):
            wb = openpyxl.load_workbook(str(path), read_only=True)
            ws = wb.active
            for r, dim in ws.row_dimensions.items():
                if dim.height:
                    heights[r] = dim.height
            wb.close()
    except Exception:  # noqa: BLE001  行高只是版面优化
        log.debug("读取行高失败：%s", path, exc_info=True)
    return heights
