# -*- coding: utf-8 -*-
"""收件人：公司模板地址库（懒加载）+ 地址文本解析。"""
from __future__ import annotations

import re
from pathlib import Path

import openpyxl

from ..common.region import region_by_state
from ..common.strings import s
from ..config import SHEET_ADDR, TEMPLATE_PATH


class AddressBook:
    """公司模板 FBA仓码地址 + 客户自带地址库（按仓码合并）。"""

    def __init__(self, path: Path = TEMPLATE_PATH):
        self.rows: dict[str, dict] = {}
        path = Path(path)
        if not path.exists():
            return
        wb = openpyxl.load_workbook(path, data_only=True)
        try:
            if SHEET_ADDR in wb.sheetnames:
                ws = wb[SHEET_ADDR]
                for r in range(2, ws.max_row + 1):
                    code = s(ws.cell(r, 1).value).upper()
                    if not code:
                        continue
                    self.rows[code] = {
                        "addr1": s(ws.cell(r, 2).value),
                        "addr2": s(ws.cell(r, 3).value),
                        "city": s(ws.cell(r, 4).value),
                        "state": s(ws.cell(r, 5).value),
                        "zip": s(ws.cell(r, 6).value),
                        "region": s(ws.cell(r, 7).value),
                    }
        finally:
            wb.close()

    def get(self, code: str) -> dict | None:
        return self.rows.get(s(code).upper())

    def merge(self, code: str, info: dict) -> None:
        key = s(code).upper()
        if key and key not in self.rows:
            self.rows[key] = info


_addr_book: AddressBook | None = None


def get_addr_book() -> AddressBook:
    """懒加载：import 阶段不读模板（避免 import 副作用 + 首次运行地址库为空）。"""
    global _addr_book
    if _addr_book is None:
        _addr_book = AddressBook()
    return _addr_book


def reload_addr_book() -> AddressBook:
    global _addr_book
    _addr_book = AddressBook()
    return _addr_book


def build_receiver(warehouse: str, addr1: str, city: str, state: str, zip_code: str,
                   region: str = "", company: str = "", name: str = "") -> dict:
    """统一收件人格式：公司/名称 = 仓码；地址1 = '仓码 - 地址 邮编 - 城市, 州 - United States'"""
    zip_code = s(zip_code)
    city = s(city)
    state = s(state)
    region = s(region) or region_by_state(state)
    if addr1:
        # 客户/地址库里的地址可能已经是完整串，避免重复拼接
        composed = re.search(r"\d{5}.*-\s*[A-Z]{2}.*United States", addr1)
        line = f"{warehouse} - {addr1}" if composed else (
            f"{warehouse} - {addr1} {zip_code} - {city}, {state} - United States")
        line = re.sub(r"\s+", " ", line)
    else:
        line = ""
    return {
        "company": s(company) or warehouse,
        "name": s(name) or warehouse,
        "addr1": line,
        "addr2": "",
        "city": city,
        "state": state,
        "zip": zip_code,
        "region": region,
    }


def receiver_from_book(warehouse: str) -> dict | None:
    info = get_addr_book().get(warehouse)
    if not info:
        return None
    return build_receiver(warehouse, info["addr1"], info["city"], info["state"],
                          info["zip"], info["region"])


def parse_address_block(raw: str, warehouse: str = "") -> dict:
    """从多行地址文本解析收件人（创颖峻 汇总.地址 / 德邦 箱单发票 L5）。"""
    lines = [s(x) for x in s(raw).splitlines() if s(x)]
    if not lines:
        return build_receiver(warehouse, "", "", "", "")
    if re.search(r"美国|United States|US$", lines[-1], re.I):
        lines = lines[:-1]
    if lines and re.fullmatch(r"[A-Z]{2}", lines[-1]):
        lines = lines[:-1]
    if warehouse:
        lines = [ln for ln in lines if ln.upper() != warehouse.upper()]

    city, state, zip_code = "", "", ""
    addr_parts = []
    for ln in lines:
        m = re.match(r"^(.*?)[,，]\s*([A-Za-z]{2})\s+([\d\-]{3,})$", ln)
        if m:
            city, state, zip_code = m.group(1).strip(), m.group(2).upper(), m.group(3)
            continue
        addr_parts.append(ln)

    # 第一行若是公司名（非纯地址），丢弃
    if addr_parts and re.search(r"(LLC|Inc\.?|Ltd|Co\.|有限公司|Corp)", addr_parts[0], re.I):
        addr_parts = addr_parts[1:]
    addr1 = " ".join(addr_parts)
    if not city and addr_parts:
        # 退化：取最后一行拆城市
        tail = addr_parts[-1]
        m = re.search(r"([A-Za-z .]+),?\s*([A-Z]{2})\s*(\d{5}(?:-\d{4})?)?", tail)
        if m:
            city, state, zip_code = m.group(1).strip(), m.group(2), m.group(3) or ""
    return build_receiver(warehouse, addr1, city, state, zip_code)
