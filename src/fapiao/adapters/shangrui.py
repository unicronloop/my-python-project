# -*- coding: utf-8 -*-
"""尚睿适配器（客户：尚睿，2026-11 新增）。

源表 `06尚睿\\客户发票.xlsx`（单 sheet，一个货件一票、一箱一行）→ 公司模板。
表头在第 17 行附近，列顺序不固定，按表头文字取列。
"""
from __future__ import annotations

import logging
import re
from pathlib import Path

from ..common.excel import load_workbook_any, safe_read
from ..common.region import region_by_state
from ..common.strings import normalize_label, num, s
from ..config import CUSTOMS_MAP, SHANGRUI_DIR, normalize_channel
from ..models.shipment import Carton, Shipment
from ..registry import register
from .base import Generated, RunContext, emit_shipment

log = logging.getLogger(__name__)

SHANGRUI_DEFAULT_CHANNEL = "美西洛杉矶卡派含税"

HEADERS = {
    "seq": ("箱序号",),
    "boxes": ("件数",),
    "battery": ("是否带电",),
    "magnet": ("是否带磁",),
    "fba": ("物品FBA ID", "物品FBAID"),
    "po": ("物品POAnumber", "物品POA"),
    "name_cn": ("物品中文名",),
    "name_en": ("物品英文名",),
    "hs": ("海关编码",),
    "weight": ("单件实重(KG)", "单箱实重/KG", "单件实重", "单箱实重"),
    "length": ("长(CM)", "长/CM"),
    "width": ("宽(CM)", "宽/CM"),
    "height": ("高(CM)", "高/CM"),
    "qty": ("单件数量", "单箱数量"),
    "unit": ("申报单价",),
    "currency": ("申报币种",),
    "total": ("单件申报总值", "单箱申报总值"),
    "brand": ("品牌",),
    "model": ("型号",),
    "mat_cn": ("材质(中文)",),
    "mat_en": ("材质(英文)",),
    "use_cn": ("用途(中文)",),
    "use_en": ("用途(英文)",),
    "net": ("单个产品净重",),
    "image": ("产品图片",),
    "link": ("销售链接",),
}

HEAD_LABELS = {
    "ref": ("客户参考号1", "客户参考号"),
    "channel": ("收货渠道",),
    "warehouse": ("FBA仓码",),
    "battery": ("是否带电",),
    "magnet": ("是否带磁",),
    "customs": ("报关类型",),
    "boxes": ("总件数", "总箱数"),
    "weight": ("预报实重",),
    "cbm": ("预报方数",),
    "value": ("申报总值",),
    "company": ("收件人公司",),
    "name": ("收件人名称",),
    "addr1": ("收件人地址1",),
    "addr2": ("收件人地址2",),
    "state": ("收件人省州",),
    "city": ("收件人城市",),
    "zip": ("收件人邮编",),
    "country": ("目的地",),
    "phone": ("收件人电话",),
    "email": ("收件人邮箱",),
}


def _header_row(ws) -> int | None:
    for r in range(1, min(ws.max_row, 60) + 1):
        texts = {normalize_label(safe_read(ws, r, c))
                 for c in range(1, min(ws.max_column, 40) + 1)}
        if "物品中文名" in texts and "海关编码" in texts:
            return r
    return None


def _col_map(ws, row: int) -> dict[str, int]:
    cells = {normalize_label(safe_read(ws, row, c)): c
             for c in range(1, min(ws.max_column, 40) + 1)}
    out: dict[str, int] = {}
    for field, aliases in HEADERS.items():
        for alias in aliases:
            if normalize_label(alias) in cells:
                out[field] = cells[normalize_label(alias)]
                break
    return out


def _label_value(ws, aliases, max_row: int = 20):
    """标签右侧 1-2 格内的值（不能贪心往右扫：右边紧邻的可能又是下一个标签）。"""
    wanted = {normalize_label(a) for a in aliases}
    for r in range(1, min(ws.max_row, max_row) + 1):
        for c in range(1, min(ws.max_column, 20) + 1):
            if normalize_label(safe_read(ws, r, c)) in wanted:
                for cc in range(c + 1, c + 3):
                    val = safe_read(ws, r, cc)
                    if s(val):
                        return val
    return None


def _flag(raw, yes: str = "是") -> str:
    text = s(raw)
    if not text:
        return "否"
    return "是" if text.strip() in (yes, "带电", "带磁", "Y", "y") else "否"


def _region(channel: str, state: str) -> str:
    for key in ("美西", "美东", "美中"):
        if key in channel:
            return key
    return region_by_state(state)


def parse_shangrui(path: Path) -> Shipment | None:
    wb = load_workbook_any(path, data_only=True)
    try:
        ws = wb.worksheets[0]
        row0 = _header_row(ws)
        if not row0:
            return None
        cols = _col_map(ws, row0)

        channel = normalize_channel(s(_label_value(ws, HEAD_LABELS["channel"])))
        warehouse = s(_label_value(ws, HEAD_LABELS["warehouse"]))
        state = s(_label_value(ws, HEAD_LABELS["state"]))
        ref = s(_label_value(ws, HEAD_LABELS["ref"])) or path.stem

        ship = Shipment(
            source_file=path.name, customer="尚睿", template="尚睿",
            ref=ref, channel=channel or SHANGRUI_DEFAULT_CHANNEL,
            warehouse=warehouse,
            battery=_flag(_label_value(ws, HEAD_LABELS["battery"])),
            magnet=_flag(_label_value(ws, HEAD_LABELS["magnet"])),
            customs=CUSTOMS_MAP.get(
                normalize_label(s(_label_value(ws, HEAD_LABELS["customs"]))), "不需报关"),
        )
        ship.total_boxes = int(num(_label_value(ws, HEAD_LABELS["boxes"])) or 0)
        ship.total_weight = num(_label_value(ws, HEAD_LABELS["weight"]))
        ship.total_cbm = num(_label_value(ws, HEAD_LABELS["cbm"]))
        ship.total_value = num(_label_value(ws, HEAD_LABELS["value"]))
        ship.receiver = {
            "company": s(_label_value(ws, HEAD_LABELS["company"])),
            "name": s(_label_value(ws, HEAD_LABELS["name"])),
            "addr1": s(_label_value(ws, HEAD_LABELS["addr1"])),
            "addr2": s(_label_value(ws, HEAD_LABELS["addr2"])),
            "state": state,
            "city": s(_label_value(ws, HEAD_LABELS["city"])),
            "zip": s(_label_value(ws, HEAD_LABELS["zip"])),
            "country": s(_label_value(ws, HEAD_LABELS["country"])),
            "phone": s(_label_value(ws, HEAD_LABELS["phone"])),
            "email": s(_label_value(ws, HEAD_LABELS["email"])),
            "region": _region(channel, state),
        }

        for r in range(row0 + 1, ws.max_row + 1):
            name_cn = s(safe_read(ws, r, cols.get("name_cn", 0)))
            fba = s(safe_read(ws, r, cols.get("fba", 0)))
            if not name_cn and not fba:
                continue
            link = s(safe_read(ws, r, cols["link"])) if "link" in cols else ""
            image = s(safe_read(ws, r, cols["image"])) if "image" in cols else ""
            ship.cartons.append(Carton(
                seq=str(len(ship.cartons) + 1),
                boxes=int(num(safe_read(ws, r, cols.get("boxes", 0))) or 1),
                battery=_flag(safe_read(ws, r, cols.get("battery", 0))),
                magnet=_flag(safe_read(ws, r, cols.get("magnet", 0))),
                fba_id=fba, po=s(safe_read(ws, r, cols.get("po", 0))) or "无",
                name_cn=name_cn or "/",
                name_en=s(safe_read(ws, r, cols.get("name_en", 0))) or "/",
                hs_code=s(safe_read(ws, r, cols.get("hs", 0))) or "/",
                weight=num(safe_read(ws, r, cols.get("weight", 0))),
                length=num(safe_read(ws, r, cols.get("length", 0))),
                width=num(safe_read(ws, r, cols.get("width", 0))),
                height=num(safe_read(ws, r, cols.get("height", 0))),
                qty=num(safe_read(ws, r, cols.get("qty", 0))),
                unit=num(safe_read(ws, r, cols.get("unit", 0))),
                currency=s(safe_read(ws, r, cols.get("currency", 0))) or "USD",
                total=num(safe_read(ws, r, cols.get("total", 0))),
                brand=s(safe_read(ws, r, cols.get("brand", 0))),
                model=s(safe_read(ws, r, cols.get("model", 0))),
                mat_cn=s(safe_read(ws, r, cols.get("mat_cn", 0))),
                mat_en=s(safe_read(ws, r, cols.get("mat_en", 0))),
                use_cn=s(safe_read(ws, r, cols.get("use_cn", 0))),
                use_en=s(safe_read(ws, r, cols.get("use_en", 0))),
                net_g=num(safe_read(ws, r, cols.get("net", 0))),
                image=image or link,
                link=link,
            ))
        if not ship.cartons:
            return None
        if not ship.total_boxes:
            ship.total_boxes = len(ship.cartons)
        return ship
    finally:
        wb.close()


class ShangruiPlugin:
    name = "尚睿"

    def discover(self, ctx: RunContext) -> list[Path]:
        if not SHANGRUI_DIR.exists():
            return []
        return [p for p in sorted(SHANGRUI_DIR.glob("*.xlsx"))
                if not p.name.startswith("~$") and "我司" not in p.name]

    def process(self, path: Path, ctx: RunContext) -> list[Generated]:
        try:
            ship = parse_shangrui(path)
        except Exception as exc:  # noqa: BLE001
            ctx.log.error("  [失败] %s: %s", path.name, exc)
            return [Generated(file=f"{self.name}/{path.name}", problems=[str(exc)])]
        if ship is None:
            return [Generated(file=f"{self.name}/{path.name}",
                              problems=["未识别模板或没有数据"])]
        return [emit_shipment(ship, ctx)]


register(ShangruiPlugin())
