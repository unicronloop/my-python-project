# -*- coding: utf-8 -*-
"""德邦交接单索引：把「交接单号 / 单号」映射到 FBA 货件号，并给出申报参考量。

两种源文件，一律按「文件名 + 表内任意 DPL 字符串」识别，不依赖列名：
- 材料2\\10月28日船期直发上海交接228件.xlsx：行内「德邦交接单号」+ 备注里
  `DPL112000161615+FBA19TWSX9DV+483件+5箱+限时达` 这样的路径串；
- 德邦材料3\\万方交接单-9.26.xlsx 等：「单号」列 + 「FBA」列。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from ..common.excel import load_workbook_any, safe_read
from ..common.strings import normalize_label, num, s

DPL_RE = re.compile(r"DPL\d{6,}", re.I)
FBA_RE = re.compile(r"FBA[0-9A-Z]{5,}", re.I)
QTY_RE = re.compile(r"(\d+)\s*件")
BOX_RE = re.compile(r"(\d+)\s*箱")


def is_handover(path: Path) -> bool:
    return "交接" in path.stem


@dataclass
class HandoverEntry:
    dpl: str
    fba: str = ""
    qty: int | None = None
    boxes: int | None = None
    channel: str = ""
    warehouse: str = ""
    src: str = ""


@dataclass
class HandoverIndex:
    entries: list[HandoverEntry] = field(default_factory=list)
    by_dpl: dict[str, list[HandoverEntry]] = field(default_factory=dict)
    by_fba: dict[str, list[HandoverEntry]] = field(default_factory=dict)

    def add(self, entry: HandoverEntry) -> None:
        self.entries.append(entry)
        self.by_dpl.setdefault(entry.dpl, []).append(entry)
        if entry.fba:
            self.by_fba.setdefault(entry.fba, []).append(entry)

    def expect_for(self, dpl: str) -> HandoverEntry | None:
        """同一 DPL 的多条记录合并成一条参考量。"""
        rows = self.by_dpl.get(dpl) or []
        if not rows:
            return None
        qty = sum(r.qty or 0 for r in rows) or None
        boxes = sum(r.boxes or 0 for r in rows) or None
        first = rows[0]
        return HandoverEntry(dpl=dpl, fba=first.fba, qty=qty, boxes=boxes,
                             channel=first.channel, warehouse=first.warehouse,
                             src=first.src)


def _cell_texts(ws, max_row: int = 200, max_col: int = 30) -> list[list[str]]:
    out = []
    for r in range(1, min(ws.max_row, max_row) + 1):
        row = []
        for c in range(1, min(ws.max_column, max_col) + 1):
            val = safe_read(ws, r, c)
            row.append(s(val) if val is not None else "")
        out.append(row)
    return out


def load_handover_index(paths: list[Path]) -> HandoverIndex:
    index = HandoverIndex()
    for path in paths:
        try:
            wb = load_workbook_any(path, data_only=True)
        except Exception:  # noqa: BLE001
            continue
        try:
            for ws in wb.worksheets:
                rows = _cell_texts(ws)
                table = _table_entries(rows, path.name)
                if table:
                    for entry in table:
                        index.add(entry)
                    continue
                for cells in rows:
                    line = " | ".join(x for x in cells if x)
                    dpls = [m.upper() for m in DPL_RE.findall(line)]
                    fbas = [m.upper() for m in FBA_RE.findall(line)]
                    if not dpls:
                        continue
                    qty = _int(QTY_RE.search(line).group(1)) if QTY_RE.search(line) else None
                    boxes = _int(BOX_RE.search(line).group(1)) if BOX_RE.search(line) else None
                    for i, dpl in enumerate(dict.fromkeys(dpls)):
                        fba = fbas[i] if i < len(fbas) else (fbas[0] if len(fbas) == 1 else "")
                        index.add(HandoverEntry(dpl=dpl, fba=fba, qty=qty,
                                                boxes=boxes, src=path.name))
        finally:
            wb.close()
    return index


def _int(text: str) -> int | None:
    value = num(text)
    return int(value) if value is not None else None


# ---------------------------------------------------------------- 表格式交接单
_HDR_HINTS = {
    "dpl": ("单号", "交接单号"),
    "fba": ("fba", "fba号", "fba标签", "fba编码"),
    "channel": ("渠道",),
    "warehouse": ("仓库",),
    "boxes": ("件数", "箱数"),
}


def _table_entries(rows: list[list[str]], src: str) -> list[HandoverEntry]:
    """识别带表头的交接单（单号 + FBA 两列同时存在），按列取值。"""
    for hi, row in enumerate(rows[:15]):
        labels = [normalize_label(c) for c in row]
        if not any(c and ("fba" in c) for c in labels):
            continue
        if not any(c and "单号" in c for c in labels):
            continue
        col: dict[str, int] = {}
        for key, hints in _HDR_HINTS.items():
            for i, lab in enumerate(labels):
                if lab and any(h in lab for h in hints):
                    col.setdefault(key, i)
                    break
        if "dpl" not in col:
            continue
        out: list[HandoverEntry] = []
        for row2 in rows[hi + 1:]:
            cell_dpl = row2[col["dpl"]] if col["dpl"] < len(row2) else ""
            m = DPL_RE.search(cell_dpl)
            dpl = m.group(0).upper() if m else ""
            if not dpl:
                # 单号列可能写的是 FBA 号，真正的 DPL 只在备注里
                joined = " | ".join(x for x in row2 if x)
                m2 = DPL_RE.search(joined)
                if not m2:
                    continue
                dpl = m2.group(0).upper()
            fba = row2[col["fba"]] if "fba" in col and col["fba"] < len(row2) else ""
            fba_m = FBA_RE.search(fba or "")
            if not fba_m:
                joined = " | ".join(x for x in row2 if x)
                fba_m = FBA_RE.search(joined)
            channel = row2[col["channel"]] if "channel" in col and col["channel"] < len(row2) else ""
            warehouse = (row2[col["warehouse"]]
                         if "warehouse" in col and col["warehouse"] < len(row2) else "")
            boxes = row2[col["boxes"]] if "boxes" in col and col["boxes"] < len(row2) else ""
            boxes_m = BOX_RE.search(boxes or "")
            out.append(HandoverEntry(
                dpl=dpl,
                fba=fba_m.group(0).upper() if fba_m else "",
                qty=_int(boxes_m.group(1)) if boxes_m else None,
                boxes=_int(boxes_m.group(1)) if boxes_m else None,
                channel=s(channel),
                warehouse=s(warehouse),
                src=src,
            ))
        if out:
            return out
    return []
