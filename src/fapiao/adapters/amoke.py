# -*- coding: utf-8 -*-
"""爱墨科技适配器（2026-11 新增）。

数据链路与其它客户不同——**两个源表 + 一对多输出**：

    订舱表（一行 = 一个 FBA 货件，带「仓库代码 / 渠道 / 报关抬头 / 总箱数 / 是否包税」）
        ↓ 用 FBA 号关联
    清关发票（一行 = 一个 SKU 箱，按「渠道 + 仓库代码」筛出来的一个文件，内含多个 FBA）
        ↓ 按「报关抬头」分组
    我司发票（一张 = 一个 仓库代码 + 箱数 + 抬头），文件名 ``RFD2+11+BY.xlsx``

三点业务约定（已与用户确认）：
1. 清关发票 = 订舱表按「仓库代码 + 渠道」筛选的结果，抬头取自订舱表，而不是清关发票；
2. 总箱数以订舱表为准，明细则一行一箱（同一个 FBA 有 M 行、B 箱时按比例摊：JL 是 1 箱 3 行）；
3. 「运单备注」写抬头，客户上传系统按 sheet 名 `T6系统发票` 取数。
"""
from __future__ import annotations

import logging
import re
from pathlib import Path

import openpyxl

from ..common.strings import num, normalize_label, s
from ..config import (AMOKE_CHANNEL_MAP, AMOKE_CUSTOMS_MAP, AMOKE_DELIVERY_WINDOW,
                      AMOKE_DIR, AMOKE_SHEET)
from ..models.address import parse_address_block, receiver_from_book
from ..models.shipment import Carton, Shipment
from ..registry import register
from .base import Generated, RunContext, emit_shipment

log = logging.getLogger(__name__)

BOOKING_DIR = AMOKE_DIR / "材料"        # 旧布局（单批次）
CLEARANCE_DIR = AMOKE_DIR / "客户发票"   # 旧布局（单批次）

# ---- 订舱表列名（normalize_label 后比较）----
BOOKING_COLS = {
    "fba": ("FBA号", "FBA编号"),
    "channel": ("渠道",),
    "warehouse": ("仓库代码",),
    "tou": ("报关抬头",),
    "boxes": ("总箱数",),
    "tax": ("是否包税",),
    "po": ("订单号",),
    "ref": ("海外清关合并编号", "报关合并编号"),
    "addr": ("仓库地址",),
    "zip": ("邮编",),
    "country": ("目的国",),
    "region": ("区域",),
}

# ---- 清关发票列名 ----
CLEARANCE_COLS = {
    "fba": ("Shipment ID", "ShipmentID"),
    "po": ("Reference ID", "ReferenceID"),
    "name_en": ("英文品名",),
    "name_cn": ("中文品名",),
    "hs": ("海关代码",),
    "brand": ("Brand品牌", "Brand"),
    "mat": ("材质",),
    "use": ("用途",),
    "link": ("ASIN/销售链接", "ASIN销售链接"),
    "unit": ("申报单价",),
    "battery": ("是否带电",),
    "model": ("型号",),
    "qty": ("总数量",),
    "net": ("单箱净重",),
    "weight": ("单箱重量(kg)", "单箱重量（kg）", "单箱重量"),
    "length": ("单箱长(cm)", "单箱长（cm）", "单箱长"),
    "width": ("单箱宽(cm)", "单箱宽（cm）", "单箱宽"),
    "height": ("单箱高(cm)", "单箱高（cm）", "单箱高"),
    "image": ("图片",),
}

# ---- 清关发票表头（标签 → 右侧取值）----
HEAD_LABELS = {
    "ref": ("客户参考号",),
    "country": ("目的国",),
    "warehouse": ("FBA仓库代码",),
    "tax": ("是否包税",),
    "boxes": ("总件数", "总箱数"),
    "currency": ("申报币种",),
    "addr": ("收件地址",),
    "company": ("收件人公司名",),
    "phone": ("收件人电话",),
    "zip": ("邮编",),
    "state": ("省/州",),
    "city": ("城市",),
}

# 文件名：清关发票_前海万方_09-24_FBA19NNKM3H7_海运-以星海派-RFD2.xls
NAME_RE = re.compile(r"_(海运-[^-_]+)-([A-Z]{3,4}\d?)\s*\.?[a-z]*$", re.I)


# ============================ 读表工具 ============================
def _read_rows(path: Path) -> list[list]:
    """把第一个 sheet 读成二维列表（.xls 走 xlrd，其余走 openpyxl）。"""
    suffix = path.suffix.lower()
    if suffix == ".xls":
        import xlrd  # 只在需要时导入

        book = xlrd.open_workbook(str(path))
        sh = book.sheet_by_index(0)
        return [[sh.cell_value(r, c) for c in range(sh.ncols)]
                for r in range(sh.nrows)]

    wb = openpyxl.load_workbook(str(path), data_only=True)
    try:
        ws = wb.worksheets[0]
        return [[ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
                for r in range(1, ws.max_row + 1)]
    finally:
        wb.close()


def _norm(v) -> str:
    return normalize_label(s(v))


def _header_index(rows: list[list], must: tuple[str, ...], limit: int = 30) -> int | None:
    """找到包含 must 全部标签的那一行（下标）。"""
    wanted = tuple(_norm(x) for x in must)
    for i, row in enumerate(rows[:limit]):
        cells = {_norm(v) for v in row}
        if all(w in cells for w in wanted):
            return i
    return None


def _col_map(row: list, spec: dict) -> dict[str, int]:
    cells = {_norm(v): idx for idx, v in enumerate(row) if _norm(v)}
    out: dict[str, int] = {}
    for field, aliases in spec.items():
        for alias in aliases:
            key = _norm(alias)
            if key in cells:
                out[field] = cells[key]
                break
        else:
            # 退一步做前缀匹配（"单箱重量(kg)" 少个括号也能认）
            for key, idx in cells.items():
                if any(key.startswith(_norm(a)[:4]) for a in aliases if len(_norm(a)) >= 4):
                    out[field] = idx
                    break
    return out


def _value_at(rows: list[list], label: str, max_scan: int = 15) -> str:
    """标签右侧（同一行）取值。"""
    want = _norm(label)
    for i, row in enumerate(rows[:max_scan]):
        for c, cell in enumerate(row):
            if _norm(cell) == want:
                for cc in range(c + 1, min(c + 4, len(row))):
                    if s(row[cc]):
                        return s(row[cc])
    return ""


# ============================ 订舱表 ============================
def load_booking(path: Path) -> dict[str, dict]:
    """FBA 号 → 订舱表行（含抬头/箱数/包税/渠道）。"""
    wb = openpyxl.load_workbook(str(path), data_only=True)
    try:
        rows = [[ws_cell for ws_cell in row]
                for row in wb.worksheets[0].iter_rows(values_only=True)]
    finally:
        wb.close()

    hi = _header_index(rows, ("FBA号", "仓库代码"))
    if hi is None:
        raise ValueError("订舱表缺少「FBA号 / 仓库代码」表头")
    cols = _col_map(rows[hi], BOOKING_COLS)

    def cell(row, key):
        idx = cols.get(key)
        return row[idx] if idx is not None and idx < len(row) else None

    out: dict[str, dict] = {}
    for row in rows[hi + 1:]:
        fba = s(cell(row, "fba")).upper()
        if not fba or fba in out:
            continue
        out[fba] = {
            "fba": fba,
            "channel": s(cell(row, "channel")),
            "warehouse": s(cell(row, "warehouse")).upper(),
            "tou": s(cell(row, "tou")),
            "boxes": int(num(cell(row, "boxes")) or 0),
            "tax": s(cell(row, "tax")),
            "po": s(cell(row, "po")),
            "ref": s(cell(row, "ref")),
            "addr": s(cell(row, "addr")),
            "zip": s(cell(row, "zip")),
            "country": s(cell(row, "country")),
            "region": s(cell(row, "region")),
        }
    return out


# ============================ 清关发票 ============================
EXCEL_SUFFIX = (".xls", ".xlsx", ".xlsm")


def _batch_dirs() -> list[Path]:
    """批次目录：`07爱墨科技/<批次>/`，一个批次 = 一张订舱表 + 若干清关发票。

    也兼容旧布局（文件直接放在 `07爱墨科技/材料/`、`07爱墨科技/客户发票/`）。
    """
    if not AMOKE_DIR.exists():
        return []
    aux = {"材料", "客户发票"}
    dirs = [p for p in sorted(AMOKE_DIR.iterdir())
            if p.is_dir() and p.name not in aux and not p.name.startswith((".", "~"))]
    # 顶层直接放订舱表（旧布局）也算一个批次
    if any("订舱" in p.name for p in AMOKE_DIR.glob("*.xlsx")):
        dirs.insert(0, AMOKE_DIR)
    return dirs or [AMOKE_DIR]


def _scan_dir(folder: Path) -> list[Path]:
    """一个目录及其 `客户发票/`、`材料/` 子目录里的 excel（不含临时文件）。"""
    if not folder.exists():
        return []
    found: list[Path] = []
    for sub in (folder, folder / "客户发票", folder / "材料"):
        if sub.is_dir():
            found += [p for p in sorted(sub.iterdir())
                      if p.is_file() and p.suffix.lower() in EXCEL_SUFFIX
                      and not p.name.startswith(("~$", "."))]
    return found


def _batch_bookings(batch: Path) -> list[Path]:
    return [p for p in _scan_dir(batch) if "订舱" in p.name and p.suffix.lower() == ".xlsx"]


def _clearance_files(batch: Path | None = None) -> list[Path]:
    """清关发票：批次目录里除订舱表以外的 excel。"""
    folders = [batch] if batch is not None else _batch_dirs()
    out: list[Path] = []
    for folder in folders:
        out += [p for p in _scan_dir(folder) if "订舱" not in p.name]
    return out


def _name_hints(path: Path) -> tuple[str, str]:
    """从文件名兜底拿 (渠道, 仓码)。"""
    m = NAME_RE.search(path.stem)
    if not m:
        return "", ""
    return m.group(1), m.group(2).upper()


def parse_clearance(path: Path, booking: dict[str, dict]) -> list[Shipment]:
    """一个清关发票 → 按「报关抬头」拆成 N 个 Shipment。"""
    rows = _read_rows(path)
    hi = _header_index(rows, ("ShipmentID", "总数量"))
    if hi is None:
        raise ValueError("未识别清关发票明细表头（Shipment ID / 总数量）")
    cols = _col_map(rows[hi], CLEARANCE_COLS)

    warehouse = _value_at(rows, "FBA仓库代码").upper()
    channel_raw = _name_hints(path)[0]
    if not warehouse:
        warehouse = _name_hints(path)[1]

    items: list[dict] = []
    for row in rows[hi + 1:]:
        def cell(key):
            idx = cols.get(key)
            return row[idx] if idx is not None and idx < len(row) else None

        fba = s(cell("fba")).upper()
        name_cn = s(cell("name_cn"))
        if not fba and not name_cn:
            continue
        items.append({
            "fba": fba,
            "po": s(cell("po")),
            "name_cn": name_cn,
            "name_en": s(cell("name_en")),
            "hs": s(cell("hs")),
            "brand": s(cell("brand")),
            "mat": s(cell("mat")),
            "use": s(cell("use")),
            "link": s(cell("link")),
            "unit": num(cell("unit")),
            "battery": s(cell("battery")),
            "model": s(cell("model")),
            "qty": num(cell("qty")),
            "weight": num(cell("weight")),
            "length": num(cell("length")),
            "width": num(cell("width")),
            "height": num(cell("height")),
            "image": s(cell("image")),
        })
    if not items:
        return []

    # 整票（文件）级带电量：人工成品里「是否带电」是按整份清关发票判定的，
    # 同一文件里只要有一行带电就填「是」（如 JL 色带与带锂电池打印机同文件）。
    file_battery = "是" if any(
        s(it["battery"]) in ("是", "带电", "Y", "YES") for it in items) else "否"
    for it in items:
        it["file_battery"] = file_battery

    # ---- 按抬头分组（抬头来自订舱表；查不到的单列出来，由调用方报告）----
    groups: dict[str, list[dict]] = {}
    for it in items:
        info = booking.get(it["fba"]) or {}
        it["tou"] = info.get("tou") or ""
        it["info"] = info
        groups.setdefault(it["tou"] or "未知抬头", []).append(it)

    ships: list[Shipment] = []
    for tou, group in groups.items():
        ship = _build_shipment(path, tou, group, warehouse, channel_raw, rows)
        if ship is not None:
            ships.append(ship)
    return ships


def _build_shipment(path: Path, tou: str, group: list[dict],
                    warehouse: str, channel_raw: str, rows: list[list]) -> Shipment | None:
    info0 = group[0].get("info") or {}
    wh = (info0.get("warehouse") or warehouse or "").upper()
    channel_src = info0.get("channel") or channel_raw
    channel = AMOKE_CHANNEL_MAP.get(channel_src, channel_src)

    # 总箱数：同一 FBA 只算一次（订舱表一行一个 FBA）
    boxes_by_fba: dict[str, int] = {}
    for it in group:
        fba = it["fba"]
        if fba and fba not in boxes_by_fba:
            boxes_by_fba[fba] = int((it.get("info") or {}).get("boxes") or 0)
    total_boxes = sum(boxes_by_fba.values())

    tax = s(info0.get("tax")) or _value_at(rows, "是否包税")
    # 客户参考号：成品里留空（清关发票表头也没有值），要用订舱表合并编号时在 config 里开
    ref = _value_at(rows, "客户参考号")
    remarks = tou if tou and tou != "未知抬头" else ""

    ship = Shipment(
        source_file=path.name,
        customer="爱墨科技",
        template="001",
        ref=ref,
        doc_no="",
        channel=channel,
        warehouse=wh,
        battery=group[0].get("file_battery") or "否",
        magnet="否",
        customs=AMOKE_CUSTOMS_MAP.get(tax, "不需报关"),
        total_boxes=total_boxes or len(group),
        remarks=remarks,
        invoice_sheet=AMOKE_SHEET,
        delivery_window=AMOKE_DELIVERY_WINDOW,
        out_name=f"{wh}+{total_boxes or len(group)}+{tou}".strip("+"),
    )
    ship.receiver = (receiver_from_book(wh)
                     or parse_address_block(info0.get("addr", ""), wh))

    # 箱序号：按明细行顺序 1..N；行数比箱数多时（1 箱装多个 SKU）封顶到箱数，
    # 与人工成品一致（JL 是 1 箱 3 行 → 1/1/1）。
    for i, it in enumerate(group):
        seq = min(i + 1, total_boxes) if total_boxes else i + 1
        ship.cartons.append(Carton(
            seq=str(seq),
            boxes=1,
            battery="是" if s(it["battery"]) in ("是", "带电") else "否",
            magnet="否",
            fba_id=it["fba"],
            po=it["po"] or "无",
            name_cn=it["name_cn"],
            name_en=it["name_en"],
            hs_code=it["hs"],
            weight=it["weight"],
            length=it["length"],
            width=it["width"],
            height=it["height"],
            qty=it["qty"],
            unit=it["unit"],
            total=None,              # 让构建器按 =O*N 出公式，与模板惯例一致
            currency="USD",
            brand=it["brand"],
            model=it["model"],
            mat_cn=it["mat"],
            use_cn=it["use"],
            image=it["image"],
            link=it["link"],
        ))
    if not ship.cartons:
        return None
    # 爱墨人工成品：空列留空（不补 "/"）；PO 空仍写「无」
    ship.empty_as_slash = False
    return ship


# ============================ 插件 ============================
class AmokePlugin:
    name = "爱墨科技"

    def discover(self, ctx: RunContext) -> list[Path]:
        """主表 = 每个批次目录里的订舱表；清关发票在 process_all 里按批次扫。"""
        bookings: list[Path] = []
        for batch in _batch_dirs():
            bookings += _batch_bookings(batch)
        return bookings

    def unit_count(self, files: list[Path], ctx: RunContext) -> int:
        return len(files) + len(_clearance_files())

    def process_all(self, files: list[Path], ctx: RunContext) -> list[Generated]:
        if not files:
            return []

        items: list[Generated] = []
        if not files:
            return [Generated(file=f"{self.name}/订舱表",
                              problems=[f"没找到订舱表（文件名要含「订舱」）：{AMOKE_DIR}"])]

        booked = {p.parent for p in files}
        for batch in _batch_dirs():
            if batch not in booked:
                ctx.log.warning("  批次 %s：没有订舱表，跳过", batch.name)
                items.append(Generated(
                    file=f"{self.name}/{batch.name}",
                    warnings=["该批次下没有订舱表（文件名要含「订舱」），已跳过"]))

        for path in files:
            batch = path.parent
            clearances = _clearance_files(batch)
            try:
                booking = load_booking(path)
            except Exception as exc:  # noqa: BLE001
                ctx.log.error("  [失败] 订舱表 %s: %s", path.name, exc)
                items.append(Generated(
                    file=f"{self.name}/{batch.name}/{path.name}", problems=[str(exc)]))
                continue
            ctx.log.info("  批次 %s：订舱表 %d 个 FBA，清关发票 %d 份",
                         batch.name, len(booking), len(clearances))
            if not clearances:
                items.append(Generated(
                    file=f"{self.name}/{batch.name}/清关发票",
                    problems=["该批次下没有清关发票"]))
                continue

            for cf in clearances:
                try:
                    ships = parse_clearance(cf, booking)
                except Exception as exc:  # noqa: BLE001
                    ctx.log.error("  [失败] %s: %s", cf.name, exc)
                    items.append(Generated(
                        file=f"{self.name}/{batch.name}/{cf.name}", problems=[str(exc)]))
                    self._tick(ctx)
                    continue

                if not ships:
                    items.append(Generated(
                        file=f"{self.name}/{batch.name}/{cf.name}",
                        problems=["清关发票没有可识别的明细行"]))
                for ship in sorted(ships, key=lambda x: (x.warehouse, x.channel, x.remarks)):
                    if ship.remarks == "":
                        ship.warnings.append("订舱表里查不到该 FBA 的报关抬头")
                    items.append(emit_shipment(ship, ctx))
                self._tick(ctx)
        return items

    @staticmethod
    def _tick(ctx: RunContext) -> None:
        progress = getattr(ctx, "progress", None)
        if progress is not None:
            progress.advance(1)

    def process(self, path: Path, ctx: RunContext) -> list[Generated]:
        return self.process_all([path], ctx)


register(AmokePlugin())
