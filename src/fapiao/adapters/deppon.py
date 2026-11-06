# -*- coding: utf-8 -*-
"""德邦适配器：客户模板家族多、版式漂移大，一律按「表头文字」定位，不写死行列号。

六个家族（旧发票1/2/3/4/5/6）版式各不相同，但语义一致：
单据头（参考号/仓码/渠道/带电/报关）+ 逐箱货品行（箱号/品名/HS/数量/单价/重量/尺寸）。

所以这里只保留**一个**解析器：
1. 用表头文字判断家族（仅用于报告和兜底渠道）；
2. 用标签定位单据头的每个字段；
3. 用表头行定位列，逐行读「箱 + 货品」；同一箱的多行货品合并成一箱。
这样换模板只改别名表，不用再写第五个 parser。
"""
from __future__ import annotations

import logging
import re
from pathlib import Path

from ..common.excel import load_workbook_any, safe_read
from ..common.strings import cn_en, normalize_label, num, s, split_dims
from ..config import CUSTOMS_MAP, DEPPON_CHANNEL, DEPPON_DIR, normalize_channel
from ..models.shipment import Carton, Shipment
from ..registry import register
from .base import Generated, RunContext, emit_shipment
from .handover import HandoverIndex, is_handover, load_handover_index

log = logging.getLogger(__name__)

DPL_RE = re.compile(r"DPL\d{6,}", re.I)
FBA_RE = re.compile(r"FBA[0-9A-Z]{5,}", re.I)
WH_RE = re.compile(r"^(?:[A-Z]{3,4}\d|[A-Z]{3,4})$")
FILENAME_WH_RE = re.compile(r"(?<![A-Za-z0-9])([A-Z]{3,4}\d)(?![A-Za-z0-9])")
TOTAL_MARKERS = ("合计", "总计", "共计", "总件数", "总箱数", "total")

# ---- 单据头标签 ----
HEAD = {
    "ref": ("客户订单号", "客户参考号1", "客户参考号", "*ref./批次跟踪号", "ref./批次跟踪号"),
    "channel": ("服务", "派送方式", "收货渠道", "备注"),
    "warehouse": ("地址库编码", "fba仓码", "仓库地址", "consignee收件人", "收货仓"),
    "battery": ("带电",),
    "magnet": ("带磁",),
    "customs": ("报关方式", "报关类型",),
    "boxes": ("箱数", "总件数", "总箱数"),
    "weight": ("预报实重",),
    "value": ("申报总值",),
    "company": ("收件人公司", "companyname公司名", "收件公司"),
    "name": ("收件人姓名", "收件人名称",),
    "addr1": ("收件人地址一", "address地址", "收件地址"),
    "city": ("收件人城市", "city城市名",),
    "state": ("收件人省份/州", "province州省",),
    "zip": ("收件人邮编", "code邮编",),
    "country": ("收件人国家代码", "country国家/区域", "目的地"),
}

# ---- 明细列 ----
COL = {
    "seq": ("箱序号", "序号"),
    "boxes": ("箱数", "箱数(ctn)", "箱", "件数（箱）"),
    "fba": ("货箱编号", "fba 标签编号", "fba标签编号", "fba编号",
            "单号/fba标签", "物品fba id", "货箱编号（货件追踪编号）"),
    "po": ("ponumber", "po number", "物品poanumber", "po号", "referenceid", "reference id",
            "货物追踪号", "货物追踪编号", "追踪号"),
    "name_cn": ("产品中文品名", "中文品名", "物品中文名", "商品名称"),
    "name_en": ("产品英文品名", "英文品名", "物品英文名"),
    "name": ("品名", "commodity品名", "fulldescriptionofgoods", "详细的商品名称"),
    "hs": ("海关编码", "hs code", "hscode", "hs编码"),
    "qty": ("产品申报数量", "数量", "total qty", "单件数量", "单箱数量", "qty"),
    "unit": ("产品申报单价", "单价", "unit value"),
    "total": ("总价", "总价值", "total amout", "单件申报总值", "单箱申报总值"),
    "weight": ("货箱重量", "单件毛重", "单箱实重", "实重", "重量"),
    "dims": ("材积cm", "尺寸(cm)", "尺寸"),
    "length": ("长(cm)", "长(cm)", "箱长"),
    "width": ("宽(cm)", "箱宽"),
    "height": ("高(cm)", "箱高"),
    "gross": ("总毛量", "总毛重"),
    "net": ("总净重", "净重"),
    "mat": ("产品材质", "材质"),
    "use": ("产品用途", "用途"),
    "brand": ("产品品牌", "品牌"),
    "model": ("产品型号", "型号", "规格型号"),
    "link": ("销售链接", "销售链接(欧洲英国需要填写)"),
    "image": ("产品图片", "产品照片"),
    "battery": ("带电",),
    "magnet": ("带磁",),
}

FINGERPRINTS = (
    ("发票3", ("*ref./批次跟踪号", "ref./批次跟踪号", "dp-发票模板")),
    ("发票6", ("fba专线出货资料模块", "fba专线出货资料模板")),
    ("发票2", ("fba 标签编号*", "fba 标签编号")),
    ("发票4", ("德邦fba海运交接资料",)),
    ("发票5", ("单号/fba标签",)),
    ("发票1", ("货箱编号*", "货箱编号")),
)


# ---------------------------------------------------------------- 基础工具
def _scan(ws, max_row: int = 80, max_col: int = 40):
    for r in range(1, min(ws.max_row, max_row) + 1):
        for c in range(1, min(ws.max_column, max_col) + 1):
            val = safe_read(ws, r, c)
            if val is not None and s(val) != "":
                yield r, c, normalize_label(val), val


def _label_rows(ws, aliases, limit: int = 60):
    """所有匹配标签的 (row, col)。"""
    wanted = [normalize_label(a) for a in aliases]
    out = []
    for r, c, text, _ in _scan(ws, limit, 30):
        for w in wanted:
            if text == w or (len(w) >= 4 and w in text):
                out.append((r, c))
                break
    return out


def _label_set() -> set[str]:
    return {normalize_label(a) for group in list(HEAD.values()) for a in group}


_LABELS = _label_set()


def _head_value(ws, aliases, limit: int = 60):
    """标签右侧（或正下方）第一个「看起来不是另一个标签」的值。"""
    for r, c in _label_rows(ws, aliases, limit):
        for rr, cc in [(r, c + 1), (r, c + 2), (r, c + 3), (r, c + 4), (r + 1, c)]:
            val = safe_read(ws, rr, cc)
            text = s(val)
            if not text or normalize_label(text) in _LABELS:
                continue
            return val
    return None


def _find_header(ws, limit: int = 60):
    """表头行：同时出现「箱/货品 + 数量/单价/HS」的那一行。"""
    best = None
    for r in range(1, min(ws.max_row, limit) + 1):
        texts = {normalize_label(safe_read(ws, r, c))
                 for c in range(1, min(ws.max_column, 40) + 1)}
        hit = sum(1 for group in COL.values() if any(
            a in texts for a in (normalize_label(x) for x in group)))
        has_name = any("品名" in t or "description" in t for t in texts)
        if hit >= 4 and has_name and (best is None or hit > best[1]):
            best = (r, hit)
    return best[0] if best else None


def _col_map(ws, header_row: int) -> dict[str, int]:
    cells: dict[str, int] = {}
    for c in range(1, min(ws.max_column, 40) + 1):
        text = normalize_label(safe_read(ws, header_row, c))
        if text and text not in cells:
            cells[text] = c
    out: dict[str, int] = {}
    for field, aliases in COL.items():
        for alias in aliases:
            key = normalize_label(alias)
            if key in cells:
                out[field] = cells[key]
                break
        if field in out:
            continue
        for text, c in cells.items():           # 近似：标签包含关系（「数量(PCS)」含「数量」）
            if any(len(normalize_label(a)) >= 2 and normalize_label(a) in text
                   for a in aliases):
                out[field] = c
                break
    return out


SUB_DIM_RE = re.compile(r"^(l|w|h)\s*\(?\s*(cm|m)\s*\)?$")


def _split_name_columns(ws, header_row: int, cols: dict) -> bool:
    """表头下一行可能是子表头（品名的「中文/英文」、规格的「L(cm)/W(cm)/H(cm)」）。

    补上列号，返回数据行是否从 header+2 开始。
    """
    sub = header_row + 1
    found = False
    name_col = cols.get("name") or cols.get("name_cn")
    for c in range(1, min(ws.max_column, 40) + 1):
        text = normalize_label(safe_read(ws, sub, c)).strip()
        if not text:
            continue
        if name_col and c in (name_col, name_col + 1):
            if text in ("中文", "中文品名"):
                cols.setdefault("name_cn", c)
                found = True
            elif text in ("英文", "英文品名"):
                cols.setdefault("name_en", c)
                found = True
            continue
        m = SUB_DIM_RE.match(text)
        if m:
            cols.setdefault({"l": "length", "w": "width", "h": "height"}[m.group(1)], c)
            found = True
    return found


CHANNEL_KEYS = ("美森限时达", "美森正班", "美森特快线", "合德卡派", "普船",
                 "限时达", "卡派", "海派", "卡航", "快线", "专线", "整拼柜")
CHANNEL_MARKS = ("卡派", "快线", "海派", "特快", "卡航", "限时", "普船", "专线", "含税", "包税")


def _channel_from(text: str) -> str:
    """从「服务/备注/文件名」里认出一个像渠道的串，认不出返回空。"""
    text = s(text)
    if not text:
        return ""
    for key in CHANNEL_KEYS:
        if key in text:
            return normalize_channel(key)
    guess = normalize_channel(text)
    if len(guess) <= 24 and any(m in guess for m in CHANNEL_MARKS):
        return guess
    return ""


def _v(ws, row: int, cols: dict, field: str):
    col = cols.get(field)
    return safe_read(ws, row, col) if col else None


def _is_total(ws, row: int, cols: dict) -> bool:
    for field in ("name_cn", "name_en", "name", "fba", "seq"):
        col = cols.get(field)
        if not col:
            continue
        text = normalize_label(safe_read(ws, row, col))
        if any(m in text for m in TOTAL_MARKERS):
            return True
    return False


def _warehouse_from(value: str) -> str:
    text = s(value).replace("\xa0", " ").strip()
    if not text:
        return ""
    head = re.split(r"\s*[-–]\s*", text, maxsplit=1)[0].strip()
    if WH_RE.match(head):
        return head
    return text if WH_RE.match(text) else ""


def _warehouse_from_name(name: str) -> str:
    for m in FILENAME_WH_RE.finditer(name):
        token = m.group(1)
        if not token.startswith(("FBA", "DPL")):
            return token
    return ""


_CJK_RE = re.compile(r"[一-鿿]")


def _cn_of(text: str) -> str:
    return text if text and _CJK_RE.search(text) else ""


def _en_of(text: str) -> str:
    return text if text and not _CJK_RE.search(text) else ""


_NOISE = {"no", "n/a", "na", "none", "null", "/", "-"}


def _header_text(ws, header_row: int, col: int | None) -> str:
    return normalize_label(safe_read(ws, header_row, col)) if col else ""


def _mat_use(ws, header_row: int, cols: dict, r: int) -> tuple[str, str, str, str]:
    """材质/用途：列表头带「英文」的写英文列，否则一律写中文列（金标如此）。"""
    out = []
    for field, en_field in (("mat", "mat_en"), ("use", "use_en")):
        text = s(_v(ws, r, cols, field))
        extra = s(_v(ws, r, cols, en_field))
        if field == "use":
            # 金标：用途一律写中文列 V（即使是英文词，如 play），英文列 W 留空
            out.append((text or extra, ""))
            continue
        head = _header_text(ws, header_row, cols.get(field))
        if not text and extra:
            out.append(("", extra))
        elif _CJK_RE.search(text):
            out.append((text, extra or ""))           # 中文材质 → 中文列，英文补到英文列
        elif text:
            out.append(("", text))                    # 纯英文 → 英文列
        else:
            out.append(("", ""))
    (m_cn, m_en), (u_cn, u_en) = out
    return m_cn, m_en, u_cn, u_en


def _brand(text: str) -> str:
    """品牌列常被填成 NO/N/A 之类的占位，金标留空。"""
    return "" if text.strip().lower() in _NOISE else text.strip()


def _raw_cell(ws, r: int, c: int | None):
    """未经合并单元格兜底的原始值（判断“新箱起点”时必须用它）。"""
    if not c:
        return None
    try:
        return ws.cell(r, c).value
    except Exception:                                    # xls 垫片可能不支持
        return safe_read(ws, r, c)


def _base_fba(fba: str) -> str:
    """FBA19QZ88PV6U000001 → FBA19QZ88PV6（金标只写基础货件号）。"""
    return re.sub(r"U\d{4,}$", "", re.sub(r"\s+", "", fba), flags=re.I)


# ---------------------------------------------------------------- 主解析
def _parse_sheet(ws, path: Path, ref: str, kind: str) -> Shipment | None:
    header_row = _find_header(ws)
    if not header_row:
        return None
    cols = _col_map(ws, header_row)
    sub_header = _split_name_columns(ws, header_row, cols)
    if "name_cn" not in cols and "name_en" not in cols and "name" not in cols:
        return None
    if "hs" not in cols and "qty" not in cols and "fba" not in cols:
        return None

    addr_raw = s(_head_value(ws, ("仓库地址", "address地址", "收件地址", "地址")))
    wh = _warehouse_from(s(_head_value(ws, HEAD["warehouse"])))
    if not wh and addr_raw:
        wh = _warehouse_from(addr_raw)
    if not wh:
        wh = _warehouse_from_name(path.name)

    channel = (_channel_from(s(_head_value(ws, ("服务", "派送方式", "收货渠道"))))
               or _channel_from(s(_head_value(ws, ("备注",))))
               or _channel_from(path.stem))

    ship = Shipment(
        source_file=path.name, customer="德邦", template=kind, ref=ref,
        channel=channel or DEPPON_CHANNEL.get("invoice1", "美森特快线"),
        warehouse=wh,
        battery="是" if s(_head_value(ws, ("带电",))) in ("是", "带电", "Y") else "否",  # 明细再兜底
        magnet="是" if s(_head_value(ws, ("带磁",))) in ("是", "带磁", "Y") else "否",
        customs=CUSTOMS_MAP.get(normalize_label(s(_head_value(ws, HEAD["customs"]))), "不需报关"),
    )

    # 收件人：源表「仓库地址」优先（客户给的是最新地址，001 地址库会过期），
    # 源表没有地址时才回落到 001 模板的 FBA仓码地址 页。
    from ..models.address import parse_address_block, receiver_from_book
    book_rc = receiver_from_book(wh) if wh else None
    if book_rc and book_rc.get("city"):
        # 地址库能命中就用库里的（城市/州/邮编齐全），源表地址只当兜底
        ship.receiver = dict(book_rc)
    elif addr_raw:
        rc = parse_address_block(addr_raw, wh)
        if rc.get("addr1"):
            ship.receiver = rc

    head_boxes = num(_head_value(ws, HEAD["boxes"]))
    head_weight = num(_head_value(ws, HEAD["weight"]))
    head_value = num(_head_value(ws, HEAD["value"]))

    cartons: list[Carton] = []
    groups: list[dict] = []          # 每个「箱」一组（金标：箱序号按组递增）
    cur: dict | None = None
    for r in range(header_row + (2 if sub_header else 1), ws.max_row + 1):
        if _is_total(ws, r, cols):
            break
        fba = s(_v(ws, r, cols, "fba"))
        if not FBA_RE.search(fba):
            fba = ""
            for c in range(1, min(ws.max_column, 40) + 1):
                raw = s(_raw_cell(ws, r, c))          # 只认原始单元格，避免合并单元格误命中
                if FBA_RE.search(raw):
                    fba = raw
                    cols.setdefault("fba", c)
                    break
        name_cn = s(_v(ws, r, cols, "name_cn"))
        name_en = s(_v(ws, r, cols, "name_en"))
        name = s(_v(ws, r, cols, "name"))
        if not (fba or name_cn or name_en or name):
            continue
        if not name_cn and not name_en and name:
            name_cn, name_en = cn_en(name)
        # 一格写了「中文/English」：按语言拆开，中文进 G、英文进 H
        for src_text in (name_cn, name_en, name):
            if src_text and _CJK_RE.search(src_text) and re.search(r"[A-Za-z]{3,}", src_text):
                parts = re.split(r"[\n/；;]+", src_text)
                zh = " ".join(x.strip() for x in parts if _CJK_RE.search(x))
                en = " ".join(x.strip() for x in parts if x.strip() and not _CJK_RE.search(x))
                name_cn, name_en = zh or name_cn, en or name_en
                break

        boxes_raw = _v(ws, r, cols, "boxes")
        # 「箱数」列里可能是 FBA 号 / 品名之类的文字，只认纯数字
        if FBA_RE.search(s(boxes_raw)) or not re.fullmatch(r"\s*\d+(?:\.\d+)?\s*", s(boxes_raw)):
            boxes_raw = None
        dims = _v(ws, r, cols, "dims")
        row_weight = num(_v(ws, r, cols, "weight")) or num(_v(ws, r, cols, "gross"))
        # 新箱起点：原始单元格里自带 FBA，且同时带箱数/重量/尺寸
        # （合并单元格会让续行“看起来”有 FBA/重量，必须读原始值）
        raw_fba = s(_raw_cell(ws, r, cols.get("fba")))
        raw_boxes = num(_raw_cell(ws, r, cols.get("boxes")))
        raw_weight = num(_raw_cell(ws, r, cols.get("weight"))) or num(_raw_cell(ws, r, cols.get("gross")))
        raw_dims = s(_raw_cell(ws, r, cols.get("dims")))
        starts = (bool(FBA_RE.search(raw_fba))
                  and s(raw_fba) != (cur or {}).get("fba_raw", "")
                  and bool(raw_boxes or raw_weight or raw_dims))
        if cur is None or starts:
            cur = {
                "seq": str(len(groups) + 1),
                "fba_raw": s(raw_fba),
                "fba": _base_fba(fba) if fba else (groups[-1]["fba"] if groups else ""),
                "boxes": int(num(boxes_raw) or 1),
                "weight": row_weight,
                "dims": s(dims),
                "length": num(_v(ws, r, cols, "length")),
                "width": num(_v(ws, r, cols, "width")),
                "height": num(_v(ws, r, cols, "height")),
                "po": s(_v(ws, r, cols, "po")),
            }
            groups.append(cur)
            if not cur["weight"] and len(groups) > 1 and groups[-2]["weight"]:
                cur["weight"] = groups[-2]["weight"]
            if not cur["dims"] and len(groups) > 1 and groups[-2]["dims"]:
                cur["dims"] = groups[-2]["dims"]
        if dims := cur.get("dims"):
            L, W, H = split_dims(dims)
        else:
            L, W, H = cur.get("length"), cur.get("width"), cur.get("height")
        boxes = cur.get("boxes") or 1
        weight = cur.get("weight")
        mat_cn, mat_en, use_cn, use_en = _mat_use(ws, header_row, cols, r)
        carton = Carton(
            seq=cur["seq"], boxes=boxes, fba_id=cur.get("fba") or _base_fba(fba),
            po=cur.get("po") or "无",
            name_cn=name_cn or "/", name_en=name_en or "/",
            hs_code=s(_v(ws, r, cols, "hs")) or "/",
            weight=weight,
            length=L, width=W, height=H,
            qty=num(_v(ws, r, cols, "qty")),
            unit=num(_v(ws, r, cols, "unit")),
            currency="USD",
            total=num(_v(ws, r, cols, "total")),
            mat_cn=mat_cn, mat_en=mat_en,
            use_cn=use_cn, use_en=use_en,
            brand=_brand(s(_v(ws, r, cols, "brand"))),
            model=s(_v(ws, r, cols, "model")),
            image=s(_v(ws, r, cols, "image")),
            link=s(_v(ws, r, cols, "link")),
        )
        if carton.qty and carton.unit and not carton.total:
            carton.total = round(carton.qty * carton.unit, 4)
        if carton.battery == "是" or normalize_label(s(_v(ws, r, cols, "battery"))) in ("是", "带电", "y"):
            carton.battery = "是"
        cartons.append(carton)

    if not cartons:
        return None

    ship.cartons = cartons
    ship.empty_as_slash = True
    ship.total_boxes = int(sum(g.get('boxes') or 1 for g in groups))
    if any(c.battery == "是" for c in cartons):
        ship.battery = "是"
    # B12 预报实重 / B13 预报方数 / B14 申报总值：金标保持空白，汇总只进校验报告
    ship.total_weight = None
    ship.total_value = None
    if head_boxes and int(head_boxes) != ship.total_boxes:
        ship.warnings.append(f"单据头箱数 {head_boxes:g} ≠ 明细箱数 {ship.total_boxes}")
    if not ship.warehouse:
        ship.warnings.append("未识别仓码（文件名/单价头/地址都没有）")
    return ship


def detect_invoice_kind(wb) -> str | None:
    """按表头文字特征猜版式：命中关键词最多的家族胜出，同分按 FINGERPRINTS 顺序。"""
    seen: list[str] = []
    for sheet in wb.worksheets:
        for _, _, text, _ in _scan(sheet, max_row=25, max_col=40):
            seen.append(normalize_label(text))
    best, best_hits = None, 0
    for kind, keys in FINGERPRINTS:
        hits = sum(1 for k in keys
                   if any(normalize_label(k) in text for text in seen))
        if hits > best_hits:
            best, best_hits = kind, hits
    return best


def parse_deppon(path: Path, ref: str) -> Shipment | None:
    wb = load_workbook_any(path, data_only=True)
    try:
        kind = detect_invoice_kind(wb) or "未知"
        for sheet in wb.worksheets:
            ship = _parse_sheet(sheet, path, ref, kind)
            if ship:
                return ship
        return None
    finally:
        wb.close()


# ---------------------------------------------------------------- 参考号解析
def resolve_ref(path: Path, handovers: HandoverIndex) -> str:
    """DPL 单号：文件名 → 表内任意单元格 → 交接单（按 FBA 反查）。"""
    m = DPL_RE.search(path.name)
    if m:
        return m.group(0).upper()

    fba_m = FBA_RE.search(path.name)
    fba = fba_m.group(0).upper() if fba_m else ""
    if fba and handovers.by_fba.get(fba):
        return handovers.by_fba[fba][0].dpl

    try:
        wb = load_workbook_any(path, data_only=True)
    except Exception:  # noqa: BLE001
        return ""
    try:
        for sheet in wb.worksheets:
            for _, _, _, raw in _scan(sheet, max_row=30, max_col=20):
                m = DPL_RE.search(s(raw))
                if m:
                    return m.group(0).upper()
    finally:
        wb.close()
    return ""


def merge_shipments(ships: list[Shipment]) -> Shipment:
    base = ships[0]
    cartons: list[Carton] = []
    warnings: list[str] = []
    group_boxes: dict[str, int] = {}          # 箱序号 → 该箱的箱数
    offset = 0
    for ship in ships:
        seqs = []
        for c in ship.cartons:
            try:
                seqs.append(int(str(c.seq).split("-")[0]))
            except ValueError:
                seqs.append(1)
        for c, raw_seq in zip(ship.cartons, seqs):
            c.seq = str(raw_seq + offset)
            group_boxes.setdefault(c.seq, c.boxes or 1)
            cartons.append(c)
        offset = max([offset] + [x + offset for x in seqs])
        warnings.extend(ship.warnings)
    merged = Shipment(
        source_file=" + ".join(x.source_file for x in ships),
        customer=base.customer, template=base.template, ref=base.ref,
        channel=base.channel, warehouse=base.warehouse,
        battery="是" if any(x.battery == "是" for x in ships) else "否",
        magnet="是" if any(x.magnet == "是" for x in ships) else "否",
        customs=base.customs, receiver=base.receiver, cartons=cartons,
        warnings=warnings,
    )
    if len({x.channel for x in ships}) > 1:
        merged.warnings.append("同一 DPL 出现多个渠道，取第一个文件")
    if len({x.warehouse for x in ships if x.warehouse}) > 1:
        merged.warnings.append("同一 DPL 出现多个仓码，取第一个文件")
    merged.empty_as_slash = True
    merged.total_boxes = int(sum(group_boxes.values()))
    merged.total_weight = None
    merged.total_value = None
    return merged


class DepponPlugin:
    name = "德邦"

    def discover(self, ctx: RunContext) -> list[Path]:
        if not DEPPON_DIR.exists():
            return []
        return [p for p in sorted(DEPPON_DIR.rglob("*.xls*"))
                if not p.name.startswith("~$") and "我司" not in p.name]

    def process_all(self, files: list[Path], ctx: RunContext) -> list[Generated]:
        handovers = load_handover_index([p for p in files if is_handover(p)])
        items: list[Generated] = []
        groups: dict[str, list[Shipment]] = {}

        for path in files:
            if is_handover(path):
                continue
            progress = getattr(ctx, "progress", None)
            if progress is not None:
                progress.advance(1)
            try:
                ref = resolve_ref(path, handovers)
                if not ref:
                    items.append(Generated(
                        file=f"{self.name}/{path.name}",
                        problems=["未找到客户参考号（文件名/交接单都没有 DPL 单号），已跳过"]))
                    continue
                ship = parse_deppon(path, ref)
            except Exception as exc:  # noqa: BLE001
                ctx.log.error("  [失败] %s: %s", path.name, exc)
                items.append(Generated(file=f"{self.name}/{path.name}",
                                       problems=[str(exc)]))
                continue
            if ship is None:
                items.append(Generated(file=f"{self.name}/{path.name}",
                                       problems=["未识别模板或没有数据"]))
                continue
            groups.setdefault(ref, []).append(ship)

        for ref, ships in groups.items():
            merged = merge_shipments(ships)
            # 重跑前清掉同一参考号的旧输出，避免上一次残留的文件混进交付
            for stale in (ctx.out_root / self.name).glob(f"{ref}_*.xlsx"):
                try:
                    stale.unlink()
                except OSError:
                    pass
            if len(ships) > 1:
                ctx.log.info("  [合并] %s ← %d 个文件", ref, len(ships))
            expect = handovers.expect_for(ref)
            if expect:
                if expect.channel:
                    merged.channel = normalize_channel(expect.channel)
                if expect.warehouse and not merged.warehouse:
                    merged.warehouse = _warehouse_from(expect.warehouse)
            if expect and expect.boxes and merged.total_boxes != expect.boxes:
                merged.warnings.append(
                    f"交接单箱数 {expect.boxes} ≠ 解析箱数 {merged.total_boxes}")
            items.append(emit_shipment(merged, ctx))
        return items

    def process(self, path: Path, ctx: RunContext) -> list[Generated]:
        return self.process_all([path], ctx)


register(DepponPlugin())
