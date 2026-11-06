# -*- coding: utf-8 -*-
"""创颖峻适配器：客户发票.xlsx（汇总 + 分箱 sheet）→ Shipment。

拆单键：交接单号 + RTD预发货单号 + FBA仓码，明细行数必须等于汇总箱数。
"""
from __future__ import annotations

import logging
import re
from collections import Counter
from pathlib import Path

import openpyxl

from ..common.strings import num, s, split_dims
from ..config import CHANNEL_MAP, CUSTOMS_MAP, CYJ_DIR
from ..models.address import parse_address_block, receiver_from_book
from ..models.shipment import Carton, Shipment
from ..registry import register
from .base import Generated, RunContext, emit_shipment

log = logging.getLogger(__name__)

# 创颖峻：单价固定 5 USD（旧样例 5/5 张一致，sum(Q)=5*sum(N) 精确成立）
CYJ_UNIT_VALUE = 5.0
CYJ_CURRENCY = "USD"
CYJ_DEFAULT_CUSTOMS = "不需报关"

# 按「货物类型」取整单常量（实证：纺织 RTD → 服装/棉/日常穿着；普货 RTD → 陶瓷/生活用品）
CYJ_CARGO_RULES: dict[str, dict] = {
    "纺织": {
        "name_cn": "服装", "name_en": "clothes",
        "hs_code": "6104420000",
        "mat_cn": "棉", "mat_en": "Cotton",
        "use_cn": "日常穿着", "use_en": "Daily wear",
    },
    "普货": {
        "name_cn": "", "name_en": "",          # 空 = 按 categoryFullName 末级类目派生
        "hs_code": "/",
        "mat_cn": "陶瓷/不锈钢", "mat_en": "",
        "use_cn": "生活用品", "use_en": "",
    },
}

# 末级类目中/英对照（缺省回退 SKU 的 productEn，再缺省写 "/"）
CATEGORY_EN = {
    "个人护理": "Personal Care",
    "陶瓷陶具": "Ceramic Pottery",
    "服装服饰": "Apparel",
    "肥人装": "Plus Size",
    "西装": "Blazer",
    "外套": "Coats",
    "上衣": "Blouses",
    "家居用品": "Home Supplies",
    "厨房用品": "Kitchen",
    "玩具": "Toys",
    "电子": "Electronics",
    "办公用品": "Office",
}


def _read_huizong(wb) -> list[dict]:
    """读「汇总」sheet，按交接单号向下填充（合计行即停）。"""
    if "汇总" not in wb.sheetnames:
        return []
    ws = wb["汇总"]
    header = {}
    for c in range(1, ws.max_column + 1):
        name = s(ws.cell(2, c).value)
        if name:
            header[name] = c

    rows = []
    ccy: dict = {}
    for r in range(3, ws.max_row + 1):
        rec = {name: ws.cell(r, c).value for name, c in header.items()}
        if s(rec.get("交接单号")).startswith("合计"):
            break
        cur = s(rec.get("交接单号"))
        if cur:
            ccy = {"交接单号": cur, "地址": rec.get("地址"), "物流方式": rec.get("物流方式"),
                   "物流渠道": rec.get("物流渠道"), "报关": rec.get("报关")}
        if not s(rec.get("预发货单号")):
            continue
        row = dict(rec)
        row["交接单号"] = ccy.get("交接单号", "")
        row["地址"] = ccy.get("地址")
        row["物流方式"] = ccy.get("物流方式")
        row["物流渠道"] = ccy.get("物流渠道")
        row["报关"] = ccy.get("报关")
        rows.append(row)
    return rows


def _parse_cyj_sheet(ws) -> tuple[list[dict], dict[str, list[dict]]]:
    """返回 (箱级行, {箱号: [明细行]})。"""
    item_hdr_row = None
    for r in range(1, min(ws.max_row, 200) + 1):
        if s(ws.cell(r, 1).value) == "出货单号":
            item_hdr_row = r
            break
    end = item_hdr_row or (ws.max_row + 1)

    cartons, cur_rtd = [], ""
    for r in range(2, end):
        rtd = s(ws.cell(r, 1).value)
        if rtd:
            cur_rtd = rtd
        box = s(ws.cell(r, 4).value)
        if not box:
            continue
        cartons.append({
            "rtd": cur_rtd,
            "box": box,
            "seq": s(ws.cell(r, 5).value),
            "dims": ws.cell(r, 6).value,
            "cargo_type": s(ws.cell(r, 8).value),
            "weight": num(ws.cell(r, 9).value),
            "qty": num(ws.cell(r, 11).value),
        })

    items: dict[str, list[dict]] = {}
    if item_hdr_row:
        hdr = {}
        for c in range(1, ws.max_column + 1):
            name = s(ws.cell(item_hdr_row, c).value)
            if name and name not in hdr:
                hdr[name] = c
        for r in range(item_hdr_row + 1, ws.max_row + 1):
            box = s(ws.cell(r, hdr.get("箱号", 2)).value)
            sku = s(ws.cell(r, hdr.get("skuId", 19)).value)
            title = s(ws.cell(r, hdr.get("title", 11)).value)
            # 显式括号：既没有箱号、也没有 SKU/标题，才算空行
            if not box or (sku == "" and title == ""):
                continue
            items.setdefault(box, []).append({
                "category": s(ws.cell(r, hdr.get("categoryFullName", 14)).value),
                "product_cn": s(ws.cell(r, hdr.get("productCn", 15)).value),
                "product_en": s(ws.cell(r, hdr.get("productEn", 16)).value),
                "material": s(ws.cell(r, hdr.get("material", 17)).value),
                "purpose": s(ws.cell(r, hdr.get("purpose", 18)).value),
                "hs": s(ws.cell(r, hdr.get("HSCODE", 27)).value),
                "link": s(ws.cell(r, hdr.get("销售链接", 32)).value),
                "qty": num(ws.cell(r, hdr.get("quantity", 7)).value),
            })
    return cartons, items


def leaf_category(cat: str) -> str:
    """categoryFullName 取末级类目。

    '交通工具配件及附件->车身件->外饰件->车门拉手->车门拉手盖' -> '车门拉手盖'
    上传系统要求「物品中文名」是末级类目名。
    """
    parts = [p.strip() for p in re.split(r"(?:->|→|＞|>)", s(cat)) if p.strip()]
    return parts[-1] if parts else ""


def parse_chuangyingjun(path: Path, unmapped: dict[str, int] | None = None) -> list[Shipment]:
    path = Path(path)
    wb = openpyxl.load_workbook(path, data_only=True)
    ships: list[Shipment] = []
    try:
        sheets = list(wb.sheetnames)
        cache: dict[str, tuple] = {}
        for rec in _read_huizong(wb):
            hz_no = s(rec.get("交接单号"))
            rtd = s(rec.get("预发货单号"))
            wh = s(rec.get("仓库"))
            fba = s(rec.get("Amazon PO"))
            po = s(rec.get("Amazon Reference ID")) or "无"
            boxes = int(num(rec.get("箱数")) or 0)
            qty = num(rec.get("件数"))
            weight = num(rec.get("实重"))
            cbm = num(rec.get("方数"))

            target = next((n for n in sheets if n.startswith(f"{hz_no} {wh}")), None)
            if not target:
                continue
            if target not in cache:
                cache[target] = _parse_cyj_sheet(wb[target])
            cartons_raw, items_map = cache[target]

            rows = [c for c in cartons_raw if c["rtd"] == rtd]
            if not rows:
                continue

            # 该单的货物类型（取箱级众数）
            cargo = Counter([r["cargo_type"] for r in rows if r["cargo_type"]]).most_common(1)
            cargo_type = cargo[0][0] if cargo else ""
            rule = dict(CYJ_CARGO_RULES.get(cargo_type, {}))

            raw_channel = s(rec.get("物流方式"))
            channel = CHANNEL_MAP.get(raw_channel, raw_channel or "美森特快线")

            ship = Shipment(
                source_file=path.name, customer="创颖峻", template="公司模板",
                ref=hz_no, doc_no=rtd, warehouse=wh, channel=channel,
                battery="否", magnet="否",
                customs=CUSTOMS_MAP.get(s(rec.get("报关")), CYJ_DEFAULT_CUSTOMS),
                total_boxes=boxes,
                total_weight=weight, total_cbm=cbm,
                remarks="",
                expect={"boxes": boxes, "qty": qty, "weight": weight, "cbm": cbm},
            )

            rc = receiver_from_book(wh) or parse_address_block(rec.get("地址"), wh)
            ship.receiver = rc

            for idx, row in enumerate(rows, start=1):
                L, W, H = split_dims(row["dims"])
                it_list = items_map.get(row["box"], [])
                hs = rule.get("hs_code") or (
                    it_list[0]["hs"] if it_list and len({i["hs"] for i in it_list}) == 1 else "/")
                hs = s(hs) or "/"
                link = it_list[0]["link"] if it_list else ""

                # 末级类目：一箱可能混多个 SKU，去重后按出现顺序拼 "/"
                leaves: list[str] = []
                for it in it_list:
                    leaf = leaf_category(it["category"])
                    if leaf and leaf not in leaves:
                        leaves.append(leaf)
                if len(leaves) > 1:
                    ship.warnings.append(
                        f"第 {row['seq'] or idx} 箱含多种末级类目：{'/'.join(leaves)}")
                for leaf in leaves:
                    if leaf not in CATEGORY_EN and unmapped is not None:
                        unmapped[leaf] = unmapped.get(leaf, 0) + 1

                leaf_cn = "/".join(leaves[:3]) + ("等" if len(leaves) > 3 else "")
                # 英文名：末级类目有对照就用对照，否则退回该末级下 SKU 的 productEn
                leaf_en_parts: list[str] = []
                for leaf in leaves[:3]:
                    en = CATEGORY_EN.get(leaf, "")
                    if not en:
                        for it in it_list:
                            if it["product_en"] and leaf_category(it["category"]) == leaf:
                                en = it["product_en"]
                                break
                    if en:
                        leaf_en_parts.append(en)
                leaf_en = " / ".join(leaf_en_parts)
                name_cn = leaf_cn or rule.get("name_cn") or "/"
                name_en = leaf_en or ("" if leaves else (rule.get("name_en") or "/")) or "/"

                ship.cartons.append(Carton(
                    seq=row["seq"] or str(idx), boxes=1,
                    battery="否", magnet="否",
                    fba_id=fba or "",
                    po=po or "无",
                    name_cn=name_cn,
                    name_en=name_en,
                    hs_code=hs,
                    weight=row["weight"],
                    length=L, width=W, height=H,
                    qty=row["qty"],
                    unit=CYJ_UNIT_VALUE,
                    currency=CYJ_CURRENCY,
                    total=round(CYJ_UNIT_VALUE * (row["qty"] or 0), 4),
                    brand="", model="",
                    mat_cn=rule.get("mat_cn") or "",
                    mat_en=rule.get("mat_en") or "",
                    use_cn=rule.get("use_cn") or "",
                    use_en=rule.get("use_en") or "",
                    image="", link=link or "",
                ))

            if not ship.channel:
                ship.channel = "美森特快线"
            ships.append(ship)
    finally:
        wb.close()
    return ships


class ChuangyingjunPlugin:
    name = "创颖峻"

    def discover(self, ctx: RunContext) -> list[Path]:
        if not CYJ_DIR.exists():
            return []
        files = [p for p in sorted(CYJ_DIR.glob("客户发票*.xlsx")) if "我司" not in p.name]
        if not files:
            files = [p for p in sorted(CYJ_DIR.glob("*.xlsx")) if "我司" not in p.name]
        return files

    def process(self, path: Path, ctx: RunContext) -> list[Generated]:
        try:
            ships = parse_chuangyingjun(path, ctx.unmapped)
        except Exception as exc:  # noqa: BLE001
            ctx.log.error("  [失败] %s: %s", path.name, exc)
            return [Generated(file=f"{self.name}/{path.name}", problems=[str(exc)])]
        if not ships:
            return [Generated(file=f"{self.name}/{path.name}",
                              problems=["未识别模板或没有数据"])]
        return [emit_shipment(ship, ctx) for ship in ships]


register(ChuangyingjunPlugin())
