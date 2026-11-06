# -*- coding: utf-8 -*-
"""Canonical 中间模型：一票货 = Shipment，一箱 = Carton。

约定（全客户统一，避免各 parser 口径不一）：
- Carton.boxes   该行代表的箱数（如 "1-3" → 3）
- Carton.weight  单箱实重（kg）；总重 = weight * boxes
- Carton.qty     单箱数量（pcs）
- Carton.total   该行申报总值；缺省时按 unit * qty 计算
"""
from __future__ import annotations

from dataclasses import dataclass, field

from ..common.strings import safe_name


@dataclass
class Carton:
    seq: str = ""                  # 箱序号（可写 "1-3"）
    boxes: int = 1                 # 箱数/CTN
    battery: str = "否"
    magnet: str = "否"
    fba_id: str = ""
    po: str = ""
    name_cn: str = ""
    name_en: str = ""
    hs_code: str = ""
    weight: float | None = None    # 单箱实重 / KG
    length: float | None = None
    width: float | None = None
    height: float | None = None
    qty: float | None = None       # 单箱数量 / PCS
    unit: float | None = None      # 申报单价
    currency: str = "USD"
    total: float | None = None     # 该行申报总值
    brand: str = ""
    model: str = ""
    mat_cn: str = ""
    mat_en: str = ""
    use_cn: str = ""
    use_en: str = ""
    net_g: float | None = None
    gross_g: float | None = None
    image: str = ""
    link: str = ""
    src_row: int = 0

    @property
    def total_value(self) -> float:
        if self.total is not None:
            return float(self.total)
        return float((self.unit or 0) * (self.qty or 0))


@dataclass
class Shipment:
    source_file: str = ""
    customer: str = ""
    template: str = ""
    ref: str = ""                  # 客户参考号1
    doc_no: str = ""               # 单号区分（创颖峻=RTD，德邦=FBA）
    channel: str = ""
    warehouse: str = ""
    battery: str = "否"
    magnet: str = "否"
    customs: str = "不需报关"
    total_boxes: int = 0
    total_weight: float | None = None
    total_cbm: float | None = None
    total_value: float | None = None
    remarks: str = ""
    receiver: dict = field(default_factory=dict)
    cartons: list[Carton] = field(default_factory=list)
    # 勾稽基准（来自客户汇总/合计行）
    expect: dict = field(default_factory=dict)
    empty_as_slash: bool = False    # 空文本列写 "/"（德邦约定）
    # 个别客户对输出还有额外约定
    out_name: str = ""              # 自定义输出文件名（不含后缀），空则用 out_stem
    invoice_sheet: str = ""         # 目标 sheet 名（默认「出货发票」）
    delivery_window: object = None  # G14 亚马逊配送时段
    warnings: list[str] = field(default_factory=list)

    @property
    def out_stem(self) -> str:
        if self.out_name:
            return safe_name(self.out_name)
        parts = [self.ref or "无参考号", self.warehouse or "无仓码",
                 self.channel or "无渠道", f"{self.total_boxes}箱"]
        if self.doc_no:                     # 同一交接单号多张发票：用 RTD / FBA 区分
            parts.insert(2, self.doc_no)
        return safe_name("_".join(parts))
