# -*- coding: utf-8 -*-
"""美国州名/区域、泽汇地址文本解析。"""
from __future__ import annotations

import re

from .strings import is_blank

# 美国州名 -> 州缩写（TikTok/私仓地址常用全称）
US_STATE_ABBR = {
    "alabama": "AL", "alaska": "AK", "arizona": "AZ", "arkansas": "AR",
    "california": "CA", "colorado": "CO", "connecticut": "CT", "delaware": "DE",
    "florida": "FL", "georgia": "GA", "hawaii": "HI", "idaho": "ID",
    "illinois": "IL", "indiana": "IN", "iowa": "IA", "kansas": "KS",
    "kentucky": "KY", "louisiana": "LA", "maine": "ME", "maryland": "MD",
    "massachusetts": "MA", "michigan": "MI", "minnesota": "MN",
    "mississippi": "MS", "missouri": "MO", "montana": "MT", "nebraska": "NE",
    "nevada": "NV", "new hampshire": "NH", "new jersey": "NJ",
    "new mexico": "NM", "new york": "NY", "north carolina": "NC",
    "north dakota": "ND", "ohio": "OH", "oklahoma": "OK", "oregon": "OR",
    "pennsylvania": "PA", "rhode island": "RI", "south carolina": "SC",
    "south dakota": "SD", "tennessee": "TN", "texas": "TX", "utah": "UT",
    "vermont": "VT", "virginia": "VA", "washington": "WA",
    "west virginia": "WV", "wisconsin": "WI", "wyoming": "WY",
    "district of columbia": "DC",
}

EAST_STATES = {
    "CT", "DE", "FL", "GA", "IN", "KY", "MA", "MD", "ME", "MI", "NC", "NH",
    "NJ", "NY", "OH", "PA", "RI", "SC", "TN", "VA", "VT", "WV", "DC", "AL",
    "MS", "LA",
}
WEST_STATES = {"CA", "WA", "OR", "NV", "AZ", "UT", "CO", "ID", "MT", "WY", "NM"}


def region_by_state(state: str) -> str:
    st = (state or "").strip().upper()[:2]
    if st in EAST_STATES:
        return "美东"
    if st in WEST_STATES:
        return "美西"
    return "美中"


def get_us_region_by_zip(zip_code) -> str:
    if is_blank(zip_code):
        return ""
    m = re.search(r"(\d)", str(zip_code).strip())
    if not m:
        return ""
    first = int(m.group(1))
    if 0 <= first <= 3:
        return "美东"
    if 4 <= first <= 7:
        return "美中"
    if 8 <= first <= 9:
        return "美西"
    return ""


def parse_zehui_address(addr1) -> dict:
    """解析泽汇收件地址，兼容两种来源格式：

    1) FBA 仓：
       RMN3 - 220 Centreport Parkway 22406 - FREDERICKSBURG, VA - United States
    2) TikTok / 私仓（逗号分隔）：
       619 E SAM HOUSTON PKWY S, Pasadena, Texas, 77503, US, ...
    """
    result = {"仓码": "", "城市": "", "省州": "", "邮编": "", "国家原文": ""}
    if is_blank(addr1):
        return result
    text = str(addr1).strip()

    # ---- 格式 1：带 " - " 的 FBA 仓地址 ----
    if " - " in text:
        parts = [p.strip() for p in re.split(r"\s+-\s*|\s*-\s+", text) if p.strip()]
        if len(parts) >= 3:
            result["仓码"] = parts[0]
            m = re.search(r"(\d{5}(?:-\d{4})?)\s*$", parts[1])
            if m:
                result["邮编"] = m.group(1)
            city_state = [x.strip() for x in re.split(r"[,，]\s*", parts[2]) if x.strip()]
            if city_state:
                result["城市"] = city_state[0]
            if len(city_state) > 1:
                state = city_state[1]
                result["省州"] = US_STATE_ABBR.get(state.lower(), state)
            if len(parts) > 3:
                result["国家原文"] = parts[3]
            return result

    # ---- 格式 2：逗号分隔的 TikTok / 私仓地址 ----
    parts = [p.strip() for p in re.split(r"[,，]\s*", text) if p.strip()]
    zip_code = ""
    zip_idx = None
    for i, part in enumerate(parts):
        if i == 0:  # 街道门牌里也可能有 5 位数字，跳过第一段
            continue
        m = re.search(r"\b(\d{5}(?:-\d{4})?)\b", part)
        if m:
            zip_code = m.group(1)
            zip_idx = i
            break

    if zip_code:
        result["邮编"] = zip_code
        seg = parts[zip_idx]
        m = re.match(r"^([A-Za-z]{2})\s+" + re.escape(zip_code) + r"$", seg.strip())
        if m:
            # "Ontario, CA 91762"
            result["省州"] = m.group(1).upper()
            if zip_idx >= 1:
                result["城市"] = parts[zip_idx - 1]
        else:
            # "Pasadena, Texas, 77503"
            if zip_idx >= 1:
                state = parts[zip_idx - 1]
                result["省州"] = US_STATE_ABBR.get(state.lower(), state)
            if zip_idx >= 2:
                result["城市"] = parts[zip_idx - 2]
    elif len(parts) >= 3:
        # 兜底：取最后几段
        result["城市"] = parts[-3]
        state = parts[-2]
        result["省州"] = US_STATE_ABBR.get(state.lower(), state)

    return result
