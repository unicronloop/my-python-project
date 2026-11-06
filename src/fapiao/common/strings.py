# -*- coding: utf-8 -*-
"""纯字符串/数值小工具：无副作用，便于单测。"""
from __future__ import annotations

import re
import time
from pathlib import Path


def is_blank(v) -> bool:
    return v is None or str(v).strip() == "" or str(v).strip().lower() == "nan"


def s(v) -> str:
    """安全转字符串（去空白、None → ''，整数型 float 去掉 .0）"""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def num(v):
    """转数字；失败返回 None"""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    txt = s(v).replace(",", "").replace("，", "")
    m = re.search(r"-?\d+(?:\.\d+)?", txt)
    return float(m.group()) if m else None


def rnd(v):
    return None if v is None else round(float(v), 4)


def normalize_label(v) -> str:
    return re.sub(r"[\s:：()（）*]", "", str(v)).lower()


def split_dims(raw) -> tuple[float | None, float | None, float | None]:
    """'58*47*34' / '36*58*32cm' / '58×47×34' → (L, W, H)"""
    if raw is None or raw == "":
        return None, None, None
    parts = [p for p in re.split(r"[*xX×/／,，\s]+", s(raw)) if re.search(r"\d", p)]
    vals = [num(p) for p in parts[:3]]
    while len(vals) < 3:
        vals.append(None)
    return vals[0], vals[1], vals[2]


_CJK = re.compile(r"[一-鿿]")


def cn_en(raw) -> tuple[str, str]:
    """中英混填的列：`中文/english` 拆成两列；含中文 → 中文列，否则 → 英文列。"""
    txt = s(raw)
    if not txt:
        return "", ""
    parts = [x.strip() for x in re.split(r"[/／|｜\r\n]+", txt) if x.strip()]
    if len(parts) >= 2:
        cn = next((x for x in parts if _CJK.search(x)), "")
        en = next((x for x in parts if x and not _CJK.search(x)), "")
        if cn or en:
            return cn, en
    return (txt, "") if _CJK.search(txt) else ("", txt)


def safe_name(name: str) -> str:
    name = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", name).strip(" .")
    return name[:120] or "未命名"


def extract_number(v):
    """从字符串/数字中提取数字部分（如 '10.10<br>USD' -> 10.1）。"""
    if v is None:
        return ""
    if isinstance(v, (int, float)):
        return v
    m = re.search(r"(\d+(?:\.\d+)?)", str(v))
    return float(m.group(1)) if m else ""


def judge_battery_magnetic(*notes):
    """从"属性备注/备注"的组合文本判断是否带电/带磁。"""
    text = " ".join(str(n).strip() for n in notes if not is_blank(n))
    if not text:
        return "是", "是"
    has_battery = any(k in text for k in ("带电", "电池", "锂电", "蓝牙"))
    has_magnetic = any(k in text for k in ("带磁", "磁"))
    if not has_battery and not has_magnetic:
        return "否", "否"
    return ("是" if has_battery else "否"), ("是" if has_magnetic else "否")


def build_output_name(customer_ref, run_date: str) -> str:
    """输出文件名：客户参考号-生成日期(yyyymmdd)-已处理.xlsx"""
    ref = "" if is_blank(customer_ref) else str(customer_ref).strip()
    if not ref:
        ref = "未命名"
    ref = re.sub(r'[\\/:*?"<>|]', "_", ref)
    return f"{ref}-{run_date}-已处理.xlsx"


def replace_with_retry(src, dst, attempts: int = 6, delay: float = 0.25) -> None:
    """Windows 下文件可能被杀软/索引服务短暂占用，os.replace 偶发 WinError 5。"""
    last_error = None
    for i in range(attempts):
        try:
            Path(src).replace(dst)
            return
        except OSError as e:
            last_error = e
            time.sleep(delay * (i + 1))
    raise last_error


def cleanup_stale_temp_files(directory: Path) -> None:
    """清理上次运行残留的临时文件，避免越积越多。"""
    if not directory.exists():
        return
    for pattern in (".__tmp__*", "*.dedupe.xlsx"):
        for p in directory.glob(pattern):
            try:
                p.unlink()
            except OSError:
                pass


def list_excel_files(directory: Path):
    """列出目录下待处理的 Excel（跳过 Office 临时锁文件）。"""
    if not directory.exists():
        return []
    return [
        f for f in sorted(directory.iterdir())
        if f.is_file()
        and f.suffix.lower() in (".xlsx", ".xls", ".xlsm")
        and not f.name.startswith("~$")
    ]


def iter_progress(items, desc: str = "", bar_width: int = 24, log=print):
    """轻量进度显示（不引第三方 progressbar）。"""
    items = list(items)
    total = len(items)
    if total == 0:
        return
    for i, item in enumerate(items, 1):
        name = getattr(item, "name", "")
        filled = int(bar_width * i / total)
        bar = "#" * filled + "-" * (bar_width - filled)
        line = f"  [{desc}] |{bar}| {i}/{total}"
        if name:
            line += f"  {name}"
        print("\r" + line[:110].ljust(110), end="", flush=True)
        yield item
    print()


def find_value_by_label(df, label, max_search_rows: int = 25, prefer_col=None) -> str:
    """在 DataFrame 里按标签找右侧第一个非空值（时易中表头用）。"""
    target = normalize_label(label)
    if prefer_col is not None and prefer_col < len(df.columns):
        for i in range(min(max_search_rows, len(df))):
            cell = df.iloc[i, prefer_col]
            if is_blank(cell):
                continue
            if normalize_label(cell) == target:
                for k in range(prefer_col + 1, len(df.columns)):
                    v = df.iloc[i, k]
                    if not is_blank(v):
                        return str(v).strip()
                return ""
    for i in range(min(max_search_rows, len(df))):
        for j in range(len(df.columns)):
            if prefer_col is not None and j == prefer_col:
                continue
            cell = df.iloc[i, j]
            if is_blank(cell):
                continue
            if normalize_label(cell) == target:
                for k in range(j + 1, len(df.columns)):
                    v = df.iloc[i, k]
                    if not is_blank(v):
                        return str(v).strip()
                return ""
    return ""
