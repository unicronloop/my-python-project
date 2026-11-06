# -*- coding: utf-8 -*-
"""校验报告 / 待补充类目清单。"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

from .plugin import Generated

REPORT_NAME = "校验报告.txt"
UNMAPPED_NAME = "类目英文待补充.txt"


def write_report(items: list[Generated], out_dir: Path, run_date: str = "") -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / REPORT_NAME
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines = [f"生成时间：{now}", ""]
    bad = [i for i in items if (not i.ok) or i.problems or i.warnings]
    for item in bad:
        lines.append(_line(item))
        for p in item.problems:
            lines.append(f"     - {p}")
        for w in item.warnings:
            lines.append(f"     * {w}")
    if not bad:
        lines.append("未发现异常。")
    path.write_text("\n".join(lines), encoding="utf-8")
    return path


def _line(item: Generated) -> str:
    flag = "OK " if item.ok else "!! "
    st = item.stat
    if st:
        return (f"{flag}{item.file}  行={st['rows']} 箱={st['boxes']} 件={st['qty']:g} "
                f"实重={st['weight']} 方数={st['cbm']} 申报值={st['value']}")
    return f"{flag}{item.file}  已生成"


def write_unmapped(unmapped: dict[str, int], out_dir: Path) -> Path | None:
    """创颖峻末级类目缺少英文对照时，输出一份待补清单（供业务同事填）。"""
    if not unmapped:
        return None
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / UNMAPPED_NAME
    rows = sorted(unmapped.items(), key=lambda kv: (-kv[1], kv[0]))
    lines = [
        "以下创颖峻末级类目还没有英文对照（输出里优先用 SKU 的 productEn，取不到写 /）。",
        "请在第三列补英文，之后加进 src/fapiao/adapters/cyj.py 的 CATEGORY_EN 即可。",
        "",
        "中文末级类目\t箱数\t英文（请补充）",
    ]
    lines += [f"{cn}\t{cnt}\t" for cn, cnt in rows]
    path.write_text("\n".join(lines), encoding="utf-8")
    return path
