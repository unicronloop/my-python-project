#!/usr/bin/env python3
"""
泽汇「报关类型」映射自检。

拿 泽汇/ 里第一份真实样表做底子，把 B11（报关类型值）改成各种写法，
跑一遍 process_zehui，再核对输出表里的报关类型是否符合预期。

以后加映射规则（新的源表写法 / 新客户模板）先改 1.py 的
ZEHUI_DECLARATION_RULES，再跑这个脚本，避免"改了映射但没生效"。

用法：python _build/check_declaration_map.py
"""

from __future__ import annotations

import importlib.util
import io
import shutil
import sys
import tempfile
from contextlib import redirect_stdout
from pathlib import Path

import openpyxl

PROJECT_ROOT = Path(__file__).resolve().parent.parent
MAIN_SCRIPT = PROJECT_ROOT / "1.py"
SAMPLE_DIR = PROJECT_ROOT / "泽汇"

# (B11 里写的原始值, 期望输出, 是否应该打印"未识别"提示)
CASES = [
    ("买单报关", "不需报关", False),
    ("买关报单", "不需报关", False),
    ("报关退税", "合并报关", False),
    ("报关 退税", "合并报关", False),      # 带空格也要能认
    ("买 单 报 关", "不需报关", False),    # 带空格也要能认
    ("某某新写法", "不需报关", True),      # 没见过的值：兜底 + 提示
    ("", "不需报关", False),               # 空值：兜底，不吵
]


def load_main_module():
    spec = importlib.util.spec_from_file_location("fapiao_main", MAIN_SCRIPT)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def pick_sample() -> Path:
    files = sorted(
        f for f in SAMPLE_DIR.iterdir()
        if f.is_file() and f.suffix.lower() == ".xlsx" and not f.name.startswith("~$")
    )
    if not files:
        raise SystemExit(f"没有样表可用：{SAMPLE_DIR}")
    return files[0]


def main() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, OSError, ValueError):
            pass

    module = load_main_module()
    sample = pick_sample()
    print(f"样表：{sample.name}")

    work = Path(tempfile.mkdtemp(prefix="decl_check_"))
    failures: list[str] = []
    try:
        for raw, expected, expect_warning in CASES:
            src = work / "input.xlsx"
            shutil.copy2(sample, src)

            wb = openpyxl.load_workbook(src)
            wb.active.cell(11, 2).value = raw if raw else None
            wb.save(src)
            wb.close()

            out = work / "output.xlsx"
            buf = io.StringIO()
            with redirect_stdout(buf):
                module.process_zehui(src, out, pickup_map=None)
            log_text = buf.getvalue()

            wb = openpyxl.load_workbook(out)
            ws = wb.active
            got = ws.cell(11, 2).value
            label = ws.cell(11, 1).value
            wb.close()

            warned = "未识别的报关类型" in log_text
            ok = got == expected and warned == expect_warning
            if not ok:
                failures.append(
                    f"输入 {raw!r}: 期望 {expected!r} / 提示={expect_warning}，"
                    f"实际 {got!r} / 提示={warned}（第 11 行标签={label!r}）"
                )
            print(
                f"  {'✓' if ok else '✗'} 输入 {raw!r:<12} → {got!r:<8}"
                f"{'（已提示）' if warned else ''}"
            )
    finally:
        shutil.rmtree(work, ignore_errors=True)

    if failures:
        print("\n[自检未通过]")
        for f in failures:
            print(f"  - {f}")
        raise SystemExit(1)
    print("\n[自检通过] 报关类型映射与兜底提示都符合预期")


if __name__ == "__main__":
    main()
