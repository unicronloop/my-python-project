#!/usr/bin/env python3
"""
用真实样例数据跑一遍便携包，确认编译后的程序不是"能启动但跑不了"。

做的事：
1. 把 _build/portable 复制到 _build/e2e
2. 从项目里的 泽汇/、时易中/、泽汇-提货表/ 各取几个真实文件放进去，
   再塞一份含「报关退税」的历史样表（验证 合并报关 分支）
3. 以双击等价方式启动 客户Excel处理工具.exe
4. 校验：退出码、无 Traceback/ERR、输出能打开、图片确实嵌进去了、
   报关类型按映射写对（买单报关→不需报关 / 报关退税→合并报关）
5. 再复制到带中文和空格的目录跑一次，确认路径解析没问题

用法：python _build/e2e_check.py [--keep]
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
import sys
import time
import zipfile
from pathlib import Path

import openpyxl

PROJECT_ROOT = Path(__file__).resolve().parent.parent
PORTABLE = PROJECT_ROOT / "_build" / "portable"
WORK = PROJECT_ROOT / "_build" / "e2e"
UNICODE_WORK = PROJECT_ROOT / "_build" / "e2e 中文 空格"
LAUNCHER_EXE = "客户Excel处理工具.exe"
OUTPUT_DIR = "生成的文件"

SAMPLES = {
    "泽汇": 3,
    "泽汇-提货表": 1,
    "时易中": 2,
}

# 历史样表目录：里面有「报关退税」的单子，用来验证合并报关分支
DECL_SOURCE_DIR = PROJECT_ROOT / "999不用到的材料" / "泽汇-材料"

# 报关类型：源表写法 -> 期望写进产出表的值
EXPECTED_DECLARATION = {
    "买单报关": "不需报关",
    "买关报单": "不需报关",
    "报关退税": "合并报关",
}


def log(msg: str = "") -> None:
    try:
        print(msg, flush=True)
    except UnicodeEncodeError:
        print(msg.encode("ascii", "backslashreplace").decode("ascii"), flush=True)


def norm(value) -> str:
    if value is None:
        return ""
    return re.sub(r"[\s:：()（）*]", "", str(value)).lower()


def base_env() -> dict:
    env = {k: v for k, v in os.environ.items() if k not in ("FAPIAO_ROOT_DIR", "APP_ROOT_DIR")}
    env["FAPIAO_NO_PAUSE"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    return env


def run_launcher(root: Path) -> tuple[int, str]:
    proc = subprocess.run(
        [str(root / LAUNCHER_EXE)], cwd=str(root), env=base_env(),
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
    )
    return proc.returncode, proc.stdout.decode("utf-8", "replace")


def read_cells(path: Path) -> tuple[object, object]:
    """读源表的 客户参考号1(B3) 和 报关类型值(B11)。"""
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        ws = wb.active
        return ws.cell(3, 2).value, ws.cell(11, 2).value
    finally:
        wb.close()


def prepare() -> None:
    if WORK.exists():
        shutil.rmtree(WORK)
    log(f"复制便携包 → {WORK.relative_to(PROJECT_ROOT)}")
    shutil.copytree(PORTABLE, WORK)

    for name, limit in SAMPLES.items():
        dst_dir = WORK / name
        dst_dir.mkdir(parents=True, exist_ok=True)
        picked = 0
        for f in sorted((PROJECT_ROOT / name).iterdir()):
            if picked >= limit:
                break
            if f.is_file() and f.suffix.lower() in (".xlsx", ".xls", ".xlsm") and not f.name.startswith("~$"):
                shutil.copy2(f, dst_dir / f.name)
                picked += 1
                log(f"  {name}/ ← {f.name}")
        if picked == 0:
            raise SystemExit(f"样例数据不足：{PROJECT_ROOT / name} 里没有 Excel 文件")

    add_declaration_sample()


def add_declaration_sample() -> None:
    """再放一份「报关退税」的历史单子，专门验证 合并报关 分支。"""
    if not DECL_SOURCE_DIR.is_dir():
        log(f"  [跳过] 没有 {DECL_SOURCE_DIR.name}，无法验证 报关退税")
        return
    for f in sorted(DECL_SOURCE_DIR.glob("*.xlsx")):
        if f.name.startswith("~$"):
            continue
        try:
            _ref, raw = read_cells(f)
        except Exception:  # noqa: BLE001
            continue
        if norm(raw) == "报关退税":
            shutil.copy2(f, WORK / "泽汇" / f.name)
            log(f"  泽汇/ ← {f.name}（报关退税，验证合并报关）")
            return
    log("  [跳过] 历史材料里没有找到报关退税样表")


def media_count(path: Path) -> int:
    """1.py 是直接把图片字节写进 xlsx 包的，openpyxl 的 ws._images 看不到，得数 zip。"""
    try:
        with zipfile.ZipFile(path) as z:
            return sum(1 for n in z.namelist() if n.startswith("xl/media/"))
    except (OSError, zipfile.BadZipFile):
        return 0


def check_workbooks(problems: list[str]) -> None:
    out_root = WORK / OUTPUT_DIR
    for kind, sheet_name in (("泽汇", None), ("时易中", "出货发票")):
        out_dir = out_root / kind
        files = sorted(out_dir.glob("*.xlsx")) if out_dir.is_dir() else []
        if not files:
            problems.append(f"{kind} 没有产出文件：{out_dir}")
            continue
        log(f"\n{kind} 产出 {len(files)} 个文件：")
        for f in files:
            try:
                wb = openpyxl.load_workbook(f)
            except Exception as exc:  # noqa: BLE001
                problems.append(f"{f.name} 打不开：{exc}")
                continue
            sheet = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active
            images = media_count(f)
            log(f"  {f.name}  {f.stat().st_size / 1024:.0f} KB  sheets={wb.sheetnames}  内嵌图片={images}")
            if sheet.max_row < 2 or sheet.max_column < 2:
                problems.append(f"{f.name} 内容看起来是空的")
            if sheet_name and not sheet.cell(4, 2).value:
                problems.append(f"{f.name} 出货发票 B4（FBA 号）为空")
            if images == 0:
                problems.append(f"{f.name} 没有内嵌图片（源数据里是有产品图的）")
            wb.close()


def check_declaration(problems: list[str]) -> None:
    """核对每个产出文件的报关类型是否等于源表映射后的期望值。"""
    sample_dir = WORK / "泽汇"
    expected: dict[str, str] = {}
    for src in sorted(sample_dir.iterdir()):
        if not src.is_file() or src.suffix.lower() != ".xlsx":
            continue
        try:
            ref, raw = read_cells(src)
        except Exception:  # noqa: BLE001
            continue
        if ref is None:
            continue
        key = re.sub(r'[\\/:*?"<>|]', "_", str(ref).strip())
        expected[key] = EXPECTED_DECLARATION.get(norm(raw), "不需报关")

    outs = sorted((WORK / OUTPUT_DIR / "泽汇").glob("*.xlsx"))
    tally: dict[str, int] = {}
    checked = 0
    for out in outs:
        match = next((k for k in expected if out.name.startswith(k + "-")), None)
        if match is None:
            continue
        want = expected[match]
        wb = openpyxl.load_workbook(out, read_only=True)
        got = wb.active.cell(11, 2).value
        wb.close()
        checked += 1
        tally[str(got)] = tally.get(str(got), 0) + 1
        if got != want:
            problems.append(f"{out.name} 报关类型应为「{want}」，实际「{got}」")
    log(f"\n报关类型核查：{checked} 份，结果 {tally}")
    if checked == 0:
        problems.append("报关类型核查没有匹配到任何产出文件")
    if "合并报关" not in tally:
        problems.append("没有验证到「报关退税 → 合并报关」这一分支")
    if "不需报关" not in tally:
        problems.append("没有验证到「买单报关 → 不需报关」这一分支")


def check_unicode_path(problems: list[str]) -> None:
    log(f"\n中文/空格路径验证：{UNICODE_WORK.name}")
    if UNICODE_WORK.exists():
        shutil.rmtree(UNICODE_WORK)
    shutil.copytree(PORTABLE, UNICODE_WORK)
    try:
        code, text = run_launcher(UNICODE_WORK)
        expected = f"程序目录：{UNICODE_WORK.resolve()}"
        if code != 0:
            problems.append(f"中文路径下退出码 {code}")
        if expected not in text:
            problems.append(f"中文路径下 BASE_DIR 不对，期望包含：{expected}")
        else:
            log("  ✓ 启动器与主程序在中文/空格路径下工作正常")
    finally:
        shutil.rmtree(UNICODE_WORK, ignore_errors=True)


def main() -> None:
    parser = argparse.ArgumentParser(description="便携包端到端验证")
    parser.add_argument("--keep", action="store_true", help="保留 _build/e2e 便于人工检查")
    args = parser.parse_args()

    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, OSError, ValueError):
            pass

    if not (PORTABLE / LAUNCHER_EXE).is_file():
        raise SystemExit("还没有便携包，请先运行 _build/build_portable.py")

    problems: list[str] = []
    prepare()

    started = time.time()
    exit_code, text = run_launcher(WORK)
    log(f"\n程序运行 {time.time() - started:.1f}s，退出码 {exit_code}")
    log("-" * 64)
    log(text.rstrip())
    log("-" * 64)

    if exit_code != 0:
        problems.append(f"退出码非 0：{exit_code}")
    for bad in ("Traceback", "ERR", "PermissionError", "ModuleNotFoundError", "ImportError"):
        if bad in text:
            problems.append(f"输出里出现「{bad}」")

    check_workbooks(problems)
    check_declaration(problems)
    check_unicode_path(problems)

    if problems:
        log("\n[核查未通过]")
        for p in problems:
            log(f"  - {p}")
        raise SystemExit(1)

    log("\n[核查通过] 便携包能正确处理 xlsx / xls（xlrd）、模板、图片、报关类型，且支持中文路径")
    if not args.keep:
        shutil.rmtree(WORK, ignore_errors=True)


if __name__ == "__main__":
    main()
