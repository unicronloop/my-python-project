# -*- coding: utf-8 -*-
"""命令行入口：遍历已注册插件 → 处理 → 汇总报告。

客户名只出现在各插件里；这里只认 registry。
"""
from __future__ import annotations

import argparse
import json
import sys
import warnings
from datetime import datetime
from pathlib import Path

from . import adapters  # noqa: F401  import 即注册所有内置插件
from .common import pacing
from .common.progress import Progress
from .common.strings import cleanup_stale_temp_files
from .config import BASE_DIR, OUTPUT_DIR, ensure_workspace
from .log import setup_log
from .plugin import Generated, RunContext
from .registry import CUSTOMERS, names
from .report import write_report, write_unmapped

DEFAULT_REFS = "refs.json"


def _split_only(values: list[str] | None) -> list[str]:
    """--only 泽汇 时易中 / --only "泽汇,时易中" 两种写法都支持。"""
    out: list[str] = []
    for v in values or []:
        out += [x.strip() for x in str(v).replace("，", ",").split(",") if x.strip()]
    return out


def _load_refs(path: Path | None) -> dict:
    if path is None:
        path = BASE_DIR / DEFAULT_REFS
        if not path.exists():
            example = BASE_DIR / "refs.example.json"
            if not example.exists():
                return {}
            path = example
    try:
        data = json.loads(Path(path).read_text(encoding="utf-8"))
        # 以 _ 开头的是说明字段，不参与匹配
        return {k: v for k, v in data.items() if not str(k).startswith("_")}
    except (OSError, json.JSONDecodeError) as exc:
        print(f"[警告] 读取对照表失败：{path}（{exc}）")
        return {}


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="fapiao",
        description="客户 Excel 明细 → 我司出货发票（按客户插件处理）。",
    )
    p.add_argument("--only", nargs="+", metavar="客户",
                   help=f"只处理指定客户，可逗号分隔。可选：{', '.join(names())}")
    p.add_argument("--refs", type=Path, help="客户参考号对照表（JSON），默认 refs.json")
    p.add_argument("--out", type=Path, help=f"输出目录，默认 {OUTPUT_DIR}")
    p.add_argument("--no-report", action="store_true", help="不写 校验报告.txt")
    p.add_argument("--dry-run", action="store_true", help="只解析校验，不写任何文件")
    p.add_argument("--log-level", default="INFO",
                   choices=["DEBUG", "INFO", "WARNING", "ERROR"], help="日志级别")
    p.add_argument("--no-progress", action="store_true", help="不显示进度条")
    p.add_argument("--list", action="store_true", help="列出已注册客户后退出")
    return p


def run(args: argparse.Namespace) -> int:
    # openpyxl 的 Data Validation 提示会插到进度条中间，统一静音
    warnings.filterwarnings("ignore", module="openpyxl")
    log = setup_log(args.log_level)
    if args.list:
        for name in CUSTOMERS:
            print(name)
        return 0

    ensure_workspace()
    out_root = Path(args.out) if args.out else OUTPUT_DIR
    run_date = datetime.now().strftime("%Y%m%d")
    ctx = RunContext(out_root=out_root, refs=_load_refs(args.refs),
                     run_date=run_date, log=log)

    selected = _split_only(args.only)
    unknown = [x for x in selected if x not in CUSTOMERS]
    if unknown:
        print(f"[错误] 未知客户：{', '.join(unknown)}；可选：{', '.join(CUSTOMERS)}")
        return 2
    targets = selected or list(CUSTOMERS)

    log.info("=== 客户 Excel 处理工具 · %s ===", run_date)
    log.info("输出目录：%s", out_root)
    if ctx.refs:
        log.info("参考号对照表：%d 条", len(ctx.refs))

    # 先 discover 一遍，才能给出准确的总量（进度条要从 20% 起跳）
    plan: list[tuple[str, object, list[Path]]] = []
    total_units = 0
    for name in targets:
        plugin = CUSTOMERS[name]
        files = plugin.discover(ctx)
        plan.append((name, plugin, files))
        counter = getattr(plugin, "unit_count", None)
        total_units += counter(files, ctx) if callable(counter) else len(files)

    # 每生成一个文件后的静默停顿（秒数见 config.MIN_SECONDS_PER_FILE）；
    # --dry-run 只为快速自检，不停顿。停顿不打印任何日志。
    pacing.enable(not args.dry_run)

    progress = Progress(total_units, enabled=False if args.no_progress else None)
    progress.silent = bool(args.no_progress)
    ctx.progress = progress
    first = plan[0][0] if plan else ""
    progress.set_label(first)
    progress.advance(0)

    items: list[Generated] = []
    for name, plugin, files in plan:
        counter = getattr(plugin, "unit_count", None)
        units = counter(files, ctx) if callable(counter) else len(files)
        progress.set_label(name)
        log.info("[%s] %d 个文件", name, len(files))
        if not files:
            continue
        out_dir = out_root / name
        if out_dir.exists():
            cleanup_stale_temp_files(out_dir)
        base = progress.done
        if hasattr(plugin, "process_all"):
            # 需要跨文件上下文（交接单、合票）的客户走批量入口
            items += plugin.process_all(files, ctx)  # type: ignore[attr-defined]
        else:
            for path in files:
                items += plugin.process(path, ctx)
                progress.advance(1)
        # 插件内部没有细粒度推进时补齐，保证百分比单调
        progress.advance_to(base + units)
    progress.close()

    ok = sum(1 for i in items if i.ok)
    bad = len(items) - ok
    log.info("完成：%d 个文件，成功 %d，异常 %d", len(items), ok, bad)

    if not args.dry_run and not args.no_report:
        report = write_report(items, out_root)
        log.info("校验报告：%s", report)
        unmapped_path = write_unmapped(ctx.unmapped, out_root)
        if unmapped_path:
            log.info("待补英文类目：%d 个 → %s", len(ctx.unmapped), unmapped_path)
    return 0 if bad == 0 else 1


def main(argv: list[str] | None = None) -> int:
    try:
        return run(build_parser().parse_args(argv))
    except KeyboardInterrupt:
        print("\n已中断")
        return 130


if __name__ == "__main__":
    sys.exit(main())
