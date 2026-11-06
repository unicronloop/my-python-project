# -*- coding: utf-8 -*-
"""泽汇适配器：上传表由旧构建器逐行改写（薄壳迁移，暂不套 Shipment）。"""
from __future__ import annotations

import logging
from pathlib import Path

from ..builders.zehui import build_zehui, load_zehui_pickup_map
from ..common.strings import build_output_name, replace_with_retry
from ..config import ZEHUI_DIR
from ..registry import register
from .base import Generated, RunContext

log = logging.getLogger(__name__)


class ZehuiPlugin:
    name = "泽汇"

    def discover(self, ctx: RunContext) -> list[Path]:
        if not ZEHUI_DIR.exists():
            return []
        return [p for p in sorted(ZEHUI_DIR.glob("*.xlsx"))
                if "提货表" not in p.name and "上传表" not in p.name
                and "已处理" not in p.name]

    def process(self, path: Path, ctx: RunContext) -> list[Generated]:
        out_dir = ctx.out_root / self.name
        out_dir.mkdir(parents=True, exist_ok=True)
        tmp = out_dir / f".{path.stem}.tmp.xlsx"
        pickup_map = load_zehui_pickup_map()
        try:
            ref = build_zehui(path, tmp, pickup_map)
        except Exception as exc:  # noqa: BLE001
            ctx.log.error("  [失败] %s: %s", path.name, exc)
            tmp.unlink(missing_ok=True)
            return [Generated(file=f"{self.name}/{path.name}", problems=[str(exc)])]

        out = out_dir / build_output_name(ref, ctx.run_date)
        n = 1
        while out.name in ctx.used:      # 同一次运行内重名才加序号，重跑直接覆盖
            n += 1
            out = out_dir / build_output_name(f"{ref}-{n}", ctx.run_date)
        ctx.used.add(out.name)
        if ctx.dry_run:                  # 只校验不落盘
            tmp.unlink(missing_ok=True)
            return [Generated(file=f"{self.name}/{path.name}", warnings=[])]
        replace_with_retry(tmp, out)   # 杀软/索引占用时自动重试
        return [Generated(file=f"{self.name}/{out.name}", path=out)]


register(ZehuiPlugin())
