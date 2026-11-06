# -*- coding: utf-8 -*-
"""时易中适配器：xls 明细 → 公司模板（旧构建器薄壳迁移）。"""
from __future__ import annotations

import logging
from pathlib import Path

from ..builders.shiyizhong import build_shiyizhong, channel_from_filename
from ..common.strings import build_output_name, replace_with_retry
from ..config import SHIYIZHONG_DEFAULTS, SHIYIZHONG_DIR, COMPANY_TEMPLATE
from ..registry import register
from .base import Generated, RunContext

log = logging.getLogger(__name__)


class ShiyizhongPlugin:
    name = "时易中"

    def discover(self, ctx: RunContext) -> list[Path]:
        if not SHIYIZHONG_DIR.exists():
            return []
        files = [*SHIYIZHONG_DIR.glob("*.xls"), *SHIYIZHONG_DIR.glob("*.xlsx")]
        return sorted(p for p in files
                      if not p.name.startswith("~$") and "已处理" not in p.name)

    def process(self, path: Path, ctx: RunContext) -> list[Generated]:
        out_dir = ctx.out_root / self.name
        out_dir.mkdir(parents=True, exist_ok=True)
        tmp = out_dir / f".{path.stem}.tmp.xlsx"

        config = dict(SHIYIZHONG_DEFAULTS)
        # 文件名里能看出渠道就用文件名，否则用默认配置
        if channel_from_filename(path):
            config["收货渠道"] = channel_from_filename(path)
        try:
            ref = build_shiyizhong(path, tmp, COMPANY_TEMPLATE, config)
        except Exception as exc:  # noqa: BLE001
            ctx.log.error("  [失败] %s: %s", path.name, exc)
            tmp.unlink(missing_ok=True)
            return [Generated(file=f"{self.name}/{path.name}", problems=[str(exc)])]

        if ref is False:
            tmp.unlink(missing_ok=True)
            log.info("  [跳过] %s：不是发票明细格式", path.name)
            return []

        out = out_dir / build_output_name(ref, ctx.run_date)
        n = 1
        while out.name in ctx.used:      # 同一次运行内重名才加序号，重跑直接覆盖
            n += 1
            out = out_dir / build_output_name(f"{ref}-{n}", ctx.run_date)
        ctx.used.add(out.name)
        if ctx.dry_run:                  # 只校验不落盘
            tmp.unlink(missing_ok=True)
            return [Generated(file=f"{self.name}/{path.name}", warnings=[])]
        replace_with_retry(tmp, out)
        return [Generated(file=f"{self.name}/{out.name}", path=out)]


register(ShiyizhongPlugin())
