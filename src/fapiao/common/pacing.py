# -*- coding: utf-8 -*-
"""每生成一个文件后的静默停顿。

业务要求：每个输出文件之间停一下（默认见 `config.MIN_SECONDS_PER_FILE`，
当前 5 秒），避免"一下子生成上百张"，也让操作员跟得上进度。

设计：
- 停顿发生在**单个文件写完之后**（builders 里 save 之后调 `after_file()`），
  所以进度条照常一张一张往前走，不会出现"100% 之后还要干等几分钟"。
- 全程不打印任何日志，操作员只看到正常的进度条。
- 默认关闭；只有 CLI 真正跑业务（非 --dry-run）时才 `enable()`，
  这样单元测试、冒烟测试不会被拖慢。
"""
from __future__ import annotations

import time

from ..config import min_seconds_per_file

_ENABLED = False


def enable(flag: bool = True) -> None:
    """CLI 入口调用：开启/关闭停顿（--dry-run 时关闭）。"""
    global _ENABLED
    _ENABLED = bool(flag)


def is_enabled() -> bool:
    return _ENABLED


def after_file() -> None:
    """一个文件写完后调用：静默等待 N 秒。

    若当前有进度条，等待期间每 `progress.SLEEP_FRAME` 秒重绘一次，
    让「剩余」时间一秒一秒往下走，而不是卡住不动。
    """
    if not _ENABLED:
        return
    seconds = min_seconds_per_file()
    if seconds <= 0:
        return

    from .progress import SLEEP_FRAME, active_progress

    progress = active_progress()
    if progress is None:
        time.sleep(seconds)
        return

    end = time.monotonic() + seconds
    while True:
        now = time.monotonic()
        left = end - now
        if left <= 0:
            break
        time.sleep(min(SLEEP_FRAME, left))
        progress.sleep_tick(min(SLEEP_FRAME, left))
    progress.sleep_tick(0)
