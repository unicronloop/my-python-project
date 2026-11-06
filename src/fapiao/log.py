# -*- coding: utf-8 -*-
"""统一日志。控制台输出格式保持和旧版一致的朴素风格。

进度条正在刷新时，先把进度行擦掉再打印日志、打完再重画，
避免日志插进进度条那一行里（`\\r` 混排的经典问题）。
"""
from __future__ import annotations

import logging
import sys


class _ProgressAwareHandler(logging.StreamHandler):
    def emit(self, record: logging.LogRecord) -> None:
        from .common.progress import active_progress

        progress = active_progress()
        try:
            if progress is not None:
                progress.clear()
            super().emit(record)
        finally:
            if progress is not None:
                progress.redraw()


def setup_log(level: str = "INFO") -> logging.Logger:
    root = logging.getLogger("fapiao")
    if root.handlers:  # 避免重复 addHandler
        root.setLevel(level.upper())
        return root
    handler = _ProgressAwareHandler(stream=sys.stdout)
    handler.setFormatter(logging.Formatter("%(message)s"))
    root.addHandler(handler)
    root.setLevel(level.upper())
    root.propagate = False
    return root


def get_logger(name: str = "fapiao") -> logging.Logger:
    return logging.getLogger(name)
