# -*- coding: utf-8 -*-
"""插件公共类型：运行上下文、处理结果、客户插件协议。

放在顶层（不放 adapters 包内）是为了避免 `registry ↔ adapters` 的循环导入：
registry 只依赖本模块，adapters 各插件依赖 registry + 本模块。
"""
from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Protocol


@dataclass
class RunContext:
    """一次运行共享的上下文。"""
    out_root: Path
    refs: dict = field(default_factory=dict)          # 客户参考号缺失时的对照表
    run_date: str = ""                               # yyyymmdd，写进输出文件名
    unmapped: dict[str, int] = field(default_factory=dict)  # 缺英文的类目 -> 箱数
    used: set[str] = field(default_factory=set)             # 本次运行已占用的输出文件名
    dry_run: bool = False                            # 只解析校验，不写文件
    log: logging.Logger = logging.getLogger("fapiao")
    progress: object | None = None                   # common.progress.Progress，可为 None


@dataclass
class Generated:
    """一条处理结果（写报告的单元）。"""
    file: str
    path: Path | None = None
    stat: dict | None = None
    problems: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.problems


class CustomerPlugin(Protocol):
    """一个客户 = 一个插件。"""
    name: str

    def discover(self, ctx: RunContext) -> list[Path]:
        """返回本次要处理的输入文件。"""
        ...

    def process(self, path: Path, ctx: RunContext) -> list[Generated]:
        """处理单个文件，写盘并返回报告条目。"""
        ...
