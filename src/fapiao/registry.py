# -*- coding: utf-8 -*-
"""客户插件注册表。

新增客户：在 `adapters/` 下写一个实现了 CustomerPlugin 的类，模块末尾
`register(MyPlugin())`，然后在 `adapters/__init__.py` 里 import 该模块即可。
`cli` / `config` / `builders` 不需要任何改动。
"""
from __future__ import annotations

from .plugin import CustomerPlugin

CUSTOMERS: dict[str, CustomerPlugin] = {}


def register(plugin: CustomerPlugin) -> CustomerPlugin:
    CUSTOMERS[plugin.name] = plugin
    return plugin


def get(name: str) -> CustomerPlugin:
    try:
        return CUSTOMERS[name]
    except KeyError:
        raise KeyError(f"未知客户：{name}（已注册：{', '.join(CUSTOMERS)}）") from None


def names() -> list[str]:
    """按名称排序，保证报告与输出顺序稳定。"""
    return sorted(CUSTOMERS)
