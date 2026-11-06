# -*- coding: utf-8 -*-
"""内置客户插件。import 本包即完成注册（顺序无关，registry 负责去重）。"""
from __future__ import annotations

from ..registry import CUSTOMERS, get, names, register
from .base import CustomerPlugin, Generated, RunContext

# 导入即注册；新增客户时在这里加一行
from . import amoke, cyj, deppon, shangrui, zehui, shiyizhong  # noqa: E402,F401

__all__ = ["CUSTOMERS", "CustomerPlugin", "Generated", "RunContext",
           "get", "names", "register"]
