# -*- coding: utf-8 -*-
"""fapiao —— 客户 Excel 统一处理工具（发票 / 装箱单生成）。

分层：

    adapters/   客户输入 → Shipment（或直接交给旧构建器）
    models/     Canonical Shipment / Carton / Totals
    builders/   公司模板 / 泽汇 / 时易中 三个输出构建器
    validate    勾稽校验
    cli         唯一入口，按 registry 分发

新增一个客户：写 adapters/<name>.py + 在 adapters/__init__.py 里 import，
不需要改 cli。
"""

__version__ = "1.0.0"
