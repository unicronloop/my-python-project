# -*- coding: utf-8 -*-
"""客户 Excel 处理工具 —— 启动入口（薄壳）。

真正的实现都在 `src/fapiao/`：
    cli        命令行
    registry   客户插件注册表
    adapters/  每客户一个文件（泽汇 / 时易中 / 创颖峻 / 德邦）
    builders/  公司模板写盘
    models/    Canonical Shipment / 汇总计算

用法：
    python 新需求.py                      # 处理全部客户
    python 新需求.py --only 德邦 创颖峻    # 只处理指定客户
    python 新需求.py --list               # 看已注册客户
"""
from __future__ import annotations

import sys
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))


def main() -> int:
    from fapiao.cli import main as cli_main
    from fapiao.config import pause_before_exit

    try:
        return cli_main()
    except Exception:  # noqa: BLE001  双击运行时留住报错信息
        print("\n[错误] 运行失败：")
        traceback.print_exc()
        # force=True：直接双击 .py 运行（未打包）时也停住，否则窗口一闪就没
        pause_before_exit(force=True)
        return 1


if __name__ == "__main__":
    code = main()
    if code == 0:
        from fapiao.config import pause_before_exit

        pause_before_exit()
    sys.exit(code)
