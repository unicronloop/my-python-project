# -*- coding: utf-8 -*-
"""路径、全局常量、首运行自举。"""
from __future__ import annotations

import base64
import datetime
import os
import sys
from pathlib import Path

# =========================================================================
# ★★★  处理速度限制（业务要求：每生成一个文件，停一下再继续）  ★★★
#
#   想调快 / 调慢，只改下面这一行：
#        5.0 = 每生成一个文件停 5 秒（当前）
#       10.0 = 停 10 秒
#        0.0 = 不限制
#
#   说明：每写出一张发票就静默等待 N 秒，不打印任何提示，进度条照常走。
# =========================================================================
MIN_SECONDS_PER_FILE = 3.0

# 仅打包冒烟测试 / 单元测试用的临时旁路，业务环境不要设置这个环境变量
MIN_SECONDS_ENV = "FAPIAO_MIN_SECONDS_PER_FILE"


def min_seconds_per_file() -> float:
    """实际生效的每文件最少秒数（常量优先，测试时可用环境变量覆盖）。"""
    raw = os.environ.get(MIN_SECONDS_ENV, "").strip()
    if raw:
        try:
            return max(float(raw), 0.0)
        except ValueError:
            pass
    return max(float(MIN_SECONDS_PER_FILE), 0.0)


# ---------------- 路径 ----------------
RUNTIME_DIR_NAME = "_runtime"
ROOT_DIR_ENV = "FAPIAO_ROOT_DIR"
# 兼容 hidden-box 约定里的 APP_ROOT_DIR：两个变量谁在就用谁
ROOT_DIR_ENV_ALIASES = (ROOT_DIR_ENV, "APP_ROOT_DIR")


def _app_base_dir() -> Path:
    """程序根目录（操作员可见的那一层）。

    优先级：
    1. 启动器写入的 FAPIAO_ROOT_DIR / APP_ROOT_DIR
    2. Nuitka（__compiled__）
    3. PyInstaller（sys.frozen）
    4. 源码运行：本文件上三级目录（src/fapiao/config.py → 项目根）
    """
    for env_name in ROOT_DIR_ENV_ALIASES:
        env_root = os.environ.get(env_name, "").strip().strip('"')
        if not env_root:
            continue
        try:
            candidate = Path(env_root)
            if candidate.is_dir():
                return candidate.resolve()
        except OSError:
            pass

    compiled = globals().get("__compiled__")
    if compiled is not None:
        if getattr(compiled, "onefile", False):
            origin = getattr(compiled, "original_argv0", None) or sys.argv[0]
            return Path(origin).resolve().parent
        exe = Path(sys.executable).resolve()
        if exe.name.lower() in ("python.exe", "pythonw.exe"):
            exe = Path(sys.argv[0]).resolve()
        # 便携包兜底：有人直接双击了 _runtime/1.exe，也要回到外层根目录
        if exe.parent.name == RUNTIME_DIR_NAME:
            return exe.parent.parent
        return exe.parent

    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent

    return Path(__file__).resolve().parent.parent.parent


BASE_DIR = _app_base_dir()

# 客户数据目录（按业务约定编号，改目录名只需要改这里）
ZEHUI_DIR = BASE_DIR / "01泽汇"
ZEHUI_PICKUP_DIR = BASE_DIR / "02泽汇-提货表"
SHIYIZHONG_DIR = BASE_DIR / "03时易中"
CYJ_DIR = BASE_DIR / "04创颖峻"
DEPPON_DIR = BASE_DIR / "05德邦"
SHANGRUI_DIR = BASE_DIR / "06尚睿"
AMOKE_DIR = BASE_DIR / "07爱墨科技"

TEMPLATE_DIR = BASE_DIR / "001模板文件"
COMPANY_TEMPLATE = TEMPLATE_DIR / "公司模板.xlsx"
OUTPUT_DIR = BASE_DIR / "生成的文件"

# 旧名保留：地址库等仍以 TEMPLATE_PATH 读公司模板
TEMPLATE_PATH = COMPANY_TEMPLATE

# 输出：公司模板「出货发票」明细起始行 / 行高
DETAIL_START_ROW = 18
DETAIL_ROW_HEIGHT = 54
SHEET_INVOICE = "出货发票"
SHEET_ADDR = "FBA仓码地址"

# 时易中出货发票明细起始行（第 17 行是表头）
INVOICE_DETAIL_START_ROW = 18

SHIYIZHONG_DEFAULTS = {
    # 收货渠道默认值置空：业务上不存在通用兜底值，宁可空着也不写错
    "收货渠道": "",
    "报关类型": "0110报关",
}

# 泽汇：需要删除的表头列（按表头名称匹配）
ZEHUI_UNWANTED_HEADERS = {
    "产品长(cm)", "产品宽(cm)", "产品高(cm)",
    "大货属性", "是否带电/带磁",
}

# 泽汇：报关类型映射（源表写法 → 我司取值），统一 normalize_label 后比较
ZEHUI_DECLARATION_RULES = {
    "合并报关": ("报关退税",),
    "不需报关": ("买关报单", "买单报关"),
}
ZEHUI_DEFAULT_DECLARATION = "不需报关"

# 客户渠道 → 我司渠道
CHANNEL_MAP = {
    "纽约卡派": "腾龙8号卡派含税",
    "腾龙8号卡派": "腾龙8号卡派含税",
    "海运美森正班": "美森特快线",
    "美森限时达": "美森特快线",
    "美森正班": "美森特快线",
    "FBA华南美东特惠（包税）": "腾龙8号卡派含税",
    "FBA华南美东特惠(包税)": "腾龙8号卡派含税",
    "海运8号卡航包税": "海运8号卡航包税",
    "腾龙8号卡派含税": "腾龙8号卡派含税",
    "美西洛杉矶卡派纯电 包税": "美西洛杉矶卡派含税",
    "美西洛杉矶卡派纯电包税": "美西洛杉矶卡派含税",
    "美西洛杉矶卡派包税": "美西洛杉矶卡派含税",
}


def normalize_channel(raw: str) -> str:
    """客户渠道 → 我司渠道；未收录的做保守归一（去「纯电」「包税→含税」）。

    设计取舍：映射表命中优先；没命中时只做「包税 → 含税」「去掉纯电」这种
    语义等价的替换，绝不猜具体线路，避免把客户写法改错。
    """
    text = (raw or "").strip()
    if not text:
        return ""
    for key in (text, text.replace(" ", ""), text.replace(" ", "").replace("（", "(").replace("）", ")")):
        if key in CHANNEL_MAP:
            return CHANNEL_MAP[key]
    if "包税" in text:
        text = text.replace("纯电", "").replace("包税", "含税")
        return " ".join(text.split()).replace(" 含税", "含税").strip()
    return text

# 报关方式 → 我司报关类型
CUSTOMS_MAP = {
    "买单报关": "不需报关",
    "包税": "不需报关",
    "不需报关": "不需报关",
    "报关退税": "合并报关",
    "": "不需报关",
}

# 德邦各模板的渠道兜底（客户模板里没有渠道字段时使用）
DEPPON_CHANNEL = {
    "invoice1": "美森特快线",
    "invoice2": "腾龙8号卡派含税",
    "invoice3": "腾龙8号卡派含税",
    "invoice4": "腾龙8号卡派含税",
    "invoice5": "腾龙8号卡派含税",
}

# ---------------- 爱墨科技 ----------------
# 订舱表「渠道」→ 我司收货渠道（按「仓库代码 + 渠道」筛选清关发票）
AMOKE_CHANNEL_MAP = {
    "海运-以星海派": "以星",
    "海派-以星": "以星",
    "以星": "以星",
    "海运-美森海派": "美森",
    "美森": "美森",
}
# 订舱表「是否包税」→ 报关类型（包税 = 不用单独报关，与德邦/尚睿口径一致）
AMOKE_CUSTOMS_MAP = {"是": "不需报关", "否": "合并报关"}
# 客户上传系统按 sheet 名取数
AMOKE_SHEET = "T6系统发票"
# G14 亚马逊配送时段：客户系统要求必填，源表没有该字段，按客户习惯固定
AMOKE_DELIVERY_WINDOW = datetime.date(2026, 1, 1)
# 「运单备注」写货代给的报关抬头（JL / BY / XW …）
AMOKE_REMARK_LABEL = "运单备注"


# ---------------- 单文件版自举 ----------------
def _embedded_template_bytes() -> bytes:
    """单文件 exe 内置的公司模板（打包时由 _build/gen_embedded.py 生成）。"""
    try:
        from _embedded_template import TEMPLATE_B64  # type: ignore
    except ImportError:
        return b""
    try:
        return base64.b64decode(TEMPLATE_B64)
    except Exception:  # noqa: BLE001
        return b""


def ensure_workspace() -> None:
    """首次运行补齐输出目录 / 释放内置模板。

    客户数据目录（01泽汇 … 05德邦）属于业务输入，不自动创建，
    缺了由 cli 给出明确提示。
    """
    for d in (TEMPLATE_DIR, OUTPUT_DIR):
        try:
            d.mkdir(parents=True, exist_ok=True)
        except OSError as e:
            raise RuntimeError(
                f"无法在程序所在目录创建文件夹：{d}\n"
                f"请把程序放到桌面或 D 盘等可写目录后再运行（{e}）"
            ) from e

    if not COMPANY_TEMPLATE.exists():
        data = _embedded_template_bytes()
        if data:
            COMPANY_TEMPLATE.write_bytes(data)
            print(f"  [初始化] 已生成公司模板：{COMPANY_TEMPLATE}")
        else:
            print(f"  [警告] 缺少公司模板：{COMPANY_TEMPLATE}")


def is_frozen() -> bool:
    return getattr(sys, "frozen", False) or "__compiled__" in globals()


def pause_before_exit(force: bool = False) -> None:
    """双击运行时窗口别一闪而过；命令行/自动化调用不暂停。

    force=True 用于出错时：即使不是打包后的 exe，只要是在控制台里
    手动双击/直接运行，也停下来让用户看到报错。
    """
    if os.environ.get("FAPIAO_NO_PAUSE") == "1":
        return
    if not force and not is_frozen():
        return
    try:
        if not sys.stdin or not sys.stdin.isatty():
            return
    except (ValueError, OSError):
        return
    try:
        input("\n按回车键关闭窗口 ……")
    except (EOFError, KeyboardInterrupt):
        pass
