# -*- coding: utf-8 -*-
"""极简启动器：把操作员看到的根目录告诉真正的主程序，然后拉起 _runtime 里的它。

只用标准库，编译出来很小；真正的业务代码全在 _runtime/新需求.exe 里。

为什么需要它：
    Nuitka --standalone 的主程序必须待在它自己的 .dist 目录里（DLL 按相对路径找），
    不能简单挪到根目录。启动器负责设置 FAPIAO_ROOT_DIR / APP_ROOT_DIR 和 PATH，
    让主程序在 _runtime 里运行、但把 BASE_DIR 认到外层根目录。
"""
from __future__ import annotations

import os
import sys
import subprocess
from pathlib import Path

RUNTIME_DIR_NAME = "_runtime"
MAIN_EXE_NAME = "新需求.exe"


def _get_root() -> Path:
    """启动器所在目录 = 操作员可见的根目录。

    不能用 Path(__file__)：Nuitka --onefile 下它指向临时解包目录。
    """
    compiled = globals().get("__compiled__")
    if compiled is not None and getattr(compiled, "onefile", False):
        origin = getattr(compiled, "original_argv0", None) or sys.argv[0]
        return Path(origin).resolve().parent

    exe = Path(sys.executable).resolve()
    if exe.name.lower() in ("python.exe", "pythonw.exe"):
        return Path(sys.argv[0]).resolve().parent
    return exe.parent


def _pause(msg: str = "") -> None:
    if msg:
        print(msg)
    try:
        input("\n按回车键关闭窗口 ……")
    except (EOFError, KeyboardInterrupt):
        pass


def main() -> int:
    root = _get_root()
    runtime = root / RUNTIME_DIR_NAME
    exe = runtime / MAIN_EXE_NAME

    if not exe.exists():
        _pause(f"[错误] 找不到主程序：{exe}\n请确认 _runtime 文件夹没有被删除或改名。")
        return 1

    env = os.environ.copy()
    # 主程序据此锁定「操作员可见的根目录」，避免 BASE_DIR 落进 _runtime
    env["FAPIAO_ROOT_DIR"] = str(root)
    env["APP_ROOT_DIR"] = str(root)
    env["PYTHONUNBUFFERED"] = "1"  # 让进度条/日志实时出来

    paths = [str(runtime), str(runtime / "lib")]
    env["PATH"] = os.pathsep.join(paths) + os.pathsep + env.get("PATH", "")

    # 不要加 creationflags=CREATE_NEW_CONSOLE：必须共用当前控制台窗口
    try:
        proc = subprocess.run([str(exe)], cwd=str(root), env=env)
    except Exception as exc:  # noqa: BLE001
        _pause(f"[错误] 启动主程序失败：{exc}")
        return 1
    return proc.returncode


if __name__ == "__main__":
    raise SystemExit(main())
