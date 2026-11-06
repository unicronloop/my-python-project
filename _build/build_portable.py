#!/usr/bin/env python3
"""
Nuitka Standalone 便携化打包脚本（hidden-box：启动器 + 隐藏运行时 + 压缩包）。

最终目录结构：

    客户Excel处理工具/                  <-- 操作员看到的根目录
    ├── 客户Excel处理工具.exe           <-- 启动器（唯一可见 exe，原生编译，约 140 KB）
    ├── 使用说明.txt
    ├── _runtime/                       <-- 隐藏，存放真实主程序与全部依赖
    │   ├── 新需求.exe                       <-- Nuitka standalone 编译的真实主程序
    │   ├── python3xx.dll
    │   └── ... (所有 DLL / pyd / lib)
    ├── 泽汇/
    ├── 时易中/
    ├── 泽汇-提货表/
    └── 001模板文件/

为什么要多一层启动器，而不是把 新需求.exe 直接挪到根目录：
Nuitka standalone 的 python3xx.dll、pyd 以及 lib 目录都是按"主程序所在目录"
定位的，把 exe 单拎出来会直接崩在 Python 初始化。让真实主程序始终待在
_runtime 里、由外层启动器设置 FAPIAO_ROOT_DIR + PATH，两个目标都不牺牲。

用法：
    python _build/build_portable.py              # 完整构建 + 冒烟测试 + 打 zip
    python _build/build_portable.py --no-smoke   # 跳过冒烟测试
    python _build/build_portable.py --skip-main  # 只重编启动器（复用已有 _runtime）
    python _build/build_portable.py --launcher=python   # 用 Nuitka 编译 launcher.py
"""

from __future__ import annotations

import argparse
import base64
import ctypes
import datetime
import os
import shutil
import subprocess
import sys
import time
import zipfile
from pathlib import Path

# ============ 配置区 ============
PROJECT_ROOT = Path(__file__).resolve().parent.parent
BUILD_ROOT = PROJECT_ROOT / "_build"
WORK_DIR = BUILD_ROOT / "work"
GEN_DIR = WORK_DIR / "generated"
MAIN_BUILD_DIR = WORK_DIR / "main"
LAUNCHER_BUILD_DIR = WORK_DIR / "launcher"
SMOKE_DIR = WORK_DIR / "smoke"
DIST_DIR = BUILD_ROOT / "portable"
CACHE_DIR = PROJECT_ROOT / "_build_cache"

MAIN_SCRIPT = PROJECT_ROOT / "新需求.py"
SRC_DIR = PROJECT_ROOT / "src"          # 包源码（fapiao/）
LAUNCHER_PY = PROJECT_ROOT / "launcher.py"
LAUNCHER_C = BUILD_ROOT / "launcher.c"
TEMPLATE_SRC_DIR = PROJECT_ROOT / "001模板文件"
TEMPLATE_FILE = TEMPLATE_SRC_DIR / "公司模板.xlsx"

RUNTIME_DIR_NAME = "_runtime"
MAIN_EXE_NAME = "新需求.exe"
LAUNCHER_EXE_NAME = "客户Excel处理工具.exe"
PROJECT_NAME = "客户Excel处理工具"
EMBEDDED_MODULE = "_embedded_template"  # 1.py 里 `from _embedded_template import TEMPLATE_B64`

PRODUCT_NAME = PROJECT_NAME
COMPANY_NAME = PROJECT_NAME
FILE_VERSION = "1.0.0.0"

# 编译期砍掉用不到的包：减小体积，也少一些被杀软盯上的二进制
NOFOLLOW_IMPORTS = [
    "pytest",
    "tests",
    "tkinter",
    "matplotlib",
    "scipy",
    "IPython",
    "jupyter",
    "notebook",
    "lib2to3",
    "pydoc_data",
    "docutils",
    "sphinx",
]

DATA_DIRS = ("01泽汇", "02泽汇-提货表", "03时易中", "04创颖峻", "05德邦", "06尚睿", "07爱墨科技")
TEMPLATE_DIR_NAME = "001模板文件"

README_NAME = "使用说明.txt"
README_TEXT = """客户Excel处理工具 使用说明
================================

一、怎么用
1. 把客户发来的 Excel 放进对应文件夹（文件夹名不要改）：
     01泽汇/            -> 泽汇发票明细（.xlsx / .xls / .xlsm）
     02泽汇-提货表/     -> 泽汇提货表（补地址、仓码）
     03时易中/          -> 时易中发票明细
     04创颖峻/          -> 创颖峻客户发票.xlsx
     05德邦/            -> 按批次分子文件夹，交接单和发票放一起
     06尚睿/            -> 尚睿客户发票.xlsx
     07爱墨科技/        -> 一个批次一个子文件夹（如 1/、2/），里面放订舱表 + 清关发票
2. 双击 客户Excel处理工具.exe。
3. 结果在 生成的文件/<客户>/ 里；有问题的单会写进 生成的文件/校验报告.txt。

二、注意
1. _runtime 是运行环境，不能删除、改名或单独拷走；分发时整个文件夹一起压缩。
2. _runtime 显示为隐藏是正常现象；刚解压时可能可见，运行一次后会自动隐藏。
3. 生成的文件如果正被 Excel 打开会写不进去，关掉后重跑即可。
4. 001模板文件/公司模板.xlsx 万一丢了，程序首次运行会自动重建。
5. 程序对每个文件至少处理 10 秒（防止出错/便于观察），所以整批会稍慢，
   请耐心等待进度条走完，不要重复双击。
"""


# ============ 小工具 ============
def log(msg: str = "") -> None:
    try:
        print(msg, flush=True)
    except UnicodeEncodeError:
        # 控制台编码太老（比如 cp936）时别因为一个 ✓ 把构建打断
        print(msg.encode("ascii", "backslashreplace").decode("ascii"), flush=True)


def step(index: int, total: int, msg: str) -> None:
    log()
    log(f"[{index}/{total}] {msg}")


def run(
    cmd,
    cwd: Path | None = None,
    env: dict | None = None,
    attempts: int = 1,
) -> None:
    """运行命令；attempts>1 时用于 Nuitka 重命名 dist 目录偶发被占用（WinError 5）。"""
    printable = " ".join(str(c) for c in cmd)
    workdir = Path(cwd or PROJECT_ROOT)
    for attempt in range(1, attempts + 1):
        log(f"  >>> {printable}" + (f"（第 {attempt}/{attempts} 次）" if attempts > 1 else ""))
        started = time.time()
        result = subprocess.run(cmd, cwd=str(workdir), env=env)
        if result.returncode == 0:
            log(f"  完成，用时 {time.time() - started:.1f}s")
            return
        if attempt < attempts:
            log("  ! 本次失败，清理 _nuitka_temp* 后重试…")
            time.sleep(3)
            for tmp in workdir.glob("_nuitka_temp*"):
                shutil.rmtree(tmp, ignore_errors=True)
    raise SystemExit(f"命令失败（退出码 {result.returncode}）：{printable}")


def human_size(num: float) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if num < 1024 or unit == "GB":
            return f"{num:.1f} {unit}"
        num /= 1024
    return f"{num:.1f} GB"


def dir_size(path: Path) -> int:
    total = 0
    for root, _dirs, files in os.walk(path):
        for name in files:
            try:
                total += (Path(root) / name).stat().st_size
            except OSError:
                pass
    return total


def nuitka_env() -> dict:
    """把 Nuitka 的缓存目录固定在工作区内，避免污染 %LOCALAPPDATA%。"""
    env = os.environ.copy()
    env["NUITKA_CACHE_DIR"] = str(CACHE_DIR)
    env.setdefault("PYTHONIOENCODING", "utf-8")
    return env


def set_hidden(path: Path) -> None:
    """给 _runtime 打 Windows 隐藏属性。失败不算致命。"""
    if sys.platform != "win32":
        return
    try:
        FILE_ATTRIBUTE_HIDDEN = 0x02
        current = ctypes.windll.kernel32.GetFileAttributesW(str(path))
        if current == -1:
            return
        ctypes.windll.kernel32.SetFileAttributesW(str(path), current | FILE_ATTRIBUTE_HIDDEN)
        log(f"  已设置隐藏属性：{path.name}")
    except OSError as exc:
        log(f"  [警告] 设置隐藏属性失败：{exc}")


def find_single(directory: Path, pattern: str) -> Path | None:
    matches = sorted(directory.glob(pattern))
    return matches[0] if matches else None


# ============ 1. 生成内嵌公司模板 ============
def gen_embedded_template() -> None:
    """
    把 001模板文件/公司模板.xlsx 编进程序，业务员就算把模板弄丢，
    ensure_workspace() 也能自己重建。生成到 work/generated，不污染项目根目录。
    """
    if not TEMPLATE_FILE.is_file():
        log(f"  [警告] 未找到模板 {TEMPLATE_FILE}，将跳过内嵌模板（首次运行需手动放模板）")
        return

    GEN_DIR.mkdir(parents=True, exist_ok=True)
    payload = base64.b64encode(TEMPLATE_FILE.read_bytes()).decode("ascii")
    target = GEN_DIR / f"{EMBEDDED_MODULE}.py"
    chunks = [payload[i:i + 100] for i in range(0, len(payload), 100)]
    body = "TEMPLATE_B64 = (\n" + "\n".join(f'    "{c}"' for c in chunks) + "\n)\n"
    target.write_text(
        '"""自动生成，请勿手改：001模板文件/公司模板.xlsx 的 base64 副本。"""\n' + body,
        encoding="utf-8",
    )
    log(f"  已生成 {target.relative_to(PROJECT_ROOT)}（{human_size(target.stat().st_size)}）")


# ============ 2. 编译真实主程序 ============
def build_main() -> Path:
    cmd = [
        sys.executable, "-m", "nuitka",
        "--standalone",
        "--enable-plugin=no-qt",
        "--windows-console-mode=force",
        "--assume-yes-for-downloads",
        "--remove-output",
        f"--output-dir={MAIN_BUILD_DIR}",
        f"--include-module={EMBEDDED_MODULE}",
        "--include-package=fapiao",
        "--include-package=xlrd",
        *[f"--nofollow-import-to={m}" for m in NOFOLLOW_IMPORTS],
        f"--company-name={COMPANY_NAME}",
        f"--product-name={PRODUCT_NAME}",
        f"--file-description={PRODUCT_NAME}",
        f"--file-version={FILE_VERSION}",
        f"--product-version={FILE_VERSION}",
        str(MAIN_SCRIPT),
    ]
    env = nuitka_env()
    # 让 Nuitka 找到 work/generated 下的 _embedded_template
    old_pythonpath = env.get("PYTHONPATH", "")
    env["PYTHONPATH"] = os.pathsep.join(
        [str(GEN_DIR), str(SRC_DIR)] + ([old_pythonpath] if old_pythonpath else []))

    # 3 次机会：Nuitka 最后一步把 _nuitka_temp 改名成 .dist 时，偶尔会被
    # Defender/索引服务短暂占用，报 WinError 5；重试基本都能过。
    run(cmd, env=env, attempts=3)

    dist = MAIN_BUILD_DIR / "新需求.dist"
    if not dist.is_dir():
        found = find_single(MAIN_BUILD_DIR, "*.dist")
        if found is None:
            raise SystemExit(f"未找到 Nuitka 输出目录：{dist}")
        dist = found
    return dist


# ============ 3. 编译启动器 ============
def find_vcvars() -> Path | None:
    """定位 vcvars64.bat（用于把 launcher.c 编成没有 Python 依赖的小 exe）。"""
    candidates: list[Path] = []

    vswhere = (
        Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"))
        / "Microsoft Visual Studio" / "Installer" / "vswhere.exe"
    )
    if vswhere.is_file():
        try:
            out = subprocess.run(
                [str(vswhere), "-latest", "-products", "*",
                 "-requires", "Microsoft.VisualStudio.Component.VC.Tools.x86.x64",
                 "-property", "installationPath"],
                capture_output=True, text=True, timeout=120,
            ).stdout
            for line in out.splitlines():
                line = line.strip()
                if line:
                    candidates.append(Path(line) / "VC" / "Auxiliary" / "Build" / "vcvars64.bat")
        except (OSError, subprocess.SubprocessError):
            pass

    for base in {
        Path(os.environ.get("ProgramFiles", r"C:\Program Files")),
        Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")),
    }:
        for pattern in ("Microsoft Visual Studio/*/*/VC/Auxiliary/Build/vcvars64.bat",
                        "Microsoft Visual Studio/*/VC/Auxiliary/Build/vcvars64.bat"):
            candidates.extend(sorted(base.glob(pattern), reverse=True))

    for candidate in candidates:
        if candidate.is_file():
            return candidate
    return None


def find_rc_exe() -> Path | None:
    """找 Windows SDK 里的 rc.exe（只用来编版本资源，不依赖 vcvars 的 INCLUDE）。"""
    bases = {
        Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "Windows Kits" / "10" / "bin",
        Path(os.environ.get("WindowsSdkDir", "")) / "bin",
    }
    for base in bases:
        if not base.is_dir():
            continue
        candidates = sorted(base.glob("*/x64/rc.exe"))
        if candidates:
            return candidates[-1]
    return None


def _short_path(path: Path) -> Path:
    """把含中文的路径转成 8.3 短路径，保证 .bat 里只剩 ASCII（cmd 读 UTF-8 批处理不可靠）。"""
    if sys.platform != "win32":
        return path
    buf = ctypes.create_unicode_buffer(32768)
    try:
        if ctypes.windll.kernel32.GetShortPathNameW(str(path), buf, 32768):
            return Path(buf.value)
    except OSError:
        pass
    return path


def build_native_launcher(vcvars: Path, dest: Path) -> None:
    LAUNCHER_BUILD_DIR.mkdir(parents=True, exist_ok=True)
    obj = LAUNCHER_BUILD_DIR / "launcher.obj"
    tmp_exe = LAUNCHER_BUILD_DIR / "launcher.exe"
    res_name = "launcher.res"

    # 版本资源：rc.exe 用相对输出名 + cwd 定位，避免 /fo:"C:\..." 的引号解析坑
    rc_file = LAUNCHER_C.with_suffix(".rc")
    res_arg = ""
    if rc_file.is_file():
        rc_exe = find_rc_exe()
        if rc_exe is not None:
            run([str(rc_exe), "/nologo", "/fo", res_name, str(rc_file)], cwd=LAUNCHER_BUILD_DIR)
            if (LAUNCHER_BUILD_DIR / res_name).is_file():
                res_arg = f" {res_name}"
            else:
                log("  [警告] 版本资源生成失败，启动器将不带版本信息")
        else:
            log("  [警告] 没找到 rc.exe，启动器将不带版本信息")

    # 走临时 .bat：
    # 1) 直接 `cmd /c call "...bat"` 走参数列表时，subprocess 会把引号转义成 \"，cmd 不认；
    # 2) 用 ASCII 短路径，避免 cmd 解析中文批处理/中文命令行时乱码。
    bat = LAUNCHER_BUILD_DIR / "build_launcher.bat"
    bat.write_text(
        "@echo off\r\n"
        'pushd "%~dp0"\r\n'
        f'call "{_short_path(vcvars)}" >nul\r\n'
        "if errorlevel 1 exit /b 1\r\n"
        f'cl /nologo /O2 /MT /W3 /utf-8 /DUNICODE /D_UNICODE "{_short_path(LAUNCHER_C)}"{res_arg} '
        f'/Fe:"{_short_path(LAUNCHER_BUILD_DIR) / tmp_exe.name}" '
        f'/Fo:"{_short_path(LAUNCHER_BUILD_DIR) / obj.name}"\r\n'
        "popd\r\n",
        encoding="ascii",
    )
    run(["cmd", "/c", str(_short_path(bat))], cwd=LAUNCHER_BUILD_DIR)
    if not tmp_exe.is_file():
        raise SystemExit(f"启动器编译失败，未生成：{tmp_exe}")

    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.exists():
        dest.unlink()
    shutil.move(str(tmp_exe), str(dest))


def build_python_launcher(dest: Path) -> None:
    cmd = [
        sys.executable, "-m", "nuitka",
        "--onefile",
        "--windows-console-mode=force",
        "--assume-yes-for-downloads",
        "--remove-output",
        f"--output-dir={LAUNCHER_BUILD_DIR}",
        f"--output-filename={LAUNCHER_EXE_NAME}",
        f"--company-name={COMPANY_NAME}",
        f"--product-name={PRODUCT_NAME}",
        f"--file-description={PRODUCT_NAME}",
        f"--file-version={FILE_VERSION}",
        f"--product-version={FILE_VERSION}",
        str(LAUNCHER_PY),
    ]
    run(cmd, env=nuitka_env())

    built = find_single(LAUNCHER_BUILD_DIR, "*.exe")
    if built is None:
        raise SystemExit(f"启动器编译失败，{LAUNCHER_BUILD_DIR} 下没有 exe")
    shutil.copy2(built, dest)


# ============ 4. 组装交付目录 ============
def prepare_data_dirs(dest: Path) -> None:
    for name in DATA_DIRS:
        (dest / name).mkdir(parents=True, exist_ok=True)
    log(f"  已创建业务文件夹：{'、'.join(DATA_DIRS)}")

    template_dst = dest / TEMPLATE_DIR_NAME
    template_dst.mkdir(parents=True, exist_ok=True)
    if TEMPLATE_SRC_DIR.is_dir():
        for f in sorted(TEMPLATE_SRC_DIR.iterdir()):
            if f.is_file() and not f.name.startswith("~$"):
                shutil.copy2(f, template_dst / f.name)
                log(f"  已复制模板文件：{TEMPLATE_DIR_NAME}/{f.name}")


def write_readme(dest: Path) -> None:
    (dest / README_NAME).write_text(README_TEXT, encoding="utf-8-sig")
    log(f"  已生成 {README_NAME}")


# ============ 5. 冒烟测试 ============
def smoke_test(dest: Path) -> None:
    """在副本里跑一遍，确认 启动器 -> _runtime/新需求.exe -> BASE_DIR 全部正确。"""
    if SMOKE_DIR.exists():
        shutil.rmtree(SMOKE_DIR)
    log("  正在复制产物到临时目录做冒烟测试 ……")
    shutil.copytree(dest, SMOKE_DIR)

    root = SMOKE_DIR.resolve()
    runtime = root / RUNTIME_DIR_NAME
    launcher = root / LAUNCHER_EXE_NAME
    template = root / TEMPLATE_DIR_NAME / "公司模板.xlsx"

    # 故意删掉模板，验证"内置模板自举"是否真的编进去了
    template.unlink(missing_ok=True)

    env = {k: v for k, v in os.environ.items()
           if k not in ("FAPIAO_ROOT_DIR", "APP_ROOT_DIR")}
    env["FAPIAO_NO_PAUSE"] = "1"
    env["FAPIAO_MIN_SECONDS_PER_FILE"] = "0"
    env["PYTHONIOENCODING"] = "utf-8"

    def run_once(exe: Path, name: str) -> str:
        # 刚链接出来的 exe 第一次运行，偶尔会被杀软/索引服务挡住：
        # 表现为 0xC0000409 之类的崩溃码，隔一两秒重跑就正常，所以这里重试。
        last_code, last_text = None, ""
        for attempt in range(1, 4):
            proc = subprocess.run(
                [str(exe)], cwd=str(root), env=env,
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            )
            text = proc.stdout.decode("utf-8", "replace")
            if proc.returncode == 0:
                if attempt > 1:
                    log(f"  （{name} 第 {attempt} 次尝试成功，前一次退出码 {last_code}，按首跑被杀软干扰处理）")
                return text
            last_code, last_text = proc.returncode, text
            if attempt < 3:
                log(f"  [重试] {name} 第 {attempt} 次退出码 {last_code}，2 秒后重试")
                time.sleep(2)
        raise SystemExit(f"[冒烟测试失败] {name} 退出码 {last_code}\n{last_text}")

    # 1) 走启动器
    text = run_once(launcher, "启动器")
    if f"输出目录：{root / "生成的文件"}" not in text:
        raise SystemExit(f"[冒烟测试失败] 通过启动器运行时 BASE_DIR 不对，实际输出：\n{text}")
    if not (root / "生成的文件").is_dir():
        raise SystemExit("[冒烟测试失败] 未生成 生成的文件 目录")
    if not template.is_file() or template.stat().st_size == 0:
        raise SystemExit("[冒烟测试失败] 公司模板没有被内置模板自动重建（_embedded_template 未编进去）")
    log("  ✓ 启动器：BASE_DIR 指向根目录，内置模板可自举")

    # 2) 直接双击 _runtime/新需求.exe（没有环境变量），靠 _runtime 兜底也应该回到根目录
    text = run_once(runtime / MAIN_EXE_NAME, "主程序")
    if f"输出目录：{root / "生成的文件"}" not in text:
        raise SystemExit(f"[冒烟测试失败] 直接运行 _runtime/新需求.exe 时 BASE_DIR 不对：\n{text}")
    log("  ✓ 主程序：直接运行 _runtime/新需求.exe 也能回落到根目录")

    if not os.environ.get("FAPIAO_KEEP_SMOKE"):
        shutil.rmtree(SMOKE_DIR, ignore_errors=True)


# ============ 6. 打压缩包 ============
def _write_dir_entry(zf: zipfile.ZipFile, arc: str, hidden: bool = False) -> None:
    """写一个目录条目；DOS 属性放在 external_attr 低 16 位，部分解压工具会照做。"""
    zi = zipfile.ZipInfo(arc.rstrip("/") + "/")
    zi.date_time = datetime.datetime.now().timetuple()[:6]
    zi.external_attr = (0o40755 << 16) | 0x10 | (0x02 if hidden else 0)
    zf.writestr(zi, b"")


def make_zip(dest: Path, zip_path: Path) -> None:
    """
    把交付目录压成 zip，压缩包里自带一层 客户Excel处理工具/ 目录，
    业务员解压后就是一个完整文件夹（不会把一堆文件散落到桌面）。

    用 Python 的 zipfile 而不是资源管理器全选压缩：_runtime 是隐藏文件夹，
    手工全选很容易漏掉，程序就残了。
    """
    log(f"  正在压缩 → {zip_path.name}")
    started = time.time()
    tmp = zip_path.with_suffix(".zip.tmp")
    if tmp.exists():
        tmp.unlink()

    empty_dirs: list[str] = []
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for root, dirs, files in os.walk(dest):
            root_path = Path(root)
            rel = root_path.relative_to(dest)
            arc_dir = Path(PROJECT_NAME) / rel if str(rel) != "." else Path(PROJECT_NAME)
            if not files and not dirs:
                empty_dirs.append(arc_dir.as_posix())
            for name in sorted(files):
                full = root_path / name
                zf.write(full, (arc_dir / name).as_posix())
        # 空目录（泽汇/ 时易中/ 等）单独建条目，否则解压后不出现
        for entry in empty_dirs:
            _write_dir_entry(zf, entry)
        # _runtime 带 DOS 隐藏属性，解压后尽量保持"看不见"；丢了也无所谓，
        # 启动器每次启动会自己补上。
        if (dest / RUNTIME_DIR_NAME).is_dir():
            _write_dir_entry(zf, f"{PROJECT_NAME}/{RUNTIME_DIR_NAME}", hidden=True)

    tmp.replace(zip_path)
    log(f"  压缩完成：{zip_path.name}（{human_size(zip_path.stat().st_size)}，"
        f"用时 {time.time() - started:.1f}s）")


# ============ 主流程 ============
def main() -> None:
    parser = argparse.ArgumentParser(description="构建客户Excel处理工具便携包")
    parser.add_argument("--launcher", choices=("auto", "native", "python"), default="auto",
                        help="启动器实现：native=cl 编译 launcher.c（推荐，约 140 KB）；"
                             "python=Nuitka 编译 launcher.py；auto=优先 native")
    parser.add_argument("--skip-main", action="store_true", help="跳过主程序编译，复用已有 _runtime")
    parser.add_argument("--no-smoke", action="store_true", help="跳过冒烟测试")
    parser.add_argument("--no-zip", action="store_true", help="只构建目录，不生成 zip")
    parser.add_argument("--keep-work", action="store_true", help="保留 _build/work 与冒烟测试目录")
    args = parser.parse_args()

    total = 8
    started = time.time()

    # 输出统一 UTF-8：管道/重定向时不会因为 GBK 编码报 UnicodeEncodeError
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, OSError, ValueError):
            pass

    log("=" * 64)
    log(f"  便携包构建：{PRODUCT_NAME}")
    log("=" * 64)

    # ---- 1. 预检 ----
    step(1, total, "环境检查")
    for required in (MAIN_SCRIPT, LAUNCHER_PY, LAUNCHER_C):
        if not required.is_file():
            raise SystemExit(f"缺少文件：{required}")
    try:
        import nuitka  # noqa: F401
    except ImportError:
        raise SystemExit("当前 Python 里没有 Nuitka，请先：pip install nuitka")
    log(f"  Python：{sys.executable}")
    log(f"  项目根：{PROJECT_ROOT}")

    if not args.keep_work and WORK_DIR.exists():
        shutil.rmtree(WORK_DIR)
    WORK_DIR.mkdir(parents=True, exist_ok=True)

    # ---- 2. 内嵌模板 ----
    step(2, total, "生成内嵌公司模板")
    gen_embedded_template()

    # ---- 3. 主程序 ----
    step(3, total, "Nuitka --standalone 编译 新需求.py（首次 5~15 分钟，请耐心等）")
    if args.skip_main:
        log("  已跳过（--skip-main），复用现有产物")
    else:
        build_main()

    # ---- 4. 组装 _runtime ----
    step(4, total, "整理运行时目录")
    runtime_dir = DIST_DIR / RUNTIME_DIR_NAME

    def clean_dist_except_runtime() -> None:
        if not DIST_DIR.exists():
            return
        for item in DIST_DIR.iterdir():
            if item.name == RUNTIME_DIR_NAME:
                continue
            if item.is_dir():
                shutil.rmtree(item)
            else:
                item.unlink()

    if args.skip_main:
        if not (runtime_dir / MAIN_EXE_NAME).is_file():
            raise SystemExit(f"--skip-main 需要已有产物：{runtime_dir / MAIN_EXE_NAME}")
        clean_dist_except_runtime()  # 只清启动器等旧文件，_runtime 原样保留
        log(f"  复用现有 _runtime（{human_size(dir_size(runtime_dir))}）")
    else:
        if DIST_DIR.exists():
            shutil.rmtree(DIST_DIR)
        DIST_DIR.mkdir(parents=True)
        dist = MAIN_BUILD_DIR / "新需求.dist"
        if not dist.is_dir():
            dist = find_single(MAIN_BUILD_DIR, "*.dist")
            if dist is None:
                raise SystemExit(f"未找到 Nuitka 输出目录：{MAIN_BUILD_DIR / '新需求.dist'}")
        shutil.move(str(dist), str(runtime_dir))
        exe = runtime_dir / MAIN_EXE_NAME
        if not exe.is_file():
            built_exe = find_single(runtime_dir, "*.exe")
            if built_exe is None:
                raise SystemExit(f"_runtime 里没有主程序 exe：{runtime_dir}")
            built_exe.rename(exe)
        log(f"  {dist.name}  →  {RUNTIME_DIR_NAME}/（{human_size(dir_size(runtime_dir))}）")
        log(f"  主程序：{RUNTIME_DIR_NAME}/{MAIN_EXE_NAME}")

    # ---- 5. 启动器 ----
    step(5, total, "编译启动器")
    launcher_exe = DIST_DIR / LAUNCHER_EXE_NAME
    vcvars = find_vcvars() if args.launcher in ("auto", "native") else None
    if vcvars is not None and args.launcher != "python":
        log(f"  原生编译：{vcvars}")
        build_native_launcher(vcvars, launcher_exe)
    else:
        if args.launcher == "native":
            raise SystemExit("找不到 vcvars64.bat，无法原生编译启动器；可改用 --launcher=python")
        log("  [提示] 没找到 MSVC，回退到 Nuitka --onefile 编译 launcher.py")
        build_python_launcher(launcher_exe)
    log(f"  启动器：{LAUNCHER_EXE_NAME}（{human_size(launcher_exe.stat().st_size)}）")

    # ---- 6. 业务目录 / 说明 / 隐藏 ----
    step(6, total, "准备业务目录与说明")
    prepare_data_dirs(DIST_DIR)
    write_readme(DIST_DIR)
    set_hidden(runtime_dir)

    # ---- 7. 冒烟测试 ----
    step(7, total, "冒烟测试")
    if args.no_smoke:
        log("  已跳过（--no-smoke）")
    else:
        smoke_test(DIST_DIR)

    if not args.keep_work:
        shutil.rmtree(WORK_DIR, ignore_errors=True)

    # ---- 8. 压缩 ----
    step(8, total, "打包 zip")
    zip_path = BUILD_ROOT / f"{PROJECT_NAME}_{datetime.datetime.now():%Y%m%d}.zip"
    if args.no_zip:
        log("  已跳过（--no-zip）")
    else:
        make_zip(DIST_DIR, zip_path)

    # ---- 完成 ----
    log()
    log("=" * 64)
    log("  构建完成")
    log("=" * 64)
    log(f"  输出目录：{DIST_DIR}")
    log(f"  启动器　：{LAUNCHER_EXE_NAME}（唯一可见 exe，双击运行）")
    log(f"  运行时　：{RUNTIME_DIR_NAME}/（已隐藏，{human_size(dir_size(runtime_dir))}）")
    log(f"  总　计　：{human_size(dir_size(DIST_DIR))}")
    if not args.no_zip:
        log(f"  压缩包　：{zip_path}（{human_size(zip_path.stat().st_size)}）")
    log(f"  耗　时　：{time.time() - started:.1f}s")
    log()
    log("  分发方式：把 zip 发给业务员；解压后业务数据放进 泽汇/、时易中/ 再双击 exe。")
    log("=" * 64)


if __name__ == "__main__":
    main()
