# 打包说明（hidden-box 方案）

`1.py` 的便携化打包：**启动器 + 隐藏运行时 + zip 交付包**。

```
_build/portable/                    <-- 交付目录（再压成 zip）
├── 客户Excel处理工具.exe           <-- 启动器（约 140 KB，只依赖 KERNEL32.dll）
├── 使用说明.txt
├── _runtime/                       <-- 隐藏；真实主程序 + 全部依赖（约 108 MB）
│   ├── 1.exe                       <-- Nuitka standalone 产物
│   ├── python312.dll
│   └── ...
├── 泽汇/  时易中/  泽汇-提货表/
└── 001模板文件/
```

## 一条命令重建

```bat
.venv\Scripts\python.exe _build\build_portable.py
```

产出：`_build/portable/` 和 `_build/客户Excel处理工具_YYYYMMDD.zip`。

| 参数 | 用途 |
| --- | --- |
| `--skip-main` | 只重编启动器，复用已有 `_runtime`（改 launcher.c 时用，几秒完成） |
| `--no-smoke` | 跳过构建后的冒烟测试 |
| `--no-zip` | 只生成目录，不压包 |
| `--launcher=python` | 没有 MSVC 时用 Nuitka `--onefile` 编译 `launcher.py`（约 1~2 MB，会自解压） |
| `--keep-work` | 保留 `_build/work` 与冒烟测试目录 |

流程：生成内嵌模板 → `nuitka --standalone 1.py` → `1.dist` 改名 `_runtime`
→ 编译启动器 → 铺业务目录 → 打隐藏属性 → 冒烟测试 → 打 zip。

## 验证脚本（发版前建议都跑）

```bat
:: 1) 报关类型映射自检（改映射表后必跑）
.venv\Scripts\python.exe _build\check_declaration_map.py

:: 2) 便携包端到端验证：真实样表 + 报关退税分支 + 中文路径
.venv\Scripts\python.exe _build\e2e_check.py
```

## 关键约定

- 真实主程序**必须**留在 `_runtime` 里。Nuitka standalone 的 `python312.dll`、
  `*.pyd`、`numpy/`、`pandas/` 是按"主程序所在目录"定位的，把 `1.exe` 单拎到
  根目录会直接崩在 Python 初始化。
- 启动器负责三件事：设 `FAPIAO_ROOT_DIR`、把 `_runtime` 前置进 `PATH`、
  以根目录为 cwd 启动 `1.exe` 并透传退出码。
- `1.py` 的 `_app_base_dir()` 优先级：`FAPIAO_ROOT_DIR` → `APP_ROOT_DIR`
  → `__compiled__`（onefile/standalone）→ `sys.frozen` → 脚本目录。
  环境变量指向的目录不存在时会退回自动推断，不会崩。
- 公司模板会以 base64 编进 exe（`work/generated/_embedded_template.py`），
  `001模板文件/公司模板.xlsx` 丢了也会在首次运行时自动重建。
- 压缩包用 Python `zipfile` 生成（不是资源管理器全选压缩）：`_runtime`
  是隐藏目录，手工全选很容易漏，漏了程序就跑不起来。

## 已知限制

- 隐藏属性只是"资源管理器默认不显示"，`dir /a`、7-Zip、Windows 资源管理器的
  "显示隐藏项目"都能看到。它是视觉整理，不是安全边界。
- 首次运行时 Windows Defender 可能对未签名 exe 做云端扫描，表现为双击后几秒
  才出窗口；分发前建议在一台干净机器上试运行，必要时加白名单。
- 需要代码签名才能真正降低误杀率；`_runtime` 约 108 MB，zip 约 40 MB。
- 启动器不做自解压，用户目录不会留临时文件夹，这也是选它而不选
  `--onefile` 的主要原因。
