# 客户 Excel 处理工具

把各客户的 Excel 明细转成「我司出货发票」（001模板文件/公司模板.xlsx）。

## 快速开始

```bash
python 新需求.py                        # 处理全部客户
python 新需求.py --only 德邦             # 只处理一个客户
python 新需求.py --only "泽汇,时易中"     # 逗号分隔
python 新需求.py --list                 # 看已注册客户
python 新需求.py --dry-run              # 只解析校验，不写文件
```

输出在 `生成的文件/<客户>/`，另有 `校验报告.txt`（逐单勾稽结果）和
`类目英文待补充.txt`（创颖峻末级类目缺英文时生成）。

## 数据放哪里

| 目录 | 放什么 |
| --- | --- |
| `01泽汇/` | 泽汇上传表（xlsx） |
| `02泽汇-提货表/` | 提货表，按「客户参考号1」补收货渠道 |
| `03时易中/` | 时易中明细（xls/xlsx） |
| `04创颖峻/` | `客户发票.xlsx`（汇总 + 分箱 sheet） |
| `05德邦/` | 德邦材料：`材料N/` 下放客户交接资料（xlsx/xls，按「交接单」关联 DPL 单号） |
| `06尚睿/` | `客户发票.xlsx`（一份发票一张出货发票） |
| `07爱墨科技/` | `材料/订舱表_*.xlsx` + `客户发票/清关发票_*.xls`（一票货多箱 → 按箱拆） |
| `001模板文件/公司模板.xlsx` | 出货发票模板 **（不要改结构）** |
| `生成的文件/` | 输出，可随时清空重跑 |

德邦的客户参考号（DPL 单号）优先从同目录的「交接单」按 FBA 反查（交接单里
DPL 号与 FBA 的对应关系两种写法都认：`德邦交接单号 → 备注`、`单号 → FBA`）；
查不到时再用 `refs.json`：复制 `refs.example.json` 为 `refs.json`，
按「源文件名 → DPL 单号」维护。

处理时顶部有 tqdm 风格进度条，带速度与预计剩余时间：

```
处理文件[德邦]:  43%|███████████▋                | 速度: 2.29it/s | 剩余: 01:12 | 已用: 00:58
```

「剩余」按真实节流估算（每个文件至少 `config.MIN_SECONDS_PER_FILE` 秒，默认写在该文件第 21 行）。

## 代码结构

```
src/fapiao/
├── cli.py            命令行：遍历插件 → 处理 → 报告
├── registry.py       客户插件注册表（只认这里，不认客户名）
├── plugin.py         插件协议 + RunContext/Generated
├── config.py         路径、渠道/报关映射、类目对照
├── report.py         校验报告 / 待补类目清单
├── validate.py       勾稽校验（箱数/件数/重量/方数）
├── common/           strings / excel / images / region / log
├── models/           Shipment（一票货）、Carton（一箱）、AddressBook、Totals
├── builders/         template（公司模板）、zehui、shiyizhong
└── adapters/         cyj（创颖峻）、deppon（德邦）、amoke（爱墨）、shangrui（尚睿）、handover（交接单）、zehui、shiyizhong
```

两条链路：

- **创颖峻 / 德邦 / 尚睿 / 爱墨**：客户表 → `Shipment`（Canonical 中间模型）→ `builders/template.py` 写公司模板 → `validate` 勾稽。
- **泽汇 / 时易中**：客户表 → `builders/zehui.py`、`builders/shiyizhong.py` 直接改表（历史逻辑，暂未套 Shipment）。

## 新增一个客户

1. 在 `src/fapiao/adapters/` 下新建 `xxx.py`：
   ```python
   from ..plugin import Generated, RunContext
   from ..registry import register
   from .base import emit_shipment

   class XxxPlugin:
       name = "客户名"
       def discover(self, ctx: RunContext) -> list[Path]: ...
       def process(self, path: Path, ctx: RunContext) -> list[Generated]: ...

   register(XxxPlugin())
   ```
2. 在 `adapters/__init__.py` 的 import 列表里加一行。

`cli.py` / `config.py` / `builders/` / 报告都不用改。

## 关键约定

- `Carton.weight` 一律是**单箱实重**，总重 = `weight * boxes`（`models/totals.py` 唯一实现）。
- 输出文件名：`<客户参考号>_<仓码>_<单号区分>_<渠道>_<N>箱.xlsx`（创颖峻单号区分=RTD，德邦=FBA）。
- 单据头 `预报方数(13)` / `申报总值(14)` 只在数据源显式给出时写，保持与手工样例一致。

## 测试

```bash
python -m unittest discover -s tests
```

## 打包

`python _build/build_portable.py`（Nuitka standalone + 原生启动器 + 隐藏 `_runtime`）。

## 版本管理

仓库只有**代码**，客户原始数据（`01泽汇/` … `07爱墨科技/`）、模板、构建产物、
生成结果都在 `.gitignore` 里，不会外传。首次拉取后按上文表格把数据目录放好即可运行。

```bash
git add -A && git commit -m "说明" && git push
```
