# 依赖配置与降级指南

本文件说明运行本 skill 需要什么、怎么装、怎么验证、缺了怎么办。依赖清单是 [skill-dependencies.json](../skill-dependencies.json)，本文件是它的操作说明；两者内容不一致时以清单为准。

本 skill 的脚本全部在本地运行，不安装系统软件、不登录账号、不需要任何密钥。**环境检查脚本不发起任何网络请求。**

## 1. 依赖总览

| 依赖 id | 类型 | 必需 | 支撑的功能 | 缺失时的影响 |
|---|---|---|---|---|
| `python-docx` | 运行时库 | 是 | 生成 Word 材料、结构审计 | 无法生成 DOCX，也无法审计版式 |
| `ms-word-com` | 本机工具 | 否 | 渲染预览 | 拿不到逐页 PNG，只能人工目视 |
| `pymupdf` | 运行时库 | 否 | 渲染预览 | PDF 无法转成逐页 PNG |
| `ppt-master-engine` | 内置归档 | 需 PPT 时必需 | 生成 PPTX | 停止 PPT 环节，标注 PPT 未生成 |
| `internet-access` | 网络 | 否 | 主题事实核实 | 只用通知给出的主题表述 |

必需项缺失时，本 skill 仍可交付 Markdown 正文，但不得声称已生成 DOCX 或已通过版式检查。

## 2. 环境检查怎么跑

```bash
python scripts/check_environment.py
```

输出 `ready` / `partial` / `needs_setup` 三态，并逐项列出结果。三个状态的含义：

- `ready`：必需项全部就绪，可走完整流程；
- `partial`：必需项就绪但可选项缺失，走降级路径，交付报告里写明缺哪一项；
- `needs_setup`：必需项缺失，按第 3 节补齐后重跑。

脚本只读，不安装、不写配置、不发网络请求。加 `--probe-word` 会额外真实启动一次 Word 进程读取版本号后立即退出，用于确认 COM 可用；不加时只读注册表。

**注意用哪个解释器**：要用装有 python-docx 的那个 Python 运行本 skill 的全部脚本。如果 `check_environment.py` 报 `python-docx` 缺失，先在下面确认该解释器：

```bash
python -c "import sys; print(sys.executable)"
python -c "import docx; print(docx.__version__)"
```

## 3. 逐项配置

### 3.1 python-docx

- 官方站点：https://python-docx.readthedocs.io/
- 官方文档：https://python-docx.readthedocs.io/en/latest/
- 要求：Python 3.9 及以上

步骤：

```bash
python --version
python -m pip install python-docx
python scripts/check_environment.py
```

验证方式：检查结果里 `python-docx` 一项为 OK，即 `import docx` 成功。
权限与凭据：纯本地库，不需要账号、Token 或网络。
安全提示：只从官方文档指向的 PyPI 渠道安装，不使用来路不明的 whl 包。

### 3.2 ms-word-com

- 官方站点：https://www.microsoft.com/microsoft-365/word
- 官方文档：https://learn.microsoft.com/en-us/office/vba/api/overview/word
- 要求：Windows，装有桌面版 Word（Microsoft 365 或 Word 2019 及以上）

步骤：

```bash
python -m pip install pywin32
python scripts/check_environment.py --probe-word
```

验证方式：检查结果里 `Word COM` 一项能读到版本号；`--probe-word` 能启动 Word 再正常退出。
权限与凭据：使用本机已登录的 Word，不需要额外账号。
安全提示：脚本以只读方式打开 DOCX 并另存为 PDF，不修改源文件、不启用宏、不打开文档里的外部链接；探测完立即退出 Word 进程。

### 3.3 pymupdf

- 官方站点：https://pymupdf.readthedocs.io/
- 官方文档：https://pymupdf.readthedocs.io/en/latest/

步骤：

```bash
python -m pip install pymupdf
python -c "import fitz; print(fitz.__doc__)"
python scripts/check_environment.py
```

验证方式：检查结果里 `PyMuPDF` 一项为 OK。
权限与凭据：本地库，不需要账号。
安全提示：只在本地打开 PDF 并导出 PNG，不上传文件、不联网。

### 3.4 ppt-master-engine

- 上游仓库：https://github.com/hugohe3/ppt-master
- 上游说明：https://github.com/hugohe3/ppt-master#readme
- 要求：ppt-master 6.4.0（无需安装，已随本 skill 内置精简归档）

**无需安装。** 引擎以精简归档内置在 `vendor/ppt-master-6.4.0.zip`（12698 个文件 / 10.28 MB），首次需要生成 PPT 时由 `scripts/bootstrap_engine.py` 校验摘要后解压到缓存目录，之后复用。归档相对上游裁去了三类与本 skill 无关的素材（AI 生图风格参考图、语音旁白音效、企业品牌素材），全部图标库、模板构件、脚本、工作流与文档完整保留。

定位与解压：

```bash
python scripts/bootstrap_engine.py --check          # 只报告状态，不写盘
python scripts/bootstrap_engine.py --verify         # 解压并运行引擎自带的完整性守卫
python scripts/bootstrap_engine.py --json           # 输出机器可读结果，engine 字段即引擎目录
```

首次解压要写入约 1.27 万个文件，视磁盘与杀毒软件设置可能耗时数分钟。过程中被杀毒软件或进程占用打断时，重新运行同一条命令即可：脚本会检出残缺缓存并重新解压修复，不会把半成品当成可用引擎。

引擎目录按以下顺序取第一个可用的，命中即用：

1. 环境变量 `PPT_MASTER_DIR` 指向的目录（已指向有效引擎时直接使用，不解压）；
2. 内置归档自举解压出的缓存目录——缓存解析顺序为 `--cache-dir` > 环境变量 `PPT_MASTER_CACHE_DIR` > skill 同级 `_engine/` > 当前目录 `_engine/`，候选目录全部落在 C 盘时脚本拒绝执行并要求显式指定。

引擎默认把项目建在 `SKILL_DIR` 上两级的 `projects/` 下；生成时一律用 `init --dir` 把项目指定到交付输出目录，不使用引擎仓库内的 `projects/`。

验证方式：`bootstrap_engine.py --verify` 退出码为 0，代表归档摘要与清单一致、解压完整、且引擎自带的 `attribution_guard.py` 通过；`check_environment.py` 的「PPTX 引擎」各项显示已就绪或内置归档完整。
权限与凭据：引擎本身不需要密钥，直接生成 .pptx 不需要任何 AI 配置。引擎的 `.env` 只用于 AI 生图和语音旁白，两者都不配也能正常导出 PPTX。
安全提示：**引擎的受保护文件不得修改**（`LICENSE`、`SPONSORS*.md`、`SKILL.md`、`scripts/`、`templates/`、`references/`、`workflows/`）。引擎自带的完整性守卫会校验 LICENSE 摘要、SKILL.md 身份字段与门禁标记、入口脚本的调用结构，改动会阻断整条 PPTX 路线。需要换引擎位置时只改 `PPT_MASTER_DIR`，不移动或编辑引擎内的文件。

### 3.5 internet-access

- 官方站点：https://www.gov.cn/
- 官方站点：https://www.moe.gov.cn/

本 skill 不内置任何网络请求脚本，也不需要安装任何东西。遇到素材库（`references/material-library.md`）里没有的主题时，核实由调用方在生成时用自己的联网检索能力完成，优先查政府网站、共青团中央、新华社等权威来源，交叉确认后再写入材料。

验证方式：材料里凡素材库知识卡以外的年度主题、口号和统计数据，都能指回具体官方来源；指不回去的一律不写。
权限与凭据：读取公开信息，不需要账号或密钥。
安全提示：不向第三方发送通知原文、活动材料或个人数据。

## 4. 缺失时的降级总表

| 缺什么 | 还能做什么 | 结论怎么标注 |
|---|---|---|
| `python-docx` | 交付 Markdown 正文；人工核对占位与章节 | 标「未生成 DOCX」「版式未验证」 |
| `ms-word-com` | 结构审计 + OOXML 检查；有 PDF 就保留 | 标「未做渲染检查」 |
| `pymupdf` | 保留 PDF 供人工翻页 | 标「未逐页检查」 |
| `ppt-master-engine` | 改到非 C 盘指定缓存目录重试自举；归档缺失或摘要不符时，停止 PPT 环节，其余交付物照常产出 | 标「PPT 未生成」 |
| `internet-access` | 只用通知与素材库已有表述 | 标「主题表述未联网核实」 |

核心正确性无法保障时停止操作，不用推测或编造填补缺失环节；降级结果一律标注限制和验证状态。

## 5. 恢复与撤销

- 依赖装错了版本：`python -m pip install --force-reinstall <包名>`，然后重跑环境检查。
- 要撤销安装：`python -m pip uninstall <包名>`。撤销后本 skill 按第 4 节降级运行，不影响已有交付件。
- 引擎要用外部副本（例如本机的完整上游克隆）：设 `PPT_MASTER_DIR` 指向该引擎目录即可，内置归档不再使用，不需要改本 skill 的任何文件。
- 引擎缓存要换位置：设 `PPT_MASTER_CACHE_DIR`，或运行 `bootstrap_engine.py --cache-dir <目录>`。删掉缓存目录不影响本 skill，下次使用时按归档重新解压。缓存不完整或上次解压中断时，重跑 `bootstrap_engine.py` 会检出并修复；`--check` 报 `needs_repair` 就是这个状态。
- 引擎要更新：本 skill 不代管引擎更新。需要升级时用上游仓库拉起新版，重新生成内置归档与清单（`vendor/ppt-master-6.4.0.manifest.json`），再跑环境检查确认摘要通过。
