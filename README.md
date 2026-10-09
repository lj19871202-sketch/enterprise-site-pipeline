# 企业官网资料包流水线

本地 Excel 驱动的企业官网批量抓取与资料包生成工具。给它一份包含企业名称的 `.xlsx`，它会在 Codex 当前 Windows 机器上完成：

- 摄入用户提交的任意结构资料（图片 / Word / Excel / PDF / PPT / txt 等），按企业匹配、按类别归档，原件备份在交付之外；
- 官网发现与校验（用户资料优先，官网补充）；
- 首页、关于、产品、工厂、资质等页面抓取；
- 四类图片归档；
- 中文简介、产品名、产品详情和主营信息提取；
- 自动中译英，生成中英双语企业简介 docx、产品清单 xlsx、汇总 xlsx；
- 质量门禁和运行清单。

不依赖 SSH、远端 worker、opencode 临时目录或常驻服务。

## 快速开始

先在 Excel 中准备至少一列企业名称；`官网` 列可选。没有官网列时，脚本会按企业名称在本机自动多引擎搜索、抓取候选首页并按命中度评分：

| 企业名称 | 官网（可选） |
|---|---|
| 示例科技有限公司 | https://example.com |
| 示例实业有限公司 | |

高/中置信度自动采用；低置信度会在门禁里列为需复核，由 Codex 打开站点确认后再交付。

企业已有资料时，用 `-Resources`（CLI：`--resources`）指向资料总目录：总目录下每家企业一个**任意命名**子文件夹，单家企业时也可直接传该企业文件夹。资料**不要求按规范结构**提交，格式除图片/Word/Excel 外还可能是 PDF、PPT、txt 等。冲突时**用户资料优先，官网只作补充**；原件会备份到 `<输出目录>\原始资料备份`（在 `deliverable` 之外，保留原目录结构）。

新环境先引导：创建技能独立 `.venv`，按 `requirements.lock.txt` 安装锁定依赖并自检（需联网；已在干净克隆上实测通过）。`run_local.ps1` 缺核心依赖时也会自动调用：

```powershell
$SkillRoot = "$env:USERPROFILE\.codex\skills\enterprise-site-pipeline"
& "$SkillRoot\bootstrap.ps1"
```

Windows PowerShell 一键运行：

```powershell
# skill 默认装在 $env:USERPROFILE\.codex\skills\enterprise-site-pipeline
# 若装在别处只改这一行；run_local.ps1 按自身路径定位 local_pipeline.py
$SkillRoot = "$env:USERPROFILE\.codex\skills\enterprise-site-pipeline"
& "$SkillRoot\scripts\run_local.ps1" `
  -Excel "D:\path\企业名录.xlsx" `
  -Out "D:\path\企业官网资料包" `
  -Limit 3
```

小样本检查通过后，去掉 `-Limit 3` 跑全量：

```powershell
$SkillRoot = "$env:USERPROFILE\.codex\skills\enterprise-site-pipeline"
& "$SkillRoot\scripts\run_local.ps1" `
  -Excel "D:\path\企业名录.xlsx" `
  -Out "D:\path\企业官网资料包"
```

也可以直接运行 Python：

```powershell
$SkillRoot = "$env:USERPROFILE\.codex\skills\enterprise-site-pipeline"
python "$SkillRoot\scripts\local_pipeline.py" `
  --excel "D:\path\企业名录.xlsx" `
  --out "D:\path\企业官网资料包" `
  --limit 3
```

环境自检（依赖 / 网络 / 翻译 / 官网可达性）：

```powershell
& "$SkillRoot\.venv\Scripts\python.exe" "$SkillRoot\scripts\local_pipeline.py" --selftest --excel "D:\path\企业名录.xlsx"
```

如果本机 `python` 不在 PATH，先运行 `bootstrap.ps1`，或用 `$env:CODEX_PYTHON` 指定解释器。手工安装依赖时优先使用锁定版本：`<python.exe> -m pip install -r "$SkillRoot\requirements.lock.txt"`；`requirements.txt` 只提供带主版本上限的可更新范围。MyMemory 429/限流在自检中是 `WARN` 可恢复告警，不代表环境不可用。

## 输出

```text
<输出目录>/
├── build/<run_id>/
│   ├── raw/                 # 每家企业抓取事实
│   ├── raw_home/            # 用户资料 + 官网图片暂存（用户图 user_ 前缀）
│   ├── deliverable/         # 企业文件夹、docx、产品清单
│   ├── en.json              # 英文层：自动翻译草稿，可被 --en 覆盖
│   ├── 英文补译清单.json     # 英文缺口清单：Codex 生成 --en 补稿的依据
│   ├── 视觉核对.json         # 图片/文档/产品的视觉核对结论
│   ├── review/              # 拼版图 + 缩略图核对表 + 核对指引.md
│   ├── 汇总.xlsx
│   ├── gates.json
│   ├── gates.log
│   └── manifest.json
├── deliverable/             # 门禁通过后发布
└── 原始资料备份/             # 用户提交的原始资料原件，保留原目录结构
```

每家企业默认生成四类图片目录，以及中英双语的 `5.企业介绍/*.docx` 和 `产品清单.xlsx`（英文来自自动翻译，`en.json` 标 `自动翻译: true`，待人工核校）。企业根目录只允许这五个文件夹和 `产品清单.xlsx`，五个文件夹内只放直接文件，不得再嵌套——无论用户提交的资料结构多乱、格式多少种，最终交付都固定为这「五个文件夹 + 一个 excel」。渲染后自动生成图片拼版与 `图片核对表.xlsx`，视觉核对结论写入 `视觉核对.json` 供门禁校验。

## 常用选项

| 选项 | 说明 |
|---|---|
| `-Limit N` / `--limit N` | 只处理前 N 家，建议先用 1–3 家试跑 |
| `-Playwright` / `--playwright on` | 强制用 Playwright 渲染 JS 站点 |
| `--playwright auto` | 默认：静态页面直接抓，内容过薄时自动改用 Playwright |
| `-NoPlaywright` / `--playwright off` | 关闭 Playwright，只用静态抓取 |
| `-NoPlaywrightInstall` | 只检测本地 Playwright，不自动安装 |
| `-NoPublish` | 先构建不发布，便于 Codex 先做视觉核对 |
| `--publish-stage <build/run_id>` | 核对与门禁通过后，把已完成构建目录发布到 `--out` |
| `-RequireVisual` / `--require-visual` | 视觉核对未完成按 error 处理 |
| `--proxy` | 本机 HTTP(S) 代理 |
| `-En <json>` / `--en <json>` | 用人工确认稿或 Codex 补翻草稿覆盖自动翻译（模型草稿必须保留 `定稿: false`、`自动翻译: true`） |
| `-HtmlDir <目录>` / `--html-dir <目录>` | 用 Codex 内置浏览器保存的离线 HTML 兜底抓取（目录内需 `manifest.json`） |
| `-Resources <目录>` / `--resources <目录>` | 用户资料总目录：每家企业一个任意命名子文件夹，或单家企业文件夹；结构与格式不限 |
| `-BackupDir <目录>` / `--backup-dir <目录>` | 原始资料备份目录，默认 `<输出目录>\原始资料备份`，位于 `deliverable` 之外 |
| `-NoTranslate` / `--no-translate` | 关闭自动翻译，只输出中文；未显式接受中文版时仍阻断发布 |
| `-AcceptNoEnglish` / `--accept-no-english` | 仅由用户明确接受中文版时使用；英文相关门禁降为告警 |
| `-NoBootstrap` | 禁止 `run_local.ps1` 自动创建 `.venv` |
| `--translate-email <邮箱>` | 可选，提高 MyMemory 匿名翻译额度 |
| `-NoVisualReview` / `--no-visual-review` | 跳过拼版、核对表和 `视觉核对.json` 生成 |
| `--strict` | 门禁不通过时返回非零退出码（`run_local.ps1` 默认启用；CLI 直跑需显式加） |
| `-AllowRed` | 关闭 `run_local.ps1` 的默认严格模式，门禁红仍返回 0 |

## 质量门禁

门禁在每次运行后自动执行，结果写入 `build/<run_id>/gates.json` 和 `gates.log`。

- 图片引用必须存在、交付目录没有 raw 未记录的孤儿图、四类图片不能全空；
- `logo` 或 `factory` 单独为空时给出告警，要求补图或标记数据边界；
- 每张图的来源页必须与官网同域（图片直链走 CDN 仅提示）；
- 每家企业必须是一层结构：四类图片目录 + `5.企业介绍/` + `产品清单.xlsx`；企业根目录无多余项，五个文件夹内不得嵌套子目录；
- docx 中文段必须等于 raw `intro_paragraphs`，产品清单数据行必须等于 raw `products`，产品详情列必须中英双语且无空值；`图片（本地连接）` 列必须指向交付目录内真实存在的 `2.企业产品图/...` 文件，官网无图时显式标注，不得保留旧占位符；
- 英文默认自动中译英并写入正文，`en.json` 标 `自动翻译: true` 待人工核校；提供 `--en` 后覆盖，人工确认稿标 `定稿: true`，MyMemory 失败时由 Codex 补翻的草稿仍标 `定稿: false`；
- 英文缺口会写入 `英文补译清单.json`；默认 `en_entry`、`en_ascii`、`product_en`、`product_detail`、`product_map` 的缺失均为 error，必须补 `--en` 后重跑。只有用户明确接受中文版时，`--accept-no-english` 才把这些英文检查降为告警；
- 官网必须找到；自动发现置信度为低/中时列出，低置信度由 Codex 打开站点复核。官网未确认或不可访问但用户资料可用时，记 `resource_only`（仅凭用户资料成档），降为告警；
- `resource_intake` 门禁检查用户资料摄入：未归类 / 抽取失败项会列出告警（原件已备份）；用户资料图跳过官网同域校验（`image_provenance`）；
- 汇总表、档案、目录和英文集合必须一致；
- 图片内容需视觉核对，结论写入 `视觉核对.json`：标“不符”即报红，未核对默认告警；加 `--require-visual` 后未核对直接报红。

## 视觉核对

流水线渲染后自动生成：

```text
build/<run_id>/review/视觉核对图/<企业>/0.总览.png    # 四类速览
build/<run_id>/review/视觉核对图/<企业>/<分类>.png     # 分类拼版（单页 ≤12 张、≤1MB；超出为 <分类>_p1.png、_p2.png…）
build/<run_id>/review/<企业>/图片核对表.xlsx           # 带缩略图，结论列可下拉
build/<run_id>/视觉核对.json                           # 结论载体
```

这一步由 Codex 自己做，不能把拼版甩给用户代看。开工前先跑 `python "$SkillRoot\scripts\session_guard.py"` 查本会话 rollout 体积，超过 20MB 先另开会话。然后读 `review/核对指引.md`：图片多时先用 `图片核对表.xlsx` 初筛，按指引标注的**序号范围**只打开可疑的 `<分类>_pN.png` 放大确认（同一分类被点名的页都要看）。拼版图进入会话后是 base64，单张过大或累计过多会撑爆请求（历史故障：19 张拼版 28.8MB 触发上游报错），所以必须短线程分批：每批 ≤10 家，一个会话做完「看图 → 回写 → 门禁 → 发布」就结束，企业多时另开会话。结论写成 `verdicts.json` 后一键回写：

```powershell
# 看图前先查会话体积（>20MB 先另开会话）
python "$SkillRoot\scripts\session_guard.py"
python "$SkillRoot\scripts\visual_review.py" `
  --json "<输出目录>\build\<run_id>\视觉核对.json" `
  --apply "<verdicts.json>"
python "$SkillRoot\scripts\gates.py" --deliverable "<...>\deliverable" --raw "<...>\raw" `
  --en "<...>\en.json" --summary "<...>\汇总.xlsx" --visual "<...>\视觉核对.json" --require-visual
```

类别结论可被单张覆盖；重跑会保留已有结论。全部核对完成并通过门禁后，用 `local_pipeline.py --publish-stage "<...>\build\<run_id>" --out "<输出目录>"` 发布。

门禁红色先修数据或补抓，不要下调阈值迁就数据。

## 已知边界

- 抓取默认 `auto`：静态页面直接抓，疑似 JS 渲染页自动用 Playwright 重抓。`run_local.ps1` 会优先复用本机已有 Playwright；渲染内核先试内置 Chromium，不可用时自动改用本机 Microsoft Edge（Windows 自带，不需要下载 Chromium），两者都没有才下载。网络受限无法安装时加 `-NoPlaywrightInstall` 显式降级为静态抓取；Chromium/Edge 都不可用时，可用 Codex 内置浏览器保存渲染后的 HTML，加 `-HtmlDir` 走离线快照兜底。
- 自动官网发现依赖搜索引擎可达性；置信度低时 Codex 必须打开候选站点复核，完全找不到才标 `no_website`。
- 英文默认走 MyMemory 免费接口自动中译英。该接口有每日匿名额度，批量较大时部分条目可能翻译失败：失败条目在 `en.json` 标 `翻译失败: true`，并写入 `英文补译清单.json`；自检中的 429 是可恢复 `WARN`，不是环境致命错误。默认门禁阻断发布，Codex 必须基于中文事实生成 `--en` 补译草稿并保留 `定稿: false`、`自动翻译: true`、`备注: 待人工核校`。只有用户明确接受中文版时才可使用 `--accept-no-english`。
- 输入格式当前以 `.xlsx` / `.xlsm` 为主；老式 `.xls` 请先另存为 `.xlsx`。
- **新环境可复现性（已实测）。** 在全新克隆、无 `.venv` 的目录上跑 `bootstrap.ps1`（从 PyPI 装锁定依赖）后再用 `run_local.ps1`，含官网场景与"仅资料、无官网"（`resource_only`）场景均门禁全绿、退出码 0；交付严格为"五文件夹 + 产品清单.xlsx"、五个文件夹内零子目录，PDF/PPT/DOCX/XLSX 原件只落在 `deliverable` 之外的 `原始资料备份/`。该结论是干净克隆模拟，不等同于全新物理机。
- **必须联网。** `bootstrap.ps1` 从 PyPI 安装依赖、官网抓取、MyMemory 翻译都依赖网络；断网环境需预先离线装好依赖与渲染内核，并自行准备离线 HTML 快照（`-HtmlDir`）。
- **PDF 抽取是尽力而为。** 环境有 `pypdf`/`PyPDF2` 时抽前 20 页；没有时 PDF 只归档并在 `资料备注` 记"PDF 未抽到文本"，不阻断主流程。PPTX 用 zipfile+XML 抽文本，不依赖 `python-pptx`。
- **必须由 Codex 复核的项**（脚本只给告警，不算失败）：官网自动发现置信度为低/中时须打开站点确认；`visual_review` 必须看图回写结论后才能发布；图片直链走 CDN 的外域需人工确认。

详细说明见 `SKILL.md` 和 `references/`。
