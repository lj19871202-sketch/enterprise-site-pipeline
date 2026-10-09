---
name: enterprise-site-pipeline
description: 在 Codex 本机把一份包含企业名称的 Excel 批量变成企业官网资料包：先摄入用户提供的任意结构资料（图片/Word/Excel/PDF/PPT 等），再自动发现官网、抓取页面补充，归档四类图片、生成中英双语企业简介 docx、中英双语产品清单 xlsx、汇总表和质量门禁结果。适用于四川源头工厂、德阳规上企业等一企一档批量产出场景；不适用于单篇文档写作、单个网页视觉设计或纯数据统计。
metadata:
  short-description: 本地企业官网资料包批量流水线
---

# 本地企业官网资料包流水线

输入一份本地 Excel，在当前 Windows 机器上完成：

**企业名称 → 用户资料摄入 → 官网发现 → 页面抓取 → 内容与图片提取 → docx/xlsx 生成 → 质量门禁 → 发布**

全程不依赖 SSH、远端 worker、opencode 临时目录或长期会话。项目事实源和交付物都落盘在当前运行目录。

## 输入与输出

输入 Excel 至少包含一列企业名称。列名可以是 `企业名称`、`公司名称`、`单位名称` 或 `名称`。建议额外提供 `官网` / `网址` / `网站` 列；没有官网列时脚本会按企业名称在本机自动搜索、抓取候选首页并按公司名命中度分级（高/中置信度自动采用，低置信度由 Codex 打开站点复核）。

输出目录结构：

```text
<输出目录>/
├── build/
│   └── <run_id>/
│       ├── raw/                     # 每家企业抓取事实
│       ├── raw_home/                # 用户资料 + 官网图片暂存
│       ├── deliverable/             # 每家企业交付目录
│       ├── en.json                  # 英文层：自动翻译草稿，可被 --en 覆盖
│       ├── 视觉核对.json             # 图片/文档/产品的视觉核对结论
│       ├── review/                  # 拼版图 + 缩略图核对表 + 核对指引.md
│       ├── 汇总.xlsx
│       ├── gates.json
│       ├── gates.log
│       └── manifest.json
├── deliverable/                     # 门禁通过后才发布
└── 原始资料备份/                     # 用户提交的原始资料原件，保留原目录结构
```

每家企业目录包含：

- `1.企业工厂图/`
- `2.企业产品图/`
- `3.企业logo/`
- `4.资质证书/`
- `5.企业介绍/<企业名>简介（日期短）.docx`
- `产品清单.xlsx`

企业根目录只允许上述五个文件夹和 `产品清单.xlsx`；五个文件夹内只放直接文件，不得再嵌套子目录。`structure` 门禁会强制检查这一层结构。无论用户提交的资料结构多乱、格式多少种，最终交付都必须是这「五个文件夹 + 一个 excel」。

## 核心原则

1. **本地单一入口。** 用 `scripts/local_pipeline.py` 或 `scripts/run_local.ps1` 执行；不要再拆成远端 SSH 命令。
2. **用户资料优先，官网补充。** 用户可能通过 `--resources` 提交资料：总目录下每家一个**任意命名**子文件夹，也可能总目录本身就是一个企业文件夹；资料**不要求按规范结构**提交，格式除图片/Word/Excel 外还可能是 PDF、PPT、txt 等。脚本递归扫描并用企业名匹配归属，图片按类别归档、文档抽文本、表格抽行列结构；冲突时**用户资料优先，官网只作补充**。原始资料一律在交付父文件夹之外备份（默认 `<输出目录>\原始资料备份`，保留用户原目录结构），便于追溯，绝不覆盖或丢弃。
3. **官网自动发现。** 只要 Excel 有企业名称，没有官网列也要自动找到官网：多引擎搜索 + 候选首页抓取 + 公司名/联系方式/备案命中打分。高、中置信度自动采用；低置信度必须在交付前由 Codex 打开站点复核，不能因为“没有官网列”就退回让用户自己找。官网找不到或不可访问但用户资料可用时，状态记为 `resource_only`，即"仅凭用户资料成档"，不算失败。
4. **渲染默认可用，且不依赖大体积下载。** `run_local.ps1` 会优先复用本机已有 Playwright；渲染内核按 **内置 Chromium → 本机 Microsoft Edge → Edge 绝对路径** 自动选择。Windows 自带 Edge 且与 Chromium 同源，通常**无需下载 100-200MB 的 Chromium**；两者都没有才下载。抓取默认 `auto`：静态页面直接抓，疑似 JS 渲染页自动用 Playwright 重抓。内核与 Edge 都不可用时，可由 Codex 用内置浏览器保存 HTML，再以 `--html-dir` 兜底。
5. **Excel 是输入源。** 企业名称和可选官网来自 Excel；不把聊天上下文当事实源。
6. **抓取有证据。** 官网、页面标题、原始文本、图片 URL 和处理状态写入 `raw/*.json`。
7. **默认中英双语，机翻须标注。** 不提供 `--en` 时，脚本自动把企业名、三段简介、产品名和产品详情翻成英文草稿并写入 docx/xlsx，同时在 `en.json` 标 `自动翻译: true`、`定稿: false`、`备注: 待人工核校`；交付时必须说明“英文为自动翻译草稿”。用户确认的英文通过 `--en` 覆盖并标记为定稿。若 MyMemory 限流导致自动翻译失败，Codex 必须基于已抓取的中文事实补出英文草稿，写成 `--en` 兼容 JSON，并显式保留 `定稿: false`、`自动翻译: true`、`备注: 待人工核校`；不得直接交付空英文，也不得把模型草稿标成人工定稿。产品清单的 `图片（本地连接）` 列写入 `2.企业产品图/` 下的真实相对路径并设为可点击链接；官网确实没有对应产品图时写中英双语无图说明，不得保留占位符。
8. **门禁即代码。** 交付前必须运行 `scripts/gates.py`；门禁未全绿不得发布为终稿。
9. **视觉核对由 Codex 自己完成。** 机器能判“文件对得上”，判不了“图到底是什么”。流水线自动生成拼版、缩略图核对表和 `review/核对指引.md`；Codex 必须用图像查看工具打开总览和每个分类拼版，逐张判断是否属于该企业、该分类，再用 `visual_review.py --apply` 回写结论。不得要求用户代看，不得把 `待核对` 留给用户后声称已发布终稿；标“不符”即门禁报红，先修图再交付。
10. **状态闭环。** 每家企业保留 `ok` / `partial` / `empty_images` / `no_website` 状态和原因，不静默丢失企业。

## 本地执行入口

优先读 [references/codex-execution.md](references/codex-execution.md)，然后运行：

**先解析 skill 根目录**（装到任意路径只需改这一行，脚本按自身位置定位其余文件）：

```powershell
$SkillRoot = "$env:USERPROFILE\.codex\skills\enterprise-site-pipeline"
```

新环境第一步先引导：创建技能独立 `.venv`，按 `requirements.lock.txt` 安装已验证版本，并运行环境自检。`run_local.ps1` 缺核心依赖时也会自动调用它：

```powershell
& "$SkillRoot\bootstrap.ps1"
```

之后可单独复跑环境自检（依赖 / 网络 / 翻译 / 官网可达性）：

```powershell
& "$SkillRoot\.venv\Scripts\python.exe" "$SkillRoot\scripts\local_pipeline.py" --selftest --excel "D:\path\企业名录.xlsx"
```

渲染内核兜底顺序（不强制下载 Chromium）：

1. `run_local.ps1` 先复用本机 Playwright；启动内核时按 **内置 Chromium → 本机 Edge** 依次尝试。用 `--browser-probe` 可确认，输出 `msedge` / `msedge-exe` 即本机 Edge 可用。
2. 输出 `none` 且用户不接受下载 Chromium 时，Codex 必须用内置浏览器打开目标页，保存渲染后的完整 HTML，并写 `manifest.json`，运行时加 `-HtmlDir`；不得因为内核缺失直接放弃抓取。
3. 只有前两步都不可用且用户同意下载时，才执行 `python -m playwright install chromium`。

一键运行（无用户资料）：

```powershell
& "$SkillRoot\scripts\run_local.ps1" `
  -Excel "D:\path\企业名录.xlsx" `
  -Out "D:\path\企业官网资料包" `
  -Limit 3
```

一键运行（含用户资料，资料目录结构不限）：

```powershell
& "$SkillRoot\scripts\run_local.ps1" `
  -Excel "D:\path\企业名录.xlsx" `
  -Out "D:\path\企业官网资料包" `
  -Resources "D:\path\企业原始资料" `
  -Limit 3
```

`<输出目录>\原始资料备份\` 会保留用户提交的全部原件及原目录结构，位于 `deliverable` 之外。

也可以直接调用 Python：

```powershell
python "$SkillRoot\scripts\local_pipeline.py" `
  --excel "D:\path\企业名录.xlsx" `
  --out "D:\path\企业官网资料包" `
  --limit 3
```

推荐：先构建不发布 → Codex 看图核对 → 门禁确认 → 发布。

```powershell
# 1. 构建（playwright 默认 auto；不加 -RequireVisual，先产出核对材料）
& "$SkillRoot\scripts\run_local.ps1" -Excel "D:\path\企业名录.xlsx" -Out "D:\path\企业官网资料包" -NoPublish

# 2. Codex 打开 <输出目录>\build\<run_id>\review\核对指引.md，逐张看图后：
python "$SkillRoot\scripts\visual_review.py" --json "<输出目录>\build\<run_id>\视觉核对.json" --apply "<verdicts.json>"

# 3. 视觉门禁必须通过
python "$SkillRoot\scripts\gates.py" --deliverable "<输出目录>\build\<run_id>\deliverable" --raw "<输出目录>\build\<run_id>\raw" --en "<输出目录>\build\<run_id>\en.json" --summary "<输出目录>\build\<run_id>\汇总.xlsx" --visual "<输出目录>\build\<run_id>\视觉核对.json" --require-visual

# 4. 发布
python "$SkillRoot\scripts\local_pipeline.py" --publish-stage "<输出目录>\build\<run_id>" --out "D:\path\企业官网资料包"
```

常用参数：

| 参数 | 用途 |
|---|---|
| `--selftest` | 只做环境自检后退出，不跑流水线（可不带 `--excel`） |
| `--limit N` | 先跑前 N 家，`0` 为全部 |
| `--playwright auto\|on\|off` | 抓取模式：`auto`=静态过薄时自动用 Playwright（默认），`on`=强制渲染，`off`=只用静态 |
| `-NoPlaywright` / `-NoPlaywrightInstall` | 关闭 Playwright；或只检测不自动安装 |
| `-NoPublish` | 先构建、不发布，等 Codex 视觉核对后再发布 |
| `--publish-stage <build/run_id>` | 视觉核对与门禁通过后，把已完成构建目录发布到 `--out` |
| `--require-visual` / `-RequireVisual` | 视觉核对未完成按 error 处理，不允许直接发布 |
| `--proxy` | 本机 HTTP(S) 代理 |
| `--en <json>` | 用人工确认英文或 Codex 补译草稿覆盖自动翻译；模型草稿必须保留 `定稿: false`、`自动翻译: true` |
| `-HtmlDir <目录>` / `--html-dir <目录>` | 用 Codex 内置浏览器保存的离线 HTML 兜底抓取（目录内需 `manifest.json`） |
| `-Resources <目录>` / `--resources <目录>` | 用户资料总目录：每家企业一个任意命名子文件夹；单家企业时也可直接传该企业文件夹。资料不要求规范结构，支持图片/Word/Excel/PDF/PPT/txt 等 |
| `-BackupDir <目录>` / `--backup-dir <目录>` | 原始资料备份目录，默认 `<输出目录>\原始资料备份`，位于 `deliverable` 之外，保留用户原目录结构 |
| `--no-translate` | 关闭自动翻译，只输出中文并把英文层留空；未显式接受中文版时仍会阻断发布 |
| `--accept-no-english` / `-AcceptNoEnglish` | 仅由用户明确接受中文版时使用；英文相关门禁降为告警，不能由 Codex 自行默认开启 |
| `-NoBootstrap` | 禁止 `run_local.ps1` 自动创建 `.venv`，用于已确认自行管理依赖的环境 |
| `--translate-email <邮箱>` | 可选，MyMemory 联系邮箱，用于提高匿名翻译额度 |
| `--no-visual-review` | 跳过拼版/核对表和视觉核对.json 生成 |
| `--strict` | 门禁未通过时返回非零退出码（`run_local.ps1` 默认启用；用 `-AllowRed` 关闭） |

## 流水线阶段

| # | 阶段 | 本地脚本行为 | 交付判据 |
|---|---|---|---|
| 0 | 读表 | 识别企业名称列和可选官网列 | 至少一条企业记录 |
| 0.5 | 用户资料摄入 | 递归扫描 `--resources`：图片按目录名/文件名关键词归入四类；文档（docx/xlsx/xlsm/csv/pptx/pdf/txt）抽文本与表格行列；原件在 `deliverable` 之外备份；未归类/抽取失败记入 `资料备注` | 各类资料尽可能归档；未归类项已列出，原始件已备份 |
| 1 | 官网发现 | 用户官网优先；否则多引擎自动发现并抓取候选首页，按公司名命中度分级；低置信度交由 Codex 复核 | 命中且置信度可接受；找不到但有用户资料时标 `resource_only`；既无官网又无资料才标 `no_website` |
| 2 | 页面抓取 | 首页 + 关于/工厂/资质页；产品分类/列表页优先并下钻一层补齐叶子分类与详情，保留标题、正文、链接、图片 | `raw/*.json` 有页面记录 |
| 3 | 内容与图片 | 抽取中文简介、产品名、产品详情、主营和地址；只有详情或产品图佐证的条目才作为产品；识别站头 logo 与 CSS 背景横幅，图片按四类落盘。**用户资料优先、官网补充**：用户图片排在官网图片之前，用户文档事实覆盖官网推断；官网置信度为"低"时不并入官网产品 | 图片引用真实存在、产品详情中英双语、产品图本地链接有效 |
| 4 | 渲染 | 生成中英双语企业 docx、产品清单 xlsx、汇总 xlsx；英文缺口写 `英文补译清单.json` | 文件可打开且结构完整；有英文缺口时补齐或取得用户明确接受 |
| 5 | 视觉核对 | 生成拼版、核对表和核对指引；Codex 自己看图，经 `visual_review.py --apply` 回写 `视觉核对.json` | 无“不符”，核对完成 |
| 6 | 门禁 | 自动调用 `gates.py` | `gates.json` 无 error |

## 开工顺序

1. **确认 Excel 与资料目录。** 先核对表头、企业数、是否已有官网列；不要猜测企业简称。没有官网列时按企业名称自动发现，不再要求用户先提供域名。若用户已有资料，用 `-Resources` 指向总目录（每家一个任意命名子文件夹，或单家企业文件夹）；资料目录结构、命名和格式都不受限制，由脚本负责匹配与分类。
2. **小样本试跑。** 用 `-Limit 1` 或 `-Limit 3 -NoPublish`，检查官网命中置信度、图片归档、docx/xlsx、`英文补译清单.json` 和门禁日志；低置信度官网由 Codex 打开确认。
3. **完整运行。** 小样本通过后去掉 `-Limit` 运行全部企业（推荐 `-NoPublish`，先构建待核对）。
4. **视觉核对（Codex 自己做）。** 打开 `review/核对指引.md`，用图像查看工具逐张看 `review/视觉核对图/<企业>/0.总览.png` 和每个分类拼版；把结论写成 `verdicts.json`，执行 `visual_review.py --apply` 回写，再跑 `gates.py --require-visual`。核对完成后用 `local_pipeline.py --publish-stage <build/run_id> --out <输出目录>` 发布。
5. **复核门禁。** 先看 `gates.json` 的 error 项，再看 `empty_images`、`no_website`、`partial` 清单。
6. **核对英文。** 默认会写入自动翻译英文草稿并标注待人工核校。MyMemory 429/限流是 `WARN` 可恢复告警，不把新环境判为不可用；英文缺口会写入 `<run_id>/英文补译清单.json`，Codex 必须基于 raw 中文事实补出 `--en` 兼容草稿后重跑。用户在当次对话中明确接受中文版时，才可加 `--accept-no-english` / `-AcceptNoEnglish`。
7. **发布。** 默认英文缺失会阻断发布；只有补 `--en` 或用户明确接受中文版后才交付。门禁红色先修数据或补抓，不下调阈值迁就数据。

## 硬性约束

- **不要 SSH 到任何远端。** 所有发现、抓取、渲染和门禁都在本机运行。
- **不要记录密码、token、cookie。** 不要求用户把凭据写入 skill、Excel、日志或命令行历史。
- **不臆造企业事实。** 产业、业务、产品、资质只能来自官网快照或用户资料。
- **用户资料优先，且必须备份。** 冲突时以用户提交的资料为准，官网只作补充；所有原件必须在 `deliverable` 之外的备份目录保留原目录结构，不得覆盖、丢弃或改动原件。
- **不要求用户改资料结构。** 用户资料可能是任意目录层级、任意命名、任意格式（图片/Word/Excel/PDF/PPT/txt…）。不得因为"结构不规范""格式不在清单里"就拒收或跳过；无法归类或抽取失败的项要记入 `资料备注` 并保留原件。
- **无官网列要自动发现，低置信度必须复核。** 输入表缺官网列时，先由脚本自动搜索、抓取和打分；高/中置信度自动采用，低置信度必须由 Codex 打开站点确认后再交付，不得直接跳过该企业。
- **依赖必须齐全且锁定。** 新环境先运行 `bootstrap.ps1`；它创建技能独立 `.venv`，按 `requirements.lock.txt` 安装 `openpyxl`、`python-docx`、`Pillow`、`playwright` 的已验证版本。`requirements.txt` 只提供带主版本上限的可更新范围。`run_local.ps1` 缺核心依赖时会自动引导。Pillow 缺失会让视觉核对材料静默消失，不能被当作“没有疑点”。该引导需联网（从 PyPI 安装）；已在无 `.venv` 的干净克隆上实测——含官网与"仅资料成档"两类场景均门禁全绿，交付严格为"五文件夹 + 产品清单.xlsx"、文件夹内零子目录，非图片原件只存 `deliverable` 之外的备份。
- **Playwright 复用优先。** `run_local.ps1` 先复用当前 Python/技能 `.venv` 中已有的 Playwright；渲染内核按内置 Chromium → 本机 Edge 自动选择，两者都不可用才考虑安装 Chromium。可用 Codex 内置浏览器保存 HTML 并经 `-HtmlDir` 兜底。
- **机翻不等于定稿。** 自动英文来自机翻，必须在 `en.json`/汇总/交付说明中保留 `自动翻译: true`、`待人工核校` 标记，不得当作人工定稿交付。
- **不把空结果当成功。** 四类图全空、官网未确认、简介不足三段都要显式记录并单独列出。
- **不跳视觉核对，也不把核对推给用户。** 图片内容是否属于该企业、该分类，只能看图判断；Codex 必须自己查看拼版并回写结论，不得在未核对时声称图片已核验，也不得把 `待核对` 留给用户后发布终稿。
- **不修改无关脚本。** 只处理当前请求涉及的小样本或输入名录。

## 参考

- [references/codex-execution.md](references/codex-execution.md) —— 本地执行、参数、故障排查
- [references/pipeline.md](references/pipeline.md) —— 各阶段行为与详细命令
- [references/data-contract.md](references/data-contract.md) —— Excel、raw JSON、英文层和交付目录契约
- [references/quality-gates.md](references/quality-gates.md) —— 门禁判据、阈值与视觉核对流程
