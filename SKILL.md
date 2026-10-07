---
name: enterprise-site-pipeline
description: 在 Codex 本机把一份包含企业名称的 Excel 批量变成企业官网资料包：自动发现官网、抓取页面、归档四类图片、生成中英双语企业简介 docx、中英双语产品清单 xlsx、汇总表和质量门禁结果。适用于四川源头工厂、德阳规上企业等一企一档批量产出场景；不适用于单篇文档写作、单个网页视觉设计或纯数据统计。
metadata:
  short-description: 本地企业官网资料包批量流水线
---

# 本地企业官网资料包流水线

输入一份本地 Excel，在当前 Windows 机器上完成：

**企业名称 → 官网发现 → 页面抓取 → 内容与图片提取 → docx/xlsx 生成 → 质量门禁 → 发布**

全程不依赖 SSH、远端 worker、opencode 临时目录或长期会话。项目事实源和交付物都落盘在当前运行目录。

## 输入与输出

输入 Excel 至少包含一列企业名称。列名可以是 `企业名称`、`公司名称`、`单位名称` 或 `名称`。建议额外提供 `官网` / `网址` / `网站` 列；没有官网列时脚本会按企业名称在本机自动搜索、抓取候选首页并按公司名命中度分级（高/中置信度自动采用，低置信度由 Codex 打开站点复核）。

输出目录结构：

```text
<输出目录>/
├── build/
│   └── <run_id>/
│       ├── raw/                     # 每家企业抓取事实
│       ├── deliverable/             # 每家企业交付目录
│       ├── en.json                  # 英文层：自动翻译草稿，可被 --en 覆盖
│       ├── 视觉核对.json             # 图片/文档/产品的视觉核对结论
│       ├── review/                  # 拼版图 + 缩略图核对表 + 核对指引.md
│       ├── 汇总.xlsx
│       ├── gates.json
│       ├── gates.log
│       └── manifest.json
└── deliverable/                     # 门禁通过后才发布
```

每家企业目录包含：

- `1.企业工厂图/`
- `2.企业产品图/`
- `3.企业logo/`
- `4.资质证书/`
- `5.企业介绍/<企业名>简介（日期短）.docx`
- `产品清单.xlsx`

## 核心原则

1. **本地单一入口。** 用 `scripts/local_pipeline.py` 或 `scripts/run_local.ps1` 执行；不要再拆成远端 SSH 命令。
2. **官网自动发现。** 只要 Excel 有企业名称，没有官网列也要自动找到官网：多引擎搜索 + 候选首页抓取 + 公司名/联系方式/备案命中打分。高、中置信度自动采用；低置信度必须在交付前由 Codex 打开站点复核，不能因为“没有官网列”就退回让用户自己找。
3. **渲染默认可用，且不依赖大体积下载。** `run_local.ps1` 会优先复用本机已有 Playwright；渲染内核按 **内置 Chromium → 本机 Microsoft Edge → Edge 绝对路径** 自动选择。Windows 自带 Edge 且与 Chromium 同源，通常**无需下载 100-200MB 的 Chromium**；两者都没有才下载。抓取默认 `auto`：静态页面直接抓，疑似 JS 渲染页自动用 Playwright 重抓。内核与 Edge 都不可用时，可由 Codex 用内置浏览器保存 HTML，再以 `--html-dir` 兜底。
4. **Excel 是输入源。** 企业名称和可选官网来自 Excel；不把聊天上下文当事实源。
5. **抓取有证据。** 官网、页面标题、原始文本、图片 URL 和处理状态写入 `raw/*.json`。
6. **默认中英双语，机翻须标注。** 不提供 `--en` 时，脚本自动把企业名、三段简介、产品名和产品详情翻成英文草稿并写入 docx/xlsx，同时在 `en.json` 标 `自动翻译: true`、`定稿: false`、`备注: 待人工核校`；交付时必须说明“英文为自动翻译草稿”。用户确认的英文通过 `--en` 覆盖并标记为定稿。若 MyMemory 限流导致自动翻译失败，Codex 必须基于已抓取的中文事实补出英文草稿，写成 `--en` 兼容 JSON，并显式保留 `定稿: false`、`自动翻译: true`、`备注: 待人工核校`；不得直接交付空英文，也不得把模型草稿标成人工定稿。产品清单的 `图片（本地连接）` 列写入 `2.企业产品图/` 下的真实相对路径并设为可点击链接；官网确实没有对应产品图时写中英双语无图说明，不得保留占位符。
7. **门禁即代码。** 交付前必须运行 `scripts/gates.py`；门禁未全绿不得发布为终稿。
8. **视觉核对由 Codex 自己完成。** 机器能判“文件对得上”，判不了“图到底是什么”。流水线自动生成拼版、缩略图核对表和 `review/核对指引.md`；Codex 必须用图像查看工具打开总览和每个分类拼版，逐张判断是否属于该企业、该分类，再用 `visual_review.py --apply` 回写结论。不得要求用户代看，不得把 `待核对` 留给用户后声称已发布终稿；标“不符”即门禁报红，先修图再交付。
9. **状态闭环。** 每家企业保留 `ok` / `partial` / `empty_images` / `no_website` 状态和原因，不静默丢失企业。

## 本地执行入口

优先读 [references/codex-execution.md](references/codex-execution.md)，然后运行：

**先解析 skill 根目录**（装到任意路径只需改这一行，脚本按自身位置定位其余文件）：

```powershell
$SkillRoot = "$env:USERPROFILE\.codex\skills\enterprise-site-pipeline"
```

环境自检（依赖 / 网络 / 翻译 / 官网可达性，建议新环境第一步就跑）：

```powershell
python "$SkillRoot\scripts\local_pipeline.py" --selftest --excel "D:\path\企业名录.xlsx"
```

渲染内核兜底顺序（不强制下载 Chromium）：

1. `run_local.ps1` 先复用本机 Playwright；启动内核时按 **内置 Chromium → 本机 Edge** 依次尝试。用 `--browser-probe` 可确认，输出 `msedge` / `msedge-exe` 即本机 Edge 可用。
2. 输出 `none` 且用户不接受下载 Chromium 时，Codex 必须用内置浏览器打开目标页，保存渲染后的完整 HTML，并写 `manifest.json`，运行时加 `-HtmlDir`；不得因为内核缺失直接放弃抓取。
3. 只有前两步都不可用且用户同意下载时，才执行 `python -m playwright install chromium`。

一键运行：

```powershell
& "$SkillRoot\scripts\run_local.ps1" `
  -Excel "D:\path\企业名录.xlsx" `
  -Out "D:\path\企业官网资料包" `
  -Limit 3
```

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
| `--en <json>` | 用人工确认的英文覆盖自动翻译 |
| `-HtmlDir <目录>` / `--html-dir <目录>` | 用 Codex 内置浏览器保存的离线 HTML 兜底抓取（目录内需 `manifest.json`） |
| `--no-translate` | 关闭自动翻译，只输出中文并把英文层留空 |
| `--translate-email <邮箱>` | 可选，MyMemory 联系邮箱，用于提高匿名翻译额度 |
| `--no-visual-review` | 跳过拼版/核对表和视觉核对.json 生成 |
| `--strict` | 门禁未通过时返回非零退出码（`run_local.ps1` 默认启用；用 `-AllowRed` 关闭） |

## 流水线阶段

| # | 阶段 | 本地脚本行为 | 交付判据 |
|---|---|---|---|
| 0 | 读表 | 识别企业名称列和可选官网列 | 至少一条企业记录 |
| 1 | 官网发现 | 用户官网优先；否则多引擎自动发现并抓取候选首页，按公司名命中度分级；低置信度交由 Codex 复核 | 命中且置信度可接受；完全找不到才标 `no_website` |
| 2 | 页面抓取 | 首页 + 关于/工厂/资质页；产品分类/列表页优先并下钻一层补齐叶子分类与详情，保留标题、正文、链接、图片 | `raw/*.json` 有页面记录 |
| 3 | 内容与图片 | 抽取中文简介、产品名、产品详情、主营和地址；只有详情或产品图佐证的条目才作为产品；识别站头 logo 与 CSS 背景横幅，图片按四类落盘 | 图片引用真实存在、产品详情中英双语、产品图本地链接有效 |
| 4 | 渲染 | 生成中英双语企业 docx、产品清单 xlsx、汇总 xlsx | 文件可打开且结构完整 |
| 5 | 视觉核对 | 生成拼版、核对表和核对指引；Codex 自己看图，经 `visual_review.py --apply` 回写 `视觉核对.json` | 无“不符”，核对完成 |
| 6 | 门禁 | 自动调用 `gates.py` | `gates.json` 无 error |

## 开工顺序

1. **确认 Excel。** 先核对表头、企业数、是否已有官网列；不要猜测企业简称。没有官网列时按企业名称自动发现，不再要求用户先提供域名。
2. **小样本试跑。** 用 `-Limit 1` 或 `-Limit 3 -NoPublish`，检查官网命中置信度、图片归档、docx/xlsx 和门禁日志；低置信度官网由 Codex 打开确认。
3. **完整运行。** 小样本通过后去掉 `-Limit` 运行全部企业（推荐 `-NoPublish`，先构建待核对）。
4. **视觉核对（Codex 自己做）。** 打开 `review/核对指引.md`，用图像查看工具逐张看 `review/视觉核对图/<企业>/0.总览.png` 和每个分类拼版；把结论写成 `verdicts.json`，执行 `visual_review.py --apply` 回写，再跑 `gates.py --require-visual`。核对完成后用 `local_pipeline.py --publish-stage <build/run_id> --out <输出目录>` 发布。
5. **复核门禁。** 先看 `gates.json` 的 error 项，再看 `empty_images`、`no_website`、`partial` 清单。
6. **核对英文。** 默认会写入自动翻译英文草稿并标注待人工核校；需要定稿时用 `--en <已确认英文.json>` 覆盖后重跑。
7. **发布。** 仅在门禁通过或用户明确接受“英文待定稿”时交付。门禁红色修数据或补抓，不下调阈值迁就数据。

## 硬性约束

- **不要 SSH 到任何远端。** 所有发现、抓取、渲染和门禁都在本机运行。
- **不要记录密码、token、cookie。** 不要求用户把凭据写入 skill、Excel、日志或命令行历史。
- **不臆造企业事实。** 产业、业务、产品、资质只能来自官网快照或用户资料。
- **无官网列要自动发现，低置信度必须复核。** 输入表缺官网列时，先由脚本自动搜索、抓取和打分；高/中置信度自动采用，低置信度必须由 Codex 打开站点确认后再交付，不得直接跳过该企业。
- **依赖必须齐全。** 首次在新环境使用先跑 `--selftest`；`openpyxl`、`python-docx`、`Pillow` 缺一即视为环境未就绪（Pillow 缺失会让视觉核对材料静默消失）。`run_local.ps1` 默认检测 Playwright 包，内置 Chromium 不可用时复用本机 Edge 作为渲染内核（两者都缺失才下载 Chromium）；网络受限无法安装时，用 `-NoPlaywrightInstall` 显式降级并说明。
- **机翻不等于定稿。** 自动英文来自机翻，必须在 `en.json`/汇总/交付说明中保留 `自动翻译: true`、`待人工核校` 标记，不得当作人工定稿交付。
- **不把空结果当成功。** 四类图全空、官网未确认、简介不足三段都要显式记录并单独列出。
- **不跳视觉核对，也不把核对推给用户。** 图片内容是否属于该企业、该分类，只能看图判断；Codex 必须自己查看拼版并回写结论，不得在未核对时声称图片已核验，也不得把 `待核对` 留给用户后发布终稿。
- **不修改无关脚本。** 只处理当前请求涉及的小样本或输入名录。

## 参考

- [references/codex-execution.md](references/codex-execution.md) —— 本地执行、参数、故障排查
- [references/pipeline.md](references/pipeline.md) —— 各阶段行为与详细命令
- [references/data-contract.md](references/data-contract.md) —— Excel、raw JSON、英文层和交付目录契约
- [references/quality-gates.md](references/quality-gates.md) —— 门禁判据、阈值与视觉核对流程