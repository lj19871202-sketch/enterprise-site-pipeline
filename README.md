# 企业官网资料包流水线

本地 Excel 驱动的企业官网批量抓取与资料包生成工具。给它一份包含企业名称的 `.xlsx`，它会在 Codex 当前 Windows 机器上完成：

- 官网发现与校验；
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
python "$SkillRoot\scripts\local_pipeline.py" --selftest --excel "D:\path\企业名录.xlsx"
```

如果本机 `python` 不在 PATH，用 `$env:CODEX_PYTHON` 指定解释器，或安装依赖：
`<python.exe> -m pip install -r "$SkillRoot\requirements.txt"`。

## 输出

```text
<输出目录>/
├── build/<run_id>/
│   ├── raw/                 # 每家企业抓取事实
│   ├── deliverable/         # 企业文件夹、docx、产品清单
│   ├── en.json              # 英文层：自动翻译草稿，可被 --en 覆盖
│   ├── 视觉核对.json         # 图片/文档/产品的视觉核对结论
│   ├── review/              # 拼版图 + 缩略图核对表 + 核对指引.md
│   ├── 汇总.xlsx
│   ├── gates.json
│   ├── gates.log
│   └── manifest.json
└── deliverable/             # 门禁通过后发布
```

每家企业默认生成四类图片目录，以及中英双语的 `5.企业介绍/*.docx` 和 `产品清单.xlsx`（英文来自自动翻译，`en.json` 标 `自动翻译: true`，待人工核校）。渲染后自动生成图片拼版与 `图片核对表.xlsx`，视觉核对结论写入 `视觉核对.json` 供门禁校验。

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
| `-En <json>` / `--en <json>` | 用人工确认稿或 Codex 补翻草稿覆盖自动翻译（模型草稿必须保留 `定稿: false`） |
| `-HtmlDir <目录>` / `--html-dir <目录>` | 用 Codex 内置浏览器保存的离线 HTML 兜底抓取（目录内需 `manifest.json`） |
| `-NoTranslate` / `--no-translate` | 关闭自动翻译，只输出中文 |
| `--translate-email <邮箱>` | 可选，提高 MyMemory 匿名翻译额度 |
| `-NoVisualReview` / `--no-visual-review` | 跳过拼版、核对表和 `视觉核对.json` 生成 |
| `--strict` | 门禁不通过时返回非零退出码（`run_local.ps1` 默认启用；CLI 直跑需显式加） |
| `-AllowRed` | 关闭 `run_local.ps1` 的默认严格模式，门禁红仍返回 0 |

## 质量门禁

门禁在每次运行后自动执行，结果写入 `build/<run_id>/gates.json` 和 `gates.log`。

- 图片引用必须存在、交付目录没有 raw 未记录的孤儿图、四类图片不能全空；
- `logo` 或 `factory` 单独为空时给出告警，要求补图或标记数据边界；
- 每张图的来源页必须与官网同域（图片直链走 CDN 仅提示）；
- 每家企业必须有四类图片目录、简介 docx、产品清单 xlsx；
- docx 中文段必须等于 raw `intro_paragraphs`，产品清单数据行必须等于 raw `products`，产品详情列必须中英双语且无空值；`图片（本地连接）` 列必须指向交付目录内真实存在的 `2.企业产品图/...` 文件，官网无图时显式标注，不得保留旧占位符；
- 英文默认自动中译英并写入正文，`en.json` 标 `自动翻译: true` 待人工核校；提供 `--en` 后覆盖，人工确认稿标 `定稿: true`，MyMemory 失败时由 Codex 补翻的草稿仍标 `定稿: false`；
- 官网必须找到；自动发现置信度为低/中时列出，低置信度由 Codex 打开站点复核；
- 汇总表、档案、目录和英文集合必须一致；
- 图片内容需视觉核对，结论写入 `视觉核对.json`：标“不符”即报红，未核对默认告警；加 `--require-visual` 后未核对直接报红。

## 视觉核对

流水线渲染后自动生成：

```text
build/<run_id>/review/视觉核对图/<企业>/0.总览.png    # 四类速览
build/<run_id>/review/视觉核对图/<企业>/<分类>.png     # 分类拼版（该类全部图片）
build/<run_id>/review/<企业>/图片核对表.xlsx           # 带缩略图，结论列可下拉
build/<run_id>/视觉核对.json                           # 结论载体
```

这一步由 Codex 自己做，不能把拼版甩给用户代看。Codex 先用图像查看工具打开总览初筛，再逐个打开分类拼版放大核对；读 `review/核对指引.md` 获取当前运行的企业、拼版绝对路径和 JSON 模板。结论写成 `verdicts.json` 后一键回写：

```powershell
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
- 英文默认走 MyMemory 免费接口自动中译英。该接口有每日匿名额度，批量较大时部分条目可能翻译失败：失败条目在 `en.json` 标 `翻译失败: true`，门禁 `en_entry` 报红，需稍后重跑（结果有本地缓存 `_translate_cache.json`）、加 `--translate-email` 提高额度，或用 `--en` 提供定稿；持续 429 时可由 Codex 基于中文事实补翻为 `--en` 草稿，但必须保留 `定稿: false`。
- 输入格式当前以 `.xlsx` / `.xlsm` 为主；老式 `.xls` 请先另存为 `.xlsx`。

详细说明见 `SKILL.md` 和 `references/`。