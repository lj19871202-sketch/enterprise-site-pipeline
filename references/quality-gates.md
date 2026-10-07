# 质量门禁

门禁是**断言**，不是报告：`scripts/gates.py` 跑完全部检查，任一 `error` 级不通过即非零退出。报告会过期、引用会变死链，断言每次都能重跑。

本地流水线每次运行都会自动调用 `gates.py`。也可以单独复跑：

```powershell
$SkillRoot = "$env:USERPROFILE\.codex\skills\enterprise-site-pipeline"
python "$SkillRoot\scripts\gates.py" `
  --deliverable "<输出>\build\<run_id>\deliverable" `
  --raw "<输出>\build\<run_id>\raw" `
  --en "<输出>\build\<run_id>\en.json" `
  --summary "<输出>\build\<run_id>\汇总.xlsx" `
  --expected 3 `
  --json "<输出>\build\<run_id>\gates.json"
```

也可以只跑指定门禁：

```powershell
& $PY "...\gates.py" --deliverable "..." --raw "..." --en "..." --summary "..." --only en_pinyin,docx_sync
```

要求视觉核对必须完成才通过：

```powershell
& $PY "...\gates.py" --deliverable "..." --raw "..." --en "..." --summary "..." --require-visual
```

## 门禁清单

| id | 级别 | 判据 | 阈值 |
|---|---|---|---|
| `site_discovery` | error/warn | 官网必须发现；自动发现置信度为中/低时由 Codex 打开候选站点复核 | 未发现 0 家；中/低置信度仅告警 |
| `coverage` | error | 档案、目录、英文、汇总四方企业集合一致，且数量等于 `expected_companies` | 差集为空 |
| `structure` | error | 每家含四类图目录 + 简介 docx + 产品清单 xlsx | 缺失数 0 |
| `images` | error | 档案里的图片引用都落到真实文件；交付目录里没有 raw 未记录的孤儿图；跨类重复图；四类图是否全空 | 失效引用 0、孤儿图 0、跨类重复组 0、四类全空 0 家 |
| `image_required` | warn | 单独检查 `logo` 与 `factory` 两类是否有图；官网确实没有素材时保留告警并写数据边界，不硬性阻断 | 空类 0 条（否则告警） |
| `image_provenance` | error | 每张图的来源页必须与官网同域；图片直链外域单独列出 | 来源页非官网 0 条（直链外域仅提示） |
| `en_entry` | error | 英文条目六字段齐全且非空；简介恰 3 段；每段 ≥60 字符；英文名/标题/品牌/简介/产品英名无中日韩字符（`子品类` 是中文分类，豁免） | 问题条目 0 |
| `en_ascii` | error | 英文段非 ASCII 字符占比 | ≤ 0.02 |
| `en_pinyin` | warn | 英文段疑似拼音/栏目词 token 占比 | ≥ 0.60 告警 |
| `product_en` | error | 产品名称列必须为“中文/型号 / 英文名”格式，且 ` / ` 右侧含至少 2 个连续拉丁字母 | 空英文行 0 |
| `product_detail` | error/warn | `产品详情` 列非空行必须同时含中文和至少 3 个连续英文词；型号字母不算英文译文；无详情行不得留空，官网真实详情覆盖率低于阈值仅告警 | 缺中/英 0 行、空详情 0 行；覆盖率 ≥ 0.30 |
| `product_image_link` | error/warn | `图片（本地连接）` 列必须指向交付目录内真实存在的相对路径；官网无图必须显式标注，旧占位符视为错误 | 无效/空/占位链接 0 行；官网无图行告警 |
| `product_rows` | error | `产品清单.xlsx` 数据行（按 ` / ` 取中文名）与 raw `products` 顺序逐行一致 | 不一致家数 0 |
| `product_map` | error | 档案产品在 `产品英名` 中的覆盖率 | ≥ 0.80 |
| `docx_sync` | error | docx 英文段与 `en.json` 英文简介逐字一致 | 不一致家数 0 |
| `docx_source` | error | docx 中文段与 raw `intro_paragraphs` 逐段一致 | 不一致家数 0 |
| `noise` | error | 中文或英文正文段不含导航/备案/联系方式/黄页词等噪声模式 | 命中 0 |
| `summary` | error | 汇总表表头与行列数符合约定 | 表头精确匹配 |
| `visual_review` | warn/error | 读取 `视觉核对.json`：结论为“不符”即 error；未回写结论按 warn，加 `--require-visual` 后按 error | 不符 0 条；`--require-visual` 时未核对 0 条 |

## 关键门禁的意义

**`site_discovery`** —— 官网是整条流水线的事实源。未发现官网直接 error；自动发现置信度为中/低时先 warn，Codex 必须打开候选站点核对，确认后再运行或补 Excel 官网列；不允许把低置信度结果静默当事实。

**`en_entry` / `en_ascii`** —— 默认英文来自自动中译英草稿，`en.json` 带 `自动翻译: true`、`定稿: false`、`翻译失败` 标记；门禁仍按可交付英文标准检查三段简介、无中日韩字符和非 ASCII 占比。门禁红说明自动草稿不达标，应补数据或用 `--en` 提供定稿，不要下调阈值。

**`product_en`** —— 只看产品名存在不够；必须能在 ` / ` 右侧读到真正的英文名，型号字母本身不能算英文已补齐。

**`product_detail`** —— `产品详情` 列不能只有中文，也不能是空白；至少要有 3 个连续英文词才算英文译文，中文详情中的 `WMS`、`GaN`、`5A` 等型号 token 不算。官网确实没有独立详情的分类/系列名行，写中英双语事实占位说明（“官网未提供独立产品详情 / No standalone product description...”），真实详情覆盖率作为告警线记录在 `企业说明` 页和门禁日志中，不允许为了凑覆盖率编造内容。

**`product_image_link`** —— 产品清单的 `图片（本地连接）` 列必须能落到 `2.企业产品图/` 下的真实文件；能按 `alt` 匹配产品名时写相对路径并设为可点击链接。官网确实没有对应产品图时写“官网未提供产品图 / No product image available...”，显式告警但不伪造图片，不允许保留“本地抓取，待核验”这类占位文字。

**`docx_sync`** —— 防止“英文 JSON 改了、docx 没重渲染”。本地流水线每次从当前 `en.json` 重新生成 docx，门禁再逐家比对，避免交付旧英文。

**`en_pinyin`** —— 防止拼音或栏目词被当成英文。它是启发式告警，最终仍需人工确认。

**`docx_source` / `product_rows`** —— 把“交付物来自哪个事实源”也变成断言：docx 中文段必须等于 raw `intro_paragraphs`，产品清单数据行必须等于 raw `products`。手工改过交付文件却不同步 raw，会在这里报红。

**`image_provenance`** —— 图片的 `from` 来源页必须与官网同域，防止把别家站点或聚合站的图当成企业自己的。图片直链走 CDN 是常见情况，只列为提示，不直接判红。

**`visual_review`** —— 机器只能判“文件和引用对得上”，判不了“这张图到底是不是工厂/产品/logo/资质”。视觉核对由 Codex 自己打开拼版完成：把结论写进 `review\verdicts.json`，用 `visual_review.py --apply` 回写 `视觉核对.json`。填“不符”直接 error；填“符合”才算完成；留空为待核对。加 `--require-visual` 可要求全部核对完成才发布。

**`images` 的四类全空** —— 目录存在不等于有图。任何企业 `factory/product/logo/cert` 四类全空都直接报 error，必须补抓或在交付说明中标记 `empty_images`。

**`image_required`** —— 单独盯 `logo` 与 `factory`。`logo` 是品牌识别的基础素材，`factory` 是源头工厂核验的重点；若其中一类为空，`gates.py` 报 warn 并把企业列出来，要求人工补图或在交付说明中标明“官网未公开该类素材”。这类情况不硬性阻断，避免把真实数据边界误判成错误。

## 视觉核对产物

`local_pipeline.py` 渲染完成后自动调用 `scripts/visual_review.py`，生成：

```text
<run_id>/review/
├── 视觉核对图/<企业>/0.总览.png        # 四类速览
├── 视觉核对图/<企业>/<分类>.png         # 分类拼版（该类全部图片）
├── <企业>/图片核对表.xlsx              # 带缩略图；结论列可下拉
└── 核对指引.md                         # Codex 看图清单、判断口径和执行命令
<run_id>/视觉核对.json                  # 结论载体，gates.py 读取
```

核对方式（Codex 直接把拼版当图片打开，逐张看，不得推给用户）：

1. 打开 `review/核对指引.md`，按其中绝对路径打开 `<企业>/0.总览.png` 做初筛；
2. 疑点打开 `<分类>.png` 放大确认；同时核对简介 docx 与产品清单内容；
3. 把结论写进 `review/verdicts.json`：可只写类别结论（该类未单独填写的图继承），也可写单张；然后执行 `visual_review.py --json "<run_id>/视觉核对.json" --apply "<run_id>/review/verdicts.json" --reviewer Codex` 回写；
4. 执行 `gates.py --require-visual` 复核；全绿后用 `local_pipeline.py --publish-stage "<run_id>" --out "<输出目录>"` 发布。

**结论值**：`符合` / `不符` / `待核对`（空等同待核对）。重跑 `visual_review.py` 会合并保留已有结论，不会覆盖。

图片内容是否真属于该类、是否属于该企业，机器判不了，只能靠这一步；`visual_review` 只是把“有没有核对过”变成可断言状态。

## 阈值

本地流水线使用以下默认值，也可以在 `gates.py` 命令行显式覆盖：

| 参数 | 默认 | 含义 |
|---|---:|---|
| `--ascii-max` | 0.02 | 英文简介非 ASCII 字符占比上限 |
| `--pinyin-max` | 0.60 | 疑似拼音 token 占比告警线 |
| `--map-min` | 0.80 | 产品英名覆盖率下限 |
| `--detail-min` | 0.30 | 官网独立产品详情覆盖率告警线 |

门禁红了先查数据，不要先改阈值。唯一例外是 `en_pinyin` 这类启发式告警；下调必须在运行说明中写明原因。

新增门禁时：

1. 先用真实企业数据复现已知问题；
2. 再确认它不会在正确数据上误报；
3. 写进 `gates.py` 的检查表和本文件，标注级别与阈值；
4. 所有 error 级结果都保留在 `gates.json`，无论通过或失败。