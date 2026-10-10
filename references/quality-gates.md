# 质量门禁

门禁是**断言**，不是报告：`scripts/gates.py` 跑完全部检查，任一 `error` 级不通过即非零退出。报告会过期、引用会变死链，断言每次都能重跑。

本地流水线每次运行都会自动调用 `gates.py`。也可以单独复跑：

```powershell
$SkillRoot = "$env:USERPROFILE\.codex\skills\enterprise-site-pipeline"
# 视觉核对开工前：先查会话体积（>20MB 先另开会话）
python "$SkillRoot\scripts\session_guard.py"
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

视觉核对默认就是 fail-closed：官网抓取图未回写结论即 error。`--require-visual` 保留兼容；只有用户明确接受风险时才用 `--skip-visual-review` 显式降为告警：

```powershell
& $PY "...\gates.py" --deliverable "..." --raw "..." --en "..." --summary "..."

# 仅用户明确接受“不做视觉核对”风险时：
& $PY "...\gates.py" --deliverable "..." --raw "..." --en "..." --summary "..." --skip-visual-review
```

仅当用户明确接受中文版时才允许英文缺失：

```powershell
& $PY "...\gates.py" --deliverable "..." --raw "..." --en "..." --summary "..." --accept-no-english
```

## 门禁清单

| id | 级别 | 判据 | 阈值 |
|---|---|---|---|
| `site_discovery` | error/warn | 官网必须确认；自动发现置信度为中/低时未复核不得进入采集/交付，未复核企业应由主流程默认跳过；`resource_only`（官网不可用但用户资料可用）单列告警 | 未发现 0 家（`resource_only` 除外）；中/低置信度未复核 error |
| `coverage` | error | 档案、目录、英文、汇总四方企业集合一致，且数量等于 `expected_companies` | 差集为空 |
| `structure` | error | 每家固定为五文件夹 + 产品清单 xlsx；`5.企业介绍` 下 docx 文件名必须等于企业文件夹名；企业根目录无多余项，五文件夹内只允许直接文件、不得嵌套子目录 | 缺失数 0、多余项 0、错误 docx 名 0、嵌套目录 0 |
| `images` | error | 档案里的图片引用都落到真实文件；交付目录里没有 raw 未记录的孤儿图；跨类重复图；四类图是否全空 | 失效引用 0、孤儿图 0、跨类重复组 0、四类全空 0 家 |
| `image_required` | warn | 单独检查 `logo` 与 `factory` 两类是否有图；官网确实没有素材时保留告警并写数据边界，不硬性阻断 | 空类 0 条（否则告警） |
| `image_provenance` | error | 每张官网图的来源页必须与官网同域；`source: 用户资料` 的图跳过同域校验；来源页为聚合/目录站或图片直链外域/聚合站均判 error | 来源页非官网 0 条；图片直链外域/聚合站 0 条（用户资料图除外） |
| `en_entry` | error/warn | 英文条目六字段齐全且非空；简介恰 3 段；每段 ≥60 字符；英文名/标题/品牌/简介/产品英名无中日韩字符（`子品类` 是中文分类，豁免）；只有显式 `--accept-no-english` 才降为 warn | 问题条目 0 |
| `en_ascii` | error/warn | 英文段非 ASCII 字符占比；只有显式 `--accept-no-english` 才降为 warn | ≤ 0.02 |
| `en_pinyin` | warn | 英文段疑似拼音/栏目词 token 占比 | ≥ 0.60 告警 |
| `product_en` | error/warn | 产品名称列必须为“中文/型号 / 英文名”格式，且 ` / ` 右侧含至少 2 个连续拉丁字母；只有显式 `--accept-no-english` 才降为 warn | 空英文行 0 |
| `product_detail` | error/warn | `产品详情` 列只填官网或企业资料中真实存在的非空内容；非空行必须同时含中文和至少 3 个连续英文词，型号字母不算英文译文；来源无详情时留空且禁止旧占位说明；真实详情覆盖率低于阈值仅告警 | 非空详情缺中/英 0 行；真实详情覆盖率 ≥ 0.30 |
| `product_image_link` | warn | `图片（本地连接）` 列要么指向交付目录内真实存在的相对路径，要么留空；来源没有对应产品图时留空，禁止写占位文字 | 无效/占位链接 0 行；空链接不报错 |
| `product_rows` | error | `产品清单.xlsx` 数据行（按 ` / ` 取中文名）与 raw `products` 顺序逐行一致 | 不一致家数 0 |
| `product_map` | error/warn | 档案产品在 `产品英名` 中的覆盖率；只有显式 `--accept-no-english` 才降为 warn | ≥ 0.80 |
| `docx_sync` | error | docx 英文段与 `en.json` 英文简介逐字一致 | 不一致家数 0 |
| `docx_source` | error | docx 中文段与 raw `intro_paragraphs` 逐段一致 | 不一致家数 0 |
| `resource_intake` | warn | 用户资料摄入情况：`资料备注` 含"未归类/抽取失败"则告警（原件已备份，需人工确认）；`resource_only` 单列 | 告警，不阻断 |
| `noise` | error | 中文或英文正文段不含导航/备案/联系方式/黄页词等噪声模式 | 命中 0 |
| `summary` | error | 汇总表表头与行列数符合约定 | 表头精确匹配 |
| `visual_review` | error/warn | 读取 `视觉核对.json`：官网抓取图必须逐张回写结论，结论三档 `符合/不符/待核对`，只有 `符合` 放行（`待核对`/空均阻断），并绑定当前图片 sha256 与 reviewed_at；用户资料图直通但“不符”仍 error；不支持类别继承。默认 error，只有显式 `--skip-visual-review` 才降 warn | 官网图未核对 0 条；不符 0 条；待核对 0 条；无 sha/reviewed_at 绑定 0 条 |

## 关键门禁的意义

**`site_discovery`** —— 官网是整条流水线的事实源。未发现官网且没有可用用户资料时 error；自动发现置信度为中/低时必须先复核，未填“决定/自定义官网”的企业由 `plan_site()` 默认跳过，不得进入 `raw` 或交付；若异常绕过主流程，`gates.py` 直接报 error。高置信度可自动采用。若官网未确认或不可访问但用户资料可用，记 `resource_only`（仅凭用户资料成档），`site_discovery` 降为告警并单列，不算失败。

**`resource_intake`** —— 用户资料可能以任意目录结构、任意格式提交（图片/Word/Excel/PDF/PPT/txt…）。脚本递归扫描、按企业名匹配、按类别归档并抽取文本，原件备份在 `deliverable` 之外。无法归类或抽取失败的项会在 `资料备注` 记录并触发本门禁告警（不阻断交付），Codex 需人工确认这些项是否可忽略或需手工补录；原件始终保留在备份目录。

**`en_entry` / `en_ascii` / `product_en` / `product_detail` / `product_map`** —— 默认英文来自自动中译英草稿，`en.json` 带 `自动翻译: true`、`定稿: false`、`翻译失败` 标记。英文缺口会先写入 `<run_id>/英文补译清单.json`，门禁默认按 error 阻断发布，正常修复路径是补 `--en`。MyMemory 429/限流在 `--selftest` 中是可恢复 `WARN`，不等于允许带英文缺口交付。只有用户在当次对话中明确接受中文版时，`--accept-no-english` 才把英文缺失降为 warn；检查仍执行，缺口仍保留在结果中。

**`product_en`** —— 只看产品名存在不够；必须能在 ` / ` 右侧读到真正的英文名，型号字母本身不能算英文已补齐。

**`product_detail`** —— `产品详情` 列只填官网或企业资料中真实存在的内容。非空详情不能只有中文，至少要有 3 个连续英文词才算英文译文，中文详情中的 `WMS`、`GaN`、`5A` 等型号 token 不算；官网与企业资料都没有详情时单元格留空，不再写“官网未提供独立产品详情 / No standalone product description...”等占位说明。真实详情覆盖率作为告警线记录在 `企业说明` 页和门禁日志中，不允许为了凑覆盖率编造内容。

**`product_image_link`** —— 产品清单的 `图片（本地连接）` 列要么能落到 `2.企业产品图/` 下的真实文件，要么留空；按 `alt`、文件名、来源路径和 URL 匹配产品名，也支持型号/产品主体包含匹配。用户资料只有一张产品图且无法精确匹配时，作为未匹配产品行的主图兜底；多张图无法唯一匹配时该单元格留空，不伪造对应关系，也不允许保留“官网未提供产品图 / No product image available...”“未匹配到本产品对应的本地图片 / No product-specific local image matched...”“本地抓取，待核验”这类占位文字。

**`docx_sync`** —— 防止“英文 JSON 改了、docx 没重渲染”。本地流水线每次从当前 `en.json` 重新生成 docx，门禁再逐家比对，避免交付旧英文。

**`en_pinyin`** —— 防止拼音或栏目词被当成英文。它是启发式告警，最终仍需人工确认。

**`docx_source` / `product_rows`** —— 把“交付物来自哪个事实源”也变成断言：docx 中文段必须等于 raw `intro_paragraphs`，产品清单数据行必须等于 raw `products`。手工改过交付文件却不同步 raw，会在这里报红。

**`image_provenance`** —— 图片的 `from` 来源页必须与官网同域，防止把别家站点或聚合站的图当成企业自己的。聚合/黄页/工商/名录/B2B 站由 `scripts/domain_rules.py` 统一识别；来源页或图片直链命中目录站、或直链为外域时直接 error，不能靠视觉结论放行。用户资料图没有官网来源页，按资料优先跳过同域校验。

**`visual_review`** —— 机器只能判“文件和引用对得上”，判不了“这张图到底是不是工厂/产品/logo/资质”。视觉核对由 Codex 自己打开拼版完成：先看核对表「锚点线索」，图片内容要能挂到本企业已证实的名称/产品/品类才判「符合」；把结论写进 `review\verdicts.json`（可用 `依据` 写清命中哪个锚点/来源），用 `visual_review.py --apply` 回写 `视觉核对.json`。默认 fail-closed：结论三档 `符合/不符/待核对`，只有 `符合` 放行；官网抓取图留空或填“待核对”即 error；填“不符”直接 error。用户资料图按资料优先直通，不要求逐张结论，但显式“不符”仍阻断。结论只认逐张图，不支持类别继承；每张结论绑定 sha256 和 reviewed_at，图片变化后旧结论自动作废。只有用户明确接受风险时，`--skip-visual-review` 才把未核对降为 warn。

**`images` 的四类全空** —— 目录存在不等于有图。任何企业 `factory/product/logo/cert` 四类全空都直接报 error，必须补抓或在交付说明中标记 `empty_images`。

**`image_required`** —— 单独盯 `logo` 与 `factory`。`logo` 是品牌识别的基础素材，`factory` 是源头工厂核验的重点；若其中一类为空，`gates.py` 报 warn 并把企业列出来，要求人工补图或在交付说明中标明“官网未公开该类素材”。这类情况不硬性阻断，避免把真实数据边界误判成错误。

## 视觉核对产物

`local_pipeline.py` 渲染完成后自动调用 `scripts/visual_review.py`，生成：

```text
<run_id>/review/
├── 视觉核对图/<企业>/0.总览.png        # 四类速览
├── 视觉核对图/<企业>/<分类>.png         # 分类拼版（单页 ≤12 张、≤1MB；超出为 <分类>_p1.png、_p2.png…）
├── <企业>/图片核对表.xlsx              # 带缩略图；含锚点线索、结论下拉、SHA256、信任层级
└── 核对指引.md                         # Codex 看图清单、判断口径和执行命令
<run_id>/视觉核对.json                  # 结论载体，gates.py 读取
```

核对方式（Codex 直接把拼版当图片打开，逐张看，不得推给用户）：

必须在短线程内完成：每批 ≤10 家，一个会话只做「查体积 → 看图 → 回写结论 → 跑门禁 → 发布」，做完即止；开工前跑 `scripts/session_guard.py`，本会话 rollout >20MB 先换会话。拼版图进入会话后是 base64，单张过大或累计过多会撑爆请求（历史故障：19 张拼版 28.8MB 触发上游报错），所以脚本把单页限制为 ≤12 张、≤1MB，图片多时按序号分页为 `<分类>_p1.png`、`<分类>_p2.png`……官网抓取图必须逐张看并逐张给结论；用户资料图标“信任层级=用户资料”，按资料优先直通，不要求逐张结论。

1. 打开 `review/核对指引.md`，图片多时先用 `<企业>/图片核对表.xlsx` 初筛（分类/序号/尺寸/图片URL/来源页面/alt/锚点线索/SHA256/信任层级）；
2. 按指引标注的序号范围打开对应 `<分类>_pN.png` 放大确认；官网抓取图逐张判定，判定时以「锚点线索」为参照，挂不上锚点的图不得判「符合」；用户资料图可跳过未标记疑点。同时核对简介 docx 与产品清单内容；
3. 把结论写进 `review/verdicts.json`，只支持逐张图结论，不支持类别继承；然后执行 `visual_review.py --json "<run_id>/视觉核对.json" --apply "<run_id>/review/verdicts.json" --reviewer Codex` 回写；
4. 执行 `gates.py` 复核（视觉默认强制）；全绿后用 `local_pipeline.py --publish-stage "<run_id>" --out "<输出目录>"` 发布，发布前会再跑一次门禁。

**结论值**：`符合` / `不符` / `待核对`（空等同待核对，只有「符合」放行，「待核对」与留空同样阻断发布；可选 `依据` 字段写清命中哪个锚点/来源）。重跑 `visual_review.py` 时，只有当旧结论记录的 sha256 与当前图片一致才保留；图片变化会清空结论并重新进入待核对。`--apply` 写入 reviewed_at，缺任一绑定都会在门禁报 error。

图片内容是否真属于该类、是否属于该企业，机器判不了，只能靠这一步；`visual_review` 只是把“官网图有没有逐张核对、结论是否仍对应当前文件”变成可断言状态。分级信任只减少用户资料图的机械核对，不降低官网抓取图的逐张要求。

## 阈值

本地流水线使用以下默认值，也可以在 `gates.py` 命令行显式覆盖：

| 参数 | 默认 | 含义 |
|---|---:|---|
| `--ascii-max` | 0.02 | 英文简介非 ASCII 字符占比上限 |
| `--pinyin-max` | 0.60 | 疑似拼音 token 占比告警线 |
| `--map-min` | 0.80 | 产品英名覆盖率下限 |
| `--detail-min` | 0.30 | 官网/用户资料真实产品详情覆盖率告警线 |

门禁红了先查数据，不要先改阈值。唯一例外是 `en_pinyin` 这类启发式告警；下调必须在运行说明中写明原因。

新增门禁时：

1. 先用真实企业数据复现已知问题；
2. 再确认它不会在正确数据上误报；
3. 写进 `gates.py` 的检查表和本文件，标注级别与阈值；
4. 所有 error 级结果都保留在 `gates.json`，无论通过或失败。
