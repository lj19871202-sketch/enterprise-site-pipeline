# 本地流水线各阶段

入口脚本：

```text
scripts/local_pipeline.py
scripts/run_local.ps1
```

脚本在 Codex 当前本机执行，输入 Excel，输出企业官网资料包。不要把这套流程拆成远端会话或手工多机复制。新环境先运行 `bootstrap.ps1` 创建技能独立 `.venv`，并按 `requirements.lock.txt` 安装锁定依赖。

## 阶段 0 · 读取 Excel

Excel 至少包含企业名称列：

| 企业名称 | 官网（可选） |
|---|---|
| 示例科技有限公司 | https://example.com |

识别规则：

- 企业名称列名：`企业名称`、`公司名称`、`单位名称`、`名称`；
- 官网列名：`官网`、`网址`、`网站`、`官网地址`；
- 如果第一行不是表头，则默认第一列为名称。

先跑小样本：

```powershell
$SkillRoot = "$env:USERPROFILE\.codex\skills\enterprise-site-pipeline"
python "$SkillRoot\scripts\local_pipeline.py" `
  --excel "D:\path\企业名录.xlsx" `
  --out "D:\path\企业官网资料包" `
  --limit 3
```

验证：控制台打印企业数和 run_id，输出目录出现 `build/<run_id>/`。

## 阶段 1 · 官网发现

- Excel 提供官网时直接使用，并记录 `置信度: 用户提供`；
- 未提供时，本机对企业全称和去地域核心名做多轮查询（`"X" 官网` / `"X" 官方网站` / `"X" ICP备案`），静态搜索 `sogou → 360 → bing → duckduckgo`；候选少于 3 个且 Playwright 可用时，再用 baidu/bing 渲染补搜；
- 候选域名按搜索引擎命中次数 + 域名与公司名在同一结果片段中的共现次数加权排序；候选页面再用企业全称/核心名命中标题正文，并叠加地址、电话、备案、关于我们、版权所有等证据打分；
- `score ≥ 80` 记 `高（自动发现）`，`50–79` 记 `中（自动发现）`，其余记 `低（自动发现，需复核）`；低/中置信度由 Codex 打开站点复核，不能直接当事实；
- https 打不开的站自动回退 http（不少国内企业站只开 http），最终记录实际访问到的地址；
- 未命中的企业保留 `status: no_website`，不能静默丢失。

验证：查看 `raw/<企业>.json` 中的 `官网`、`置信度`、`errors`。

## 阶段 2 · 页面抓取

默认抓取：

- 首页；
- 关于/公司介绍；
- 产品/业务/解决方案；
- 工厂/车间/生产/基地；
- 资质/证书/荣誉。

选页顺序为“产品分类/列表页优先，其次企业信息页，最后产品详情页”，并在首批页面抓取后再下钻一层，补齐叶子分类与产品详情页——列表页能一次带出更多“产品名 + 详情 + 产品图”，避免只抓到首页几条产品。

默认 `--playwright auto`：静态抓取优先，页面正文少于 400 字、链接少于 5、图片少于 3 时自动用本机 Playwright 渲染重抓；`on` 强制渲染，`off` 纯静态。`run_local.ps1` 默认复用本机 Playwright；渲染内核先试内置 Chromium，不可用时自动改用本机 Microsoft Edge（Windows 自带、与 Chromium 同源），两者都不可用时才下载 Chromium。离线可用 `-NoPlaywrightInstall` 降级，或用 `--html-dir` 吃 Codex 内置浏览器保存的离线 HTML。

验证：

- `raw/<企业>.json` 的 `pages` 有 URL、标题和分类；
- `home_text`、`about_text`、`product_text` 有原文快照。

## 阶段 3 · 内容与图片

中文处理：

- 按句子切分首页和关于页面，生成最多三段简介；
- 从产品标题、列表和链接提取产品名；导航栏目词（产品中心/关于我们/荣誉资质等）与图标字符会被过滤；只有出现在产品详情里、或与已下载产品图 `alt` 精确/系列名匹配的条目才保留为产品，避免把“储能系统”这类分类词当产品；
- 从产品卡片链接文本提取“产品名 + 简介”，剥离“查看详情”后写入 raw 的 `product_details`；首页省略完整型号时按唯一“系列名”前缀回填；
- 提取主营、产业和地址候选；
- 噪声导航、备案、联系方式不进入正文。

图片处理：

- `factory` → `1.企业工厂图`
- `product` → `2.企业产品图`
- `logo` → `3.企业logo`
- `cert` → `4.资质证书`

下载失败、非图片响应会被跳过；同内容只归档一次，但产品图允许同内容、不同 `alt` 各存一份（如官网用同一张照片对应“美标/欧标”产品），任何内容都不跨类别重复。图片记录来源 URL 和本地文件名。产品清单的 `图片（本地连接）` 列按图片 `alt` 与产品名匹配，写入 `2.企业产品图/<文件名>` 相对路径并设为可点击链接；匹配不到时写“官网未提供产品图 / No product image available on the official website”，不伪造图片。

导航/收起/联系我们/二维码/关注/喜报/揭牌等 UI 元素与图标（按 alt/src 关键词识别）会在分类前丢弃，避免混入产品图或证书；alt 为公司名/有限公司的图同样丢弃（无法区分顶部 logo 与二维码）。

同时补充两类易漏图片：

- `logo`：位于指向首页的链接内、且是该链接第一张图的 `<img>` 视为站头 logo（即使无 alt 或 alt 恰为公司名），优先于 alt 噪声过滤；
- `factory`：解析 CSS `background-image:url(...)`，首页/关于页中命中 `about`、`company`、`factory`、`厂区`、`简介` 等线索的背景横幅作为工厂/厂区候选；其余背景横幅不归档。

验证：

- `raw/<企业>.json` 的四类图片数组与本地文件一致；
- 交付目录没有 raw 未记录的孤儿图；
- 每张图的 `from` 来源页与官网同域；
- `empty_images` 企业单独列入报告。

图片分类只能靠 URL/alt 关键词和所在页面推测，属启发式，必须进入阶段 6 的视觉核对。

## 阶段 4 · 渲染

生成：

- `<企业>/5.企业介绍/<企业名>简介（YYYYMMDD短）.docx`
- `<企业>/产品清单.xlsx`
- `<run_id>/汇总.xlsx`
- `<run_id>/en.json`
- `<run_id>/英文补译清单.json`（存在英文缺口时）
- `<run_id>/review/视觉核对图/<企业>/*.png`
- `<run_id>/review/<企业>/图片核对表.xlsx`
- `<run_id>/review/核对指引.md`
- `<run_id>/视觉核对.json`

默认自动把企业名、三段简介、产品名和产品详情翻成英文草稿（`en.json` 标 `自动翻译: true`、`定稿: false`），`英文简介` 非空即写入 docx 英文段；产品详情以“中文\nEnglish”写入 `产品清单.xlsx` 的 `产品详情` 列；官网没有独立详情的分类/系列名行写中英双语占位说明，不编造内容。提供 `--en` 时覆盖。MyMemory 429/限流不再把自检判死，但英文缺口会写入 `英文补译清单.json`，默认门禁阻断发布，等待 Codex 基于中文事实补 `--en`。

验证：

- 每家企业有四个图片目录；
- docx 和 xlsx 可打开；
- 产品清单 `图片（本地连接）` 每行要么指向真实本地文件，要么显式标注官网无图；
- 汇总表行数与输入企业数一致。

## 阶段 6 · 视觉核对

渲染后自动生成：

```text
<run_id>/review/视觉核对图/<企业>/0.总览.png
<run_id>/review/视觉核对图/<企业>/<分类>.png
<run_id>/review/<企业>/图片核对表.xlsx
<run_id>/review/核对指引.md
<run_id>/视觉核对.json
```

流程由 Codex 自己执行，不得推给用户：

1. Codex 打开 `review/核对指引.md`，按绝对路径逐家看图；
2. 打开 `0.总览.png` 做初筛，疑点打开 `<分类>.png` 放大确认；同时核对简介 docx 和产品清单与原始页面是否一致；
3. 把结论写进 `review/verdicts.json`，执行 `visual_review.py --json "<run_id>/视觉核对.json" --apply "<run_id>/review/verdicts.json" --reviewer Codex` 回写；
4. 执行 `gates.py --require-visual` 复核，全绿后用 `local_pipeline.py --publish-stage "<run_id>" --out "<输出目录>"` 发布。

结论可写在类别层（该类未单独填写的图片继承），也可写到单张图。重跑 `visual_review.py` 会保留已有结论。

跳过生成用 `--no-visual-review`；要求核对必须完成才发布，用 `gates.py --require-visual`。构建阶段可加 `local_pipeline.py --no-publish`，保证 Codex 先看图再发布。

## 阶段 7 · 门禁与发布

脚本自动调用同目录 `gates.py`：

```powershell
$SkillRoot = "$env:USERPROFILE\.codex\skills\enterprise-site-pipeline"
python "$SkillRoot\scripts\local_pipeline.py" `
  --excel "D:\path\企业名录.xlsx" `
  --out "D:\path\企业官网资料包" `
  --strict
```

`run_local.ps1` 默认透传 `--strict`（门禁红灯返回非零退出码）；确需保留红灯产物时加 `-AllowRed`。

门禁结果：

```text
<输出目录>\build\<run_id>\gates.json
<输出目录>\build\<run_id>\gates.log
```

处理规则：

- 所有 error 项通过后，才把 stage 发布到输出目录 `deliverable/`；
- 官网未发现时 `site_discovery` 为 error；自动发现置信度为中/低时先 warn，Codex 必须打开候选站点复核；
- `visual_review` 标“不符”必红；未核对默认告警，`--require-visual` 时按 error；
- 自动翻译失败或英文缺失时 `en_entry`/`en_ascii`/`product_en`/`product_detail`/`product_map` 默认报红，并生成 `英文补译清单.json`；正常修复是用 `--en` 提供补译草稿或人工定稿后重跑；
- 只有用户在当次对话中明确接受中文版时，才可加 `--accept-no-english` 将英文相关门禁降为告警；
- 图片为空先补抓或在 Excel 补充官网，不修改门禁标准；
- `no_website`、`empty_images`、`partial` 必须进入交付说明。

## 交付前检查

1. 输入 Excel 行数与 `manifest.json` 的企业数一致；
2. `gates.json` 没有未解释的 error；
3. `no_website`、`empty_images`、`partial` 清单可追溯；自动发现的官网置信度为高，或中/低已由 Codex 打开复核；
4. 默认交付必须英文完整：`英文补译清单.json` 无未处理项，英文是“已确认定稿”，或是已明确标注“自动翻译·待人工核校”的完整草稿；只有用户明确接受中文版时，才可保留 `accept_no_english` 的告警状态；
5. 交付目录中的图片引用全部存在，图片、文档、产品清单均已在 `视觉核对.json` 标为“符合”；
6. docx 中文段与 raw `intro_paragraphs`、产品清单行与 raw `products` 一致；
7. 没有把密码、token 或 cookie 写入任何输出。