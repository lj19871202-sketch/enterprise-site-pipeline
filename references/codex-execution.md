# Codex 本地执行说明

本流水线以 Codex 当前的 Windows 机器为执行机。输入 Excel、抓取结果、图片、docx、xlsx、日志和门禁结果全部落盘到用户指定的本地目录。

不使用远端登录、远端 worker、opencode 临时目录或常驻 crawler 服务。

## 一、环境模型

| 角色 | 位置 | 用途 |
|---|---|---|
| 执行机 | 当前 Codex Windows 主机 | 运行 Python、请求网页、下载图片、生成文件 |
| 输入 | 用户提供的本地 Excel | 至少包含企业名称列；可选官网列 |
| 用户资料 | 用户提供的本地资料目录（`--resources`） | 每家企业一个**任意命名**子文件夹，或单家企业文件夹；结构与格式不限 |
| 临时构建 | `<输出目录>\build\<run_id>\` | 本次抓取、raw、render 和门禁的暂存区 |
| 发布目录 | `<输出目录>\deliverable\` | 门禁通过后的交付副本（每家企业固定五文件夹 + 产品清单 xlsx） |
| 原件备份 | `<输出目录>\原始资料备份\` | 用户提交的原始资料原件，保留原目录结构；在 `deliverable` 之外，可用 `--backup-dir` 改位置 |

先解析 skill 根目录，后续命令都以它为准（装到任意路径只需改这一行）：

```powershell
$SkillRoot = "$env:USERPROFILE\.codex\skills\enterprise-site-pipeline"
```

解释器查找顺序：`$env:CODEX_PYTHON` → 技能目录 `.venv\Scripts\python.exe` → PATH 里的 `python` / `py -3` → Codex 主运行时
`$env:USERPROFILE\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe`。
`run_local.ps1` 会自动完成查找；核心依赖缺失时自动调用 `bootstrap.ps1` 创建技能独立 `.venv`。

脚本默认只依赖：

- Python 3.10+
- `openpyxl`
- `python-docx`
- `Pillow`（视觉核对拼版必需；缺失时视觉材料会静默消失，因此列为必需项）

安装方式：先运行 `& "$SkillRoot\bootstrap.ps1"`。它按 `requirements.lock.txt` 安装已验证版本；`requirements.txt` 只提供带主版本上限的可更新范围。

网络请求和 HTML 解析使用标准库，不要求安装 `requests` 或 `beautifulsoup4`。
`playwright` 是默认正式依赖，用于官网自动发现补搜和 JS 渲染页面。`run_local.ps1` 会优先复用技能 `.venv` 或当前 Python 环境里已有的 Playwright；缺失时才安装。锁定版本以 `requirements.lock.txt` 为准。

渲染内核按 **内置 Chromium → 本机 Microsoft Edge（`channel=msedge`）→ Edge 绝对路径** 依次自动选择。Windows 自带 Edge，且与 Chromium 同源、渲染能力一致，所以绝大多数新环境**不需要下载 100-200MB 的 Chromium**；只有 Chromium 和 Edge 都不可用时才会执行 `python -m playwright install chromium`。离线或受控环境可用 `-NoPlaywrightInstall` 跳过安装，流水线自动降级为静态抓取。

查看当前实际生效的内核：`python scripts/local_pipeline.py --browser-probe`，输出 `chromium` / `msedge` / `msedge-exe` / `none`。

Chromium 与 Edge 都不可用时，用 Codex 内置浏览器兜底：

1. 用 Codex 内置浏览器打开目标官网页面，把**渲染后的完整 HTML** 保存为 `<目录>/<名称>.html`，不要只保存可见文本。
2. 在同一目录写 `manifest.json`，建立“抓取 URL -> HTML 文件”的映射，例如 `{"https://a.com/":"a.html"}`。
3. 运行时加 `-HtmlDir "<目录>"`（CLI 加 `--html-dir "<目录>"`）。脚本会先读离线快照，未覆盖的 URL 才继续走网络/静态抓取。

## 二、一键运行

```powershell
$SkillRoot = "$env:USERPROFILE\.codex\skills\enterprise-site-pipeline"
& "$SkillRoot\scripts\run_local.ps1" `
  -Excel "D:\path\企业名录.xlsx" `
  -Out "D:\path\企业官网资料包" `
  -Limit 3
```

有用户资料时加 `-Resources`（结构与格式不限，脚本负责匹配与分类）：

```powershell
& "$SkillRoot\scripts\run_local.ps1" `
  -Excel "D:\path\企业名录.xlsx" `
  -Out "D:\path\企业官网资料包" `
  -Resources "D:\path\企业原始资料" `
  -BackupDir "D:\path\企业原始资料备份" `
  -Limit 3
```

`run_local.ps1` 会自动寻找同时包含 `openpyxl`、`python-docx`、`Pillow` 的 Python 解释器，然后调用：

```text
scripts/local_pipeline.py
```

也可以跳过包装脚本：

```powershell
$SkillRoot = "$env:USERPROFILE\.codex\skills\enterprise-site-pipeline"
python "$SkillRoot\scripts\local_pipeline.py" `
  --excel "D:\path\企业名录.xlsx" `
  --out "D:\path\企业官网资料包" `
  --limit 3
```

## 三、运行前检查

0. **先引导并自检**（新环境、换机器、升级 Codex 后必做）：

   ```powershell
   $SkillRoot = "$env:USERPROFILE\.codex\skills\enterprise-site-pipeline"
   & "$SkillRoot\bootstrap.ps1" -Excel "D:\path\企业名录.xlsx"
   ```

   引导会创建技能独立 `.venv`、按 `requirements.lock.txt` 安装依赖并运行自检。自检覆盖 Python 依赖（openpyxl / python-docx / Pillow / playwright）、渲染内核（内置 Chromium 或本机 Edge）、本机出网、MyMemory 中译英、Excel 可读性、前 3 家官网可达性或自动发现结果。任何 `[FAIL]` 都先修再跑全量。MyMemory 429/限流是 `[WARN]` 可恢复告警，不阻止对新环境的基本判定；正式生成如仍失败，会写 `英文补译清单.json`，由 Codex 补 `--en`。
1. Excel 第一个工作表包含表头；企业名称列名可用 `企业名称`、`公司名称`、`单位名称` 或 `名称`。
2. 用户资料可选，通过 `-Resources` 指向总目录；每家企业一个任意命名子文件夹（脚本先剥离"资料/文件/材料"等词、再用企业全称/核心名/品牌词匹配），单家企业时总目录本身可就是企业文件夹。资料不要求按规范结构提交，格式除图片/Word/Excel 外还可能是 PDF/PPT/txt；原件在 `deliverable` 之外备份。冲突时用户资料优先，官网只作补充。
3. 官网列可选，列名可用 `官网`、`网址`、`网站`、`官网地址`。没有官网列时，脚本会用企业全称和去地域核心名做多引擎搜索，并按域名与公司名共现、站点内容命中综合打分自动发现官网；低/中置信度结果由 Codex 打开候选站点复核后再进入正式交付，不能把低置信度结果直接当事实。
4. JS 渲染站默认走 `--playwright auto`：普通站先静态抓取，页面内容过薄且本机 Playwright 可用时自动渲染重抓；需要强制渲染用 `on`，离线排障用 `off`。
5. 输出目录不要指向 Excel 所在文件本身或已有重要交付目录。脚本会在输出目录下创建 `build/<run_id>/`。

## 四、常用参数

| 参数 | 作用 |
|---|---|
| `--selftest` | 只做环境自检后退出；可不带 `--excel` |
| `--limit N` | 只跑前 N 家企业；`0` 为全部 |
| `--timeout N` | 单次请求超时秒数，默认 20 |
| `--max-pages N` | 每家企业最多抓多少页，默认 20；产品分类/列表页优先，并下钻一层补齐叶子分类 |
| `--max-products N` | 最多保留多少产品名，默认 60 |
| `--max-images N` | 最多归档多少张图片，默认 140 |
| `--playwright {auto,on,off}` | `auto` 默认：静态优先，内容过薄时渲染重抓；`on` 强制渲染；`off` 纯静态 |
| `--browser-probe` | 只探测可用渲染内核后退出，输出 `chromium`/`msedge`/`msedge-exe`/`none` |
| `-HtmlDir <目录>` / `--html-dir <目录>` | 用 Codex 内置浏览器保存的离线 HTML 兜底抓取（目录内需 `manifest.json`，形如 `{"https://a.com/":"a.html"}`） |
| `-Resources <目录>` / `--resources <目录>` | 用户资料总目录：每家企业一个任意命名子文件夹，或单家企业文件夹；结构与格式不限（图片/Word/Excel/PDF/PPT/txt…） |
| `-BackupDir <目录>` / `--backup-dir <目录>` | 原始资料备份目录，默认 `<输出目录>\原始资料备份`，位于 `deliverable` 之外，保留原目录结构 |
| `--require-visual` | 视觉核对未完成时按 error 处理，发布前应开启 |
| `--publish-stage <build/run_id>` | 不重抓，直接把已完成的 stage 发布到 `--out` |
| `--no-publish` | 门禁通过也不发布，仅保留 build 目录，供 Codex 看图核对后再发布 |
| `--proxy URL` | 使用本机 HTTP(S) 代理 |
| `--insecure` | 跳过 TLS 校验，仅用于用户明确承认的测试环境 |
| `--en FILE` | 用人工确认的英文 JSON 覆盖自动翻译 |
| `--no-translate` | 关闭自动翻译，英文层留空；默认仍阻断发布 |
| `--accept-no-english` | 仅由用户明确接受中文版时使用；英文相关门禁降为告警 |
| `--translate-email MAIL` | 可选，MyMemory 联系邮箱，提高匿名翻译额度 |
| `--translate-delay SEC` | 每次翻译调用后的间隔秒数，默认 0.2 |
| `--no-visual-review` | 跳过拼版、图片核对表和 `视觉核对.json` 生成 |
| `--strict` | 门禁未通过时返回非零退出码；`run_local.ps1` 默认透传，`-AllowRed` 可关闭 |

## 五、运行结果怎么看

一次运行结束后先看：

```text
<输出目录>\build\<run_id>\gates.json
<输出目录>\build\<run_id>\gates.log
<输出目录>\build\<run_id>\manifest.json
```

企业状态：

- `ok`：页面和图片基本抓取成功；
- `partial`：有页面错误或中文简介不足三段；
- `empty_images`：未抓到可归档图片；
- `no_website`：本机搜索和校验未确认官网。

`no_website` 和 `empty_images` 都可以是正常的数据边界，但必须在交付说明里列出，不能当作全量成功。

## 六、中英双语与英文定稿

不提供 `--en` 时，脚本默认自动生成英文草稿：企业英文名、英文标题、品牌、三段英文简介、每个产品的英文名和产品详情，写入 docx 和 xlsx。`en.json` 用 `自动翻译: true`、`定稿: false`、`翻译失败`、`备注` 记录状态：

```json
{
  "英文名": "Tongwei Group Co., Ltd.",
  "英文标题": "Tongwei Group Co., Ltd.",
  "品牌": "Tongwei",
  "英文简介": ["段1", "段2", "段3"],
  "产品英名": {"多晶硅": "Polysilicon"},
  "产品详情英": {"多晶硅": "Polysilicon is ..."},
  "子品类": "新能源",
  "定稿": false,
  "自动翻译": true,
  "翻译引擎": "MyMemory",
  "翻译失败": false,
  "备注": "英文为自动翻译草稿，待人工核校"
}
```

翻译走 MyMemory 免费接口（`zh-CN → en`），结果缓存在 `<输出目录>\_translate_cache.json`，重复运行不重复计费。可用 `--translate-email` 提高额度，`--no-translate` 关闭自动翻译。

需要人工定稿时，用已确认英文覆盖：

```powershell
& $PY "...\local_pipeline.py" `
  --excel "D:\path\企业名录.xlsx" `
  --out "D:\path\企业官网资料包" `
  --en "D:\path\已确认英文.json"
```

自动英文是机翻草稿，交付时必须标注“自动翻译，待人工核校”，不得当作人工定稿。

MyMemory 429/限流时，自检只给 `[WARN]`；正式生成会在 `<run_id>/英文补译清单.json` 列出问题。默认英文缺失会阻断发布，正常修复路径是 Codex 基于 raw 中文事实生成 `--en` 兼容草稿并重跑。仅在用户在当次对话中明确接受中文版时，才可加 `--accept-no-english` 将英文相关门禁降为告警。

## 七、视觉核对

渲染后自动生成：

```text
<run_id>\review\视觉核对图\<企业>\0.总览.png    # 四类速览
<run_id>\review\视觉核对图\<企业>\<分类>.png     # 分类拼版，单页 ≤12 张、≤1MB；超出为 <分类>_p1.png、_p2.png…
<run_id>\review\<企业>\图片核对表.xlsx           # 带缩略图，结论列可下拉
<run_id>\review\核对指引.md                       # Codex 看图清单、判断口径和执行命令
<run_id>\视觉核对.json                            # 结论载体
```

Codex 必须自己完成看图核对，不得把判断推给用户。拼版图进入会话后是 base64，会持续撑大会话历史（历史故障：19 张拼版 28.8MB 触发上游报错），因此必须短线程分批：每批 ≤10 家，一个会话做完「查体积 → 看图 → 回写 → 门禁 → 发布」就结束，企业多时另开会话。开工前先跑 `python "$SkillRoot\scripts\session_guard.py"`，本会话 rollout 超过 20MB 先换会话。

步骤：

1. 先跑 `scripts\session_guard.py` 查会话体积（>20MB 换会话），再打开 `review\核对指引.md`，按其中列出的绝对路径逐家看图；
2. 图片多时先用 `<企业>/图片核对表.xlsx` 初筛（分类/序号/尺寸/图片URL/来源页面/alt），按指引标注的序号范围打开对应 `<分类>_pN.png` 逐张核对；图片少时打开 `0.总览.png` 初筛、疑点再看分类拼版；同时确认简介 docx 与产品清单是否和原始页面一致；
3. 把结论写进 `review\verdicts.json`（类别层或单张图均可），再用 `visual_review.py --apply` 回写 `视觉核对.json`；
4. 用 `gates.py --require-visual` 复核，全绿后用 `local_pipeline.py --publish-stage` 发布。

回写与发布命令示例：

```powershell
$SkillRoot = "$env:USERPROFILE\.codex\skills\enterprise-site-pipeline"
$Run = "D:\path\企业官网资料包\build\20261007-204839"
python "$SkillRoot\scripts\visual_review.py" `
  --json "$Run\视觉核对.json" `
  --apply "$Run\review\verdicts.json" `
  --reviewer Codex
python "$SkillRoot\scripts\gates.py" `
  --deliverable "$Run\deliverable" `
  --raw "$Run\raw" `
  --en "$Run\en.json" `
  --summary "$Run\汇总.xlsx" `
  --require-visual
python "$SkillRoot\scripts\local_pipeline.py" `
  --publish-stage "$Run" `
  --out "D:\path\企业官网资料包"
```

`视觉核对.json` 结构：

```jsonc
{
  "示例科技有限公司": {
    "类别": {
      "1.企业工厂图": {"结论": "符合", "说明": ""},
      "2.企业产品图": {"结论": "不符", "说明": "混入车间照片，需人工重分类"}
    },
    "图片": {
      "2.企业产品图/01_prod.jpg": {"结论": "不符", "说明": "实际是车间照片"}
    },
    "文档": {"结论": "符合", "说明": ""},
    "产品清单": {"结论": "待核对", "说明": ""}
  }
}
```

规则：

- 类别结论非空时，该类未单独标注的图片继承类别结论；单张结论优先。
- 填“不符”会让 `visual_review` 门禁报 error，即使其他门禁全绿也不会发布。
- 留空即待核对，默认只告警；加 `--require-visual` 后按 error。
- 重跑 `visual_review.py` 或整条流水线会合并保留已有结论。

## 八、故障排查

| 现象 | 先查什么 | 处理 |
|---|---|---|
| 找不到 Python 包 | 先运行 `bootstrap.ps1` 或 `--selftest` 看缺哪一项 | 用技能 `.venv`，或用 `$env:CODEX_PYTHON` 指定解释器后安装 `requirements.lock.txt` |
| 视觉核对材料没生成 | `visual_review` 门禁是否报错、Pillow 是否可导入 | 缺失 Pillow 时视觉材料会静默消失；装上 Pillow 后重跑，别把它当"没有疑点" |
| 门禁通过但发布报错/没发布 | 控制台是否出现"发布目录被占用" | 被占用的 xlsx 关掉后重跑；或直接用自动改发的 `<输出>\publish_<run_id>\` |
| 门禁红灯但退出码是 0 | 是否用了 `-AllowRed` 或没走 `run_local.ps1` | 去掉 `-AllowRed`，或给 `local_pipeline.py` 显式加 `--strict` |
| 官网自动发现低/中置信度 | `gates.json` 的 `site_discovery`、`raw/<企业>.json` 的候选与分数 | 由 Codex 打开候选站点核对；确认后补 Excel 官网列或保留确认记录，不把低置信度结果直接当事实 |
| 图片为空 | 页面是否 JS 渲染、图片是否要求 Referer | 默认保持 `--playwright auto`；`run_local.ps1` 会复用或自动安装 Playwright，仍抓不到再人工补图并标注 |
| 英文门禁红 | `en.json` 的 `翻译失败`、`英文补译清单.json`、`en_entry`、`en_ascii`、`product_detail` | 正常修复是 Codex 基于 raw 中文事实生成 `--en` 草稿并重跑（保留 `定稿: false`、`自动翻译: true`）；只有用户明确接受中文版才加 `--accept-no-english` |
| 产品详情缺英文译文 | MyMemory 返回 HTTP 429 限流 | 429 在自检中为可恢复 `WARN`；正式生成会写英文补译清单。可稍后重跑、加 `--translate-email`，或由 Codex 补 `--en` 草稿 |
| 页面抓取慢 | 超时、最大页数、页面数量 | 先小样本，必要时调小 `--max-pages` |
| 网络请求失败 | 代理、TLS、目标站点限制 | 使用 `--proxy` 或 `--insecure` 仅做明确测试 |
| 视觉核对红 | `视觉核对.json` 的“不符”/“待核对”条目 | “不符”先重新归类或补图；未回写时由 Codex 看图后写 `review\verdicts.json` 并执行 `visual_review.py --apply`，再用 `gates.py --require-visual` 复核 |
| 渲染内核不可用 | `--browser-probe` 输出 `none` | 确认本机 Edge 存在（Windows 默认自带）；仍无则用 `-NoPlaywrightInstall` 降级静态抓取，或联网后执行 `python -m playwright install chromium` |
| Chromium 下载卡住 | `playwright install chromium` 长时间停在 0% | 不必下载：本机 Edge 会被自动复用；确认 `--browser-probe` 输出 `msedge` 即可 |
| 内核和 Edge 都没有 | 反病毒/精简系统裁掉 Edge 的离线机 | 由 Codex 用内置浏览器打开目标页并保存 HTML，写 `manifest.json` 后加 `--html-dir` 兜底抓取 |
| 视觉结论没回写 | `visual_review` 是否仍为 warn | 由 Codex 完成看图，写 `review\verdicts.json` 后执行 `visual_review.py --apply`；`--require-visual` 会把未完成核对判为 error |
| 图片来源非官网 | `image_provenance` 的 offenders | 确认是否为企业自有站点或可信 CDN；别家站点图片必须剔除 |
| 产品清单行数不符 | `product_rows`、raw `products` | 手工改过 xlsx 就重跑渲染，不要只改交付文件 |
| 产品图本地链接为空/失效 | `product_image_link`、产品清单 `图片（本地连接）` | 先确认产品来自产品分类/列表页且已下钻；检查 raw `product` 的 `alt` 与产品名是否精确/系列名匹配；官网无图会显式标注，不要写旧占位符 |

## 九、安全约束

- 不把密码、token、cookie 写入 skill、Excel、日志或命令历史；
- 不要求为了本流水线配置远端账号；
- 不修改用户未指定的业务文件；
- 不在门禁红色时把结果标成终稿；
- 输入表缺官网列时先自动发现；低/中置信度域名必须由 Codex 打开站点复核，确认前不进入正式交付，不把低置信度结果直接当事实；
- 视觉核对由 Codex 自己完成，不得把看图判断转交给用户；未回写结论的条目在 `--require-visual` 下按 error 处理。