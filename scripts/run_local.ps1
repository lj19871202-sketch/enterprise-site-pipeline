param(
    [Parameter(Mandatory = $true)]
    [string]$Excel,

    [string]$Out = "",
    [string]$En = "",
    [string]$HtmlDir = "",          # Codex 内置浏览器保存的离线 HTML 目录（含 manifest.json）
    [string]$Resources = "",        # 用户资料总目录：每家一个任意命名子文件夹，或单家企业文件夹
    [string]$BackupDir = "",        # 原始资料备份目录（默认 <输出目录>\原始资料备份，在 deliverable 之外）
    [string]$SiteDecisions = "",    # 官网候选复核表：决定优先；中/低置信度未复核默认跳过
    [switch]$DiscoverOnly,          # 只做官网发现，产出「官网候选复核表.xlsx」后退出
    [switch]$AcceptNoEnglish,       # 用户明确接受中文版；英文相关门禁降为告警
    [int]$Limit = 0,
    [string]$TranslateEmail = "",
    [switch]$NoTranslate,
    [switch]$NoVisualReview,
    [switch]$Playwright,            # 强制 JS 渲染（等价 --playwright on）
    [switch]$NoPlaywright,          # 关闭 JS 渲染（等价 --playwright off）
    [switch]$NoPlaywrightInstall,   # 只检测，不自动安装 Playwright/Chromium
    [switch]$RequireVisual,         # 兼容旧参数；视觉核对默认已强制
    [switch]$SkipVisualReview,      # 显式跳过视觉核对门禁（仅调试/用户明确授权）
    [switch]$NoPublish,             # 先只构建，Codex 核对后再 --publish-stage
    [int]$PlaywrightInstallTimeoutSec = 600,  # 单次 Playwright/Chromium 安装超时（秒）
    [switch]$NoBootstrap,          # 缺核心依赖时不自动创建 .venv
    [switch]$AllowBuilderCdn,      # 放行建站平台自有 CDN（faiusr/faisys/508sys），默认关闭；仅来源页同域时生效
    # 默认严格模式：门禁出现 error 时返回非零退出码；缺失英文须用 -En 补稿或显式 -AcceptNoEnglish。
    # 仅在明确接受"红灯产物"的调试场景加 -AllowRed 关闭。
    [switch]$AllowRed
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$skillRoot = Split-Path -Parent $scriptDir
$pipeline = Join-Path $scriptDir "local_pipeline.py"
$bootstrap = Join-Path $skillRoot "bootstrap.ps1"
$venvPython = Join-Path $skillRoot ".venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $pipeline)) {
    throw "找不到 local_pipeline.py：$pipeline"
}
if (-not (Test-Path -LiteralPath $Excel)) {
    throw "找不到 Excel：$Excel"
}
if (-not $Out) {
    $Out = Join-Path (Get-Location).Path "enterprise-site-output"
}

$excelPath = (Resolve-Path -LiteralPath $Excel).Path
$outPath = [IO.Path]::GetFullPath($Out)

function Invoke-NativeProbe {
    # 运行探测类原生命令，并把 stderr 一并吞掉。
    # PS 5.1 在 $ErrorActionPreference="Stop" 下会把原生命令的 stderr 当作
    # NativeCommandError 终止整个脚本，所以这里临时降级错误策略。
    param(
        [string]$FilePath,
        [string[]]$Arguments
    )
    $prev = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        & $FilePath @Arguments 2>&1 | Out-Null
        return $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $prev
    }
}

function Invoke-ProcessWithTimeout {
    param(
        [string]$FilePath,
        [string[]]$Arguments,
        [int]$TimeoutSec
    )
    $proc = Start-Process -FilePath $FilePath -ArgumentList $Arguments -PassThru -NoNewWindow
    if (-not $proc.WaitForExit($TimeoutSec * 1000)) {
        & taskkill.exe /PID $proc.Id /T /F 2>$null | Out-Null
        try { $proc.Kill() } catch { }
        throw "命令超时（${TimeoutSec}s）：$FilePath $($Arguments -join ' ')"
    }
    return $proc.ExitCode
}

# ---- 找 Python：优先技能独立 .venv，其次核心依赖 + 本地已有 Playwright ----
$candidates = @()
if (Test-Path -LiteralPath $venvPython) { $candidates += $venvPython }
if ($env:CODEX_PYTHON) { $candidates += $env:CODEX_PYTHON }
foreach ($name in @("python.exe", "python3.exe")) {
    $cmd = Get-Command $name -ErrorAction SilentlyContinue
    if ($cmd) { $candidates += $cmd.Source }
}
$candidates += (Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe")
$candidates = $candidates | Where-Object { $_ } | Select-Object -Unique

$python = $null
$pythonHasPlaywright = $false
# 第一轮：核心依赖齐全，且本机已有 Playwright
foreach ($candidate in $candidates) {
    if (-not (Test-Path -LiteralPath $candidate)) { continue }
    $probeExit = Invoke-NativeProbe -FilePath $candidate -Arguments @("-c", "import openpyxl, docx, PIL, playwright")
    if ($probeExit -eq 0) {
        $python = $candidate
        $pythonHasPlaywright = $true
        break
    }
}
# 第二轮：核心依赖齐全，Playwright 可以随后安装
if (-not $python) {
    foreach ($candidate in $candidates) {
        if (-not (Test-Path -LiteralPath $candidate)) { continue }
        $probeExit = Invoke-NativeProbe -FilePath $candidate -Arguments @("-c", "import openpyxl, docx, PIL")
        if ($probeExit -eq 0) {
            $python = $candidate
            break
        }
    }
}
if (-not $python) {
    if (-not $NoBootstrap) {
        if (-not (Test-Path -LiteralPath $bootstrap)) {
            throw "缺少核心依赖，且找不到 bootstrap.ps1：$bootstrap"
        }
        Write-Host "[bootstrap] 未找到完整依赖，正在创建技能独立环境..."
        & $bootstrap -SkipSelftest
        if ($LASTEXITCODE -ne 0) { throw "bootstrap.ps1 失败，退出码 $LASTEXITCODE" }
        if (Test-Path -LiteralPath $venvPython) {
            $probeExit = Invoke-NativeProbe -FilePath $venvPython -Arguments @("-c", "import openpyxl, docx, PIL, playwright")
            if ($probeExit -eq 0) {
                $python = $venvPython
                $pythonHasPlaywright = $true
            }
        }
    }
    if (-not $python) {
        throw ("未找到同时含 openpyxl / python-docx / Pillow 的 Python。" +
               "请运行 bootstrap.ps1，或设置 CODEX_PYTHON 后安装 requirements.lock.txt。")
    }
}

# ---- Playwright：检测本地 -> 缺失则安装 Playwright；渲染内核优先本机 Edge ----
if (-not $NoPlaywright) {
    if (-not $pythonHasPlaywright) {
        $probeExit = Invoke-NativeProbe -FilePath $python -Arguments @("-c", "import playwright")
        if ($probeExit -eq 0) { $pythonHasPlaywright = $true }
    }
    if (-not $pythonHasPlaywright) {
        if ($NoPlaywrightInstall) {
            Write-Warning "当前 Python 未安装 Playwright；auto 模式将降级为静态抓取。"
        } else {
            Write-Host "[playwright] 未检测到，正在安装到当前 Python：$python"
            $env:PLAYWRIGHT_DOWNLOAD_CONNECTION_TIMEOUT = "60000"
            try {
                $pipExit = Invoke-ProcessWithTimeout -FilePath $python `
                    -Arguments @("-m", "pip", "install", "--disable-pip-version-check", "playwright>=1.40") `
                    -TimeoutSec $PlaywrightInstallTimeoutSec
                if ($pipExit -ne 0) { throw "pip 返回退出码 $pipExit" }
                $pythonHasPlaywright = $true
            } catch {
                if ($Playwright) {
                    throw "Playwright 安装失败（强制渲染模式）：$($_.Exception.Message)"
                }
                Write-Warning "Playwright 安装失败，auto 模式将降级为静态抓取：$($_.Exception.Message)"
                Write-Warning "可稍后重试： $python -m pip install --disable-pip-version-check playwright"
            }
        }
    }
    if ($pythonHasPlaywright) {
        # 内置 Chromium 缺失时优先复用本机 Edge，避免下载 100-200MB。
        $probeOut = & $python $pipeline --browser-probe 2>$null
        $pwKernel = (@($probeOut) | Select-Object -Last 1)
        if ($pwKernel) { $pwKernel = "$pwKernel".Trim() }
        if ($pwKernel -eq "chromium") {
            Write-Host "[playwright] 渲染内核：Playwright 内置 Chromium"
        } elseif ($pwKernel -eq "msedge" -or $pwKernel -eq "msedge-exe") {
            Write-Host "[playwright] 渲染内核：本机 Microsoft Edge（无需下载 Chromium）"
        } else {
            $pwKernel = ""
            if ($NoPlaywrightInstall) {
                Write-Warning "未检测到 Chromium 或 Edge；auto 模式将降级为静态抓取。"
                $pythonHasPlaywright = $false
            } else {
                Write-Host "[playwright] 未检测到 Chromium/Edge，正在安装 Chromium（首次约 100-200MB）..."
                $env:PLAYWRIGHT_DOWNLOAD_CONNECTION_TIMEOUT = "60000"
                try {
                    $pwExit = Invoke-ProcessWithTimeout -FilePath $python `
                        -Arguments @("-m", "playwright", "install", "chromium") `
                        -TimeoutSec $PlaywrightInstallTimeoutSec
                    if ($pwExit -ne 0) { throw "playwright install 返回退出码 $pwExit" }
                    $pwKernel = "chromium"
                } catch {
                    if ($Playwright) {
                        throw "渲染内核不可用（强制渲染模式）：$($_.Exception.Message)"
                    }
                    Write-Warning "Chromium 安装失败，auto 模式将降级为静态抓取：$($_.Exception.Message)"
                    Write-Warning "可安装 Edge 后自动启用，或稍后重试： $python -m playwright install chromium"
                    $pythonHasPlaywright = $false
                }
            }
        }
    }
}

$pwMode = "auto"
if ($NoPlaywright) { $pwMode = "off" }
elseif ($Playwright) { $pwMode = "on" }

$runArgs = @($pipeline, "--excel", $excelPath, "--out", $outPath, "--playwright", $pwMode)
if ($En) {
    if (-not (Test-Path -LiteralPath $En)) { throw "找不到英文 JSON：$En" }
    $runArgs += @("--en", (Resolve-Path -LiteralPath $En).Path)
}
if ($HtmlDir) {
    if (-not (Test-Path -LiteralPath $HtmlDir)) { throw "找不到离线 HTML 目录：$HtmlDir" }
    $runArgs += @("--html-dir", (Resolve-Path -LiteralPath $HtmlDir).Path)
}
if ($Resources) {
    if (-not (Test-Path -LiteralPath $Resources)) { throw "找不到资料目录：$Resources" }
    $runArgs += @("--resources", (Resolve-Path -LiteralPath $Resources).Path)
}
if ($BackupDir) { $runArgs += @("--backup-dir", [IO.Path]::GetFullPath($BackupDir)) }
if ($SiteDecisions) {
    if (-not (Test-Path -LiteralPath $SiteDecisions)) { throw "找不到官网复核表：$SiteDecisions" }
    $runArgs += @("--site-decisions", (Resolve-Path -LiteralPath $SiteDecisions).Path)
}
if ($DiscoverOnly) { $runArgs += "--discover-only" }
if ($Limit -gt 0) { $runArgs += @("--limit", [string]$Limit) }
if ($TranslateEmail) { $runArgs += @("--translate-email", $TranslateEmail) }
if ($NoTranslate) { $runArgs += "--no-translate" }
if ($AcceptNoEnglish) { $runArgs += "--accept-no-english" }
if ($NoVisualReview) { $runArgs += "--no-visual-review" }
if ($SkipVisualReview) {
    Write-Host "警告：已显式跳过视觉核对门禁，仅用于调试或用户明确授权。" -ForegroundColor Yellow
    $runArgs += "--skip-visual-review"
} elseif ($RequireVisual) {
    $runArgs += "--require-visual"
}
if ($NoPublish) { $runArgs += "--no-publish" }
if ($AllowBuilderCdn) { $runArgs += "--allow-builder-cdn" }
if (-not $AllowRed) { $runArgs += "--strict" }

& $python @runArgs
exit $LASTEXITCODE
