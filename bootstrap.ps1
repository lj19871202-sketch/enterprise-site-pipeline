param(
    [string]$Python = "",
    [switch]$Force,
    [switch]$SkipSelftest,
    [string]$Excel = ""
)

$ErrorActionPreference = "Stop"
$SkillRoot = $PSScriptRoot
$VenvRoot = Join-Path $SkillRoot ".venv"
$VenvPython = Join-Path $VenvRoot "Scripts\python.exe"
$LockFile = Join-Path $SkillRoot "requirements.lock.txt"
$SelfTest = Join-Path $SkillRoot "scripts\local_pipeline.py"

function Invoke-NativeProbe {
    param([string]$FilePath, [string[]]$Arguments)
    $prev = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        & $FilePath @Arguments 2>&1 | Out-Null
        return $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $prev
    }
}

function Test-Python {
    param([string]$FilePath)
    if (-not $FilePath -or -not (Test-Path -LiteralPath $FilePath)) { return $false }
    $exitCode = Invoke-NativeProbe -FilePath $FilePath -Arguments @(
        "-c", "import sys; raise SystemExit(0 if sys.version_info >= (3, 9) else 1)"
    )
    return $exitCode -eq 0
}

function Resolve-PyLauncher {
    $py = Get-Command "py.exe" -ErrorAction SilentlyContinue
    if (-not $py) { return "" }
    $prev = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        $resolved = & $py.Source -3 -c "import sys; print(sys.executable)" 2>$null
        if ($LASTEXITCODE -eq 0 -and $resolved) {
            return ([string](@($resolved) | Select-Object -Last 1)).Trim()
        }
    } catch {
    } finally {
        $ErrorActionPreference = $prev
    }
    return ""
}

if (-not (Test-Path -LiteralPath $LockFile)) {
    throw "找不到依赖锁定文件：$LockFile"
}
if (-not (Test-Path -LiteralPath $SelfTest)) {
    throw "找不到 local_pipeline.py：$SelfTest"
}

$candidates = New-Object System.Collections.Generic.List[string]
if ($Python) { $candidates.Add($Python) }
if ($env:CODEX_PYTHON) { $candidates.Add($env:CODEX_PYTHON) }
if (Test-Path -LiteralPath $VenvPython) { $candidates.Add($VenvPython) }
foreach ($name in @("python.exe", "python3.exe")) {
    $cmd = Get-Command $name -ErrorAction SilentlyContinue
    if ($cmd) { $candidates.Add($cmd.Source) }
}
$pyResolved = Resolve-PyLauncher
if ($pyResolved) { $candidates.Add($pyResolved) }

$runtimeRoot = Join-Path $env:USERPROFILE ".cache\codex-runtimes"
if (Test-Path -LiteralPath $runtimeRoot) {
    $fixed = Join-Path $runtimeRoot "codex-primary-runtime\dependencies\python\python.exe"
    if (Test-Path -LiteralPath $fixed) { $candidates.Add($fixed) }
    Get-ChildItem -LiteralPath $runtimeRoot -Directory -ErrorAction SilentlyContinue |
        ForEach-Object {
            $candidate = Join-Path $_.FullName "dependencies\python\python.exe"
            if (Test-Path -LiteralPath $candidate) { $candidates.Add($candidate) }
        }
}

$basePython = ""
foreach ($candidate in ($candidates | Where-Object { $_ } | Select-Object -Unique)) {
    if (Test-Python -FilePath $candidate) {
        $basePython = $candidate
        break
    }
}
if (-not $basePython) {
    throw ("未找到 Python 3.9+。请设置 CODEX_PYTHON，或先安装 Python 后重试。" +
           "Codex 桌面环境通常自带 runtime Python。")
}

if ($Force -and (Test-Path -LiteralPath $VenvRoot)) {
    Write-Host "[bootstrap] --Force：重建 $VenvRoot"
    $exitCode = Invoke-NativeProbe -FilePath $basePython -Arguments @("-m", "venv", "--clear", $VenvRoot)
    if ($exitCode -ne 0) { throw "创建/清理 venv 失败，退出码 $exitCode" }
} elseif (-not (Test-Path -LiteralPath $VenvPython)) {
    Write-Host "[bootstrap] 创建独立环境：$VenvRoot"
    $exitCode = Invoke-NativeProbe -FilePath $basePython -Arguments @("-m", "venv", $VenvRoot)
    if ($exitCode -ne 0) { throw "创建 venv 失败，退出码 $exitCode" }
}

Write-Host "[bootstrap] 安装锁定依赖：$LockFile"
$pipExit = -1
$prev = $ErrorActionPreference
$ErrorActionPreference = "Continue"
try {
    & $VenvPython -m pip install --disable-pip-version-check --requirement $LockFile
    $pipExit = $LASTEXITCODE
} finally {
    $ErrorActionPreference = $prev
}
if ($pipExit -ne 0) {
    throw "依赖安装失败，退出码 $pipExit。检查网络或手动执行： $VenvPython -m pip install -r $LockFile"
}

$probeExit = Invoke-NativeProbe -FilePath $VenvPython -Arguments @(
    "-c", "import openpyxl, docx, PIL, playwright"
)
if ($probeExit -ne 0) {
    throw "依赖校验失败：openpyxl / python-docx / Pillow / playwright 未全部导入成功"
}

Write-Host "[bootstrap] Python：$VenvPython"
if (-not $SkipSelftest) {
    $testArgs = @($SelfTest, "--selftest")
    if ($Excel) {
        if (-not (Test-Path -LiteralPath $Excel)) { throw "找不到 Excel：$Excel" }
        $testArgs += @("--excel", (Resolve-Path -LiteralPath $Excel).Path)
    }
    Write-Host "[bootstrap] 运行环境自检"
    & $VenvPython @testArgs
    if ($LASTEXITCODE -ne 0) {
        throw "环境自检未通过。按上面的 FAIL/WARN 修复后重试。"
    }
}
Write-Host "[bootstrap] 完成。可运行： & `"$SkillRoot\scripts\run_local.ps1`" -Excel `"<企业名录.xlsx>`""
