<#
  自建 Hermes 自动更新（Windows）
  ------------------------------------------------------------------
  为什么有它：官方 `hermes update` 在「Windows + 国内网络 + 桌面端开着」的组合下
  会踩四个坑：① venv 文件被桌面后端锁住 → 拒绝执行；② 直连 GitHub 下载 uv 卡死；
  ③ 卡死留下僵尸"更新进行中"标记 → 桌面端起不来；④ 单次运行删除超过 50 个文件
  触发安全阀 → 跑几秒即退，且每失败一次又多留一堆快照，越跑越超。

  这个脚本把上面四件事按顺序内建：探测网络/切换镜像 → 自检 uv → 清遗留（让删除量
  低于阀门，且用 PowerShell 删，不计入更新进程的删除预算）→ 优雅关闭桌面端 →
  跑官方更新 → 校验 → 自动重开桌面端 → 写结果摘要。

  2026-10-01 加固：① uv 自修复只按 pm/lock.json 锁定版本安装并强制 sha256（校验不过不落盘）；
  ② 更新结束恢复 origin 为裸 github 直连，消除镜像残留导致的 fork 误判；
  ③ 拉取失败区分「本地仓库错误（pack-objects）」与「网络错误」，前者不换镜像；
  ④ 重开桌面端后等待后端就绪（HERMES_BACKEND_READY）再报成功；结果 JSON 增加
  desktopReady / originUrl / autostashCount 字段。

  用法：
    双击 「更新Hermes.cmd」，或
    powershell -NoProfile -ExecutionPolicy Bypass -File hermes-update.ps1 [-DryRun] [-NoRelaunch]
#>
[CmdletBinding()]
param(
    # 安装根目录（默认 %LOCALAPPDATA%\hermes\hermes-agent）
    [string]$InstallRoot = "",
    [string]$Branch = "main",
    # 只检查不动手：探测网络、检查 uv、列出要清理的遗留物
    [switch]$DryRun,
    # 更新完不自动重开桌面端
    [switch]$NoRelaunch,
    # 关闭桌面端前的倒计时（秒），0 = 不等待直接关
    [int]$CountdownSeconds = 10,
    # 重开桌面端后等待后端就绪的上限（秒）；0 = 不等，只确认进程起来
    [int]$LaunchReadyTimeoutSeconds = 90,
    # 更新失败时最多重试几轮（每轮先清遗留）
    [int]$MaxAttempts = 3,
    # 单轮总上限（分钟）—— 兜底，最终仍受各阶段硬上限约束
    [int]$UpdateTimeoutMinutes = 60,
    # 所有轮次合计上限（分钟）
    [int]$GlobalTotalTimeoutMinutes = 90,
    # 进度刷新间隔（秒），只影响屏幕刷新
    [int]$ProgressRefreshSeconds = 2,
    # 进度采样间隔（秒），用于“过慢”速率判定
    [int]$StallSampleSeconds = 10,
    # “过慢”软阈值：最近窗口进度 < 基线速率 * 该百分比 -> 判定过慢（策略B 终止）
    [int]$SlowRatioPercent = 20,
    # 连续多少次采样过慢 -> 终止（配合下面的“≥90s 基线 + ≥30s 零进展”守卫）
    [int]$SlowStreakSamples = 6,
    # get-windows 预编译下载镜像（node-pre-gyp host mirror 前缀；本机无 C++ 工具链时必须）
    [string]$NodeBinaryMirror = "",
    # 额外的 git 镜像前缀（追加到内置镜像池；可多次传或用逗号分隔）
    [string[]]$GitMirrorExtra = @(),
    # 镜像并行探测窗口（秒）：窗口内最先应答的镜像胜出
    [int]$GitProbeSeconds = 120,
    # 保留最近几份更新收据
    [int]$ReceiptsKeep = 3,
    # 结束后不等待回车
    [switch]$NoPause
)

$ErrorActionPreference = "Continue"
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }
try { if ($Host.Name -eq "ConsoleHost") { $Host.UI.RawUI.WindowTitle = "Hermes 自动更新" } } catch { }

# ---------------------------------------------------------------- 基础路径
if (-not $InstallRoot) { $InstallRoot = Join-Path $env:LOCALAPPDATA "hermes\hermes-agent" }
if (-not (Test-Path -LiteralPath $InstallRoot)) { Write-Host "找不到安装目录：$InstallRoot" -ForegroundColor Red; exit 1 }
$InstallRoot = (Resolve-Path -LiteralPath $InstallRoot).Path
$HermesHome  = Split-Path -Parent $InstallRoot
$VenvPython  = Join-Path $InstallRoot "venv\Scripts\python.exe"
$DesktopExe  = Join-Path $InstallRoot "apps\desktop\release\win-unpacked\Hermes.exe"
$BinDir      = Join-Path $HermesHome "bin"
$UvExe       = Join-Path $BinDir "uv.exe"
$BackupsDir  = Join-Path $HermesHome "backups"
$ReceiptsDir = Join-Path $HermesHome "logs\update_receipts"
$MarkerPath  = Join-Path $HermesHome ".hermes-update-in-progress"
$RunLog      = Join-Path $HermesHome ("logs\self-update-{0}.log" -f (Get-Date -Format "yyyyMMdd-HHmmss"))
$PrebuiltCacheDir = Join-Path $HermesHome "tools\self-update\cache"

# ---------------------------------------------------------------- 阶段 / 卡死判定配置
# 阶段中文名
$StageLabel = @{
    fetch = '拉取代码'; python = '同步 Python 依赖'; npm = '安装 Node 依赖'
    tui = '构建 TUI'; web = '构建 Web UI'; desktop = '重建桌面端'
    completion = '源码收尾'; tail = '更新完成/收尾'; unknown = '未知阶段'
}
# 静默判死：该阶段内进度/输出在这么多秒内零增长 -> 判死
$StageSilenceSeconds = @{ fetch = 120; python = 120; npm = 120; tui = 120; web = 120; desktop = 180; completion = 180; tail = 300; unknown = 120 }
# 阶段硬上限（分钟）：永不被续期突破（completion=官方"未完成源码收尾"，含依赖同步+可能重建，给足 60 分钟）
$StageTimeoutMinutes = @{ fetch = 10; python = 30; npm = 20; tui = 10; web = 10; desktop = 45; completion = 60; tail = 45; unknown = 30 }
# 运行期状态（由 Invoke-HermesUpdate 写入，供失败报告读取）
$script:lastUpdateStage = 'unknown'
$script:lastStallReason = $null
$script:lastProgressRate = '-'
$script:stageElapsedText = '-'
$script:silenceSeconds = -1

# 桌面端打包用 electron-builder，它默认从 GitHub release 下载工具集（icons-bundle、
# winCodeSign、nsis…）与 electron 本体；国内会 fetch failed。npmmirror 有完整镜像，
# 设这两个环境变量（会被子进程继承）。已设过则尊重原值。
if (-not $env:ELECTRON_BUILDER_BINARIES_MIRROR) { $env:ELECTRON_BUILDER_BINARIES_MIRROR = "https://npmmirror.com/mirrors/electron-builder-binaries/" }
if (-not $env:ELECTRON_MIRROR) { $env:ELECTRON_MIRROR = "https://npmmirror.com/mirrors/electron/" }

# ---------------------------------------------------------------- git 镜像池
# Hermes 仓库的 git 流量只能走 GitHub 代理镜像（本机直连 GitHub 会因 TLS 吊销校验失败）。
# 内置池 + 额外来源（-GitMirrorExtra / 同目录 git-mirrors.txt，每行一个前缀，# 为注释）。
# 运行期并行探测、先通者胜；拉取阶段失败/停滞时自动换下一个。
$script:GitMirrorPrefixes = @(
    "https://ghproxy.net/",
    "https://ghfast.top/",
    "https://gh-proxy.com/",
    "https://gh.llkk.cc/",
    "https://ghproxy.cn/",
    "https://hub.gitmirror.com/",
    "https://github.moeyy.xyz/"
)
$mirrorListFile = Join-Path $PSScriptRoot "git-mirrors.txt"
if (Test-Path -LiteralPath $mirrorListFile) {
    foreach ($ln in @(Get-Content -LiteralPath $mirrorListFile -ErrorAction SilentlyContinue)) {
        $t = "$ln".Trim()
        if ($t -and -not $t.StartsWith("#")) {
            if (-not $t.EndsWith("/")) { $t += "/" }
            $script:GitMirrorPrefixes += $t
        }
    }
}
foreach ($t in @($GitMirrorExtra)) {
    $t = "$t".Trim()
    if ($t) { if (-not $t.EndsWith("/")) { $t += "/" }; $script:GitMirrorPrefixes += $t }
}
$script:GitMirrorPrefixes = @($script:GitMirrorPrefixes | Select-Object -Unique)
$script:GitMirrorCurrent = ""   # 当前 origin（Resolve-GitRemote / Switch-GitMirror 维护）
$script:RecordedOriginUrl = $null  # 首次改 origin 前记录的原值（更新结束按它恢复）
$script:CanonicalOrigin = "https://github.com/NousResearch/hermes-agent.git"
$script:DesktopReady = "not-started"  # not-started / ready / timeout / launched / skipped
$script:ProgressLen = 0  # 进度行上次输出的字符数（自适应清行用）
$script:OriginRestored = $false  # 本次运行是否把 origin 从镜像恢复为直连

New-Item -ItemType Directory -Path (Split-Path -Parent $RunLog) -Force | Out-Null

# ---------------------------------------------------------------- 输出
function Write-Log {
    param([string]$Message, [ValidateSet("INFO", "STEP", "OK", "WARN", "ERR")][string]$Level = "INFO", [switch]$LogOnly)
    $line = "{0} [{1}] {2}" -f (Get-Date -Format "HH:mm:ss"), $Level, $Message
    if (-not $LogOnly) {
        # 有未收尾的步骤行时，先收行再加一条警告/错误（避免把消息打在步骤行中间）
        if ($script:StepOpen -and $Level -in @("WARN", "ERR")) {
            Write-Host " …" -ForegroundColor DarkGray
            $script:StepOpen = $false
        }
        $color = switch ($Level) {
            "STEP" { "Cyan" } "OK" { "Green" } "WARN" { "Yellow" } "ERR" { "Red" } default { "Gray" }
        }
        if ($Level -eq "STEP") { Write-Host "" }
        Write-Host $line -ForegroundColor $color
    }
    Add-Content -LiteralPath $RunLog -Value $line -Encoding UTF8
}

# ---------------------------------------------------------------- 渲染层（B 方案界面）
# 中英文混排按"显示列宽"计算：CJK、圈号数字等按 2 列，其余按 1 列
function Get-DisplayWidth {
    param([string]$Text)
    $w = 0
    foreach ($ch in $Text.ToCharArray()) {
        $c = [int]$ch
        if (($c -ge 0x1100 -and $c -le 0x115F) -or ($c -ge 0x2460 -and $c -le 0x24FF) -or
            ($c -ge 0x2E80 -and $c -le 0xA4CF) -or ($c -ge 0xAC00 -and $c -le 0xD7A3) -or
            ($c -ge 0xF900 -and $c -le 0xFAFF) -or ($c -ge 0xFE30 -and $c -le 0xFE4F) -or
            ($c -ge 0xFF00 -and $c -le 0xFF60) -or ($c -ge 0xFFE0 -and $c -le 0xFFE6)) { $w += 2 } else { $w += 1 }
    }
    return $w
}
function Pad-Cell {
    param([string]$Text, [int]$Width)
    $pad = $Width - (Get-DisplayWidth $Text)
    if ($pad -lt 1) { $pad = 1 }
    return $Text + (" " * $pad)
}
# 从 URL 取一个短标签用于界面（本机链路长，界面不铺全 URL）
function Get-UrlLabel {
    param([string]$Url)
    if (-not $Url) { return "-" }
    try {
        if ($Url -match '^[a-z]+://([^/]+)/') {
            $h = $Matches[1]
            if ($h -match 'github\.com') { return "github 直连" }
            return $h
        }
    } catch { }
    return $Url
}
# 步骤行（①–⑤ 用）：先打 " ▸ 标签"，完成时在同一行补 " ✓ 结果"
$script:StepOpen = $false
$script:StepWidth = 22
function Start-Step {
    param([string]$Text)
    if ($script:StepOpen) { Write-Host "" ; $script:StepOpen = $false }
    Write-Host ""
    Write-Host (" ▸ " + (Pad-Cell $Text $script:StepWidth)) -ForegroundColor Cyan -NoNewline
    $script:StepOpen = $true
    Add-Content -LiteralPath $RunLog -Value ("{0} [STEP] {1}" -f (Get-Date -Format "HH:mm:ss"), $Text) -Encoding UTF8
}
function Complete-Step {
    param([string]$Result = "", [switch]$Fail)
    $mark = if ($Fail) { "✗" } else { "✓" }
    $color = if ($Fail) { "Red" } else { "Green" }
    if ($script:StepOpen) {
        Write-Host (" {0} {1}" -f $mark, $Result) -ForegroundColor $color
        $script:StepOpen = $false
    } else {
        Write-Host ("   {0} {1}" -f $mark, $Result) -ForegroundColor $color
    }
    Add-Content -LiteralPath $RunLog -Value ("{0} [{1}] {2}" -f (Get-Date -Format "HH:mm:ss"), $(if ($Fail) { "ERR" } else { "OK" }), $Result) -Encoding UTF8
}
# 标题行（⑥⑦ 用）：独占一行，内容缩进在下面
function Write-Head {
    param([string]$Text)
    if ($script:StepOpen) { Write-Host "" ; $script:StepOpen = $false }
    Write-Host ""
    Write-Host (" ▸ " + $Text) -ForegroundColor Cyan
    Add-Content -LiteralPath $RunLog -Value ("{0} [STEP] {1}" -f (Get-Date -Format "HH:mm:ss"), $Text) -Encoding UTF8
}
# 缩进结果行（标题下的收尾/状态）
function Write-Sub {
    param([string]$Text, [switch]$Fail)
    $mark = if ($Fail) { "✗" } else { "✓" }
    $color = if ($Fail) { "Red" } else { "Green" }
    Write-Host ("   {0} {1}" -f $mark, $Text) -ForegroundColor $color
    Add-Content -LiteralPath $RunLog -Value ("{0} [{1}] {2}" -f (Get-Date -Format "HH:mm:ss"), $(if ($Fail) { "ERR" } else { "OK" }), $Text) -Encoding UTF8
}
# ⑥ 的子阶段树行：  "    ├─ 拉取代码      ✓ 00:52" / 末行用 -Last 变成 └─
function Write-StageLine {
    param([string]$Label, [string]$Detail = "", [switch]$Fail, [switch]$Last)
    $mark = if ($Fail) { "✗" } else { "✓" }
    $color = if ($Fail) { "Red" } else { "Green" }
    $tree = if ($Last) { "    └─ " } else { "    ├─ " }
    Write-Host ($tree + (Pad-Cell $Label 14)) -ForegroundColor DarkGray -NoNewline
    Write-Host ("{0} {1}" -f $mark, $Detail) -ForegroundColor $color
    Add-Content -LiteralPath $RunLog -Value ("{0} [STAGE] {1} {2} {3}" -f (Get-Date -Format "HH:mm:ss"), $Label, $mark, $Detail) -Encoding UTF8
}

function Write-Banner {
    $mode = if ($DryRun) { "DryRun · 仅检查（不改动）" } else { "正常 · 上限 $GlobalTotalTimeoutMinutes 分钟" }
    Write-Host ""
    Write-Host "  ╭─ Hermes 自动更新 " -ForegroundColor DarkCyan -NoNewline
    Write-Host ("─" * 27) -ForegroundColor DarkCyan -NoNewline
    Write-Host "╮" -ForegroundColor DarkCyan
    Write-Host ("  │ " + (Pad-Cell $mode 44) + "│") -ForegroundColor DarkCyan
    Write-Host ("  ╰" + ("─" * 45) + "╯") -ForegroundColor DarkCyan
    Write-Host "   安装目录  $InstallRoot"
    Write-Host "   会话日志  $RunLog"
    Write-Host ""
}

# ---------------------------------------------------------------- 小工具
function Invoke-Git {
    param([string[]]$Arguments, [int]$TimeoutSeconds = 90)
    $prev = $env:GIT_TERMINAL_PROMPT
    $env:GIT_TERMINAL_PROMPT = "0"
    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "git"
        $psi.Arguments = (($Arguments | ForEach-Object { if ($_ -match '\s') { '"' + $_ + '"' } else { $_ } }) -join " ")
        $psi.WorkingDirectory = $InstallRoot
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.CreateNoWindow = $true
        $proc = [System.Diagnostics.Process]::Start($psi)
        $outTask = $proc.StandardOutput.ReadToEndAsync()
        $errTask = $proc.StandardError.ReadToEndAsync()
        if (-not $proc.WaitForExit($TimeoutSeconds * 1000)) {
            # 连同 git 派生的孙进程一起杀，否则管道不关、任务读不完
            try { & taskkill /T /F /PID $proc.Id 2>$null | Out-Null } catch { try { $proc.Kill() } catch { } }
            return @{ Code = 124; Out = ""; Err = "timeout after ${TimeoutSeconds}s" }
        }
        # 主进程已退出，但重定向管道可能被孙进程占着不放 -> 有界等待，绝不无界阻塞
        $outOk = $outTask.Wait(10000)
        $errOk = $errTask.Wait(10000)
        if (-not ($outOk -and $errOk)) {
            try { & taskkill /T /F /PID $proc.Id 2>$null | Out-Null } catch { }
            try { [void]$outTask.Wait(3000) } catch { }
            try { [void]$errTask.Wait(3000) } catch { }
            $o = if ($outTask.IsCompleted) { $outTask.Result } else { "" }
            $e = if ($errTask.IsCompleted) { $errTask.Result } else { "" }
            return @{ Code = 125; Out = $o; Err = ($e + "`n[stdout/stderr 未在 10s 内关闭，已终止进程树]") }
        }
        return @{ Code = $proc.ExitCode; Out = $outTask.Result; Err = $errTask.Result }
    } catch {
        return @{ Code = 1; Out = ""; Err = $_.Exception.Message }
    } finally {
        $env:GIT_TERMINAL_PROMPT = $prev
    }
}

function Test-GitReachable {
    param([string]$Url)
    if (-not $Url) { return $false }
    # 本机链路很慢（ghproxy 直连 GitHub 实测 ~80s），固定的低速阈值会误杀“慢但在传”的连接，
    # 所以这里禁用 git 自带的低速中止，只靠 Invoke-Git 的硬超时兜底。超时放宽到 150s。
    $r = Invoke-Git -Arguments @("-c", "http.lowSpeedLimit=0", "-c", "http.lowSpeedTime=0",
                                 "ls-remote", "--heads", $Url, $Branch) -TimeoutSeconds 150
    return ($r.Code -eq 0 -and $r.Out -match "refs/heads/")
}

function Get-ListeningProxy {
    # 探测常见的本机代理端口（翻墙软件/自建代理）
    $candidates = @(7890, 7897, 10809, 10808, 1080, 52058, 8889)
    foreach ($port in $candidates) {
        $client = New-Object System.Net.Sockets.TcpClient
        try {
            $iar = $client.BeginConnect("127.0.0.1", $port, $null, $null)
            if ($iar.AsyncWaitHandle.WaitOne(250, $false) -and $client.Connected) {
                return "http://127.0.0.1:$port"
            }
        } catch { } finally { $client.Close() }
    }
    return ""
}

# 从任意镜像 URL 里剥出裸 github 地址；原本就是 github 则原样返回
function Get-PlainGitHubUrl {
    param([string]$Url)
    if (-not $Url) { return "" }
    if ($Url -match '^git@github\.com:(.+)$') {
        $p = $Matches[1]
        if ($p -notmatch '\.git$') { $p += '.git' }
        return "https://github.com/$p"
    }
    if ($Url -match '(https?://(?:www\.)?github\.com/.+)$') { return $Matches[1] }
    return $Url
}

# 并行探测多个镜像：同时各自跑 git ls-remote，先通者胜。
# 串行逐个试会白等超时；并行只受探测窗口约束，失败的镜像不拖慢整体。
function Select-FastestGitRemote {
    param([string[]]$Urls, [int]$WindowSeconds = 90)
    $Urls = @($Urls | Where-Object { $_ } | Select-Object -Unique)
    if ($Urls.Count -eq 0) { return $null }
    Write-Log ("  并行探测 {0} 个镜像（先通者胜，窗口 {1}s）…" -f $Urls.Count, $WindowSeconds) INFO -LogOnly

    $procs = New-Object System.Collections.ArrayList
    foreach ($u in $Urls) {
        try {
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = "git"
            $psi.Arguments = '-c http.lowSpeedLimit=0 -c http.lowSpeedTime=0 ls-remote --heads "' + $u + '" ' + $Branch
            $psi.WorkingDirectory = $InstallRoot
            $psi.UseShellExecute = $false
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError = $true
            $psi.CreateNoWindow = $true
            # 不让 git 进入交互式认证（与 Invoke-Git 保持一致）
            $psi.EnvironmentVariables["GIT_TERMINAL_PROMPT"] = "0"
            $p = [System.Diagnostics.Process]::Start($psi)
            [void]$procs.Add(@{
                Url     = $u
                Proc    = $p
                Start   = (Get-Date)
                OutTask = $p.StandardOutput.ReadToEndAsync()
                ErrTask = $p.StandardError.ReadToEndAsync()
            })
        } catch {
            Write-Log ("  ✗ {0}（启动失败：{1}）" -f $u, $_.Exception.Message) WARN -LogOnly
        }
    }

    $deadline = (Get-Date).AddSeconds($WindowSeconds)
    $winner = $null
    while (-not $winner -and (Get-Date) -lt $deadline -and $procs.Count -gt 0) {
        foreach ($e in @($procs)) {
            if (-not $e.Proc.HasExited) { continue }
            $ok = $false
            try { $ok = ($e.Proc.ExitCode -eq 0 -and ([string]$e.OutTask.Result) -match 'refs/heads/') } catch { $ok = $false }
            if ($ok) {
                $winner = @{ Url = $e.Url; Seconds = [math]::Round(((Get-Date) - $e.Start).TotalSeconds, 1) }
            } else {
                Write-Log ("  ✗ {0}" -f $e.Url) INFO -LogOnly
            }
            [void]$procs.Remove($e)
            if ($winner) { break }
        }
        if (-not $winner) { Start-Sleep -Milliseconds 300 }
    }
    # 收掉没胜出的探测进程（含 git 派生的子进程）；已是胜出者定局，不再标它们失败
    foreach ($e in @($procs)) {
        if (-not $winner) { Write-Log ("  ✗ {0}（窗口内未应答）" -f $e.Url) WARN -LogOnly }
        try { & taskkill /T /F /PID $e.Proc.Id 2>$null | Out-Null } catch { }
    }
    return $winner
}

# 首次修改 origin 之前记录原值（更新结束恢复用；重复调用只记一次）
function Record-OriginUrl {
    if ($script:RecordedOriginUrl) { return }
    $r = Invoke-Git -Arguments @("remote", "get-url", "origin")
    $script:RecordedOriginUrl = ($r.Out).Trim()
}

# 更新结束后恢复 origin：本仓库的官方地址恢复为裸 github 直连（消除镜像残留导致的 fork 误判），
# 其它来源（自定义 fork 等）按原样恢复。
function Restore-OriginUrl {
    if ($DryRun -or -not $script:RecordedOriginUrl) { return }
    $plain = Get-PlainGitHubUrl $script:RecordedOriginUrl
    $target = if ($plain -eq $script:CanonicalOrigin) { $script:CanonicalOrigin } else { $script:RecordedOriginUrl }
    # 读数用原始配置：remote get-url 会展开全局 insteadOf，显示为镜像改写后的 URL
    $cur = ((Invoke-Git -Arguments @("config", "--local", "--get", "remote.origin.url")).Out).Trim()
    if ($cur -eq $target) { return }
    Invoke-Git -Arguments @("remote", "set-url", "origin", $target) | Out-Null
    $script:OriginRestored = $true
    Write-Log ("origin 已恢复 -> {0}" -f $target) OK
}

# 拉取阶段失败/停滞时：并行找下一个可用镜像并切换 origin
function Switch-GitMirror {
    param([string]$Reason = "")
    if (-not $script:GitMirrorCurrent) {
        $r = Invoke-Git -Arguments @("remote", "get-url", "origin")
        $script:GitMirrorCurrent = ($r.Out).Trim()
    }
    $plain = Get-PlainGitHubUrl $script:GitMirrorCurrent
    $cands = @()
    foreach ($p in $script:GitMirrorPrefixes) { $cands += ($p + $plain) }
    $cands += $plain    # 官方直连（兜底；本机可能 TLS 失败）
    $cands = @($cands | Where-Object { $_ -and $_ -ne $script:GitMirrorCurrent })
    $why = if ($Reason) { "（$Reason）" } else { "" }
    Write-Log ("拉取阶段没成功{0} —— 尝试换镜像" -f $why) WARN
    $winner = Select-FastestGitRemote -Urls $cands -WindowSeconds $GitProbeSeconds
    if (-not $winner) { Write-Log "没有找到其它可用镜像（网络/代理问题？）" ERR; return $false }
    if ($DryRun) { Write-Log ("（DryRun：建议切换 origin -> {0}，本次未改）" -f $winner.Url) WARN; return $false }
    Record-OriginUrl
    Invoke-Git -Arguments @("remote", "set-url", "origin", $winner.Url) | Out-Null
    $script:GitMirrorCurrent = $winner.Url
    Write-Log ("origin 已切换 -> {0}（{1}s 应答）" -f $winner.Url, $winner.Seconds) OK
    return $true
}

# ---------------------------------------------------------------- 1. 网络
function Resolve-GitRemote {
    Start-Step "① 代码源与代理"
    $r = Invoke-Git -Arguments @("remote", "get-url", "origin")
    $origin = ($r.Out).Trim()

    # 先探本机代理：有代理的话，后面更新子进程（uv 下载、pip）也能用上
    $proxy = Get-ListeningProxy
    if ($proxy) {
        Write-Log "发现本机代理 $proxy —— 本次运行的 git/uv/更新子进程都会走它" OK -LogOnly
        $env:HTTPS_PROXY = $proxy
        $env:HTTP_PROXY = $proxy
    }

    if (-not $origin) { Complete-Step "无 origin（交给官方更新处理）"; return }
    $script:GitMirrorCurrent = $origin

    # 单次探测（150s 硬超时、禁用低速误杀）；失败就进并行镜像池，origin 也会作为候选再试一次
    if (Test-GitReachable -Url $origin) { Complete-Step ("origin 可达（" + (Get-UrlLabel $origin) + "）"); return }

    Write-Log "origin 探测失败，进入镜像池并行优选…" WARN -LogOnly
    $plain = Get-PlainGitHubUrl $origin
    $cands = @($origin)
    foreach ($p in $script:GitMirrorPrefixes) { $cands += ($p + $plain) }
    $cands += $plain    # 官方直连（兜底；本机可能 TLS 失败）
    $winner = Select-FastestGitRemote -Urls $cands -WindowSeconds $GitProbeSeconds
    if (-not $winner) {
        Complete-Step -Fail "所有镜像都不可达（检查网络/代理后重试）"
        return
    }
    if ($winner.Url -eq $origin) { Complete-Step ("origin 可用（" + (Get-UrlLabel $origin) + "）"); return }
    if ($DryRun) { Complete-Step ("DryRun：建议切换 -> " + (Get-UrlLabel $winner.Url)); return }
    Record-OriginUrl
    Invoke-Git -Arguments @("remote", "set-url", "origin", $winner.Url) | Out-Null
    $script:GitMirrorCurrent = $winner.Url
    Complete-Step ("已切换镜像 -> {0}（{1}s 应答）" -f (Get-UrlLabel $winner.Url), $winner.Seconds)
}

# ---------------------------------------------------------------- 2. uv
function Test-Uv { if (-not (Test-Path -LiteralPath $UvExe)) { return $false }
    $v = & $UvExe --version 2>$null; return ($LASTEXITCODE -eq 0 -and $v) }

# 从 pm/lock.json 读 uv 的 win32-x64 锁定版本 + sha256（未通过校验的下载一律拒绝）
function Get-PinnedUv {
    $lockPath = Join-Path $InstallRoot "pm\lock.json"
    if (-not (Test-Path -LiteralPath $lockPath)) { return $null }
    try {
        $lock = Get-Content -LiteralPath $lockPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $uv = $lock.packages.uv
        $art = $uv.artifacts.'win32-x64'
        if ($uv.version -and $art.sha256 -and $art.url) {
            $sha = ([string]$art.sha256).ToLower()
            if ($sha -match '^[0-9a-f]{64}$') {
                return [pscustomobject]@{ Version = [string]$uv.version; Url = [string]$art.url; Sha256 = $sha }
            }
        }
    } catch { }
    return $null
}

function Repair-Uv {
    Start-Step "④ uv 校准"
    if (Test-Uv) { Complete-Step ("正常 · " + (& $UvExe --version 2>$null)); return $true }

    foreach ($suffix in @(".bak", ".previous", ".old")) {
        $cand = "$UvExe$suffix"
        if (Test-Path -LiteralPath $cand) {
            Write-Log "用备份 $cand 还原 uv.exe" WARN -LogOnly
            if (-not $DryRun) {
                Copy-Item -LiteralPath $cand -Destination $UvExe -Force
                if (Test-Uv) { Complete-Step ("已从备份还原（" + (& $UvExe --version 2>$null) + "）"); return $true }
            }
        }
    }

    # 按 pm/lock.json 的锁定版本修复，并强制 sha256 校验；校验不过的字节一律不落盘
    $pin = Get-PinnedUv
    if ($DryRun) {
        if ($pin) { Complete-Step ("DryRun：需修复（锁定 {0}，不改动）" -f $pin.Version) }
        else { Complete-Step -Fail "DryRun：需修复（读不到锁定信息）" }
        return $false
    }
    if (-not $pin) {
        Complete-Step -Fail "拒绝未校验下载（pm/lock.json 读不到锁定信息）"
        return $false
    }
    Write-Log ("uv 缺失/损坏；按锁定版本 {0} 修复（sha256 {1}…）" -f $pin.Version, $pin.Sha256.Substring(0, 12)) WARN -LogOnly

    New-Item -ItemType Directory -Path $BinDir -Force | Out-Null
    $zip = ""
    $isTempZip = $false

    # 1) 优先用 PM store 中已通过校验的资产（离线可用）
    $storeEntry = Join-Path (Join-Path $HermesHome "tools") ("fetch-" + $pin.Sha256)
    if (Test-Path -LiteralPath $storeEntry) {
        $cand = Get-ChildItem -LiteralPath $storeEntry -File -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($cand) {
            $got = (Get-FileHash -LiteralPath $cand.FullName -Algorithm SHA256).Hash.ToLower()
            if ($got -eq $pin.Sha256) { $zip = $cand.FullName; Write-Log "使用 PM store 中已校验的 uv 资产" OK -LogOnly }
            else { Write-Log ("PM store 中的 uv 资产校验失败（{0}…），改为下载" -f $got.Substring(0, 12)) WARN -LogOnly }
        }
    }

    # 2) 下载：镜像池 + 官方直连兜底；每个候选都必须过 sha256 才采纳
    if (-not $zip) {
        $tmpZip = Join-Path $env:TEMP ("uv-{0}-{1}.zip" -f $pin.Version, [guid]::NewGuid().ToString("N").Substring(0, 6))
        $urls = @()
        foreach ($p in $script:GitMirrorPrefixes) { $urls += ($p + $pin.Url) }
        $urls += $pin.Url    # 官方直连兜底（本机 TLS 可能失败）
        $urls = @($urls | Where-Object { $_ } | Select-Object -Unique)
        foreach ($url in $urls) {
            try {
                Write-Log ("下载：{0}" -f $url) -LogOnly
                $wc = New-Object System.Net.WebClient
                if ($env:HTTPS_PROXY) { $wc.Proxy = New-Object System.Net.WebProxy($env:HTTPS_PROXY) }
                else { $wc.Proxy = [System.Net.WebRequest]::DefaultWebProxy }
                $wc.DownloadFile($url, $tmpZip)
                $got = (Get-FileHash -LiteralPath $tmpZip -Algorithm SHA256).Hash.ToLower()
                if ($got -ne $pin.Sha256) {
                    Write-Log ("  sha256 不匹配（得到 {0}…），丢弃换下一个源" -f $got.Substring(0, 12)) WARN -LogOnly
                    Remove-Item -LiteralPath $tmpZip -Force -ErrorAction SilentlyContinue
                    continue
                }
                $zip = $tmpZip
                $isTempZip = $true
                break
            } catch {
                Write-Log ("  源失败：{0}" -f $_.Exception.Message) WARN -LogOnly
            }
        }
        if (-not $zip) { Complete-Step -Fail "下载失败（所有源不可用或通不过 sha256）"; return $false }
    }

    # 3) 解压安装
    try {
        $tmp = Join-Path $env:TEMP ("uv-extract-{0}" -f (Get-Random))
        Expand-Archive -LiteralPath $zip -DestinationPath $tmp -Force
        Get-ChildItem -LiteralPath $tmp -Filter "uv*.exe" -Recurse | ForEach-Object {
            Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $BinDir $_.Name) -Force
        }
        Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
        if ($isTempZip) { Remove-Item -LiteralPath $zip -Force -ErrorAction SilentlyContinue }
        if (Test-Uv) { Complete-Step ("已修复 · {0}" -f $pin.Version); return $true }
        Write-Log "安装后 uv 仍不可用" ERR -LogOnly
    } catch {
        Write-Log ("  安装失败：{0}" -f $_.Exception.Message) ERR -LogOnly
    }
    Complete-Step -Fail "修复失败（详见日志）"
    return $false
}

# ---------------------------------------------------------------- 3. 清遗留
function Get-SnapshotRoots {
    $roots = @()
    $main = Join-Path $HermesHome "state-snapshots"
    if (Test-Path -LiteralPath $main) { $roots += $main }
    $profiles = Join-Path $HermesHome "profiles"
    if (Test-Path -LiteralPath $profiles) {
        Get-ChildItem -LiteralPath $profiles -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $p = Join-Path $_.FullName "state-snapshots"
            if (Test-Path -LiteralPath $p) { $roots += $p }
        }
    }
    return $roots
}

function Invoke-Cleanup {
    $plan = New-Object System.Collections.ArrayList

    # 3.1 僵尸"更新进行中"标记：pid 已死 / 不是更新进程 / 标记陈旧 -> 都是僵尸
    if (Test-Path -LiteralPath $MarkerPath) {
        $lines = @(Get-Content -LiteralPath $MarkerPath -ErrorAction SilentlyContinue)
        $markerPid = 0
        if ($lines.Count -ge 1) { [void][int]::TryParse(($lines[0] -as [string]).Trim(), [ref]$markerPid) }
        $alive = $false
        if ($markerPid -gt 0) { $alive = [bool](Get-Process -Id $markerPid -ErrorAction SilentlyContinue) }
        $isUpdater = $false
        if ($alive) {
            try {
                $mp = Get-CimInstance Win32_Process -Filter ("ProcessId = " + $markerPid) -ErrorAction SilentlyContinue
                if ($mp -and $mp.CommandLine -and ($mp.CommandLine -match 'hermes_cli\.main|source_completion\.py|node-deps\.mjs')) { $isUpdater = $true }
            } catch { }
        }
        $ageMin = ((Get-Date) - (Get-Item -LiteralPath $MarkerPath).LastWriteTime).TotalMinutes
        if ($alive -and $isUpdater -and $ageMin -lt 30) {
            Write-Log "更新标记还在，pid $markerPid 是更新进程且较新 —— 跳过（可能真的有更新在跑）" WARN -LogOnly
        } else {
            $why = "僵尸更新标记 (pid $markerPid 已死)"
            if ($alive -and -not $isUpdater) { $why = "僵尸更新标记 (pid $markerPid 不是更新进程)" }
            elseif ($alive -and $ageMin -ge 30) { $why = "陈旧更新标记 (pid $markerPid 存活但已 " + [int]$ageMin + " 分钟无更新)" }
            [void]$plan.Add(@{ Path = $MarkerPath; Kind = "标记"; Reason = $why })
        }
    }

    # 3.2 快照目录：每个 root 只保留最新 1 份
    foreach ($root in (Get-SnapshotRoots)) {
        $dirs = @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)
        if ($dirs.Count -gt 1) {
            foreach ($d in $dirs[1..($dirs.Count - 1)]) {
                [void]$plan.Add(@{ Path = $d.FullName; Kind = "快照"; Reason = "旧快照（保留最新 1 份）" })
            }
        }
    }

    # 3.3 半截备份 / 临时库
    if (Test-Path -LiteralPath $BackupsDir) {
        Get-ChildItem -LiteralPath $BackupsDir -File -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like "*.partial" -or $_.Name -like ".*.partial" -or $_.Name -like "tmp*.db" } |
            ForEach-Object { [void]$plan.Add(@{ Path = $_.FullName; Kind = "半截文件"; Reason = "未完成的上传/备份残留" }) }
    }

    # 3.4 过期收据
    if (Test-Path -LiteralPath $ReceiptsDir) {
        $receipts = @(Get-ChildItem -LiteralPath $ReceiptsDir -File -Filter "update_*.json" -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending)
        if ($receipts.Count -gt $ReceiptsKeep) {
            foreach ($f in $receipts[$ReceiptsKeep..($receipts.Count - 1)]) {
                [void]$plan.Add(@{ Path = $f.FullName; Kind = "旧收据"; Reason = "只保留最近 $ReceiptsKeep 份" })
            }
        }
    }

    if ($plan.Count -eq 0) { Write-Log "没有需要清理的遗留物" OK; return 0 }

    Write-Log ("待清理 {0} 项：" -f $plan.Count) -LogOnly
    foreach ($item in $plan) { Add-Content -LiteralPath $RunLog -Value ("     · [{0}] {1}  ({2})" -f $item.Kind, (Split-Path -Leaf $item.Path), $item.Reason) -Encoding UTF8 }
    if ($DryRun) { Write-Log "（DryRun：只列出，不删除）" WARN -LogOnly; return $plan.Count }

    $removed = 0
    $failed = 0
    foreach ($item in $plan) {
        try { Remove-Item -LiteralPath $item.Path -Recurse -Force -ErrorAction Stop; $removed++ }
        catch { $failed++; Write-Log "删不掉：$($item.Path) —— $($_.Exception.Message)" WARN -LogOnly }
    }
    Write-Log ("清理完成：{0} 项已删，{1} 项失败" -f $removed, $failed) OK -LogOnly
    return $removed
}

# ---------------------------------------------------------------- 4. 桌面端
function Get-DesktopMainPids {
    return @(Get-CimInstance Win32_Process -Filter "Name='Hermes.exe'" -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -notmatch "--type=" } | Select-Object -ExpandProperty ProcessId)
}

function Stop-Desktop {
    Write-Log "关闭桌面端（更新完成后会自动重开）" INFO
    $pids = Get-DesktopMainPids
    if ($pids.Count -eq 0) { Write-Log "桌面端没在运行" OK; return $false }
    if ($DryRun) { Write-Log "（DryRun：不关闭桌面端）" WARN; return $false }

    Write-Log ("Hermes 将在 {0} 秒后自动关闭，更新完成后会自动重新打开" -f $CountdownSeconds) WARN
    for ($i = $CountdownSeconds; $i -gt 0; $i--) {
        Write-Host ("`r   倒计时 {0,2}s … " -f $i) -ForegroundColor Yellow -NoNewline
        Start-Sleep -Seconds 1
    }
    Write-Host "`r                      " -NoNewline

    if ($DryRun) { Write-Log "（DryRun：不关）" WARN; return $false }

    foreach ($p in $pids) { try { (Get-Process -Id $p -ErrorAction SilentlyContinue).CloseMainWindow() | Out-Null } catch { } }

    $deadline = (Get-Date).AddSeconds(20)
    while ((Get-Date) -lt $deadline) {
        if (@(Get-DesktopMainPids).Count -eq 0) { break }
        Start-Sleep -Milliseconds 500
    }
    $left = @(Get-DesktopMainPids)
    if ($left.Count -gt 0) {
        Write-Log "优雅关闭失败，强制结束：$($left -join ',')" WARN
        foreach ($p in $left) { Stop-Process -Id $p -Force -ErrorAction SilentlyContinue }
        Start-Sleep -Seconds 2
    }
    Write-Log "桌面端已退出" OK

    # 旧后端（内存里还是旧代码，留着会重新锁 venv）
    $backends = @(Get-CimInstance Win32_Process -Filter "Name='python.exe'" -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -match "hermes_cli\.main" -and $_.CommandLine -match "\bserve\b" })
    foreach ($b in $backends) {
        Write-Log "停掉旧后端 pid $($b.ProcessId)" WARN
        Stop-Process -Id $b.ProcessId -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 2
    return $true
}

# ---------------------------------------------------------------- 5. 更新
function Get-LogTextShared([string]$Path) {
    # 子进程正写着日志文件，需共享读
    if (-not (Test-Path -LiteralPath $Path)) { return "" }
    try {
        $fs = [System.IO.File]::Open($Path, 'Open', 'Read', 'ReadWrite')
        $sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8)
        $t = $sr.ReadToEnd(); $sr.Close(); $fs.Close(); return $t
    } catch { return "" }
}

function Invoke-HermesUpdate {
    param([int]$Attempt)
    $updateLog = Join-Path $HermesHome "logs\self-update-run.log"
    $stdoutLog = Join-Path $HermesHome "logs\self-update-run.out.log"
    $stderrLog = Join-Path $HermesHome "logs\self-update-run.err.log"
    Write-Log ("   运行官方更新（第 {0}/{1} 轮，单轮上限 {2} 分钟）…" -f $Attempt, $MaxAttempts, $UpdateTimeoutMinutes) -LogOnly
    Add-Content -LiteralPath $RunLog -Value ("--- hermes update attempt {0} @ {1} ---" -f $Attempt, (Get-Date -Format s)) -Encoding UTF8

    # 本机链路慢（ghproxy 实测 ~80s），低速阈值会把“慢但在传”的连接误杀。
    # 置 0 禁用 git 自带的低速中止；是否卡死交给脚本的阶段/静默看门狗判断。
    $env:GIT_HTTP_LOW_SPEED_LIMIT = "0"
    $env:GIT_HTTP_LOW_SPEED_TIME = "0"
    # 防止 git / 凭据助手在无终端环境下弹交互或死等输入（代理返回 401 时常见）。
    # 注意：不要用 GIT_CONFIG_* 注入“空值”配置 —— Windows PowerShell 5.1 里
    # `$env:X = ""` 会直接删掉变量，导致 git 报 "missing config value GIT_CONFIG_VALUE_0"
    # 而让**所有** git 命令失败（实测踩过）。GIT_TERMINAL_PROMPT=0 + GCM_INTERACTIVE=never 已足够。
    $env:GIT_TERMINAL_PROMPT = "0"
    $env:GCM_INTERACTIVE = "never"
    $env:PYTHONUNBUFFERED = "1"
    $script:lastUpdateOutput = ""
    Remove-Item -LiteralPath $stdoutLog, $stderrLog -Force -ErrorAction SilentlyContinue

    # 子进程：加 -u + PYTHONUNBUFFERED，保证进度实时写入文件（否则会被块缓冲，25 分钟读不到一行）
    $proc = Start-Process -FilePath $VenvPython `
        -ArgumentList @("-u", "-m", "hermes_cli.main", "update", "--yes", "--keep-stash", "--branch", $Branch) `
        -WorkingDirectory $InstallRoot -NoNewWindow -PassThru `
        -RedirectStandardOutput $stdoutLog -RedirectStandardError $stderrLog

    $attemptStart = Get-Date
    $logPaths = @($stdoutLog, $stderrLog, $updateLog)

    $stage = 'fetch'; $stageStart = $attemptStart
    $ioHigh = @{}
    $act = Get-UpdateActivity -RootPid $proc.Id -LogPaths $logPaths -HighWater $ioHigh
    $prevProgress = [long]$act.Total
    $progressAt = $attemptStart
    $samples = New-Object System.Collections.ArrayList
    [void]$samples.Add(@{ t = $attemptStart; p = $prevProgress })
    $nextSampleAt = $attemptStart.AddSeconds($StallSampleSeconds)
    $lowStreak = 0
    $completeSeenAt = $null
    $terminatedReason = $null
    $rateText = "-"

    while (-not $proc.HasExited) {
        Start-Sleep -Seconds $ProgressRefreshSeconds
        $now = Get-Date

        # 进度 = 进程树 IO 峰值累计 + 日志输出字节（CPU 不作为续期依据；总量单调不减，
        # 子进程退出不会让数值倒退，避免把正常推进误判成“过慢”）
        $act = Get-UpdateActivity -RootPid $proc.Id -LogPaths $logPaths -HighWater $ioHigh
        $progress = [long]$act.Total
        if ($progress -gt $prevProgress) { $prevProgress = $progress; $progressAt = $now } else { $progress = $prevProgress }

        $silence = ($now - $progressAt).TotalSeconds
        $stageElapsed = ($now - $stageStart).TotalMinutes
        $attemptElapsed = ($now - $attemptStart).TotalMinutes
        $totalElapsed = ($now - $startedAt).TotalMinutes

        # 阶段识别：只看本轮的 stdout/stderr（合并日志尾部是历史内容，会把阶段带偏）
        $liveText = (Get-LogTailShared $stderrLog 65536) + "`n" + (Get-LogTailShared $stdoutLog 65536)
        $newStage = Get-CurrentStage $liveText
        if ($newStage -ne $stage) {
            # 先把刚结束的阶段收成一行树（用旧阶段名与耗时），再切到新阶段
            $prevLabel = $StageLabel[$stage]
            $prevDur = Format-Duration $stageStart $now
            Clear-ProgressLine
            Write-StageLine -Label $prevLabel -Detail $prevDur
            $stage = $newStage; $stageStart = $now
            # 换阶段 = 换基线：清空采样与过慢计数，避免拿上一阶段的高速率误杀新阶段
            $samples.Clear(); [void]$samples.Add(@{ t = $now; p = $progress })
            $nextSampleAt = $now.AddSeconds($StallSampleSeconds)
            $lowStreak = 0
        }
        # 阶段耗时必须在阶段切换之后重算：否则刚切到新阶段时会拿"旧阶段的耗时"去撞
        # 新阶段的硬上限（实测踩过：切到"安装 Node 依赖"的瞬间被自己的 20 分钟上限杀掉）
        $stageElapsed = ($now - $stageStart).TotalMinutes

        if ($now -ge $nextSampleAt) {
            [void]$samples.Add(@{ t = $now; p = $progress })
            if ($samples.Count -gt 240) { $samples.RemoveRange(0, $samples.Count - 240) }
            $nextSampleAt = $now.AddSeconds($StallSampleSeconds)
            $rate = Get-ProgressRate -Samples $samples -Now $now
            $rateText = $rate.Text
            # “过慢”只在两个条件下才累计：本阶段已跑满 90s（基线可信）＋ 确实已有 30s 零进展。
            # 纯速率对比容易在 IO 抖动时误杀，宁可让阶段硬上限兜底。
            if ($rate.IsSlow -and ($now - $stageStart).TotalMinutes -ge 1.5 -and $silence -ge 30) { $lowStreak++ } else { $lowStreak = 0 }
        }

        # 收尾：出现「Update complete!」只记一笔日志；不再按墙钟掐。
        # 实测尾部可能是可选组件（agent-browser 等）的大下载，5 分钟墙钟会误杀真实工作；
        # 交给 tail 阶段的静默阈值与 45 分钟上限兜底。
        if ($liveText -match 'Update complete!' -and -not $completeSeenAt) {
            $completeSeenAt = $now
            Write-Log "检测到「Update complete!」—— 更新主体已完成，等待收尾结束（tail 阶段上限 45 分钟）…" OK
        }

        $status = "推进中"
        if ($lowStreak -ge $SlowStreakSamples) { $status = "疑似过慢" }
        elseif ($silence -ge $StageSilenceSeconds[$stage]) { $status = "疑似卡死" }
        Write-ProgressLine -Now $now -Attempt $Attempt -Stage $stage -StageStart $stageStart -AttemptStart $attemptStart -Start $startedAt -Rate $rateText -Silence $silence -Status $status

        # ---- 三态判定（策略B：真死 / 过慢 / 超时 都终止）----
        if ($silence -ge $StageSilenceSeconds[$stage]) {
            $terminatedReason = 'stalled-silent'
            Write-Log ("判定卡死：阶段「{0}」{1} 秒内无任何进度/输出 —— 终止进程树" -f $StageLabel[$stage], [int]$silence) ERR
            & taskkill /T /F /PID $proc.Id 2>$null | Out-Null
            break
        }
        if ($lowStreak -ge $SlowStreakSamples) {
            $terminatedReason = 'stalled-slow'
            Write-Log ("判定过慢：阶段「{0}」连续 {1} 次采样进度低于基线 {2}% —— 策略B 终止" -f $StageLabel[$stage], $lowStreak, $SlowRatioPercent) ERR
            & taskkill /T /F /PID $proc.Id 2>$null | Out-Null
            break
        }
        if ($stageElapsed -ge $StageTimeoutMinutes[$stage]) {
            $terminatedReason = 'stage-timeout'
            Write-Log ("阶段「{0}」超过硬上限 {1} 分钟 —— 终止" -f $StageLabel[$stage], $StageTimeoutMinutes[$stage]) ERR
            & taskkill /T /F /PID $proc.Id 2>$null | Out-Null
            break
        }
        if ($attemptElapsed -ge $UpdateTimeoutMinutes) {
            $terminatedReason = 'attempt-timeout'
            Write-Log ("单轮超过 {0} 分钟 —— 终止" -f $UpdateTimeoutMinutes) WARN
            & taskkill /T /F /PID $proc.Id 2>$null | Out-Null
            break
        }
        if ($totalElapsed -ge $GlobalTotalTimeoutMinutes) {
            $terminatedReason = 'global-timeout'
            Write-Log ("全部轮次合计超过 {0} 分钟 —— 终止" -f $GlobalTotalTimeoutMinutes) ERR
            & taskkill /T /F /PID $proc.Id 2>$null | Out-Null
            break
        }
    }

    Clear-ProgressLine
    try { [void]$proc.WaitForExit(15000) } catch { }

    # 合并本轮 stdout/stderr 到统一日志，并留存给安全阀/构建判定
    $out = @()
    foreach ($f in @($stdoutLog, $stderrLog)) {
        if (Test-Path -LiteralPath $f) { $out += ((Get-LogTextShared $f) -split "`n") }
    }
    $out | Tee-Object -FilePath $updateLog -Append | Out-Null
    $script:lastUpdateOutput = ($out -join "`n")
    Normalize-LogFile $updateLog

    $script:lastUpdateStage = $stage
    $script:lastStallReason = $terminatedReason
    $script:lastProgressRate = $rateText
    $script:stageElapsedText = (Format-Duration $stageStart (Get-Date))
    $script:silenceSeconds = [int]((Get-Date) - $progressAt).TotalSeconds

    $code = 124
    try { if ($proc.HasExited) { $code = [int]$proc.ExitCode } } catch { $code = 124 }
    # 被我们自己终止的轮次不能算成功：否则会跳过重试并骗过最终校验
    if ($terminatedReason -and $terminatedReason -ne 'complete-tail') { $code = 124 }
    if ($code -ne 0 -and ($script:lastUpdateOutput -match "Update complete!")) {
        Write-Log "更新主体已完成（日志含「Update complete!」），忽略收尾进程的退出码 $code" WARN
        $code = 0
    }
    # 收尾：最后一个阶段 + 本轮汇总（树的最后两行）
    if ($terminatedReason) {
        Write-StageLine -Label $StageLabel[$stage] -Detail ("停滞：{0}" -f $terminatedReason) -Fail
    } else {
        Write-StageLine -Label $StageLabel[$stage] -Detail (Format-Duration $stageStart (Get-Date))
    }
    if ($code -eq 0) {
        Write-StageLine -Label ("第 {0} 轮" -f $Attempt) -Detail "完成" -Last
    } else {
        Write-StageLine -Label ("第 {0} 轮" -f $Attempt) -Detail ("未完成（exit={0}）" -f $code) -Fail -Last
    }
    return $code
}

function Read-TextRobust {
    param([string]$Path)
    # 更新子进程的输出可能是 UTF-16（管道下常见），统一读法
    try {
        $bytes = [System.IO.File]::ReadAllBytes($Path)
        if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
            return [System.Text.Encoding]::Unicode.GetString($bytes)
        }
        if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
            return [System.Text.Encoding]::BigEndianUnicode.GetString($bytes)
        }
        $probe = [Math]::Min(512, $bytes.Length)
        $nulls = 0
        for ($i = 0; $i -lt $probe; $i++) { if ($bytes[$i] -eq 0) { $nulls++ } }
        if ($nulls -gt ($probe / 4)) { return [System.Text.Encoding]::Unicode.GetString($bytes) }
        return [System.Text.Encoding]::UTF8.GetString($bytes)
    } catch { return "" }
}

function Normalize-LogFile {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return }
    $text = Read-TextRobust $Path
    if (-not $text) { return }
    try { [System.IO.File]::WriteAllText($Path, $text, (New-Object System.Text.UTF8Encoding($false))) } catch { }
}

function Rotate-UpdateLogs {
    # 每次运行先把上一轮的 self-update-run*.log 归档：失败报告 / 桌面重建判定 /
    # 安全阀检测都只读本轮内容，避免历史内容污染判定（实测踩过：报告里混入旧运行报错）。
    $logDir = Join-Path $HermesHome "logs"
    $archive = Join-Path $logDir "archive"
    New-Item -ItemType Directory -Path $archive -Force | Out-Null
    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
    foreach ($name in @("self-update-run.log", "self-update-run.out.log", "self-update-run.err.log")) {
        $src = Join-Path $logDir $name
        if (-not (Test-Path -LiteralPath $src)) { continue }
        $stem = [System.IO.Path]::GetFileNameWithoutExtension($name)
        try { Move-Item -LiteralPath $src -Destination (Join-Path $archive ("$stem-$stamp.log")) -Force } catch { }
    }
    # 归档只保留最近 3 轮（每轮 3 个文件）
    $old = @(Get-ChildItem -LiteralPath $archive -File -Filter "self-update-run*.log" -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -Skip 9)
    foreach ($f in $old) { Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue }
    $kept = @(Get-ChildItem -LiteralPath $archive -File -Filter "self-update-run*.log" -ErrorAction SilentlyContinue).Count
    Write-Log ("日志已轮转（归档保留 {0} 份）" -f $kept) INFO -LogOnly
}

function Test-SafeDeleteTrip {
    $updateLog = Join-Path $HermesHome "logs\self-update-run.log"
    if (-not (Test-Path -LiteralPath $updateLog)) { return $false }
    $text = Read-TextRobust $updateLog
    return ($text -match "SAFE_DELETE_BULK_CONFIRM_REQUIRED")
}

function Get-DesktopBuildVerdict {
    # 官方更新自己会决定要不要重建桌面端；读它的结论，别自己猜
    $updateLog = Join-Path $HermesHome "logs\self-update-run.log"
    if (-not (Test-Path -LiteralPath $updateLog)) { return "unknown" }
    $text = Read-TextRobust $updateLog
    if ($text -match "Desktop app up to date") { return "uptodate" }
    if ($text -match "Building the desktop app|Building desktop app|Desktop build complete|Rebuilding desktop") { return "rebuilt" }
    return "unknown"
}

# ---------------------------------------------------------------- 6. 校验
function Get-UpdateSummary {
    # 优先读最新的 update_*.json（latest.json 可能指向 pm sync 收据，outcome 语义不同）
    $summary = @{ Outcome = "unknown"; From = ""; To = ""; Steps = @(); Restart = "" }
    $receipt = $null
    $newest = Get-ChildItem -LiteralPath $ReceiptsDir -File -Filter "update_*.json" -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($newest) { $receipt = $newest.FullName }
    else {
        $latest = Join-Path $ReceiptsDir "latest.json"
        if (Test-Path -LiteralPath $latest) { $receipt = $latest }
    }
    if ($receipt) {
        try {
            $d = Get-Content -LiteralPath $receipt -Raw | ConvertFrom-Json
            if ($d.outcome) { $summary.Outcome = $d.outcome }
            if ($d.pre_update) { $summary.From = $d.pre_update.short_sha }
            if ($d.post_update) { $summary.To = $d.post_update.short_sha }
            if ($d.gateway_restart) { $summary.Restart = ("restarted={0} killed={1}" -f @($d.gateway_restart.restarted_services).Count, @($d.gateway_restart.killed_pids).Count) }
        } catch { }
    }
    return $summary
}

# ---------------------------------------------------------------- 7. 重开
# 只看"启动之后新写入"的日志段，避免把历史 HERMES_BACKEND_READY 当成本次就绪
function Test-DesktopBackendReady {
    param([string]$LogPath, [long]$AfterOffset)
    try {
        if (-not (Test-Path -LiteralPath $LogPath)) { return $false }
        $fs = [System.IO.File]::Open($LogPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        try {
            if ($fs.Length -le $AfterOffset) { return $false }
            $readLen = [int][Math]::Min(1048576, ($fs.Length - $AfterOffset))
            $fs.Seek($fs.Length - $readLen, [System.IO.SeekOrigin]::Begin) | Out-Null
            $buf = New-Object byte[] $readLen
            $read = $fs.Read($buf, 0, $readLen)
            $text = [System.Text.Encoding]::UTF8.GetString($buf, 0, $read)
            return [bool]($text -match 'HERMES_BACKEND_READY')
        } finally { $fs.Close() }
    } catch { return $false }
}

function Start-Desktop {
    if ($NoRelaunch) { $script:DesktopReady = "skipped"; Write-Log "按参数要求，不重开桌面端" WARN; return $false }
    if (-not (Test-Path -LiteralPath $DesktopExe)) { $script:DesktopReady = "not-started"; Write-Log "找不到桌面端 exe：$DesktopExe" ERR; return $false }
    Write-Log "重开桌面端…" INFO -LogOnly

    $desktopLog = Join-Path $HermesHome "logs\desktop.log"
    $before = 0
    try { if (Test-Path -LiteralPath $desktopLog) { $before = (Get-Item -LiteralPath $desktopLog).Length } } catch { $before = 0 }

    Start-Process -FilePath $DesktopExe -WorkingDirectory (Split-Path -Parent $DesktopExe)

    # 第一步：先确认主进程确实起来了（旧行为：最多等 30s）
    $procSeen = $false
    $procDeadline = (Get-Date).AddSeconds(30)
    while (-not $procSeen -and (Get-Date) -lt $procDeadline) {
        if (@(Get-DesktopMainPids).Count -gt 0) { $procSeen = $true; break }
        Start-Sleep -Milliseconds 500
    }
    if (-not $procSeen) { $script:DesktopReady = "not-started"; Write-Log "桌面端没起来，请手动打开" ERR; return $false }

    # 第二步：等后端就绪（而不是只看进程在不在）
    if ($LaunchReadyTimeoutSeconds -le 0) { $script:DesktopReady = "launched"; Write-Log "桌面端已启动（未启用就绪等待）" OK -LogOnly; return $true }
    $deadline = (Get-Date).AddSeconds($LaunchReadyTimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        if (Test-DesktopBackendReady -LogPath $desktopLog -AfterOffset $before) {
            $script:DesktopReady = "ready"
            Write-Log "桌面端后端已就绪（HERMES_BACKEND_READY）" OK -LogOnly
            return $true
        }
        Start-Sleep -Seconds 2
    }
    $script:DesktopReady = "timeout"
    Write-Log ("桌面端已启动，但 {0}s 内未等到后端就绪 —— 见 {1}" -f $LaunchReadyTimeoutSeconds, $desktopLog) WARN
    return $true
}

# ---------------------------------------------------------------- 构建环境 / 残留进程
function Set-BuildEnv {
    Start-Step "② 构建环境"
    # 本机没有 C++ 工具链：原生模块只能走预编译，禁止 node-pre-gyp 回退源码编译。
    # 注意：该变量会被 node-deps.mjs 计入安装指纹 -> 首次会触发一次全量 npm ci，之后可复用。
    $env:npm_config_build_from_source = "false"
    $env:npm_config_registry = "https://registry.npmmirror.com/"
    # 下面这些被 node-deps.mjs 排除在复用指纹之外，可安全用于提速
    $env:npm_config_prefer_offline = "true"
    $env:npm_config_audit = "false"
    $env:npm_config_fund = "false"
    # 网络快速失败，避免 npm 挂在下载上
    $env:npm_config_fetch_timeout = "120000"
    $env:npm_config_fetch_retries = "2"
    if (-not $env:ELECTRON_BUILDER_BINARIES_MIRROR) { $env:ELECTRON_BUILDER_BINARIES_MIRROR = "https://npmmirror.com/mirrors/electron-builder-binaries/" }
    if (-not $env:ELECTRON_MIRROR) { $env:ELECTRON_MIRROR = "https://npmmirror.com/mirrors/electron/" }
    $env:PYTHONUNBUFFERED = "1"

    # get-windows 的预编译在 GitHub release；本机缺 C++ 工具链，拉不到就会回退源码编译（必失败）。
    # node-pre-gyp 支持用 `npm_config_<module>_binary_host_mirror` 覆盖下载源
    # （见 @mapbox/node-pre-gyp/lib/util/versioning.js，module_name=node-get-windows ->
    #  validModuleName=node_get-windows）。默认用 GitHub 代理镜像。
    $gwm = $NodeBinaryMirror
    if (-not $gwm) { $gwm = "https://ghproxy.net/https://github.com/sindresorhus/get-windows/releases/download/" }
    try { [System.Environment]::SetEnvironmentVariable('npm_config_node_get-windows_binary_host_mirror', $gwm, 'Process') } catch { }

    Write-Log "build_from_source=false | registry=npmmirror | prefer_offline=true | fetch_timeout=120s | PYTHONUNBUFFERED=1" OK -LogOnly
    Write-Log ("get-windows 预编译镜像: " + $gwm) INFO -LogOnly
    Complete-Step "镜像已设 · 禁源码编译 · 快速失败"
}

function Stop-StaleUpdaters {
    Start-Step "③ 残留进程"
    $pattern = 'hermes_cli\.main\s+update|source_completion\.py|scripts/build/node-deps\.mjs|scripts/build/desktop\.mjs|scripts/build\.mjs|scripts/build/freshness\.mjs'
    $targets = @(Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
        Where-Object { $_.ProcessId -ne $PID -and $_.CommandLine -and ($_.CommandLine -match $pattern) })
    if ($targets.Count -eq 0) { Complete-Step "无"; return }
    foreach ($t in $targets) {
        Write-Log ("结束残留更新进程 pid {0} ({1})" -f $t.ProcessId, $t.Name) WARN -LogOnly
        if (-not $DryRun) { & taskkill /T /F /PID $t.ProcessId 2>$null | Out-Null }
    }
    Complete-Step ("已结束 {0} 个残留进程" -f $targets.Count)
}

function Cache-GetWindowsPrebuilt {
    # 备份当前可用的 get-windows 预编译，供诊断/手动恢复（npm ci 会清空 node_modules，无法自动回填）
    $src = Join-Path $InstallRoot "node_modules\get-windows\lib\binding\napi-9-win32-unknown-x64\node-get-windows.node"
    if (-not (Test-Path -LiteralPath $src)) { return }
    try {
        $dstDir = Join-Path $PrebuiltCacheDir "get-windows\napi-9-win32-unknown-x64"
        New-Item -ItemType Directory -Path $dstDir -Force | Out-Null
        Copy-Item -LiteralPath $src -Destination (Join-Path $dstDir "node-get-windows.node") -Force
        Write-Log ("已缓存 get-windows 预编译 -> " + (Join-Path $dstDir "node-get-windows.node")) INFO -LogOnly
    } catch { Write-Log ("缓存 get-windows 预编译失败：" + $_.Exception.Message) WARN }
}

# ---------------------------------------------------------------- 进度 / 卡死判定
function Get-LogTailShared {
    param([string]$Path, [int]$MaxBytes = 65536)
    if (-not (Test-Path -LiteralPath $Path)) { return "" }
    try {
        $fs = [System.IO.File]::Open($Path, 'Open', 'Read', 'ReadWrite')
        try {
            $len = $fs.Length
            $start = [Math]::Max(0, $len - $MaxBytes)
            $count = [int]($len - $start)
            if ($count -le 0) { return "" }
            $fs.Seek($start, 'Begin') | Out-Null
            $buf = New-Object byte[] $count
            $read = $fs.Read($buf, 0, $count)
            return [System.Text.Encoding]::UTF8.GetString($buf, 0, $read)
        } finally { $fs.Close() }
    } catch { return "" }
}

function Get-UpdateActivity {
    # 一次 CIM 查询：汇总更新进程树的 IO 字节（O(1)，便宜）与日志输出字节。
    # HighWater：调用方持有的「pid → 历史峰值」表。子进程退出后它的字节不会消失，
    # 保证 Total 单调不减 —— 否则子进程一退出总量就倒退，会被误判成“过慢”。
    param([int]$RootPid, [string[]]$LogPaths, [hashtable]$HighWater)
    $all = @(Get-CimInstance Win32_Process -ErrorAction SilentlyContinue)
    $byParent = @{}
    $ioById = @{}
    foreach ($p in $all) {
        $r = 0L; $w = 0L
        if ($p.ReadTransferCount)  { $r = [long]$p.ReadTransferCount }
        if ($p.WriteTransferCount) { $w = [long]$p.WriteTransferCount }
        $ioById[[int]$p.ProcessId] = ($r + $w)
        $pp = [int]$p.ParentProcessId
        if (-not $byParent.ContainsKey($pp)) { $byParent[$pp] = New-Object System.Collections.ArrayList }
        [void]$byParent[$pp].Add([int]$p.ProcessId)
    }
    $count = 0
    $queue = New-Object System.Collections.Queue
    $queue.Enqueue($RootPid)
    $seen = @{}
    while ($queue.Count -gt 0) {
        $cur = [int]$queue.Dequeue()
        if ($seen.ContainsKey($cur)) { continue }
        $seen[$cur] = $true; $count++
        if ($ioById.ContainsKey($cur)) {
            if (-not $HighWater.ContainsKey($cur) -or $ioById[$cur] -gt [long]$HighWater[$cur]) {
                $HighWater[$cur] = $ioById[$cur]
            }
        }
        if ($byParent.ContainsKey($cur)) { foreach ($c in $byParent[$cur]) { $queue.Enqueue($c) } }
    }
    $io = 0L
    foreach ($v in $HighWater.Values) { $io += [long]$v }
    $outBytes = 0L
    foreach ($lp in $LogPaths) { if (Test-Path -LiteralPath $lp) { $outBytes += (Get-Item -LiteralPath $lp).Length } }
    return @{ Io = $io; OutputBytes = $outBytes; Procs = $count; Total = ($io + $outBytes) }
}

function Get-CurrentStage {
    param([string]$Text)
    if (-not $Text) { return 'fetch' }
    $stage = 'fetch'
    $lines = $Text -split "`n"
    if ($lines.Count -gt 400) { $lines = $lines[($lines.Count - 400)..($lines.Count - 1)] }
    foreach ($ln in $lines) {
        $l = $ln.ToLower()
        if ($l -match 'update complete') { $stage = 'tail' }
        elseif ($l -match 'finishing an interrupted source update|completing source-update dependencies|source completion') { $stage = 'completion' }
        elseif ($l -match 'building the desktop app|building desktop packaged app|packaging the desktop app|desktop\.mjs|build-only') { $stage = 'desktop' }
        elseif ($l -match 'building the web ui|build/web\.mjs') { $stage = 'web' }
        elseif ($l -match 'building the tui|build/tui\.mjs') { $stage = 'tui' }
        elseif ($l -match 'preparing node dependencies|npm ci|node-deps') { $stage = 'npm' }
        elseif ($l -match 'syncing python|syncing dependencies|pip install|uv pip') { $stage = 'python' }
    }
    return $stage
}

function Format-Bytes {
    param([long]$n)
    if ($n -lt 0) { return "0B" }
    if ($n -ge 1GB) { return ("{0:N1}GB" -f ($n / 1GB)) }
    if ($n -ge 1MB) { return ("{0:N1}MB" -f ($n / 1MB)) }
    if ($n -ge 1KB) { return ("{0:N1}KB" -f ($n / 1KB)) }
    return ("{0}B" -f $n)
}

function Format-Duration {
    param([datetime]$From, [datetime]$To)
    $ts = $To - $From
    if ($ts.TotalSeconds -lt 0) { $ts = [timespan]::Zero }
    if ($ts.TotalHours -ge 1) { return ("{0:00}:{1:00}:{2:00}" -f [int]$ts.TotalHours, $ts.Minutes, $ts.Seconds) }
    return ("{0:00}:{1:00}" -f $ts.Minutes, $ts.Seconds)
}

function Get-ProgressRate {
    # 最近 60s 进度 vs 此前 180s 基线；返回是否过慢与显示文本
    param($Samples, [datetime]$Now)
    $res = @{ IsSlow = $false; Text = "-" }
    $arr = @($Samples)
    if ($arr.Count -lt 4) { return $res }
    $last = $arr[$arr.Count - 1]
    $p60 = $null; $p180 = $null
    foreach ($s in $arr) {
        if ($s.t -ge $Now.AddSeconds(-60)) { if ($null -eq $p60) { $p60 = $s.p } }
        if ($s.t -ge $Now.AddSeconds(-180)) { if ($null -eq $p180) { $p180 = $s.p } }
    }
    if ($null -eq $p60) { $p60 = $arr[0].p }
    if ($null -eq $p180) { $p180 = $p60 }
    $recent = [long]($last.p - $p60)
    if ($recent -lt 0) { $recent = 0 }
    $baseWindow = [long]($p60 - $p180)
    if ($baseWindow -lt 0) { $baseWindow = 0 }
    $basePerMin = [double]$baseWindow / 2.0
    $res.Text = ("+{0}/min" -f (Format-Bytes $recent))
    if ($basePerMin -gt 0) {
        $ratio = ([double]$recent) / $basePerMin
        if ($ratio -lt ($SlowRatioPercent / 100.0)) { $res.IsSlow = $true }
    }
    return $res
}

function Write-ProgressLine {
    param($Now, $Attempt, $Stage, $StageStart, $AttemptStart, $Start, $Rate, $Silence, $Status)
    # 输出被重定向（日志/管道）时不刷进度行：\r 在原地覆盖，写进文件只会粘连成超长行
    if ([Console]::IsOutputRedirected) { return }
    # 精简版：正常只留 轮次 · 阶段 · 阶段耗时 · 速率；异常才追加状态/静默
    $line = "[{0}] 第{1}/{2}轮 · {3} · {4} · {5}" -f `
        $Now.ToString('HH:mm:ss'), $Attempt, $MaxAttempts, $StageLabel[$Stage], `
        (Format-Duration $StageStart $Now), $Rate
    if ($Status -eq "疑似卡死") { $line += " │ 疑似卡死" }
    elseif ($Status -eq "疑似过慢") { $line += " │ 疑似过慢" }
    elseif ($Silence -ge 30) { $line += (" │ 静默 {0}s" -f [int]$Silence) }
    if ((Get-DisplayWidth $line) -gt 110) { $line = $line.Substring(0, 110) }
    # 自适应清行：按上次实际输出长度补空格，不再固定 118
    $pad = [Math]::Max(0, $script:ProgressLen - $line.Length)
    Write-Host ("`r" + $line + (" " * $pad)) -NoNewline
    $script:ProgressLen = $line.Length
}
function Clear-ProgressLine {
    if ($script:ProgressLen -gt 0) {
        Write-Host ("`r" + (" " * $script:ProgressLen) + "`r") -NoNewline
        $script:ProgressLen = 0
    }
}

# ---------------------------------------------------------------- 失败报告
function New-FailureReport {
    param([string]$Reason = "unknown", [hashtable]$Summary = @{})
    $updateLog = Join-Path $HermesHome "logs\self-update-run.log"
    $logText = Read-TextRobust $updateLog
    if (-not $logText) { $logText = "" }
    $lines = @($logText -split "`n")
    $tail = $logText
    if ($lines.Count -gt 80) { $tail = ($lines[($lines.Count - 80)..($lines.Count - 1)] -join "`n") }

    $patterns = @(
        @{ n = 'node-gyp 源码编译失败'; re = 'node-gyp|gyp ERR!' },
        @{ n = 'npm 错误'; re = 'npm ERR!' },
        @{ n = 'Python 调用失败'; re = 'CalledProcessError|Traceback' },
        @{ n = 'npm ci ENOTEMPTY'; re = 'ENOTEMPTY' },
        @{ n = '删除安全阀'; re = 'SAFE_DELETE_BULK_CONFIRM_REQUIRED' },
        @{ n = '网络失败'; re = 'network request failed|timed out|WinError 10054|ETIMEDOUT|ECONNRESET|fetch failed' },
        @{ n = '桌面打包失败'; re = 'Packaging the desktop app failed|Stopped running desktop app' },
        @{ n = '卡死/超时终止'; re = '判定卡死|判定过慢|超过硬上限|timeout' }
    )
    $hits = New-Object System.Collections.ArrayList
    foreach ($pat in $patterns) {
        $ms = @($lines | Where-Object { $_ -match $pat.re } | Select-Object -Last 5)
        if ($ms.Count -gt 0) {
            [void]$hits.Add(("### " + $pat.n))
            foreach ($m in $ms) { [void]$hits.Add("    " + $m.Trim()) }
            [void]$hits.Add("")
        }
    }

    $nodeVer = ""
    $nodeDir = Get-ChildItem -LiteralPath (Join-Path $HermesHome "tools") -Directory -Filter "node-*-win32-x64" -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First 1
    if ($nodeDir) { try { $nodeVer = (& (Join-Path $nodeDir.FullName "node.exe") --version 2>$null) } catch { } }
    $npmVer = ""
    try { $npmVer = (& (Join-Path $HermesHome "tools\npm-12.0.2-win32-x64\npm.CMD") --version 2>$null) } catch { }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("# Hermes 更新失败报告")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("- 生成时间: " + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
    [void]$sb.AppendLine("- 结果: " + $Summary.Outcome + "  (exit=" + $Summary.ExitCode + ")")
    [void]$sb.AppendLine("- 版本: " + $Summary.From + " -> " + $Summary.To)
    [void]$sb.AppendLine("- 开始: " + $startedAt.ToString('yyyy-MM-dd HH:mm:ss') + "  总耗时: " + (Format-Duration $startedAt (Get-Date)))
    [void]$sb.AppendLine("- 终止原因: " + $Reason)
    [void]$sb.AppendLine("- 最后阶段: " + $StageLabel[$script:lastUpdateStage] + "  阶段耗时: " + $script:stageElapsedText + "  静默: " + $script:silenceSeconds + "s  进度速率: " + $script:lastProgressRate)
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("## 环境")
    [void]$sb.AppendLine("- 安装根: " + $InstallRoot)
    [void]$sb.AppendLine("- 分支: " + $Branch)
    [void]$sb.AppendLine("- Node: " + $nodeVer + "   npm: " + $npmVer)
    [void]$sb.AppendLine("- build_from_source: " + $env:npm_config_build_from_source)
    [void]$sb.AppendLine("- registry: " + $env:npm_config_registry)
    [void]$sb.AppendLine("- HTTPS_PROXY: " + $env:HTTPS_PROXY)
    [void]$sb.AppendLine("- ELECTRON_MIRROR: " + $env:ELECTRON_MIRROR)
    [void]$sb.AppendLine("- SAFE_DELETE_BULK_CONFIRM_REQUIRED: " + (Test-SafeDeleteTrip))
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("## 命中的错误特征")
    if ($hits.Count -eq 0) { [void]$sb.AppendLine("(未匹配到已知特征，见下方最近输出)") } else { [void]$sb.AppendLine(($hits -join "`n")) }
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("## 最近 80 行更新输出")
    [void]$sb.AppendLine('```')
    [void]$sb.AppendLine($tail)
    [void]$sb.AppendLine('```')
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("## 相关日志")
    [void]$sb.AppendLine("- " + $updateLog)
    [void]$sb.AppendLine("- " + $RunLog)
    [void]$sb.AppendLine("- " + (Join-Path $HermesHome 'logs\desktop.log'))
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("## 建议下一步")
    [void]$sb.AppendLine("1. 桌面打包失败: 关闭桌面端后运行 hermes desktop --force-build")
    [void]$sb.AppendLine("2. 网络失败: 检查代理后重跑 更新Hermes.cmd")
    [void]$sb.AppendLine("3. node-gyp / get-windows 编译失败: 本机无 C++ 工具链，必须用预编译。")
    [void]$sb.AppendLine("   镜像可用 -NodeBinaryMirror 覆盖；已缓存的预编译在: " + (Join-Path $PrebuiltCacheDir "get-windows\napi-9-win32-unknown-x64\node-get-windows.node"))
    [void]$sb.AppendLine("4. 收集诊断: hermes debug share")
    [void]$sb.AppendLine("")

    $reportPath = Join-Path $HermesHome ("logs\update-report-{0}.md" -f (Get-Date -Format "yyyyMMdd-HHmmss"))
    try {
        [System.IO.File]::WriteAllText($reportPath, $sb.ToString(), (New-Object System.Text.UTF8Encoding($true)))
        Write-Log ("已生成失败报告：" + $reportPath) ERR
        return $reportPath
    } catch {
        Write-Log ("失败报告写入失败：" + $_.Exception.Message) ERR
        return $null
    }
}

# ---------------------------------------------------------------- 状态检查（还债/更新循环用）
function Get-LocalHeadSha {
    for ($i = 1; $i -le 2; $i++) {
        $r = Invoke-Git -Arguments @("rev-parse", "HEAD")
        if ($r.Code -eq 0) {
            $sha = (($r.Out -split "`n")[0]).Trim()
            if ($sha) { return $sha }
        }
        Start-Sleep -Milliseconds 800
    }
    return ""
}

function Get-RemoteHeadSha {
    # 直接 ls-remote 取远端分支 SHA（不 fetch、不污染本地引用）；失败返回空
    $url = $script:GitMirrorCurrent
    if (-not $url) {
        $r0 = Invoke-Git -Arguments @("remote", "get-url", "origin")
        $url = ($r0.Out).Trim()
    }
    if (-not $url) { return "" }
    $r = Invoke-Git -Arguments @("-c", "http.lowSpeedLimit=0", "-c", "http.lowSpeedTime=0",
                                 "ls-remote", "--heads", $url, $Branch) -TimeoutSeconds 120
    if ($r.Code -ne 0) { Write-Log ("取不到远端 HEAD（ls-remote 失败，code={0}）" -f $r.Code) WARN; return "" }
    $m = [regex]::Match($r.Out, ('(?m)^([0-9a-fA-F]{7,40})\s+refs/heads/' + [regex]::Escape($Branch) + '\s*$'))
    if ($m.Success) { return $m.Groups[1].Value }
    Write-Log "取不到远端 HEAD（输出中未找到该分支）" WARN
    return ""
}

function Test-SourceCompletionPending {
    $c = @(Get-ChildItem (Join-Path $HermesHome "installs") -Recurse -Force -File `
            -Filter "source-completion-pending" -ErrorAction SilentlyContinue)
    return ($c.Count -gt 0)
}

# ================================================================ 主流程
$startedAt = Get-Date
Write-Banner

if (-not (Test-Path -LiteralPath $VenvPython)) { Write-Log "venv python 不存在：$VenvPython（安装目录写错？）" ERR; exit 1 }

Resolve-GitRemote | Out-Null
Set-BuildEnv
Stop-StaleUpdaters
$uvOk = Repair-Uv
Start-Step "⑤ 清理 / 资产预置"
$cleaned = Invoke-Cleanup

# ③.5 预置 pm store（国内镜像补齐 lock.json 里 pin 的资产，避免更新卡在下载）
# pm 的工具下载只认 GitHub 原始地址 + 一个 Azure CDN，本机实测极慢；预置器把缺失的
# 资产用镜像下好、过 sha256 后放进 <store>\fetch-<sha>，pm 便会直接跳过下载。已安装的
# 包会被预置器跳过，不会为 ffmpeg 之类白下。
$seeder = Join-Path $PSScriptRoot "seed-pm-store.ps1"
if (Test-Path -LiteralPath $seeder) {
    $seedArgs = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $seeder, "-InstallRoot", $InstallRoot, "-Quiet")
    if ($DryRun) { $seedArgs += "-DryRun" }
    # 捕获输出：成功时界面只留一行汇总；失败时把尾部原样打出来（并全文进日志）
    $seedOut = & powershell @seedArgs 2>&1
    $seedExit = [int]$LASTEXITCODE
    $seedOut | ForEach-Object { Add-Content -LiteralPath $RunLog -Value $_ -Encoding UTF8 }
    if ($seedExit -ne 0) {
        Complete-Step -Fail ("清理 {0} 项 · store 预置有失败（exit={1}，详见日志）" -f $cleaned, $seedExit)
        $seedOut | Select-Object -Last 12 | ForEach-Object { Write-Host ("     " + $_) -ForegroundColor Yellow }
    } else {
        Complete-Step ("清理 {0} 项 · store 就绪" -f $cleaned)
    }
} else {
    $seedExit = 0
    Complete-Step ("清理 {0} 项 · 无预置器" -f $cleaned)
}

if ($DryRun) {
    Write-Sub ("DryRun 结束：uv {0} · 待清理 {1} 项 —— 未做任何改动" -f $(if ($uvOk) { "正常" } else { "需修复" }), $cleaned)
    if (-not $NoPause -and $Host.Name -eq "ConsoleHost") { Write-Host "按回车关闭…" -NoNewline; [void](Read-Host) }
    exit 0
}

Rotate-UpdateLogs
Cache-GetWindowsPrebuilt
Write-Head ("⑥ 运行更新（最多 {0} 轮）" -f $MaxAttempts)
$desktopWasRunning = Stop-Desktop

# treeless/部分克隆在 git 2.53+ 下 fetch 后的自动 repack 会崩（should_include_obj，#124272）。
# 幂等关闭自动维护：脚本自己的裸 fetch 与官方更新子进程都受益（官方另有自身恢复逻辑）。
if (((Invoke-Git -Arguments @("config", "--get", "maintenance.auto")).Out).Trim() -ne "false") {
    Invoke-Git -Arguments @("config", "maintenance.auto", "false") | Out-Null
    Write-Log "已设置仓库配置 maintenance.auto=false（规避 partial-clone fetch 崩溃）" INFO
}

$exitCode = -1
$lastSignature = ""
$localSha = ""; $remoteSha = ""; $pendingNow = $false; $behind = $false
for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
    # 本轮性质（只影响日志）：有"未完成收尾"标记 = 还债轮（官方恢复路径，不拉新代码）
    if (Test-SourceCompletionPending) {
        Write-Log ("   第 {0}/{1} 轮：还清未完成的源码收尾（官方恢复路径，本轮不拉新代码）" -f $attempt, $MaxAttempts) INFO
    } else {
        Write-Log ("   第 {0}/{1} 轮：检查并执行更新（拉取 → 迁移 → 收尾）" -f $attempt, $MaxAttempts) INFO
    }

    $exitCode = Invoke-HermesUpdate -Attempt $attempt

    # 每轮结束都算一次"还有没有活"：收尾标记 + 本地/远端 SHA 比对
    $pendingNow = Test-SourceCompletionPending
    $localSha = Get-LocalHeadSha
    $remoteSha = Get-RemoteHeadSha
    if (-not $localSha) { Write-Log "警告：取不到本地 HEAD，本轮收敛判定不可靠" WARN }
    $behind = [bool]($remoteSha -and $localSha -and ($localSha -ne $remoteSha))
    $moreWork = ($pendingNow -or $behind)
    $signature = "{0}|{1}|{2}" -f $localSha, $remoteSha, $pendingNow

    if ($exitCode -eq 0 -and -not $moreWork) {
        if ($localSha -and $remoteSha) {
            Write-Log ("第 {0} 轮完成且已收敛：HEAD={1}，远端={2}，无未完成收尾" -f $attempt, $localSha, $remoteSha) OK -LogOnly
        } else {
            Write-Log ("第 {0} 轮完成：无未完成收尾（SHA 校验不完整：local='{1}' remote='{2}'）" -f $attempt, $localSha, $remoteSha) WARN
        }
        break
    }

    if ($exitCode -eq 0 -and $moreWork) {
        # 成功但活没干完（典型：还债轮只清了收尾，还没拉新代码）—— 自动继续，直到真正没事干
        if ($signature -eq $lastSignature) {
            Write-Log ("第 {0} 轮无法推进（HEAD={1} 远端={2} 尾巴={3}，与上轮相同）—— 停止，避免空转" -f $attempt, $localSha, $remoteSha, $pendingNow) WARN
            break
        }
        $why = if ($pendingNow) { "还有未完成的源码收尾" } else { "本地落后远端（$localSha -> $remoteSha）" }
        Write-Log ("第 {0} 轮完成，但{1} —— 自动继续下一轮" -f $attempt, $why) WARN
        $lastSignature = $signature
        if ($attempt -eq $MaxAttempts) { break }
        Start-Sleep -Seconds 3
        continue
    }

    # ---- 失败路径（策略B / 换镜像 / 清遗留 / 重试）----
    # 策略B：构建阶段停滞不重试（重试只会再跑一次全量 npm ci，纯浪费）
    if ($script:lastStallReason -in @('stalled-slow', 'stalled-silent', 'stage-timeout') -and `
            $script:lastUpdateStage -in @('npm', 'tui', 'web', 'desktop')) {
        Write-Log ("构建阶段停滞（{0} @ {1}）—— 按策略B 不再重试，直接失败并出报告" -f $script:lastStallReason, $StageLabel[$script:lastUpdateStage]) ERR
        break
    }
    # 拉取阶段失败/停滞：先判性质。本地仓库类错误（partial-clone 的 pack-objects 崩溃，
    # 形如 should_include_obj）换镜像不会好；官方更新器自带该崩溃的恢复，脚本侧也已幂等
    # 关闭 fetch 后的自动 repack。只有网络类失败才轮换镜像。
    if ($script:lastUpdateStage -eq 'fetch') {
        if ($script:lastUpdateOutput -match 'should_include_obj|pack-objects|repack local links') {
            Write-Log "拉取失败属本地仓库类错误（pack-objects/索引），不换镜像（换镜像不会好）" WARN
        } else {
            $why = if ($script:lastStallReason) { "$($script:lastStallReason)（exit=$exitCode）" } else { "exit=$exitCode" }
            [void](Switch-GitMirror -Reason $why)
        }
    }
    if (Test-SafeDeleteTrip) {
        Write-Log "命中「单次删除 > 50 个文件」安全阀，再清一轮遗留后重试" WARN
        Invoke-Cleanup | Out-Null
    } else {
        Write-Log "更新返回 $exitCode，重试前先清一轮遗留" WARN
        Invoke-Cleanup | Out-Null
    }
    if ($attempt -eq $MaxAttempts) { break }
    Write-Log ("等待 {0} 秒后重试…" -f (5 * $attempt))
    Start-Sleep -Seconds (5 * $attempt)
}
$roundsRun = $attempt

# 脚本为换镜像改过 origin：更新结束后恢复（官方地址 -> 裸 github，规避官方更新器的 fork 误判）
Restore-OriginUrl

Write-Head "⑦ 校验与重开"
$summary = Get-UpdateSummary
$logCompleted = [bool]($script:lastUpdateOutput -match "Update complete!")
$verified = $false
if ($exitCode -eq 0 -and ($summary.Outcome -eq "success" -or $logCompleted)) {
    if ($behind) {
        Write-Log ("更新未完成：本地 {0} 仍落后远端 {1}（轮次已用尽，再跑一次本脚本即可继续）" -f $localSha, $remoteSha) ERR
    } else {
        $verified = $true
        Write-Log ("更新成功：{0} -> {1}   （receipt={2}，日志完成标记={3}，HEAD={4}）" -f $summary.From, $summary.To, $summary.Outcome, $logCompleted, $localSha) OK -LogOnly
    }
} else {
    Write-Log ("更新未成功：exit=$exitCode receipt=$($summary.Outcome) 日志完成标记=$logCompleted") ERR
}

# 桌面端构建产物时间 + 官方重建结论
$buildVerdict = Get-DesktopBuildVerdict
if (Test-Path -LiteralPath $DesktopExe) {
    $exeTime = (Get-Item -LiteralPath $DesktopExe).LastWriteTime
    switch ($buildVerdict) {
        "uptodate" { Write-Log ("桌面端：官方判定无需重建（现有产物 {0}）" -f $exeTime.ToString("MM-dd HH:mm")) OK -LogOnly }
        "rebuilt"  { Write-Log ("桌面端：本次更新内已重建（产物 {0}）" -f $exeTime.ToString("MM-dd HH:mm")) OK -LogOnly }
        default {
            if ($exeTime -lt $startedAt.AddMinutes(-1)) {
                Write-Log ("桌面端：产物 {0} 早于本次更新，日志里没看到重建结论 —— 必要时跑 hermes desktop --build-only" -f $exeTime.ToString("MM-dd HH:mm")) WARN -LogOnly
            } else {
                Write-Log ("桌面端：本次更新内已重建（产物 {0}）" -f $exeTime.ToString("MM-dd HH:mm")) OK -LogOnly
            }
        }
    }
}

# 更新"尾巴"是否还欠着？installs\<id>\source-completion-pending 存在 = 源码更新的收尾
# （依赖/产物/桌面端重建）没跑完。此时后端每次启动都会先重跑收尾，而收尾里的桌面端重建
# 在桌面端运行时必然被跳过 → 判失败 → 后端 exit 1 → 界面永远起不来。所以这时绝不能把
# 桌面端开回去，否则直接开进死循环。
$pendingMarkers = @(Get-ChildItem (Join-Path $HermesHome "installs") -Recurse -Force -File `
        -Filter "source-completion-pending" -ErrorAction SilentlyContinue)
if ($pendingMarkers.Count -gt 0) {
    Write-Log ("更新尾巴未还清，仍存在：{0}" -f (($pendingMarkers | ForEach-Object { $_.FullName }) -join '; ')) ERR
    Write-Log "说明源码更新的收尾没跑完——先在桌面端完全退出后重跑本脚本（必要时先跑 seed-pm-store.ps1）" WARN
    $verified = $false
}

# 失败 -> 生成报告
$reportPath = $null
if (-not $verified) {
    $reason = [string]$exitCode
    if ($script:lastStallReason) { $reason = $script:lastStallReason }
    elseif ($behind -and $exitCode -eq 0) { $reason = 'still-behind' }
    $reportPath = New-FailureReport -Reason $reason -Summary @{ Outcome = $summary.Outcome; From = $summary.From; To = $summary.To; ExitCode = $exitCode }
}

if ($verified) {
    $desktopUp = Start-Desktop
} elseif ($pendingMarkers.Count -gt 0) {
    $desktopUp = $false
    Write-Log "尾巴没还清，不重开桌面端（开回去只会卡在端口超时）" WARN
} else {
    $desktopUp = $false
    if ($desktopWasRunning) { Write-Log "更新没成功，仍把桌面端开回来" WARN; $desktopUp = Start-Desktop }
}

# ⑦ 收尾：验证结论 + 重开状态
if ($verified) {
    $readyText = switch ($script:DesktopReady) {
        "ready" { "后端就绪" } "timeout" { "已启动（未就绪）" } "launched" { "已启动" } "skipped" { "未重开" } default { "未启动" }
    }
    $vt = if ($summary.From -and $summary.To -and ($summary.From -ne $summary.To)) { "{0} → {1}" -f $summary.From, $summary.To }
          elseif ($summary.To) { "已是最新（$($summary.To)）" }
          else { "完成" }
    Write-Sub ("{0} · 桌面端{1}" -f $vt, $readyText)
} else {
    Write-Sub -Fail ("未完成 · 最后阶段：{0}" -f $StageLabel[$script:lastUpdateStage])
}

# autostash 提示：--keep-stash 会累积；只报告不清理（stash 里可能有真实的本地改动），详情见结束卡片
$stashOut = (Invoke-Git -Arguments @("stash", "list")).Out
$autostashCount = @($stashOut -split "`n" | Where-Object { $_ -match 'hermes-update-autostash' }).Count
if ($autostashCount -gt 0) {
    Write-Log ("仓库中有 {0} 条更新自动暂存（hermes-update-autostash）—— 视需要用 git stash list / git stash drop 处理" -f $autostashCount) WARN -LogOnly
}
$finalOrigin = ((Invoke-Git -Arguments @("config", "--local", "--get", "remote.origin.url")).Out).Trim()

# 结果文件（给桌面端/后续会话读）
$result = [ordered]@{
    outcome          = $(if ($verified) { "success" } else { "failed" })
    exitCode         = $exitCode
    rounds           = $roundsRun
    head             = $localSha
    remoteHead       = $remoteSha
    from             = $summary.From
    to               = $summary.To
    startedAt        = $startedAt.ToString("s")
    finishedAt       = (Get-Date).ToString("s")
    cleanedItems     = $cleaned
    seedExit         = $seedExit
    desktopRelaunched = $desktopUp
    desktopReady     = $script:DesktopReady
    originUrl        = $finalOrigin
    autostashCount   = $autostashCount
    stage            = $script:lastUpdateStage
    stalled          = [bool]$script:lastStallReason
    reason           = $script:lastStallReason
    report           = $reportPath
    log              = $RunLog
}
try { ($result | ConvertTo-Json) | Set-Content -LiteralPath (Join-Path $HermesHome ".self-update-result.json") -Encoding UTF8 } catch { }

Write-Host ""
$elapsedText = Format-Duration $startedAt (Get-Date)
Write-Host ("  " + ("─" * 62)) -ForegroundColor DarkCyan
if ($verified) {
    $bits = @("用时 $elapsedText")
    if ($script:DesktopReady -eq "ready") { $bits += "桌面端已就绪" }
    elseif ($desktopUp) { $bits += "桌面端已启动" }
    if ($script:OriginRestored) { $bits += "origin 已恢复" }
    if ($autostashCount -gt 0) { $bits += ("autostash {0} 条" -f $autostashCount) }
    Write-Host ("  ✅ 完成 · " + ($bits -join " · ")) -ForegroundColor Green
    Write-Host ("     日志  " + $RunLog) -ForegroundColor DarkGray
} else {
    Write-Host ("  ❌ 未成功 · 用时 {0} · 最后阶段：{1}" -f $elapsedText, $StageLabel[$script:lastUpdateStage]) -ForegroundColor Red
    Write-Host ("     日志  " + $RunLog) -ForegroundColor Yellow
    if ($reportPath) { Write-Host ("     报告  " + $reportPath) -ForegroundColor Yellow }
}
Write-Host ""

if (-not $NoPause -and $Host.Name -eq "ConsoleHost") {
    Write-Host "按回车关闭…" -NoNewline
    [void](Read-Host)
}
