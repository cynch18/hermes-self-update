<#
  pm store 预置器（Windows）
  ------------------------------------------------------------------
  为什么有它：`hermes update` 在同步 venv 之前会调 pm.ensure_tools_for_sync()，
  它要把 lockfile 里"必需"的工具全部发布到 PATH —— 而 pm 的下载器
  (pm/downloader.py) 只认两个源：
     ① lock.json 里的原始 URL（基本全是 github.com / nodejs.org / registry.npmjs.org）
     ② 镜像 https://hermes-assets.nousresearch.com/upstream/sha256/<sha256>
  在国内这两个源都会慢到不可用（实测 Azure CDN 约 47 KB/s，120 MB 要 43 分钟，
  且多连接反复断）。pm 的注释写得很直白：
     "An update is an explicit user action ... a failed download fails the update."
  于是更新永远卡在下载 → 死掉 → 留下 .hermes-update-in-progress 与
  source-completion-pending → 桌面端起不来。形成死循环。

  这个脚本做的事只有一件：**把缺的资产用国内镜像下回来，验签，放进 pm 的
  store 缓存**。放进去之后 pm 会走 pm/downloader.py:293 的
  `_existing_dest_ok()` 分支，把该文件当作"全段已覆盖"，**整段跳过下载**。

  store 布局（pm/store.py:fetch_many）：
     <store>\fetch-<sha256>\<原始文件名>      ← 该目录必须"恰好一个文件"
  其中 <store> = installs\<install-id>（即含 facts.json 的那个目录）。

  安全保证：**每份下载都必须通过 lock.json 里的 sha256 校验才会落盘**。
  所以镜像给错文件、给了旧版本、或者被中间人换了内容，都只会被丢弃并继续试下一个源，
  绝不会污染你的安装。

  用法：
     powershell -NoProfile -ExecutionPolicy Bypass -File seed-pm-store.ps1
     powershell ... -File seed-pm-store.ps1 -DryRun          # 只列出缺什么、打算从哪下
     powershell ... -File seed-pm-store.ps1 -Packages node,npm
     powershell ... -File seed-pm-store.ps1 -All             # 不限于必需工具，全部尝试

  注意：本文件必须保存为「UTF-8 with BOM」。你的 更新Hermes.cmd 调的是 Windows
  PowerShell 5.1，它会把无 BOM 的 UTF-8 当 GBK 读，中文注释直接变成乱码并导致语法错误。
  用记事本/VS Code 另存时选 "UTF-8 with BOM"，或：
     $p='...\seed-pm-store.ps1'; $b=[IO.File]::ReadAllBytes($p)
     if(-not($b[0]-eq0xEF-and$b[1]-eq0xBB-and$b[2]-eq0xBF)){[IO.File]::WriteAllBytes($p,([byte[]](0xEF,0xBB,0xBF)+$b))}
#>
[CmdletBinding()]
param(
    [string]$InstallRoot = "",
    # 只处理这些包（默认：必需工具集 + 它们会拉取的构建依赖）
    [string[]]$Packages = @("ffmpeg", "node", "npm", "python", "ripgrep", "git", "uv"),
    # 处理 lock.json 里全部资产（含 llama.cpp / playwright 等大件，很费流量）
    [switch]$All,
    # 单个候选源的最长下载时间（秒）
    [int]$PerSourceTimeoutSeconds = 150,
    # 只检查不动手
    [switch]$DryRun,
    [string]$WorkDir = "",
    # 跳过 get-windows 预编译预置
    [switch]$SkipGetWindows,
    # get-windows 预编译下载用的 GitHub 代理前缀（默认 ghproxy.net）
    [string]$GetWindowsMirror = "",
    # 供上级脚本调用：收起 banner 与逐包"已就绪/跳过"行，只输出变化、下载与失败
    [switch]$Quiet
)

$ErrorActionPreference = "Continue"
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

if (-not $InstallRoot) { $InstallRoot = Join-Path $env:LOCALAPPDATA "hermes\hermes-agent" }
if (-not (Test-Path -LiteralPath $InstallRoot)) { Write-Host "找不到安装目录：$InstallRoot" -ForegroundColor Red; exit 1 }
$InstallRoot = (Resolve-Path -LiteralPath $InstallRoot).Path
$HermesHome  = Split-Path -Parent $InstallRoot
if (-not $WorkDir) { $WorkDir = Join-Path $HermesHome "cache\seed-tmp" }
New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null

$CN_MIRROR_ROOT = "https://registry.npmmirror.com/-/binary"
$CDN_ROOT       = "https://hermes-assets.nousresearch.com/upstream/sha256"
$GH_PROXIES     = @("https://ghproxy.net/", "https://gh-proxy.com/", "https://ghfast.top/", "https://gh.llkk.cc/", "https://ghproxy.cn/", "https://github.moeyy.xyz/")
# 高校 github-release 镜像根（完整形式：<根>/<owner>/<repo>/<tag>/<文件>）
$GH_RELEASE_MIRRORS = @(
    "https://mirrors.tuna.tsinghua.edu.cn/github-release",
    "https://mirror.nju.edu.cn/github-release",
    "https://mirrors.ustc.edu.cn/github-release"
)

function Write-Log {
    param([string]$Message, [ValidateSet("INFO", "STEP", "OK", "WARN", "ERR", "SKIP")][string]$Level = "INFO")
    $color = switch ($Level) {
        "STEP" { "Cyan" } "OK" { "Green" } "WARN" { "Yellow" } "ERR" { "Red" } "SKIP" { "DarkGray" } default { "Gray" }
    }
    if ($Level -eq "STEP") { Write-Host "" }
    Write-Host ("[{0}] {1}" -f $Level, $Message) -ForegroundColor $color
}

# ---------------------------------------------------------------- 定位 store
function Get-StoreRoot {
    # 镜像 pm 的 writable_store_root()：有 payload manifest.json 时用它的 store，
    # 否则是 <hermes-home>\tools（pm/paths.py）。fetch-<sha> 缓存必须落在这个目录，
    # 落错地方（例如 installs\<id>）pm 根本不会读，等于白下。
    $manifest = Join-Path $HermesHome "manifest.json"
    if (Test-Path -LiteralPath $manifest) {
        try {
            $m = Get-Content -LiteralPath $manifest -Raw | ConvertFrom-Json
            if ($m.store) { return (Join-Path $HermesHome $m.store) }
        } catch { Write-Log "manifest.json 解析失败，回退到 tools" WARN }
    }
    return (Join-Path $HermesHome "tools")
}

# ---------------------------------------------------------------- 候选源
function Get-CandidateUrls {
    param([string]$Url, [string]$Sha256)
    $out = New-Object System.Collections.ArrayList
    $add = { param($u) if ($u -and -not $out.Contains($u)) { [void]$out.Add($u) } }

    $file = $Url.Split('/')[-1]

    # GitHub release → 国内二进制镜像（npmmirror 按 repo 名分目录）
    if ($Url -match '^https://github\.com/([^/]+)/([^/]+)/releases/download/([^/]+)/(.+)$') {
        $owner, $repo, $tag, $rest = $Matches[1], $Matches[2], $Matches[3], $Matches[4]
        & $add ("$CN_MIRROR_ROOT/$repo/$tag/$rest")
        & $add ("$CN_MIRROR_ROOT/$repo/$rest")
        # 少数仓库在镜像里用 owner-repo 命名
        & $add ("$CN_MIRROR_ROOT/$owner-$repo/$tag/$rest")
        # 高校镜像的 github-release 布局：/<owner>/<repo>/<tag>/<file>
        # 有些仓库的最新版放在 LatestRelease/ 下，两条都试（反正要过 sha256）
        foreach ($m in $GH_RELEASE_MIRRORS) {
            & $add ("$m/$owner/$repo/$tag/$rest")
            & $add ("$m/$owner/$repo/LatestRelease/$rest")
        }
    }
    # nodejs.org → npmmirror 的 node 目录
    if ($Url -match '^https://nodejs\.org/dist/([^/]+)/(.+)$') {
        & $add ("$CN_MIRROR_ROOT/node/$($Matches[1])/$($Matches[2])")
    }
    # npm registry → npmmirror registry（完整镜像）
    if ($Url -match '^https://registry\.npmjs\.org/(.+)$') {
        & $add ("https://registry.npmmirror.com/$($Matches[1])")
    }
    # 通用 GitHub 代理
    if ($Url -match '^https://github\.com/') {
        foreach ($p in $GH_PROXIES) { & $add ($p + $Url) }
    }
    # 官方 CDN 镜像（慢但通常可达，放最后）
    & $add ("$CDN_ROOT/$Sha256")
    # 原始 URL 最后兜底
    & $add $Url
    return $out
}

# ---------------------------------------------------------------- 校验/落盘
function Test-Seeded {
    param([string]$Store, [string]$Sha256)
    $entry = Join-Path $Store "fetch-$Sha256"
    if (-not (Test-Path -LiteralPath $entry)) { return $false }
    $files = @(Get-ChildItem -LiteralPath $entry -File -ErrorAction SilentlyContinue)
    if ($files.Count -ne 1) { return $false }
    try { return ((Get-FileHash -LiteralPath $files[0].FullName -Algorithm SHA256).Hash.ToLower() -eq $Sha256) }
    catch { return $false }
}

function Save-Artifact {
    param([string]$Store, [string]$Sha256, [string]$Url, [string]$FileName)
    $tmp = Join-Path $WorkDir ("dl-" + $Sha256.Substring(0, 12) + "-" + [guid]::NewGuid().ToString("N").Substring(0, 6))
    foreach ($cand in (Get-CandidateUrls -Url $Url -Sha256 $Sha256)) {
        if ($DryRun) { Write-Log ("      [DryRun] 会尝试 {0}" -f $cand) "SKIP"; continue }
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
        $host_ = try { ([uri]$cand).Host } catch { "?" }
        & curl.exe -sS -L --fail --max-time $PerSourceTimeoutSeconds --retry 1 -o $tmp $cand 2>$null
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $tmp)) {
            Write-Log ("      {0} 失败（curl {1}）" -f $host_, $LASTEXITCODE) "SKIP"
            continue
        }
        $got = (Get-FileHash -LiteralPath $tmp -Algorithm SHA256).Hash.ToLower()
        if ($got -ne $Sha256) {
            Write-Log ("      {0} 内容不匹配（得到 {1}…）—— 丢弃" -f $host_, $got.Substring(0, 12)) WARN
            continue
        }
        $entry = Join-Path $Store "fetch-$Sha256"
        New-Item -ItemType Directory -Path $entry -Force | Out-Null
        # fetch_many 要求该目录"恰好一个文件"
        Get-ChildItem -LiteralPath $entry -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
        Move-Item -LiteralPath $tmp -Destination (Join-Path $entry $FileName) -Force
        $size = (Get-Item (Join-Path $entry $FileName)).Length
        Write-Log ("      {0} 命中并验签通过（{1:N1} MB）" -f $host_, ($size / 1MB)) OK
        return $true
    }
    return $false
}

# ================================================================ 主流程
if (-not $Quiet) {
    Write-Host ""
    Write-Host "  ╭──────────────────────────────────────────────╮" -ForegroundColor DarkCyan
    Write-Host "  │   pm store 预置器（锁定哈希 + 国内镜像）      │" -ForegroundColor DarkCyan
    Write-Host "  ╰──────────────────────────────────────────────╯" -ForegroundColor DarkCyan
    Write-Host "   安装目录: $InstallRoot"
    if ($DryRun) { Write-Host "   模式    : 仅检查（DryRun）" -ForegroundColor Yellow }
    Write-Host ""
}

$store = Get-StoreRoot
if (-not $store) { Write-Log "无法解析 pm store 根目录" ERR; exit 1 }
Write-Log "store 根目录：$store" OK

# 已装清单：<store>\facts.json。已装且版本一致的包，pm 的 ensure() 不会再去
# 下载，也就无需为它预置 fetch 缓存——否则会为 184MB 的 ffmpeg 之类白下。
$facts = $null
$factsPath = Join-Path $store "facts.json"
if (Test-Path -LiteralPath $factsPath) {
    try { $facts = Get-Content -LiteralPath $factsPath -Raw | ConvertFrom-Json } catch { $facts = $null }
}
function Test-Installed([string]$Pkg, [string]$Version) {
    if (-not $facts -or -not $facts.packages) { return $false }
    $prop = $facts.packages.PSObject.Properties[$Pkg]
    if (-not $prop) { return $false }
    $rec = $prop.Value
    if ($Version -and $rec.version -and ($rec.version -ne $Version)) { return $false }
    if (-not $rec.entry) { return $false }
    return (Test-Path -LiteralPath (Join-Path $store $rec.entry))
}

$lockPath = Join-Path $InstallRoot "pm\lock.json"
if (-not (Test-Path -LiteralPath $lockPath)) { Write-Log "找不到 $lockPath" ERR; exit 1 }
$lock = Get-Content -LiteralPath $lockPath -Raw | ConvertFrom-Json
Write-Log "lockfile   ：$lockPath" OK

# 收集 lock.json 里全部 {sha256,url}，并记住它属于哪个包
$found = New-Object System.Collections.ArrayList
function Walk($node, [string]$path) {
    if ($null -eq $node) { return }
    if ($node -is [System.Management.Automation.PSCustomObject]) {
        $props = $node.PSObject.Properties.Name
        if ($props -contains 'sha256' -and $props -contains 'url') {
            # 路径形如 lock/packages/<包名>/artifacts/<目标>[/索引]
            $pkgName = ($path -split '/')[2]
            [void]$script:found.Add([pscustomobject]@{ pkg = $pkgName; sha = $node.sha256; url = $node.url; path = $path })
            return
        }
        foreach ($p in $node.PSObject.Properties) { Walk $p.Value "$path/$($p.Name)" }
    } elseif ($node -is [System.Collections.IEnumerable] -and $node -isnot [string]) {
        $i = 0; foreach ($e in $node) { Walk $e "$path[$i]"; $i++ }
    }
}
Walk $lock 'lock'

# 只保留本机目标 + any
$targets = @('win32-x64', 'any')
$items = @($found | Where-Object { $t = $_.path; ($targets | Where-Object { $t -match [regex]::Escape("/$_") }) })
if (-not $All) { $items = @($items | Where-Object { $Packages -contains $_.pkg }) }

# 去重（同一 sha 可能被多个包/多目标引用）
$seen = @{}
$plan = New-Object System.Collections.ArrayList
foreach ($it in $items) {
    if ($seen.ContainsKey($it.sha)) { continue }
    $seen[$it.sha] = $true
    [void]$plan.Add($it)
}

if (-not $Quiet) { Write-Log ("待处理 {0} 项（{1}）" -f $plan.Count, $(if ($All) { "全部资产" } else { "包: " + ($Packages -join ', ') })) STEP }

$seeded = 0; $already = 0; $failed = @()
foreach ($it in $plan) {
    $file = $it.url.Split('/')[-1]
    if (Test-Seeded -Store $store -Sha256 $it.sha) {
        if (-not $Quiet) { Write-Log ("[{0}] {1}  已就绪" -f $it.pkg, $file) "SKIP" }
        $already++
        continue
    }
    $ver = $null
    if ($lock.packages -and $lock.packages.PSObject.Properties[$it.pkg]) {
        $ver = $lock.packages.PSObject.Properties[$it.pkg].Value.version
    }
    if (Test-Installed -Pkg $it.pkg -Version $ver) {
        if (-not $Quiet) { Write-Log ("[{0}] {1}  已安装，跳过" -f $it.pkg, $file) "SKIP" }
        $already++
        continue
    }
    Write-Log ("[{0}] {1}  缺失 → 开始找镜像" -f $it.pkg, $file) "INFO"
    if (Save-Artifact -Store $store -Sha256 $it.sha -Url $it.url -FileName $file) {
        $seeded++
    } else {
        Write-Log ("[{0}] {1}  所有源都失败" -f $it.pkg, $file) ERR
        $failed += "$($it.pkg)/$file"
    }
}

Remove-Item -LiteralPath $WorkDir -Recurse -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Log ("完成：新增 {0} 项，已存在 {1} 项，失败 {2} 项" -f $seeded, $already, $failed.Count) $(if ($failed.Count -eq 0) { "OK" } else { "WARN" })
if ($failed.Count -gt 0) {
    Write-Host "  以下资产拿不到，更新仍会卡在它们上：" -ForegroundColor Yellow
    $failed | ForEach-Object { Write-Host "    · $_" -ForegroundColor Yellow }
    Write-Host "  可选：找一个能用的代理后设置 HTTPS_PROXY 再跑一次，或手动把文件放到" -ForegroundColor Yellow
    Write-Host "        <store>\fetch-<sha256>\<文件名>（sha256 必须匹配 lock.json）" -ForegroundColor Yellow
}

# ---------------------------------------------------------------- get-windows 预编译预置
# get-windows 没有 C++ 工具链时无法源码编译（本机 vswhere 查不到 VC 工具），其预编译在 GitHub release。
# 这里把预编译 tarball 预取到 cache，并备份当前 node_modules 里可用的 .node。
# 注意：node-pre-gyp 只接受 https 源，吃不了本地 file://（Windows 会被解析成 /C:/... 非法路径），
# 真正生效靠 hermes-update.ps1 设置的 npm_config_node_get-windows_binary_host_mirror。
function Seed-GetWindowsPrebuilt {
    $gwPkg = Join-Path $InstallRoot "node_modules\get-windows\package.json"
    $ver = $null
    if (Test-Path -LiteralPath $gwPkg) { try { $ver = (Get-Content -LiteralPath $gwPkg -Raw | ConvertFrom-Json).version } catch { } }
    if (-not $ver) { $ver = "9.3.0" }
    $cacheRoot = Join-Path $HermesHome "cache\get-windows"
    $napiDir = "napi-9-win32-unknown-x64"
    $tarball = Join-Path $cacheRoot ("v{0}\{1}.tar.gz" -f $ver, $napiDir)

    # 1) 备份当前可用的 .node（供诊断/手动恢复）
    $src = Join-Path $InstallRoot ("node_modules\get-windows\lib\binding\{0}\node-get-windows.node" -f $napiDir)
    if (-not (Test-Path -LiteralPath $src)) {
        if (-not $Quiet) { Write-Log "[get-windows] 当前 node_modules 无预编译可备份" "SKIP" }
    } elseif ($DryRun) {
        if (-not $Quiet) { Write-Log ("[get-windows] 可缓存预编译: " + $src + "（DryRun 不写）") "SKIP" }
    } else {
        $dstDir = Join-Path $cacheRoot ("v{0}\{1}" -f $ver, $napiDir)
        New-Item -ItemType Directory -Path $dstDir -Force | Out-Null
        Copy-Item -LiteralPath $src -Destination (Join-Path $dstDir "node-get-windows.node") -Force
        Write-Log ("[get-windows] 已缓存当前预编译 -> " + (Join-Path $dstDir 'node-get-windows.node')) OK
    }

    # 2) 预取 release tarball
    if (Test-Path -LiteralPath $tarball) { if (-not $Quiet) { Write-Log ("[get-windows] tarball 已就绪: v{0}" -f $ver) "SKIP" }; return }
    if ($DryRun) { Write-Log ("[get-windows] 缺少 tarball: v{0}/{1}.tar.gz（DryRun 不下载）" -f $ver, $napiDir) INFO; return }
    $prefix = $GetWindowsMirror
    if (-not $prefix) { $prefix = "https://ghproxy.net/" }
    $mirrors = @(
        ($prefix.TrimEnd('/') + "/https://github.com/sindresorhus/get-windows/releases/download"),
        "https://ghfast.top/https://github.com/sindresorhus/get-windows/releases/download",
        "https://gh-proxy.com/https://github.com/sindresorhus/get-windows/releases/download"
    )
    New-Item -ItemType Directory -Path (Split-Path -Parent $tarball) -Force | Out-Null
    foreach ($m in $mirrors) {
        $u = "{0}/v{1}/{2}.tar.gz" -f $m.TrimEnd('/'), $ver, $napiDir
        try {
            Write-Log ("[get-windows] 试 {0}" -f $u)
            $wc = New-Object System.Net.WebClient
            if ($env:HTTPS_PROXY) { $wc.Proxy = New-Object System.Net.WebProxy($env:HTTPS_PROXY) }
            $wc.Headers.Add('User-Agent', 'seed-pm-store')
            $tmp = "$tarball.part"
            $wc.DownloadFile($u, $tmp)
            $b = [IO.File]::ReadAllBytes($tmp)
            if (-not ($b.Length -ge 2 -and $b[0] -eq 0x1F -and $b[1] -eq 0x8B)) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue; throw "不是 gzip 包" }
            Move-Item -LiteralPath $tmp -Destination $tarball -Force
            Write-Log ("[get-windows] 已获取 tarball -> " + $tarball) OK
            return
        } catch {
            Write-Log ("[get-windows] 源失败: " + $_.Exception.Message) WARN
        }
    }
    Write-Log "[get-windows] 所有镜像都拿不到 tarball（不影响其它资产；更新时会走 node-pre-gyp 镜像）" WARN
}

if (-not $SkipGetWindows) { Seed-GetWindowsPrebuilt }

Write-Host ""
if ($DryRun) { Write-Host "（DryRun：未做任何改动）" -ForegroundColor Yellow }

# 有资产没预置成功 -> 非 0 退出码，主脚本据此告警（更新很可能卡在下载）
if ($failed.Count -gt 0) { exit 2 }
exit 0
