# Hermes 自建自动更新（Windows）

官方 `hermes update` 在「Windows + 国内网络 + 桌面端开着」时会踩四个坑，这个脚本把它们全部内建。

## 用法

- 双击 **`更新Hermes.cmd`**（推荐），或
- 命令行：

```powershell
& "$env:LOCALAPPDATA\hermes\tools\self-update\hermes-update.ps1"
```

### 参数

| 参数 | 说明 |
|---|---|
| `-DryRun` | 只检查不动手：探测网络、自检 uv、列出要清理的遗留物 |
| `-NoRelaunch` | 更新完不自动重开桌面端 |
| `-CountdownSeconds N` | 关闭桌面端前的倒计时（默认 10 秒，0 = 立刻关） |
| `-LaunchReadyTimeoutSeconds N` | 重开桌面端后等待后端就绪的上限（默认 90 秒，0 = 只确认进程起来，不等就绪） |
| `-MaxAttempts N` | 最多跑几轮（默认 3）：失败重试 + 还债/更新续跑都算在内 |
| `-UpdateTimeoutMinutes N` | 单轮总上限（默认 60 分钟） |
| `-GlobalTotalTimeoutMinutes N` | 所有轮次合计上限（默认 90 分钟） |
| `-ProgressRefreshSeconds N` | 屏幕进度刷新间隔（默认 2 秒） |
| `-StallSampleSeconds N` | 进度采样间隔（默认 10 秒，用于“过慢”判定） |
| `-SlowRatioPercent N` | “过慢”软阈值：最近窗口进度 < 基线速率×该百分比（默认 20） |
| `-SlowStreakSamples N` | 连续多少次采样过慢即终止（默认 6；同时要求阶段 ≥90s、无活动 ≥30s，且窗口内仍有进展） |
| `-NodeBinaryMirror <url>` | get-windows 预编译下载镜像前缀（默认 ghproxy.net） |
| `-GitMirrorExtra <前缀>` | 追加 git 镜像前缀（可多次传/逗号分隔；也可写进同目录 `git-mirrors.txt`） |
| `-GitProbeSeconds N` | 镜像并行探测窗口，先通者胜（默认 120 秒） |
| `-ReceiptsKeep N` | 保留最近几份更新收据（默认 3） |
| `-Branch main` | 更新到哪个分支 |
| `-InstallRoot <路径>` | 安装目录（默认 `%LOCALAPPDATA%\hermes\hermes-agent`） |
| `-NoPause` | 结束后不等待回车（给无人值守/计划任务用） |

## 它做了哪七件事

1. **网络探测（多镜像池 + 并行优选 + 慢则切换）**：先用 `git ls-remote` 单探 origin（150s 硬超时、禁用低速误杀）；不通则把「origin + 内置 7 个 GitHub 镜像 + 官方直连」一起**并行探测，先通者胜**，胜者设为 origin。更新中若拉取阶段失败/停滞，自动并行找下一个可用镜像并切换后重试。可用 `-GitMirrorExtra` 或同目录 `git-mirrors.txt`（每行一个前缀）自行扩充镜像池。顺带探测本机代理端口（7890/7897/10809/10808/1080/52058 等）。
2. **自检 uv（强制校验）**：`bin\uv.exe` 缺失或损坏时，先用 `.bak/.previous` 备份还原；不行则按 `pm/lock.json` 的锁定版本，优先用 PM store 中已校验的资产、再从**镜像池**（与 git 同一组）下载安装，**强制 sha256 校验**——校验不过的字节一律不落盘；读取不到锁定信息时拒绝安装。官方更新卡 6 分钟的那一步。
3. **清遗留**：僵尸「更新进行中」标记（先验 pid 死活）、每个 profile 只保留最新 1 份快照、半截备份文件、过期收据。**关键**：这些删除由 PowerShell 完成，不占更新进程"单次 ≤50 个文件"的删除预算。
4. **关桌面端**：倒计时提示 → 优雅关闭（`CloseMainWindow`）→ 20 秒不退才强杀 → 顺带停掉内存里还跑着旧代码的 `serve` 后端（它们会重新锁住 venv）。
5. **跑更新（自动跑到底）**：`venv\Scripts\python.exe -m hermes_cli.main update --yes --keep-stash`，输出同时写到 `logs\self-update-run.log`。主循环最多 3 轮：每轮结束检查「未完成收尾标记 + 本地/远端 SHA」，只要还有活就自动继续——**还债轮跑完会自动进入真正的拉取更新轮，不需要再点一次**；命中删除安全阀会自动清遗留后重试。
6. **校验**：读 `logs\update_receipts\latest.json`（outcome / 版本 from→to）+ 桌面端构建产物时间。
7. **收尾**：自动重开桌面端并等待后端就绪（`desktop.log` 出现 `HERMES_BACKEND_READY`，默认上限 90s；超时只告警，不谎报就绪），写 `.self-update-result.json`（新增 `desktopReady` / `originUrl` / `autostashCount` 字段），并留一份带时间戳的日志 `logs\self-update-<时间>.log`。

## 界面（B 方案 · 结构化清单）

运行界面按步骤 ①–⑦ 线性呈现；①–⑤ 为单行（步骤名 + 结果），⑥⑦ 为标题行，子阶段以树形展开，结尾是结果卡片：

```
 ▸ ① 代码源与代理      ✓ origin 可达（gh-proxy.com）
 ▸ ② 构建环境          ✓ 镜像已设 · 禁源码编译 · 快速失败
 ▸ ③ 残留进程          ✓ 无
 ▸ ④ uv 校准           ✓ 正常 · uv 0.12.3
 ▸ ⑤ 清理 / 资产预置    ✓ 清理 0 项 · store 就绪
 ▸ ⑥ 运行更新（最多 3 轮）
    ├─ 拉取代码      ✓ 00:52
    ├─ Node 依赖     ✓ 01:28
    └─ 第 1 轮        ✓ 完成
 ▸ ⑦ 校验与重开        ✓ 4e740313 → 234badf4 · 桌面端后端就绪
  ──────────────────────────────────────────────────────────────
  ✅ 完成 · 用时 17:21 · 桌面端已就绪 · origin 已恢复 · autostash 2 条
```

更新期间屏幕下方是一行精简进度（每 2 秒刷新；异常时追加“静默/疑似过慢”）：

```
[17:12:45] 第2/3轮 · 拉取代码 · 00:40 · +1.0GB/min
```

屏幕只保留步骤、阶段与结论；每个镜像探测明细、清理条目、uv 下载与校验过程等全部写入会话日志与官方日志。`seed-pm-store.ps1` 在被本脚本调用时带 `-Quiet`（收起 banner 与逐包“已就绪”行），单独运行时行为不变。

阶段按官方输出自动识别：拉取代码 / 同步 Python 依赖 / 安装 Node 依赖 / 构建 TUI / 构建 Web UI / 重建桌面端 / 更新完成收尾。

## 卡死判定（三态 + 以“进度”续期）

重点：**CPU 在跑不算活着**。判定用的是“真实进度”是否增长——进程树 IO 字节（`ReadTransferCount/WriteTransferCount`）＋ 日志输出增量。

进度按「每进程 IO 峰值」累计，**总量只增不减**：子进程退出后它的字节不会被抹掉（子进程退出导致总量倒退会被误判成“过慢”，这是实测踩过的坑）。

三种状态：

- **真死**：阶段静默窗口内进度/输出/IO 全零增长
- **过慢（含 CPU 满但进度不动的打转）**：连续多次采样进度低于基线速率的设定百分比；且必须同时满足「本阶段已跑满 90 秒（基线可信）」＋「确实已有 30 秒无 IO/输出」＋「最近 60 秒内仍有进展」。完全零进展不计入过慢，统一交给静默判死（120-300 秒），避免 60 秒软阈值抢在静默阈值前误杀构建收尾；纯速率对比遇到 IO 抖动容易误杀，宁可交给阶段硬上限兜底
- **正常**：速率达标才续期

换阶段会清空采样与过慢计数：否则上一阶段的高速率基线会瞬间把新阶段判成过慢。

策略为**激进止损（B）**：真死 / 过慢 / 阶段超时，一律终止；其中**构建阶段停滞不再重试**（重试只会再跑一次全量 `npm ci`）。

| 阶段 | 静默判死 | 过慢软阈值 | 硬上限 |
|---|---|---|---|
| 拉取代码 | 120s | 速率 < 1MB/s 持续 3min | 10 min |
| 安装 Node 依赖 | 120s | 5min 增量 < 前5min 的 20% | 20 min |
| 构建 TUI / Web | 120s | 同上 | 各 10 min |
| 重建桌面端 | 180s | 同上 | 45 min |
| 源码收尾（"finishing an interrupted source update"） | 180s | 同上 | 60 min |
| 收尾（Code updated! / Update complete 后） | 300s | — | 45 min |
| 单轮兜底 | — | — | 60 min |
| 3 轮合计 | — | — | 90 min |

## 失败报告

失败或未验证时，生成 `logs\update-report-<时间>.md`，含：概要（版本/退出码/耗时/终止原因）、卡在哪个阶段、进度速率与静默时长、环境（Node/npm 版本、`build_from_source`、registry、代理）、命中的错误特征（node-gyp / npm ERR! / CalledProcessError / 网络失败…）、最近 80 行输出、相关日志路径、建议命令。报告路径同时写入 `.self-update-result.json` 的 `report` 字段。

## 其它改动

- **界面（B 方案）**：步骤线性编号 ①–⑦、对齐的单行结果、⑥ 子阶段树、精简进度行（按实际长度清行，重定向到文件不再粘连）、结束卡片汇总（用时/桌面就绪/origin/autostash）；屏幕与日志分轨——日志保留全部明细与时间戳。
- **origin 不留镜像**：更新结束把 origin 恢复为裸 `https://github.com/NousResearch/hermes-agent.git`（官方地址套过镜像前缀的）；自定义 fork 按原样恢复。避免官方更新器把安装误判成 fork，也不留下随时会死的镜像 URL。
- **拉取失败分类**：命中 `should_include_obj` / `pack-objects` / `repack local links` 等本地仓库类错误时不换镜像（换镜像不会好），并在更新前幂等写入 `git config maintenance.auto false`（规避 git 2.53+ partial-clone 的 fetch 后自动 repack 崩溃，#124272）；网络类失败才走镜像轮换。
- **autostash 提示**：`--keep-stash` 累积的自动暂存会在日志与结果 JSON（`autostashCount`）中提示数量，不自动清理。
- **一键到底（还债 + 更新）**：官方更新遇到"未完成收尾"标记时只续跑收尾、不拉新代码；脚本会在同一轮循环里自动继续，直到「无收尾标记 且 本地 HEAD == 远端 HEAD」才算完成。收敛保护：连续两轮状态无变化即停（防空转）；本地/远端比对用 `git ls-remote`，不污染本地引用。
- **日志轮转**：每次运行先把上一轮的 `logs\self-update-run*.log` 归档到 `logs\archive\`（保留最近 3 轮）。失败报告、桌面重建判定、安全阀检测只读本轮内容，不会被历史日志污染。
- **预置器退出码**：`seed-pm-store.ps1` 有资产没预置成功会以非 0 退出；主脚本据此告警并写进结果 JSON（`seedExit`）——不再"失败也照跑"。
- **禁源码编译**：设 `npm_config_build_from_source=false`。本机无 C++ 工具链（vswhere 查不到 VC 工具），原生模块必须用预编译；下载失败会**快速失败**而非卡在编译。首次会因指纹变化触发一次全量 `npm ci`。
- **禁用 git 低速误杀**：给 git 子进程设 `GIT_HTTP_LOW_SPEED_LIMIT/TIME=0`，不再用固定低速阈值中止连接（本机链路慢，慢但在传会被误杀）；是否卡死交给阶段/静默看门狗判断。
- **get-windows 预编译**：官方支持 `npm_config_<module>_binary_host_mirror`，脚本默认用 `ghproxy.net` 代理 GitHub release（可用 `-NodeBinaryMirror` 覆盖）。`seed-pm-store.ps1` 会把预编译 tarball 与当前可用 `.node` 备份到 `%LOCALAPPDATA%\hermes\cache\get-windows\`。
- **清残留更新进程**：开始前按完整命令行结束上一轮遗留的 `hermes_cli.main update` / `source_completion.py` / `node-deps.mjs` / `desktop.mjs`。
- **僵尸标记收紧**：更新标记的 pid 若“不是更新进程”或“超过 30 分钟”，也按僵尸清理。

## 文件

```
tools\self-update\
├── 更新Hermes.cmd          # 双击入口
├── hermes-update.ps1       # 主脚本
├── seed-pm-store.ps1       # 预置器（含 get-windows 预编译预取/备份）
├── git-mirrors.txt         # 可选：自定义 git 镜像前缀（每行一个，# 为注释）
├── cache\                  # get-windows 预编译缓存（运行时生成）
└── README.md               # 本文件
```

脚本放在 `%LOCALAPPDATA%\hermes\tools\` 下（repo 之外），不会被 Hermes 自己的更新覆盖。
