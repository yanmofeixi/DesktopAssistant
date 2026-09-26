# DesktopAssistant

Windows 系统托盘工具：显示器控制 + 游戏AHK自动启动 + 定时提醒

## 功能

| 功能 | 说明 |
|------|------|
| **显示器控制** | 托盘菜单一键关闭/开启显示器 |
| **游戏AHK自动启动** | 检测游戏启动后自动运行对应AHK脚本 |
| **书房壁灯快捷键** | 全局 `Ctrl+F1`：切换二楼书房壁灯，开灯亮度固定为 100% |
| **提醒功能** | 读取桌面 ToDo.json，定时弹窗提醒 |
| **空闲自动静音** | 检测到系统空闲 30 分钟自动静音，恢复操作后自动取消静音（可通过托盘菜单启用/禁用） |
| **屏幕截图与远程控制** | 内置 HTTP 服务（端口 18888）：截图（区域、缩放）、读取窗口文字/元素/网址、按坐标或元素名称点击、粘贴与按键、等待条件、批量执行 |

## SSH 远程操作

DesktopAssistant 在登录桌面中运行一个只监听 `127.0.0.1:18888` 的 HTTP 服务，SSH 登录后有两种用法：

- **HTTP 接口**：SSH 默认进入 CMD，直接调用 Windows 自带的 `curl.exe`。不经过 PowerShell、不需要控制端安装脚本，适合脚本和 AI Agent。`GET /help` 返回全部操作说明。
- **PowerShell 快捷命令**：profile 加载 `RemoteControl.ps1`，适合人工交互（见下文）。

### HTTP 接口

请求 `http://127.0.0.1:18888/<op>`。参数可写在查询串，或作为 JSON 对象请求体；含中文等非 ASCII 字符的参数用 JSON 请求体经标准输入发送，避免命令行编码问题。结果按类型返回：操作结果为单行 JSON（中文不转义、省略空字段）；`text`、`elements`、`url`、`windows`、`help` 为纯文本；截图默认为图片二进制。带 `Origin` 或 `Sec-Fetch-Mode` 请求头的请求（浏览器网页发出的请求）一律拒绝，防止网页借本机端口操控键鼠。

```bash
# 截图直接写到控制端文件；region 只截一块，maxWidth 限制宽度
ssh windows-desktop 'curl.exe -s "http://127.0.0.1:18888/screenshot?window=active&maxWidth=1280"' > /tmp/w.png
# 读窗口可见文字、浏览器网址、可点击元素及坐标
ssh windows-desktop 'curl.exe -s "http://127.0.0.1:18888/text?window=active"'
ssh windows-desktop 'curl.exe -s "http://127.0.0.1:18888/url?window=active"'
ssh windows-desktop 'curl.exe -s "http://127.0.0.1:18888/elements?window=active&type=button,tabitem"'
# 一串操作一次完成；最后一步是截图时输出就是图片
ssh windows-desktop 'curl.exe -s --data-binary @- http://127.0.0.1:18888/batch' > /tmp/w.png <<'EOF'
[{"op":"focus","window":"chrome"},{"op":"click","name":"确定","type":"button"},
 {"op":"wait","text":"已保存","timeout":60},{"op":"screenshot","window":"active","maxWidth":1280}]
EOF
# 长文本原样粘贴到当前焦点
ssh windows-desktop 'curl.exe -s -H "Content-Type: text/plain; charset=utf-8" --data-binary @- http://127.0.0.1:18888/paste' < /tmp/input.txt
```

| 操作 | 参数 | 说明 |
|---|---|---|
| `windows` | | 每行一个窗口：句柄、程序、标题、位置与尺寸 |
| `focus` | `window` | 恢复最小化窗口并切到前台 |
| `screenshot` | `window` `region=x,y,宽,高` `maxWidth` `format=png\|jpg` `save=1` | 默认返回图片；`save=1` 保存到截图目录并返回路径和坐标 JSON |
| `text` | `window` `all=1` `max=20000` | 读窗口文字，默认只读屏幕上可见的部分 |
| `elements` | `window` `name` `type=button,edit\|all` `limit=200` | 可见元素，每行 `类型 "名称" value="值" @中心x,y` |
| `url` | `window` | 浏览器当前完整网址 |
| `click` | `x y` 或 `name`（`type` `index` `window`）`button` `double=1` `shot=1` | 点击坐标，或按名称找到元素后点击其中心 |
| `move` / `drag` / `scroll` | `x y` / `x1 y1 x2 y2 duration` / `delta`（`x y` `horizontal=1`） | 移动、拖拽、滚轮（正数向上，一格约 120） |
| `paste` | `text` `keys=ctrl+v` | 经桌面剪贴板一次粘贴；`text` 也可作为 `text/plain` 请求体 |
| `type` | `text` | 逐字模拟 Unicode 按键，每字约 50ms |
| `key` | `combo` | `enter`、`ctrl+w`；逗号分隔依次按下，如 `ctrl+m,m` |
| `wait` | `text`（`gone=1` `all=1` `window`）或 `change=1`（`window` `region` `threshold=3`）`timeout=300` `interval=1000` | 等文字出现/消失或画面变化，超时返回 408 |
| `sleep` | `ms` | 暂停 |
| `batch` | POST JSON 数组，每项含 `op` 及其参数 | 依次执行，失败即停止并在 `details.completed` 返回已完成步骤；最后一步是未 `save` 的截图时响应就是图片 |
| `status` | | 屏幕、光标、服务运行时间与最新截图路径 |

**窗口选择**：`window` 按 `active`、`0x` 十六进制句柄、程序名（可带 `.exe`）、完整标题、标题片段依次匹配。省略时截图为全屏，其余操作取前台窗口。多个窗口匹配时返回候选列表要求改用句柄；找不到、完全移出屏幕，或（除 `focus` 外）已最小化时报错。窗口截图截取屏幕上的可见矩形，不激活窗口，遮挡内容也会出现在图中。

**坐标**：一律为桌面物理像素。截图响应头 `X-Screenshot-Left`、`X-Screenshot-Top`、`X-Screenshot-Scale` 给出换算：桌面坐标 = 偏移 + 图像坐标 / scale。`click`、`move`、`drag`、`scroll` 加 `shot=1` 时，`x`、`y` 直接取上一张返回截图上的像素，由服务按该图换算。

**文字读取**基于 Windows UI Automation。Chromium 浏览器在首次被访问时才构建网页的辅助功能树，服务查不到网页文档时会等待 1 秒重试一次。`text` 去掉嵌入对象占位符并压缩空行。能读多少取决于程序是否提供辅助功能数据：浏览器和使用标准控件的程序（如微信的会话列表、聊天内容，以 `listitem` 形式出现）可读；自绘界面（如 Directory Opus 文件列表）、游戏和 canvas 绘制的内容读不到，需要截图后按坐标点击。本地程序没有网页文档，`text` 只能按树的顺序列出元素，读内容时优先用 `elements` 指定 `type`。`click` 的 `name` 先找名称完全相同的元素，再找名称或值包含该文字的元素；匹配多个时返回 409 和候选列表，用 `type` 或从 0 开始的 `index` 区分。

**按键**：组合键用 `+` 连接。查询串里的 `+` 会被解码成空格，服务把空格也当作分隔符。`+`、`,` 本身用键名 `plus`、`comma`，其他标点键名为 `minus` `period` `slash` `semicolon` `quote` `backquote` `bracketleft` `bracketright` `backslash`。

**粘贴**：`paste` 在 STA 线程写入 Unicode 剪贴板再发送 `keys`（默认 `ctrl+v`；xterm 等网页终端用 `ctrl+shift+v`），支持中文、emoji、引号和换行。新内容保留在剪贴板中，不自动恢复，以免目标程序延迟读取时粘贴到旧内容；不追加 Enter。`text/plain` 请求体从不按 JSON 解析，粘贴以 `{` 开头的代码也安全。不要在 SSH 中用 `Set-Clipboard` 代替：SSH 会话的剪贴板与登录桌面不同。

**等待变化**：`change=1` 比较区域的 64 像素宽灰度缩略图，平均差异超过 `threshold`（0–255）即返回；光标闪烁、动画区域会误触发，应配合 `region` 限定范围。

### PowerShell 快捷命令

执行 `powershell.exe -NoLogo` 或 `ssh -t windows-desktop powershell.exe -NoLogo` 时自动加载；已打开的窗口可执行 `. C:\Code\DesktopAssistant\RemoteControl.ps1` 重载。命令把参数作为 UTF-8 JSON 请求体发送，输出按行返回；失败时打印一行红色错误 JSON 并把 `$LASTEXITCODE` 设为 1。只有交互式 SSH 终端显示速查提示。

| 操作 | 命令 |
|---|---|
| 截图并保存（返回路径与坐标） | `shot` / `shot active -Region 0,0,800,600 -MaxWidth 800` / `shot 0x123456 -Jpg` |
| 窗口列表 / 切到前台 | `shot -ListWindows` / `focus notepad` |
| 读文字 / 元素 / 网址 | `txt chrome` / `txt chrome -All` / `el chrome -Name 保存 -Type button` / `url chrome` |
| 点击 | `c 500 300` / `c 250 150 -Shot` / `c -Name 确定 -Type button -Window chrome` / `dc` / `rc`（参数同 `c`） |
| 移动 / 滚动 | `m 500 300` / `sc -240` / `sc 240 800 400` |
| 粘贴 / 逐字输入 | `paste '你好 "Windows"'` / `paste -File D:\Temp\input.txt` / `paste 'ls' -Keys ctrl+shift+v` / `t '逐字'` |
| 按键 | `k enter` / `k ctrl+w` / `k ctrl+m,m` |
| 等待 | `waitfor -Text 完成 -Window chrome -Timeout 600` / `waitfor -Text 加载中 -Gone` / `waitfor -Change -Region 0,0,400,300` |
| 批量 | `batch -File D:\Temp\steps.json` |

截图文件在 `%LOCALAPPDATA%\DesktopAssistant\Screenshots`，只保留最新 10 张。

## 新机器部署

使用 `Build-Package.ps1` 生成 `dist\DesktopAssistant-Windows-x64.zip`。发布包内置 .NET 8 Windows Desktop 运行时；在目标 Windows 管理员桌面解压后双击 `Install.cmd`，配置 SSH 公钥、登录启动任务和 PowerShell 快捷命令。截图默认写入当前用户的 LocalAppData，不要求 D 盘。

完整安装、更新和卸载说明见 [deployment/README.md](deployment/README.md)。Home Assistant、AHK、实时翻译等机器专有配置不进入部署包；缺少壁灯配置时静默跳过。

## 游戏AHK映射

| 游戏进程 | AHK脚本 |
|---------|--------|
| GenshinImpact.exe | C:\Code\AHK\Genshin.ahk |
| StarRail.exe | C:\Code\AHK\StarRail.ahk |

## 项目结构

```
DesktopAssistant/
 Program.cs           # 入口，托盘图标与菜单
 DisplayManager.cs    # 显示器开关（Windows API）
 GameAhkManager.cs    # 游戏进程监控，自动启动AHK
 StudyLightHotkey.cs  # Ctrl+F1 经 Home Assistant 控制二楼书房壁灯
 ReminderManager.cs   # ToDo.json 定时提醒
 IdleMuteManager.cs   # 空闲自动静音（Win32 + Core Audio API）
 ScreenshotManager.cs # 全屏/窗口/区域截图、缩放、坐标映射与旧图自动清理
 WindowCapture.cs     # 可见窗口枚举、选择与物理像素边界
 RemoteInputServer.cs # 远程控制 HTTP 服务：请求解析与响应输出
 RemoteCommands.cs    # 远程控制操作实现（单次请求与 /batch 共用）
 CommandParameters.cs # 操作参数解析、结果与错误类型
 InputSimulator.cs    # 鼠标、键盘、剪贴板粘贴与窗口前台切换
 UiAutomationReader.cs # UI Automation：窗口文字、可见元素、浏览器网址
 RemoteControl.ps1    # SSH 中使用的 PowerShell 快捷命令
 DesktopAssistant.csproj
 app.manifest
 icon.ico
 publish/             # 发布输出
```

## 开发与发布

```powershell
.\Build-Package.ps1 -PublicKeyPath 'C:\path\to\controller.pub'
```

## 书房壁灯快捷键

全局 `Ctrl+F1` 调用 Home Assistant 专用 webhook。HA 根据实体
`light.jim_s_office_jim_s_offc_wall_sconces` 的实际状态判断：亮着时关闭，关着时以
`brightness_pct: 100` 打开；状态不可用时不执行。长按不会重复触发，请求进行中和随后
500 毫秒内忽略重复按键；网络失败不会自动重试，避免重复切换。

连接配置位于 `%LOCALAPPDATA%\DesktopAssistant\HomeAssistant.json`，包含 `WebhookUrl`
字符串。该 URL 含专用随机标识，应只保存在用户配置中，不提交到源码或写入日志。
HA 自动化 `windows_ctrl_f1_study_wall_light` 仅接受本地 POST 请求，且只控制这一盏灯。
注册冲突、配置错误或网络错误通过托盘通知和日志报告；HTTP 成功仅表示 HA 收到请求。

## ToDo.json 格式

```json
[{"TaskDescription": "开会", "ReminderTime": "2025-01-15T14:00:00"}]
```

## 技术栈

- .NET 8.0-windows（WinForms）
- Windows API P/Invoke
- System.Management（进程监控）
- System.Windows.Automation（UI Automation 客户端，经 `UseWPF` 引用）
- System.Text.Json
