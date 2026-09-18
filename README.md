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
| **屏幕截图与远程控制** | 内置轻量 HTTP 服务（端口 18888），支持 SSH 快捷命令进行全屏/窗口截图、鼠标点击/拖拽/滚轮、长文本快速粘贴、Unicode 逐键输入与快捷键 |

## SSH 远程操作

Mac 运行 `ssh -t windows-desktop powershell.exe -NoLogo`。Windows PowerShell profile 加载 `C:\Code\DesktopAssistant\RemoteControl.ps1`；已经打开的 PowerShell 可执行 `. C:\Code\DesktopAssistant\RemoteControl.ps1` 重载。

| 操作 | 命令 |
|---|---|
| 当前窗口截图 | `shot -Window active` |
| 按程序名或窗口标题截图 | `shot -Window notepad` / `shot -Window '标题片段'` |
| 列出窗口 / 按句柄截图 | `shot -ListWindows` / `shot -Window 0x123456` |
| 全屏截图 | `shot` |
| 单击 / 双击 / 右键 / 移动 | `c 500 300` / `dc 500 300` / `rc 500 300` / `m 500 300` |
| 向下 / 向上滚动 | `sc -240` / `sc 240` |
| 快速输入文字（默认） | `paste '你好 "Windows" 世界'` |
| 快速输入 UTF-8 文件内容 | `paste -File 'D:\Temp\input.txt'` |
| 逐字模拟按键（按需） | `t '你好 "Windows" 世界'` |
| 按键 / 组合键 | `k enter` / `k ctrl+w` / `k alt+tab` |

`shot` 返回截图文件路径和 `left`、`top`、`width`、`height`。窗口截图坐标 `(x,y)` 对应桌面坐标 `(left+x, top+y)`，鼠标命令始终使用桌面坐标。截图文件在 `D:\Temp\ScreenshotMonitoring`，只保留最新 10 张；使用 Mac `scp` 按返回路径下载。

窗口选择按 `active`、十六进制句柄、程序名（可带 `.exe`）、完整标题、标题片段匹配。多个窗口匹配时返回候选列表并要求选择句柄；找不到、已最小化或完全移出屏幕时明确报错。窗口截图截取该窗口在屏幕上的可见矩形，不激活窗口，遮挡内容仍会出现在截图中。

快捷命令自动编码查询参数，以 UTF-8 请求体发送文字和组合键；`k ctrl+w` 直接使用 `+`。`sc` 的 `Set-Content` 内置别名由脚本移除，让滚轮函数正常生效。连接或 HTTP 错误会抛出异常。

### 快速输入大段文字

普通文字、代码和多行内容默认使用 `paste`，先把焦点放到目标输入框。它由登录桌面中的 DesktopAssistant 在 STA 线程写入 Unicode 剪贴板，再一次性发送 `Ctrl+V`，支持中文、emoji、引号、制表符和换行。文本保留在桌面剪贴板中，会覆盖原内容；不自动恢复，以免目标程序延迟读取时粘贴到旧内容。命令不追加 Enter，需要提交时另用 `k enter`。

大段文字优先保存为 UTF-8 文件后使用 `paste -File`，避免 PowerShell 字符串转义和命令行长度限制。也可从 Mac 把本地文件原样通过 SSH 标准输入发送，省去逐字输入和远端临时文件：

```bash
ssh -T windows-desktop 'curl.exe --silent --show-error --fail-with-body -H "Content-Type: text/plain; charset=utf-8" --data-binary @- http://127.0.0.1:18888/paste' < /tmp/input.txt
```

仅在需要模拟逐个按键、触发每次按键事件或目标不支持粘贴时使用 `t`；它保留原有逐字 Unicode 输入方式，每个字符等待约 50ms。单键和组合键仍使用 `k`。不要在 SSH 中用 `Set-Clipboard` 代替 `paste`：SSH 会话的剪贴板与登录桌面不同。

快速输入 API 为 `POST /paste`，请求体是 `text/plain; charset=utf-8` 原文；成功返回 `status`、`action: paste` 和 UTF-16 长度 `length`，表示已发送粘贴操作，目标程序是否接受需通过页面或截图确认。仅支持 POST，空文本或含 NUL 的文本返回错误。旧 `/type` 保持逐字输入行为。

HTTP 服务仅监听 `127.0.0.1:18888`：`GET /windows` 列出窗口，`GET /screenshot?window=active&save=true` 返回路径和坐标。`window` 可省略；不设 `save` 默认返回 PNG，`format=jpg` 返回 JPEG，`base64=true` 返回带图像数据的 JSON。二进制响应通过 `X-Screenshot-Left` / `X-Screenshot-Top` 返回桌面偏移。

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
 ScreenshotManager.cs # 全屏/窗口截图、坐标元数据与旧图自动清理
 WindowCapture.cs     # 可见窗口枚举、选择与物理像素边界
 RemoteControl.ps1    # SSH 中使用的 PowerShell 快捷命令
 RemoteInputServer.cs # 远程键鼠控制与即时截图 HTTP 服务
 DesktopAssistant.csproj
 app.manifest
 icon.ico
 publish/             # 发布输出
```

## 开发与发布

```powershell
dotnet publish -c Release -r win-x64 --self-contained false -p:PublishSingleFile=true -o ./publish
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
- System.Text.Json
