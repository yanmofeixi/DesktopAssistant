# DesktopAssistant

Windows 系统托盘工具：显示器控制 + 游戏AHK自动启动 + 热键管理 + 定时提醒

## 功能

| 功能 | 说明 |
|------|------|
| **显示器控制** | 托盘菜单一键关闭/开启显示器 |
| **游戏AHK自动启动** | 检测游戏启动后自动运行对应AHK脚本 |
| **热键管理** | Against the Storm 存档快捷键（Ctrl+1/2/3） |
| **提醒功能** | 读取桌面 ToDo.json，定时弹窗提醒 |

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
 HotkeyManager.cs     # 全局热键（Ctrl+1/2/3）
 ReminderManager.cs   # ToDo.json 定时提醒
 DesktopAssistant.csproj
 app.manifest
 icon.ico
 publish/             # 发布输出
```

## 开发与发布

```powershell
dotnet publish -c Release -r win-x64 --self-contained false -p:PublishSingleFile=true -o ./publish
```

## 热键

| 热键 | 功能 | 条件 |
|------|------|------|
| `Ctrl+1` | 保存存档到 SL 文件夹 | Against the Storm 窗口激活时 |
| `Ctrl+2` | 从 SL 文件夹读取存档 | Against the Storm 窗口激活时 |
| `Ctrl+3` | 从备份恢复存档 | Against the Storm 窗口激活时 |

## ToDo.json 格式

```json
[{"TaskDescription": "开会", "ReminderTime": "2025-01-15T14:00:00"}]
```

## 技术栈

- .NET 8.0-windows（WinForms）
- Windows API P/Invoke
- System.Management（进程监控）
- System.Text.Json

