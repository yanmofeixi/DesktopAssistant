# DesktopAssistant Windows 部署包（1.1.0，x64）

把本 ZIP 复制到新的 Windows 10/11 x64 电脑，即可安装 DesktopAssistant、配置 SSH 公钥登录和登录启动。包内包含 .NET 8 Windows Desktop 运行时，不需要另装 .NET 或 SDK。

## 新机器安装

1. 使用将来要被操作的 **Windows 管理员账号**登录桌面。
2. 将 ZIP **完整解压到本地文件夹**，例如桌面的 `DesktopAssistant-Windows-x64`。不要直接在 ZIP 或共享盘中运行。
3. 双击 **`Install.cmd`**，接受 Windows 的管理员权限提示。
4. 等待完成。首次安装 OpenSSH Server 可能需要通过 Windows Update 下载组件；如果系统提示需要重启，重启后重新运行安装脚本。
5. 新开一个 **Windows PowerShell** 窗口，执行：

   ```powershell
   Invoke-RestMethod http://127.0.0.1:18888/status
   shot -Window active
   ```

`status` 应为 `ok`，`shot` 应返回用户目录中的 PNG 文件路径。程序应出现在系统托盘中。

程序本体可以离线运行；新 Windows 若尚未安装 OpenSSH Server，首次配置 SSH 需要系统组件来源。已经配置好 SSH 时可在管理员 PowerShell 执行 `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1 -SkipSsh`，只安装桌面程序和启动配置。

## 安装脚本做什么

- 安装到 `%ProgramFiles%\DesktopAssistant\app\`，不依赖原来电脑的 `C:\Code` 目录。
- 校验包内文件 SHA-256，确认 .NET 自包含运行时和控制端公钥完整。
- 安装/启动 Windows OpenSSH Server，将 `sshd` 设为自动启动；保留已有 SSH 配置、主机密钥和其他公钥。
- 把包内 `controller.pub` 加入当前账号实际使用的 authorized keys 文件，重复安装不重复添加。按照 Windows OpenSSH 的默认管理员配置，这通常是 `%ProgramData%\ssh\administrators_authorized_keys`，其公钥对使用该文件的管理员账号生效。
- 新建的 SSH 防火墙规则默认只允许 `LocalSubnet`；已有其他 SSH 防火墙规则保留其原有范围。HTTP 服务继续只监听 `127.0.0.1:18888`，不开放 HTTP 入站端口。
- 在当前用户的 Windows PowerShell 全主机 profile 中加载 `RemoteControl.ps1`，提供 `shot`、`txt`、`c`、`paste`、`k` 等快捷命令。
- 创建 `DesktopAssistant-<用户SID>` 计划任务，在该用户登录时以最高权限启动，任务使用 **Interactive** 登录类型，操作真实登录桌面。安装结束也会尝试启动并检查 HTTP 状态。
- 更新已有安装时备份旧程序、profile、授权公钥文件及其 ACL、旧任务定义。备份位于 `%LOCALAPPDATA%\DesktopAssistant\InstallBackups\`，按安装时间区分。

安装默认针对当前已登录的管理员账号。如果 Windows 的提权过程切换成了另一个账号，脚本会停止，请先登录目标管理员账号后再安装。无需提供或保存 Windows 登录密码。

脚本没有替换系统的默认 SSH shell。SSH 默认进入 CMD；执行 `powershell.exe -NoLogo` 后加载快捷命令。不要加 `-NoProfile`，或手动执行：

```powershell
. "$env:ProgramFiles\DesktopAssistant\RemoteControl.ps1"
```

## 公钥和连接

本包的 `controller.pub` 是生成部署包时指定的**控制端公钥**，本次包已使用现有控制端连接 Windows 的公钥。拥有对应私钥的控制端可登录新机器。包内不含私钥、Windows 密码、旧机器的 SSH 主机密钥或 Home Assistant 配置。

首次部署后，使用新机器的实际用户名和 IP，在控制端配置独立 SSH 别名，例如：

```sshconfig
Host windows-new
    HostName 192.168.0.200
    User new_windows_username
    IdentityFile ~/.local/ssh/windows_desktop_ed25519
    IdentitiesOnly yes
    ServerAliveInterval 60
```

IP 和用户名是示例，不能沿用旧机器的地址或 `yanmo`，除非新机器确实使用它们。首次连接时将显示的主机指纹与安装脚本输出的指纹比较。

```bash
ssh windows-new
ssh windows-new 'powershell.exe -NoLogo -NonInteractive -Command "shot -Window active"'
```

如果需要让不同控制端使用另一把 SSH 公钥，应重新生成包并传入另一份 `.pub` 文件；包有内容校验，直接修改 `controller.pub` 后校验会失败。现有授权公钥不会在更新时被删除。

从 VPN/Tailscale 等非本地子网连接时，按实际控制端地址指定新增规则的范围，例如在管理员 PowerShell 执行：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1 -SshRemoteAddress '100.100.100.100'
```

这个地址只是示例。已有自定义 SSH `Match`、认证组合、监听地址或防火墙策略会继续生效；安装脚本不会绕过这些设置。

## 截图和输入

截图目录已改成：

```text
%LOCALAPPDATA%\DesktopAssistant\Screenshots
```

例如当前登录用户是 `Alice`，通常为 `C:\Users\Alice\AppData\Local\DesktopAssistant\Screenshots`。目录在首次启动时自动创建，保留最新 10 张；新机器不需要 D 盘。旧机器的 D 盘历史截图不会自动搬迁或删除。

| 操作 | PowerShell 命令 |
|---|---|
| 截图并保存 / 窗口列表 | `shot active -MaxWidth 1280` / `shot -ListWindows` |
| 读窗口文字 / 元素 / 网址 | `txt active` / `el active -Type button` / `url active` |
| 点击坐标 / 按名称点击 | `c 500 300` / `c -Name 确定 -Type button` |
| 滚动 / 按键 | `sc -240` / `k enter` / `k ctrl+w` |
| 中文、emoji、多行快速粘贴 | `paste '你好，Windows 👋'` / `paste -File "$env:TEMP\input.txt"` |

坐标与句柄为示例，坐标一律为桌面物理像素。控制端也可不经 PowerShell，直接通过 SSH 调用 `curl.exe` 访问 HTTP 接口，例如截图直接写到控制端文件、长文本经标准输入粘贴：

```bash
ssh windows-new 'curl.exe -s "http://127.0.0.1:18888/screenshot?window=active&maxWidth=1280"' > /tmp/w.png
ssh windows-new 'curl.exe -s -H "Content-Type: text/plain; charset=utf-8" --data-binary @- http://127.0.0.1:18888/paste' < /tmp/input.txt
```

全部操作（读文字、按名称点击、等待条件、批量执行等）见项目 [README](../README.md#ssh-远程操作)，或在目标机器执行 `curl.exe -s http://127.0.0.1:18888/help`。

## 更新与排查

更新时退出旧版程序，将新 ZIP 解压到独立文件夹，重新运行 `Install.cmd`。对于本安装器安装的版本，安装器也会在替换文件前停止该安装路径的进程，再启动新版。

若以前直接从 `C:\Code\DesktopAssistant\publish` 或其他目录启动，需先退出旧实例，并取消其原有登录启动设置。新安装器遇到其他程序占用 18888 会报错，不会结束其他路径的程序，也不会删除旧启动任务。

- **程序无法启动**：保持 `app` 内所有运行时文件完整；不要只复制 EXE。检查 `%LOCALAPPDATA%\DesktopAssistant\Logs`。
- **SSH 超时**：确认 IP、网络可达性、`sshd` 服务状态及防火墙允许的来源地址。SSH 未认证前的超时与公钥内容无关。
- **Permission denied**：检查用户名是否正确、控制端是否使用对应私钥。自定义 SSH 认证规则仍然生效。
- **SSH 能连但截图/输入失败**：让该用户保持登录、桌面解锁、电脑唤醒。该程序不替代登录界面，也不能作为 Session 0 的 Windows 服务来控制桌面。
- **找不到 shot/paste**：新开 Windows PowerShell，或手动加载上面的 `RemoteControl.ps1`。已有 profile 提前 `return`、执行策略或公司策略可能阻止 profile 加载；不要为此全局关闭系统策略。
- **安装中途失败**：查看明确的错误，解决后可重复安装。OpenSSH 安装等已完成的系统步骤不会假装回滚，原文件备份保存在 InstallBackups 中。

游戏 AHK、壁灯和实时翻译为可选功能，其脚本和机器配置不随包迁移。没有 `HomeAssistant.json` 时静默跳过壁灯快捷键；需要这些附加功能时在新机器单独配置。

## 高级验证和卸载

只校验解压包，不进行安装：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1 -ValidateOnly
```

管理员可用 `-SkipSsh`、`-SkipProfile`、`-SkipStartup`、`-SkipStart` 分别跳过对应配置。自定义 `-InstallDirectory` 创建最高权限启动任务时必须位于 Program Files 下。

卸载时退出 DesktopAssistant，删除任务计划程序中的 `DesktopAssistant-<用户SID>`，删除 PowerShell profile 中 `BEGIN/END DesktopAssistant managed profile` 区块，再删除 `%ProgramFiles%\DesktopAssistant`。用户截图、配置和备份位于 `%LOCALAPPDATA%\DesktopAssistant`，按需保留。

SSH 是系统共享服务，其他任务可能使用它；卸载桌面程序不会要求删除 sshd、系统 SSH 配置或其他授权密钥。若需要撤销本控制端访问，仅删除相应 authorized keys 文件中与 `controller.pub` 匹配的那一行，并按需要删除 `DesktopAssistant-SSH-In-TCP` 防火墙规则。
