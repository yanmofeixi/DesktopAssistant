# DesktopAssistant Windows 部署包（1.2.0，x64）

把本 ZIP 复制到 Windows 10/11 x64 电脑并运行 `Install.cmd`，即可安装 DesktopAssistant、配置 SSH 公钥登录和登录启动。包内包含 .NET 8 Windows Desktop 运行时，不需要另装 .NET 或 SDK。

## 安装

先在控制这台电脑的 Mac 上生成 SSH 密钥，并准备好公钥（一行以 `ssh-ed25519` 开头的文字，步骤见配套的 Mac 连接说明）。

1. 用将来要被操作的 **Windows 账号**登录桌面。标准账号也可以，运行安装时 UAC 会要求输入一个管理员账号的密码。
2. 将 ZIP **完整解压到本地文件夹**，例如桌面的 `DesktopAssistant-Windows-x64`。不要直接在 ZIP 或共享盘中运行。
3. 双击 **`Install.cmd`**，在 Windows 的权限提示中选「是」或输入管理员账号密码。
4. 在弹出的管理员窗口中看到 `Public key:` 时，粘贴 Mac 的公钥并回车，出现 `Key accepted.` 即成功；有多台 Mac 就逐行粘贴，最后直接按回车继续。
5. 等待完成。首次安装 OpenSSH Server 可能需要通过 Windows Update 下载组件；如果系统提示需要重启，重启后重新运行 `Install.cmd`。
6. 管理员窗口最后会显示：
   - 主机指纹（`SHA256:...`），Mac 首次连接时核对；
   - 本机局域网 IPv4 地址和一段 `Host windows-desktop ...` SSH 配置，照抄到 Mac 的 `~/.ssh/config`。

   记下这些内容后按回车关闭窗口。
7. 新开一个 **Windows PowerShell** 窗口，执行：

   ```powershell
   Invoke-RestMethod http://127.0.0.1:18888/status
   shot -Window active
   ```

   `status` 应为 `ok`，`shot` 应返回用户目录中的 PNG 文件路径。程序图标出现在系统托盘中。

建议在路由器上为这台电脑设置固定 IP（DHCP 保留），否则 IP 变化后需要修改 Mac 上的 `HostName`。

程序本体可以离线运行；新 Windows 若尚未安装 OpenSSH Server，首次配置 SSH 需要系统组件来源。已经配置好 SSH 时可在管理员 PowerShell 执行 `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1 -SkipSsh`，只安装桌面程序和启动配置。

## 安装脚本做什么

- 安装到 `%ProgramFiles%\DesktopAssistant\app\`。
- 校验包内文件 SHA-256，确认 .NET 自包含运行时和控制端公钥完整。
- 安装/启动 Windows OpenSSH Server，将 `sshd` 设为自动启动；保留已有 SSH 配置、主机密钥和其他公钥。
- 把包内 `controller.pub` 和安装时粘贴的公钥加入目标账号实际使用的 authorized keys 文件，重复安装不重复添加。按照 Windows OpenSSH 的默认配置，管理员账号使用 `%ProgramData%\ssh\administrators_authorized_keys`（对所有使用该文件的管理员账号生效），标准账号使用 `%USERPROFILE%\.ssh\authorized_keys`。
- 新建的 SSH 防火墙规则只允许 `LocalSubnet`；已有其他 SSH 防火墙规则保留其原有范围。HTTP 服务只监听 `127.0.0.1:18888`，不开放 HTTP 入站端口。
- 在当前用户的 Windows PowerShell 全主机 profile 中加载 `RemoteControl.ps1`，提供 `shot`、`txt`、`c`、`paste`、`k` 等快捷命令。
- 创建 `DesktopAssistant-<用户SID>` 计划任务，在该用户登录时以最高权限启动，任务使用 **Interactive** 登录类型，操作真实登录桌面。安装结束会启动程序并检查 HTTP 状态。
- 更新已有安装时备份旧程序、profile、授权公钥文件及其 ACL、旧任务定义。备份位于 `%LOCALAPPDATA%\DesktopAssistant\InstallBackups\`，按安装时间区分。

安装的目标账号是**双击 `Install.cmd` 的账号**：SSH 公钥、登录启动任务、PowerShell profile、截图和备份目录都属于它。UAC 提权时用的是另一个管理员账号也没关系，窗口开头会显示 `Configuring desktop account ...`。目标账号是标准账号时，SSH 登录后拿到的是标准权限，DesktopAssistant 也以标准权限运行，无法操作以管理员身份运行的窗口。安装器不保存任何 Windows 密码。

脚本不替换系统的默认 SSH shell。SSH 默认进入 CMD；执行 `powershell.exe -NoLogo` 后加载快捷命令。不要加 `-NoProfile`，或手动执行：

```powershell
. "$env:ProgramFiles\DesktopAssistant\RemoteControl.ps1"
```

## 公钥和连接

- `controller.pub` 是生成部署包时内置的控制端公钥，拥有对应私钥的控制端可直接登录。包内不含任何私钥、Windows 密码、SSH 主机密钥或 Home Assistant 配置。
- 其他控制端的公钥在安装时粘贴。之后要增加一台 Mac，重新运行 `Install.cmd` 粘贴它的公钥即可；已有安装会原地更新，已授权的公钥不会被删除。
- 无人值守安装可在管理员 PowerShell 中用参数传入一把额外公钥：

  ```powershell
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1 -AdditionalPublicKey 'ssh-ed25519 AAAA... new-mac'
  ```

- 只接受 `.pub` 公钥行；私钥或带 `restrict`、`command=` 等 authorized_keys 选项的行会被拒绝。

在 Mac 上按安装结束时显示的配置添加 SSH 别名（IP 和用户名以安装窗口输出为准）：

```sshconfig
Host windows-desktop
    HostName 192.168.1.50
    User alice
    IdentityFile ~/.ssh/windows_desktop_ed25519
    IdentitiesOnly yes
    ServerAliveInterval 30
```

首次连接时将显示的主机指纹与安装窗口输出的指纹比较，一致再输入 `yes`。

```bash
ssh windows-desktop
ssh windows-desktop 'curl.exe -s http://127.0.0.1:18888/status'
```

从 VPN/Tailscale 等非本地子网连接时，按实际控制端地址指定新增规则的范围，例如在管理员 PowerShell 执行：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1 -SshRemoteAddress '100.64.0.0/10'
```

已有自定义 SSH `Match`、认证组合、监听地址或防火墙策略会继续生效；安装脚本不会绕过这些设置。

## 截图和输入

截图目录为 `%LOCALAPPDATA%\DesktopAssistant\Screenshots`，例如当前用户是 `Alice` 时通常为 `C:\Users\Alice\AppData\Local\DesktopAssistant\Screenshots`。目录在首次启动时自动创建，保留最新 10 张。

| 操作 | PowerShell 命令 |
|---|---|
| 截图并保存 / 窗口列表 | `shot active -MaxWidth 1280` / `shot -ListWindows` |
| 读窗口文字 / 元素 / 网址 | `txt active` / `el active -Type button` / `url active` |
| 点击坐标 / 按名称点击 | `c 500 300` / `c -Name 确定 -Type button` |
| 滚动 / 按键 | `sc -240` / `k enter` / `k ctrl+w` |
| 中文、emoji、多行快速粘贴 | `paste '你好，Windows 👋'` / `paste -File "$env:TEMP\input.txt"` |

坐标与句柄为示例，坐标一律为桌面物理像素。控制端通常不经 PowerShell，直接通过 SSH 调用 `curl.exe` 访问 HTTP 接口，例如截图直接写到控制端文件、长文本经标准输入粘贴：

```bash
ssh windows-desktop 'curl.exe -s "http://127.0.0.1:18888/screenshot?window=active&maxWidth=1280"' > /tmp/w.png
ssh windows-desktop 'curl.exe -s -H "Content-Type: text/plain; charset=utf-8" --data-binary @- http://127.0.0.1:18888/paste' < /tmp/input.txt
```

全部接口（读文字、按名称点击、等待条件、批量执行等）在目标机器执行 `curl.exe -s http://127.0.0.1:18888/help` 查看。

## 更新与排查

更新时将新 ZIP 解压到独立文件夹，重新运行 `Install.cmd`。安装器会在替换文件前停止本安装路径的进程，再启动新版。

若这台电脑上已有从其他目录启动的 DesktopAssistant，需先退出该实例，并取消其登录启动设置。安装器遇到其他程序占用 18888 会报错，不会结束其他路径的程序，也不会删除其他启动任务。

- **程序无法启动**：保持 `app` 内所有运行时文件完整；不要只复制 EXE。检查 `%LOCALAPPDATA%\DesktopAssistant\Logs`。
- **粘贴公钥被拒绝**：确认复制的是 `.pub` 文件的完整一行（`ssh-ed25519 AAAA... 注释`），中间没有换行。
- **SSH 超时**：确认 IP、网络可达性、`sshd` 服务状态及防火墙允许的来源地址。SSH 未认证前的超时与公钥内容无关。
- **Permission denied**：检查用户名是否与安装窗口显示的一致、控制端是否使用对应私钥、该公钥是否已在安装时粘贴。自定义 SSH 认证规则仍然生效。
- **SSH 能连但截图/输入失败**：让该用户保持登录、桌面解锁、电脑唤醒。该程序不替代登录界面，也不能作为 Session 0 的 Windows 服务来控制桌面。
- **找不到 shot/paste**：新开 Windows PowerShell，或手动加载上面的 `RemoteControl.ps1`。已有 profile 提前 `return`、执行策略或公司策略可能阻止 profile 加载；不要为此全局关闭系统策略。
- **安装中途失败**：查看明确的错误，解决后可重复安装。OpenSSH 安装等已完成的系统步骤不会回滚，原文件备份保存在 InstallBackups 中。

游戏 AHK、壁灯和实时翻译为可选功能，其脚本和机器配置不随包部署。没有 `HomeAssistant.json` 时静默跳过壁灯快捷键；需要这些附加功能时单独配置。

## 高级验证和卸载

只校验解压包，不进行安装：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1 -ValidateOnly
```

管理员可用 `-SkipSsh`、`-SkipProfile`、`-SkipStartup`、`-SkipStart` 分别跳过对应配置。自定义 `-InstallDirectory` 创建最高权限启动任务时必须位于 Program Files 下。

卸载时退出 DesktopAssistant，删除任务计划程序中的 `DesktopAssistant-<用户SID>`，删除 PowerShell profile 中 `BEGIN/END DesktopAssistant managed profile` 区块，再删除 `%ProgramFiles%\DesktopAssistant`。用户截图、配置和备份位于 `%LOCALAPPDATA%\DesktopAssistant`，按需保留。

SSH 是系统共享服务，其他任务可能使用它；卸载桌面程序不会删除 sshd、系统 SSH 配置或授权密钥。若需要撤销某个控制端的访问，只删除 authorized keys 文件中对应的那一行，并按需要删除 `DesktopAssistant-SSH-In-TCP` 防火墙规则。
