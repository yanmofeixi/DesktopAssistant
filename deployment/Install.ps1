#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$InstallDirectory = (Join-Path $env:ProgramFiles 'DesktopAssistant'),
    [string]$SshRemoteAddress = 'LocalSubnet',
    [switch]$ValidateOnly,
    [switch]$SkipSsh,
    [switch]$SkipProfile,
    [switch]$SkipStartup,
    [switch]$SkipStart,
    [string]$AdditionalPublicKey,
    [switch]$PauseOnExit,
    [string]$TargetUser
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
. (Join-Path $PSScriptRoot 'Deployment.Common.ps1')

function Save-Original {
    param([string]$Path, [string]$Name)
    if (Test-Path -LiteralPath $Path) {
        Copy-Item -LiteralPath $Path -Destination (Join-Path $backupDir $Name) -Recurse -Force
        Get-Acl -LiteralPath $Path | Export-Clixml -LiteralPath (Join-Path $backupDir ($Name + '.acl.xml'))
    }
}

function Complete-Installer {
    param([int]$ExitCode)
    # The elevated console closes on exit; keep fingerprints and errors readable.
    if ($PauseOnExit) { $null = Read-Host 'Press Enter to close this window' }
    exit $ExitCode
}

function Get-CurrentTargetUser {
    $current = [Security.Principal.WindowsIdentity]::GetCurrent()
    return [pscustomobject]@{
        Sid = $current.User.Value
        Name = $current.Name
        UserName = $env:USERNAME
        ProfileDirectory = $env:USERPROFILE
        LocalAppData = $env:LOCALAPPDATA
        PowerShellProfile = $PROFILE.CurrentUserAllHosts
    }
}

function ConvertTo-TargetUserArgument {
    param([Parameter(Mandatory)]$User)
    return [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(($User | ConvertTo-Json -Compress)))
}

function ConvertFrom-TargetUserArgument {
    param([Parameter(Mandatory)][string]$Encoded)
    return [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($Encoded)) | ConvertFrom-Json
}

function Read-AdditionalPublicKeys {
    $keys = @()
    if ($AdditionalPublicKey) { $keys += ConvertFrom-PublicKeyLine $AdditionalPublicKey }
    if ($AdditionalPublicKey -or [Console]::IsInputRedirected) { return $keys }
    Write-Host ''
    Write-Host 'Paste the public key of each Mac that should control this PC'
    Write-Host '(one line starting with ssh-ed25519, printed by the Mac setup steps).'
    Write-Host 'Press Enter on an empty line when done.'
    while ($true) {
        $line = Read-Host 'Public key'
        if (-not $line.Trim()) { return $keys }
        try {
            $keys += ConvertFrom-PublicKeyLine $line
            Write-Host 'Key accepted.'
        } catch {
            Write-Host "Rejected: $($_.Exception.Message) Paste the whole .pub line again." -ForegroundColor Yellow
        }
    }
}

function Write-ConnectionSummary {
    param([int]$Port)
    $addresses = @(Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object {
        $_.IPAddress -ne '127.0.0.1' -and -not $_.IPAddress.StartsWith('169.254.') -and $_.PrefixOrigin -in @('Dhcp', 'Manual')
    } | ForEach-Object { $_.IPAddress })
    $hostName = if ($addresses.Count -gt 0) { $addresses[0] } else { '<this PC IPv4 address>' }
    Write-Host ''
    Write-Host ('LAN IPv4 addresses: ' + $(if ($addresses.Count -gt 0) { $addresses -join ', ' } else { 'none found' }))
    Write-Host 'Add this to ~/.ssh/config on the Mac (pick the address on the Mac''s network):'
    Write-Host ''
    Write-Host 'Host windows-desktop'
    Write-Host "    HostName $hostName"
    if ($Port -ne 22) { Write-Host "    Port $Port" }
    Write-Host "    User $($target.UserName)"
    Write-Host ''
}

function Install-SshAccess {
    param([object[]]$Keys)
    $defaultRuleBefore = Get-NetFirewallRule -Name 'OpenSSH-Server-In-TCP' -ErrorAction SilentlyContinue
    $sshDir = Join-Path $env:WINDIR 'System32\OpenSSH'
    $sshd = Join-Path $sshDir 'sshd.exe'
    if (-not (Get-Service sshd -ErrorAction SilentlyContinue)) {
        Write-Host 'Installing Windows OpenSSH Server (Windows Update may be required)...'
        $result = Add-WindowsCapability -Online -Name 'OpenSSH.Server~~~~0.0.1.0'
        if ($result.RestartNeeded) { throw 'Windows requires a restart. Restart, then run this installer again.' }
    }
    # Respect a separately installed OpenSSH service instead of silently replacing it.
    $service = Get-CimInstance Win32_Service -Filter "Name='sshd'"
    if ($service.PathName -match '^"([^"]+)"' -or $service.PathName -match '^(\S+\.exe)(?:\s|$)') {
        $sshd = $Matches[1]
        $sshDir = Split-Path -Parent $sshd
    }
    if (-not (Test-Path -LiteralPath $sshd)) { throw 'Cannot locate the installed sshd.exe.' }
    $sshState = Join-Path $env:ProgramData 'ssh'
    $null = New-Item -ItemType Directory -Path $sshState -Force
    $sshConfig = Join-Path $sshState 'sshd_config'
    if (-not (Test-Path -LiteralPath $sshConfig)) {
        Copy-Item -LiteralPath (Join-Path $sshDir 'sshd_config_default') -Destination $sshConfig
    }
    & (Join-Path $sshDir 'ssh-keygen.exe') -A
    if ($LASTEXITCODE -ne 0) { throw 'Generating missing SSH host keys failed.' }
    & (Join-Path $sshDir 'ssh-keygen.exe') -l -f (Join-Path $PSScriptRoot 'controller.pub')
    if ($LASTEXITCODE -ne 0) { throw 'The supplied controller public key is invalid.' }
    & $sshd -t
    if ($LASTEXITCODE -ne 0) { throw 'Existing SSH configuration did not pass sshd -t; it was not overwritten.' }
    $effective = @(& $sshd -T -C ("user={0},host=localhost,addr=127.0.0.1" -f $target.UserName.ToLowerInvariant()))
    if ($LASTEXITCODE -ne 0) { throw 'Cannot inspect effective SSH configuration.' }
    $keySetting = @($effective | Where-Object { $_ -match '^authorizedkeysfile ' })
    if ($keySetting.Count -ne 1) { throw 'Cannot determine AuthorizedKeysFile.' }
    $keyPath = ($keySetting[0] -split '\s+')[1]
    if ($keyPath -eq 'none') { throw 'Existing SSH configuration disables authorized key files.' }
    $keyPath = $keyPath.Replace('__PROGRAMDATA__', $env:ProgramData).Replace('%h', $target.ProfileDirectory).Replace('%u', $target.UserName)
    if (-not [IO.Path]::IsPathRooted($keyPath)) { $keyPath = Join-Path $target.ProfileDirectory $keyPath }
    $keyPath = [IO.Path]::GetFullPath($keyPath)
    Save-Original $keyPath 'authorized_keys'
    $null = New-Item -ItemType Directory -Path (Split-Path -Parent $keyPath) -Force
    $text = if (Test-Path -LiteralPath $keyPath) { [IO.File]::ReadAllText($keyPath) } else { '' }
    foreach ($controllerKey in $Keys) { $text = Add-ControllerKeyText $text $controllerKey }
    [IO.File]::WriteAllText($keyPath, $text, (New-Object Text.UTF8Encoding($false)))
    $admins = New-Object Security.Principal.SecurityIdentifier('S-1-5-32-544')
    $system = New-Object Security.Principal.SecurityIdentifier('S-1-5-18')
    $globalKeys = $keyPath.Equals((Join-Path $sshState 'administrators_authorized_keys'), [StringComparison]::OrdinalIgnoreCase)
    $acl = New-Object Security.AccessControl.FileSecurity
    $acl.SetAccessRuleProtection($true, $false)
    $targetSid = New-Object Security.Principal.SecurityIdentifier($target.Sid)
    $acl.SetOwner($(if ($globalKeys) { $admins } else { $targetSid }))
    foreach ($sid in @($admins, $system)) {
        $acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule($sid, 'FullControl', 'Allow')))
    }
    if (-not $globalKeys) {
        $acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule($targetSid, 'FullControl', 'Allow')))
    }
    Set-Acl -LiteralPath $keyPath -AclObject $acl
    $portLine = @($effective | Where-Object { $_ -match '^port ' })
    if ($portLine.Count -ne 1) { throw 'Expected a single SSH port; configure the firewall manually for this SSH setup.' }
    $port = [int]($portLine[0] -split '\s+')[1]
    # An OpenSSH capability install may create an unrestricted rule. Scope only that newly created rule.
    if (-not $defaultRuleBefore -and (Get-NetFirewallRule -Name 'OpenSSH-Server-In-TCP' -ErrorAction SilentlyContinue)) {
        Set-NetFirewallRule -Name 'OpenSSH-Server-In-TCP' -RemoteAddress $SshRemoteAddress | Out-Null
    }
    $ruleName = 'DesktopAssistant-SSH-In-TCP'
    if (Get-NetFirewallRule -Name $ruleName -ErrorAction SilentlyContinue) {
        Set-NetFirewallRule -Name $ruleName -Enabled True -Direction Inbound -Action Allow -Protocol TCP -LocalPort $port -RemoteAddress $SshRemoteAddress | Out-Null
    } else {
        New-NetFirewallRule -Name $ruleName -DisplayName 'DesktopAssistant SSH' -Enabled True -Direction Inbound -Action Allow -Protocol TCP -LocalPort $port -RemoteAddress $SshRemoteAddress -Profile Any | Out-Null
    }
    Set-Service sshd -StartupType Automatic
    Start-Service sshd
    Write-Host "SSH ready: $($target.UserName)@$($env:COMPUTERNAME), port $port; key file: $keyPath ($($Keys.Count) controller key(s) ensured)"
    Write-Host 'Host fingerprints (compare these when first connecting):'
    Get-ChildItem -LiteralPath $sshState -Filter 'ssh_host_*_key.pub' | ForEach-Object {
        & (Join-Path $sshDir 'ssh-keygen.exe') -l -f $_.FullName
    }
    Write-ConnectionSummary $port
}

try {
    Test-DesktopPayload $PSScriptRoot
    $key = Read-ControllerPublicKey (Join-Path $PSScriptRoot 'controller.pub')
    if ($ValidateOnly) { Write-Host 'PASS: payload hashes, self-contained runtime and controller public key.'; Complete-Installer 0 }
    if (-not [Environment]::Is64BitProcess) { throw 'Run the installer using 64-bit Windows PowerShell.' }
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    # The account that launched the installer owns the desktop; UAC may elevate with a different administrator.
    $target = if ($TargetUser) { ConvertFrom-TargetUserArgument $TargetUser } else { Get-CurrentTargetUser }
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        $arguments = @('-NoLogo', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (Quote-ProcessArgument $PSCommandPath),
            '-TargetUser', (ConvertTo-TargetUserArgument $target), '-InstallDirectory', (Quote-ProcessArgument $InstallDirectory),
            '-SshRemoteAddress', (Quote-ProcessArgument $SshRemoteAddress), '-PauseOnExit')
        if ($AdditionalPublicKey) { $arguments += @('-AdditionalPublicKey', (Quote-ProcessArgument $AdditionalPublicKey)) }
        foreach ($option in @('SkipSsh', 'SkipProfile', 'SkipStartup', 'SkipStart')) {
            if (Get-Variable -Name $option -ValueOnly) { $arguments += '-' + $option }
        }
        $child = Start-Process powershell.exe -Verb RunAs -ArgumentList ($arguments -join ' ') -Wait -PassThru
        exit $child.ExitCode
    }
    if ($target.Sid -ne $identity.User.Value) { Write-Host "Configuring desktop account $($target.Name) (elevated as $($identity.Name))." }
    $InstallDirectory = [IO.Path]::GetFullPath($InstallDirectory).TrimEnd('\')
    if ($InstallDirectory -eq [IO.Path]::GetPathRoot($InstallDirectory).TrimEnd('\') -or $InstallDirectory.StartsWith('\\')) {
        throw 'Choose a dedicated local installation directory.'
    }
    if (-not $SkipStartup -and -not $InstallDirectory.StartsWith($env:ProgramFiles.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
        throw 'A task with highest privileges must run from an administrator-protected Program Files directory.'
    }
    if ($InstallDirectory.Equals($PSScriptRoot.TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase)) { throw 'Extract the ZIP somewhere other than the installation destination.' }
    $installedExe = Join-Path $InstallDirectory 'app\DesktopAssistant.exe'
    if (-not $SkipStart) {
        foreach ($listener in @(Get-NetTCPConnection -State Listen -LocalPort 18888 -ErrorAction SilentlyContinue)) {
            $ownerProcess = Get-Process -Id $listener.OwningProcess -ErrorAction SilentlyContinue
            if (-not $ownerProcess -or $ownerProcess.Path -ne $installedExe) {
                throw 'Port 18888 is already in use by another installation. Exit it and disable its old login startup before installing.'
            }
        }
    }
    if ((Test-Path -LiteralPath $InstallDirectory) -and
        @(Get-ChildItem -LiteralPath $InstallDirectory -Force).Count -gt 0 -and
        -not (Test-Path -LiteralPath (Join-Path $InstallDirectory 'installed.json'))) { throw 'The destination contains unmanaged files. Choose an empty directory.' }
    $sshKeys = if ($SkipSsh) { @() } else { @($key) + @(Read-AdditionalPublicKeys) }
    $state = Join-Path $target.LocalAppData 'DesktopAssistant'
    $backupDir = Join-Path $state ('InstallBackups\' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    $null = New-Item -ItemType Directory -Path $backupDir -Force
    if (Test-Path -LiteralPath $InstallDirectory) { Save-Original $InstallDirectory 'previous-installation' }
    foreach ($running in @(Get-Process -Name DesktopAssistant -ErrorAction SilentlyContinue)) {
        if ($running.Path -eq $installedExe) { Stop-Process -Id $running.Id -Force; $running.WaitForExit(10000) | Out-Null }
    }
    $null = New-Item -ItemType Directory -Path $InstallDirectory -Force
    # Copy the named package payload only; user settings and screenshots remain in LOCALAPPDATA.
    foreach ($name in @('app', 'RemoteControl.ps1', 'Deployment.Common.ps1', 'Install.ps1', 'Install.cmd', 'README.md', 'controller.pub', 'SHA256SUMS.txt')) {
        if ($name -eq 'app' -and (Test-Path -LiteralPath (Join-Path $InstallDirectory 'app'))) {
            Remove-Item -LiteralPath (Join-Path $InstallDirectory 'app') -Recurse -Force
        }
        Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination $InstallDirectory -Recurse -Force
    }
    $taskName = 'DesktopAssistant-' + $target.Sid
    $version = ([Version](Get-Item -LiteralPath $installedExe).VersionInfo.FileVersion).ToString(3)
    @{ version = $version; user = $target.Name; task = $taskName; installedAt = (Get-Date).ToString('o') } |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $InstallDirectory 'installed.json') -Encoding UTF8
    if (-not $SkipProfile) {
        $profilePath = $target.PowerShellProfile
        Save-Original $profilePath 'powershell-profile.ps1'
        $null = New-Item -ItemType Directory -Path (Split-Path -Parent $profilePath) -Force
        $existing = if (Test-Path -LiteralPath $profilePath) { Get-Content -LiteralPath $profilePath -Raw } else { '' }
        $text = Add-DesktopProfileText $existing (Join-Path $InstallDirectory 'RemoteControl.ps1')
        [IO.File]::WriteAllText($profilePath, $text, (New-Object Text.UTF8Encoding($true)))
    }
    if (-not $SkipSsh) { Install-SshAccess $sshKeys }
    if (-not $SkipStartup) {
        if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
            Export-ScheduledTask -TaskName $taskName | Set-Content -LiteralPath (Join-Path $backupDir 'previous-task.xml') -Encoding UTF8
        }
        $action = New-ScheduledTaskAction -Execute $installedExe -WorkingDirectory (Split-Path -Parent $installedExe)
        $trigger = New-ScheduledTaskTrigger -AtLogOn -User $target.Name
        $taskPrincipal = New-ScheduledTaskPrincipal -UserId $target.Name -LogonType Interactive -RunLevel Highest
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -ExecutionTimeLimit ([TimeSpan]::Zero) -MultipleInstances IgnoreNew
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $taskPrincipal -Settings $settings -Force | Out-Null
    }
    if (-not $SkipStart) {
        if ($SkipStartup) { Start-Process -FilePath $installedExe -WorkingDirectory (Split-Path -Parent $installedExe) | Out-Null }
        else { Start-ScheduledTask -TaskName $taskName }
        $ready = $false
        for ($attempt = 0; $attempt -lt 20; $attempt++) {
            Start-Sleep -Milliseconds 500
            try {
                $info = Invoke-RestMethod 'http://127.0.0.1:18888/status' -TimeoutSec 2
                if ($info.status -eq 'ok') { $ready = $true; break }
            } catch { }
        }
        if (-not $ready) { throw 'Files installed, but the desktop API is not ready. Check the logged-in desktop and LOCALAPPDATA\DesktopAssistant\Logs.' }
    }
    Write-Host "Installed: $InstallDirectory"
    Write-Host "Backups: $backupDir"
    Write-Host 'Open a new Windows PowerShell window and run: shot -Window active'
    Write-Host 'Screenshot folder: %LOCALAPPDATA%\DesktopAssistant\Screenshots'
    Write-Host 'API: 127.0.0.1:18888 (no inbound HTTP firewall port is opened)'
    Complete-Installer 0
} catch {
    Write-Error $_ -ErrorAction Continue
    Complete-Installer 1
}
