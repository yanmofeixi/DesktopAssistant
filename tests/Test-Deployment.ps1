#Requires -Version 5.1
[CmdletBinding()]
param([Parameter(Mandatory)][string]$PackageDirectory)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\deployment\Deployment.Common.ps1')
$key = Read-ControllerPublicKey (Join-Path $PackageDirectory 'controller.pub')
$old = "# existing key`r`nrestrict " + $key.Line + "`r`n"
if ((Add-ControllerKeyText $old $key) -cne $old) { throw 'Existing key restrictions were changed.' }
$first = Add-ControllerKeyText '# untouched comment' $key
if ((Add-ControllerKeyText $first $key) -cne $first) { throw 'Duplicate public key was added.' }
$profile = "Write-Output 'existing profile'`r`n"
$path = "C:\Program Files\Owner's tools\RemoteControl.ps1"
$first = Add-DesktopProfileText $profile $path
$second = Add-DesktopProfileText $first $path
if ($first -cne $second -or -not $first.Contains($profile.Trim())) { throw 'Profile update did not preserve/idempotently merge content.' }
$tokens, $parseErrors = $null, $null
$null = [Management.Automation.Language.Parser]::ParseInput($second, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw 'Profile quoting failed.' }
foreach ($file in @(Get-ChildItem (Join-Path $PSScriptRoot '..') -Filter '*.ps1' -Recurse | Where-Object { $_.FullName -notmatch '\\(dist|artifacts)\\' })) {
    $null = [Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count -gt 0) { throw "$($file.Name): $($parseErrors[0])" }
}
Test-DesktopPayload $PackageDirectory
$temporary = Join-Path $env:TEMP ('DesktopAssistant-deployment-test-' + [guid]::NewGuid().ToString('N'))
try {
    $null = New-Item -ItemType Directory -Path $temporary
    $badKey = Join-Path $temporary 'private-key.txt'
    [IO.File]::WriteAllText($badKey, '-----BEGIN OPENSSH PRIVATE KEY-----')
    $rejected = $false
    try { $null = Read-ControllerPublicKey $badKey } catch { $rejected = $true }
    if (-not $rejected) { throw 'Private-key input was not rejected.' }
    # Exercise installation/copy/update while leaving SSH, profiles, tasks and GUI untouched.
    $destination = Join-Path $temporary 'Installed App'
    for ($attempt = 0; $attempt -lt 2; $attempt++) {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PackageDirectory 'Install.ps1') `
            -InstallDirectory $destination -SkipSsh -SkipProfile -SkipStartup -SkipStart
        if ($LASTEXITCODE -ne 0) { throw 'Isolated install/update failed.' }
    }
    Test-DesktopPayload $destination
    Write-Host 'PASS: key validation/preservation, profile quoting/idempotency, PS 5.1 syntax, self-contained payload and isolated install/update.'
} finally {
    if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Recurse -Force }
}
