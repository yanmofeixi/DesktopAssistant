Set-StrictMode -Version Latest

function Read-ControllerPublicKey {
    param([Parameter(Mandatory)][string]$Path)
    $lines = @([IO.File]::ReadAllLines($Path) | Where-Object { $_.Trim() -and -not $_.Trim().StartsWith('#') })
    if ($lines.Count -ne 1) { throw 'controller.pub must contain exactly one OpenSSH public key.' }
    $line = $lines[0].Trim()
    if ($line -notmatch '^(ssh-ed25519|ssh-rsa|ecdsa-sha2-nistp(?:256|384|521))\s+([A-Za-z0-9+/]+={0,2})(?:\s+.*)?$') {
        throw 'Expected a public .pub key; private keys and authorized_keys options are not accepted.'
    }
    $algorithm, $base64 = $Matches[1], $Matches[2]
    $blob = [Convert]::FromBase64String($base64)
    if ($blob.Length -lt 12 -or $blob.Length -gt 32768) { throw 'Invalid SSH public key length.' }
    $length = [int]$blob[0] * 16777216 + [int]$blob[1] * 65536 + [int]$blob[2] * 256 + [int]$blob[3]
    if ($length -ne $algorithm.Length -or $length + 4 -ge $blob.Length -or
        [Text.Encoding]::ASCII.GetString($blob, 4, $length) -ne $algorithm) { throw 'Invalid SSH public key encoding.' }
    if ($algorithm -eq 'ssh-ed25519' -and ($blob.Length -ne 51 -or
        $blob[15] -ne 0 -or $blob[16] -ne 0 -or $blob[17] -ne 0 -or $blob[18] -ne 32)) {
        throw 'Invalid Ed25519 public key.'
    }
    return [pscustomobject]@{ Line = $line; Blob = $base64; Algorithm = $algorithm }
}

function Add-ControllerKeyText {
    param([AllowEmptyString()][string]$Existing, [Parameter(Mandatory)]$Key)
    # Preserve existing restrictions/comments if this key is already registered.
    if ($Existing -match ('(?m)(?:^|\s)' + [regex]::Escape($Key.Blob) + '(?:\s|$)')) { return $Existing }
    if ($Existing -and -not $Existing.EndsWith("`n")) { $Existing += "`r`n" }
    return $Existing + $Key.Line + "`r`n"
}

function Add-DesktopProfileText {
    param([AllowEmptyString()][string]$Existing, [Parameter(Mandatory)][string]$ScriptPath)
    $escaped = $ScriptPath.Replace("'", "''")
    $block = "# BEGIN DesktopAssistant managed profile`r`n. '$escaped'`r`n# END DesktopAssistant managed profile"
    $pattern = '(?ms)^# BEGIN DesktopAssistant managed profile\r?\n.*?^# END DesktopAssistant managed profile[^\r\n]*'
    if ([regex]::IsMatch($Existing, $pattern)) {
        return [regex]::Replace($Existing, $pattern, [Text.RegularExpressions.MatchEvaluator]{ param($m) $block })
    }
    return $Existing.TrimEnd("`r", "`n") + "`r`n`r`n" + $block + "`r`n"
}

function Test-DesktopPayload {
    param([Parameter(Mandatory)][string]$Directory)
    foreach ($name in @('app\DesktopAssistant.exe', 'app\DesktopAssistant.dll', 'app\coreclr.dll',
        'app\hostfxr.dll', 'app\System.Private.CoreLib.dll', 'app\System.Windows.Forms.dll',
        'RemoteControl.ps1', 'controller.pub', 'SHA256SUMS.txt')) {
        if (-not (Test-Path -LiteralPath (Join-Path $Directory $name) -PathType Leaf)) { throw "Missing package file: $name" }
    }
    $config = Get-Content -LiteralPath (Join-Path $Directory 'app\DesktopAssistant.runtimeconfig.json') -Raw | ConvertFrom-Json
    if ($config.runtimeOptions.PSObject.Properties.Name -contains 'frameworks' -or
        $config.runtimeOptions.PSObject.Properties.Name -contains 'framework') { throw 'Package requires an external .NET runtime.' }
    $base = [IO.Path]::GetFullPath($Directory).TrimEnd('\') + '\'
    foreach ($entry in [IO.File]::ReadAllLines((Join-Path $Directory 'SHA256SUMS.txt'))) {
        if ($entry -notmatch '^([0-9a-fA-F]{64})  (.+)$') { throw 'Invalid SHA256 manifest entry.' }
        $hash, $relative = $Matches[1], $Matches[2]
        if ([IO.Path]::IsPathRooted($relative) -or $relative.Contains(':')) { throw 'Invalid payload path.' }
        $file = [IO.Path]::GetFullPath((Join-Path $Directory $relative))
        if (-not $file.StartsWith($base, [StringComparison]::OrdinalIgnoreCase)) { throw 'Payload path escapes package.' }
        if ((Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash -ne $hash) { throw "Package checksum mismatch: $relative" }
    }
    $null = Read-ControllerPublicKey (Join-Path $Directory 'controller.pub')
}

function Quote-ProcessArgument {
    param([string]$Value)
    # Windows CommandLineToArgvW quoting, including a trailing backslash.
    return '"' + [regex]::Replace([regex]::Replace($Value, '(\\*)"', '$1$1\"'), '(\\+)$', '$1$1') + '"'
}
