#Requires -Version 5.1
[CmdletBinding()]
param([Parameter(Mandatory)][string]$PublicKeyPath)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
. (Join-Path $PSScriptRoot 'deployment\Deployment.Common.ps1')
$key = Read-ControllerPublicKey ([IO.Path]::GetFullPath($PublicKeyPath))
$dist = Join-Path $PSScriptRoot 'dist'
$stage = Join-Path $dist ('stage-' + [guid]::NewGuid().ToString('N'))
$package = Join-Path $stage 'DesktopAssistant-Windows-x64'
$null = New-Item -ItemType Directory -Path (Join-Path $package 'app') -Force
try {
    & dotnet publish (Join-Path $PSScriptRoot 'DesktopAssistant.csproj') -c Release -r win-x64 --self-contained true `
        -p:PublishSingleFile=false -p:PublishTrimmed=false -o (Join-Path $package 'app')
    if ($LASTEXITCODE -ne 0) { throw 'dotnet publish failed.' }
    foreach ($name in @('Install.ps1', 'Install.cmd', 'Deployment.Common.ps1', 'README.md')) {
        Copy-Item -LiteralPath (Join-Path $PSScriptRoot ('deployment\' + $name)) -Destination $package
    }
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'RemoteControl.ps1') -Destination $package
    [IO.File]::WriteAllText((Join-Path $package 'controller.pub'), $key.Line + "`n", (New-Object Text.UTF8Encoding($false)))
    $manifest = @(Get-ChildItem -LiteralPath $package -File -Recurse | Sort-Object FullName | ForEach-Object {
        $relative = $_.FullName.Substring($package.Length + 1)
        (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant() + '  ' + $relative
    })
    [IO.File]::WriteAllLines((Join-Path $package 'SHA256SUMS.txt'), [string[]]$manifest, (New-Object Text.UTF8Encoding($false)))
    Test-DesktopPayload $package
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Add-Type -AssemblyName System.IO.Compression
    $temporaryZip = Join-Path $dist ('package-' + [guid]::NewGuid().ToString('N') + '.zip')
    # Windows PowerShell's .NET compatibility settings can make CreateFromDirectory
    # write backslashes in entry names. ZIP paths must use forward slashes.
    $archive = [IO.Compression.ZipFile]::Open($temporaryZip, [IO.Compression.ZipArchiveMode]::Create)
    try {
        Get-ChildItem -LiteralPath $stage -File -Recurse | Sort-Object FullName | ForEach-Object {
            $entry = $_.FullName.Substring($stage.Length + 1).Replace('\', '/')
            [IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, $_.FullName, $entry, [IO.Compression.CompressionLevel]::Optimal) | Out-Null
        }
    } finally { $archive.Dispose() }
    $zip = Join-Path $dist 'DesktopAssistant-Windows-x64.zip'
    Move-Item -LiteralPath $temporaryZip -Destination $zip -Force
    (Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash.ToLowerInvariant() + '  DesktopAssistant-Windows-x64.zip' |
        Set-Content -LiteralPath ($zip + '.sha256') -Encoding ASCII
    Write-Host "PACKAGE: $zip"
    Write-Host "EXPANDED: $package"
    Write-Host ('ZIP BYTES: ' + (Get-Item -LiteralPath $zip).Length)
    # Keep this unique staging folder for validation; never publish from a stale output directory.
} catch { Write-Error $_ -ErrorAction Continue; exit 1 }
