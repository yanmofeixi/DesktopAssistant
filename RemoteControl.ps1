#region DesktopAssistant 远程控制快捷命令与 SSH 欢迎提示
# Parameters travel as a UTF-8 JSON body so Chinese text and quotes need no URL encoding.
function global:Invoke-DesktopAssistant {
    param(
        [Parameter(Mandatory = $true, Position = 0)][string]$Op,
        [Parameter(Position = 1)][hashtable]$Params = @{},
        [string]$RawJson,
        [int]$TimeoutSec = 120
    )
    $body = $RawJson
    if (-not $body) {
        $clean = @{}
        foreach ($entry in $Params.GetEnumerator()) {
            if ($null -ne $entry.Value -and [string]$entry.Value -ne '') { $clean[$entry.Key] = [string]$entry.Value }
        }
        $body = $clean | ConvertTo-Json -Compress
    }
    $request = @{
        Uri = 'http://127.0.0.1:18888/' + $Op; Method = 'POST'; UseBasicParsing = $true; ErrorAction = 'Stop'
        TimeoutSec = $TimeoutSec; ContentType = 'application/json; charset=utf-8'
        Body = [Text.Encoding]::UTF8.GetBytes($body)
    }
    try { $response = Invoke-WebRequest @request }
    catch {
        # One red line instead of a multi-line error record; use /batch when later steps must stop.
        $detail = $_.ErrorDetails.Message
        if (-not $detail) { $detail = $_.Exception.Message }
        Write-Host $detail.Trim() -ForegroundColor Red
        $global:LASTEXITCODE = 1
        return
    }
    $global:LASTEXITCODE = 0
    # One line per pipeline item so Select-Object / Select-String work on text results.
    [Text.Encoding]::UTF8.GetString($response.RawContentStream.ToArray()).TrimEnd() -split '\r?\n'
}
function global:shot {
    [CmdletBinding(DefaultParameterSetName = 'Capture')]
    param(
        [Parameter(Position = 0, ParameterSetName = 'Capture')][string]$Window,
        # Untyped: PowerShell parses 0,0,400,200 as an array.
        [Parameter(ParameterSetName = 'Capture')]$Region,
        [Parameter(ParameterSetName = 'Capture')][int]$MaxWidth,
        [Parameter(ParameterSetName = 'Capture')][switch]$Jpg,
        [Parameter(ParameterSetName = 'List', Mandatory = $true)][switch]$ListWindows
    )
    if ($ListWindows) { return Invoke-DesktopAssistant windows }
    $params = @{ save = 'true'; window = $Window; region = ($Region -join ',') }
    if ($MaxWidth) { $params.maxWidth = $MaxWidth }
    if ($Jpg) { $params.format = 'jpg' }
    Invoke-DesktopAssistant screenshot $params
}
function global:c {
    param(
        [Parameter(Position = 0)]$X, [Parameter(Position = 1)]$Y,
        [string]$Name, [string]$Type, $Index, [string]$Window,
        [string]$Button = 'left', [switch]$Double, [switch]$Shot
    )
    $params = @{ x = $X; y = $Y; name = $Name; type = $Type; index = $Index; window = $Window; button = $Button }
    if ($Double) { $params.double = 'true' }
    if ($Shot) { $params.shot = 'true' }
    Invoke-DesktopAssistant click $params
}
function global:dc { c @args -Double }
function global:rc { c @args -Button right }
function global:m {
    param([Parameter(Mandatory = $true, Position = 0)]$X, [Parameter(Mandatory = $true, Position = 1)]$Y, [switch]$Shot)
    $params = @{ x = $X; y = $Y }
    if ($Shot) { $params.shot = 'true' }
    Invoke-DesktopAssistant move $params
}
# sc is a built-in alias for Set-Content; remove it so the scroll function is callable.
Remove-Item Alias:sc -Force -ErrorAction SilentlyContinue
function global:sc {
    param([Parameter(Mandatory = $true, Position = 0)][int]$Delta, [Parameter(Position = 1)]$X, [Parameter(Position = 2)]$Y,
        [switch]$Horizontal, [switch]$Shot)
    $params = @{ delta = $Delta; x = $X; y = $Y }
    if ($Horizontal) { $params.horizontal = 'true' }
    if ($Shot) { $params.shot = 'true' }
    Invoke-DesktopAssistant scroll $params
}
function global:paste {
    [CmdletBinding(DefaultParameterSetName = 'Text')]
    param(
        [Parameter(Position = 0, Mandatory = $true, ParameterSetName = 'Text')]
        [ValidateNotNullOrEmpty()][string]$Text,
        [Parameter(Mandatory = $true, ParameterSetName = 'File')][string]$File,
        [string]$Keys
    )
    if ($PSCmdlet.ParameterSetName -eq 'File') {
        $Text = Get-Content -LiteralPath $File -Raw -Encoding UTF8 -ErrorAction Stop
    }
    Invoke-DesktopAssistant paste @{ text = $Text; keys = $Keys }
}
function global:t([Parameter(Mandatory = $true)][string]$Text) {
    Invoke-DesktopAssistant type @{ text = $Text } -TimeoutSec ([int][Math]::Max(120, [Math]::Ceiling($Text.Length * 0.1) + 30))
}
function global:k([Parameter(Mandatory = $true)][string]$Combo) { Invoke-DesktopAssistant key @{ combo = $Combo } }
function global:focus([Parameter(Mandatory = $true)][string]$Window) { Invoke-DesktopAssistant focus @{ window = $Window } }
function global:txt {
    param([Parameter(Position = 0)][string]$Window, [switch]$All, [int]$Max)
    $params = @{ window = $Window }
    if ($All) { $params.all = 'true' }
    if ($Max) { $params.max = $Max }
    Invoke-DesktopAssistant text $params
}
function global:el {
    param([Parameter(Position = 0)][string]$Window, [string]$Name, [string]$Type, [int]$Limit)
    $params = @{ window = $Window; name = $Name; type = $Type }
    if ($Limit) { $params.limit = $Limit }
    Invoke-DesktopAssistant elements $params
}
function global:url([Parameter(Position = 0)][string]$Window) { Invoke-DesktopAssistant url @{ window = $Window } }
function global:waitfor {
    param([string]$Text, [switch]$Gone, [switch]$All, [switch]$Change, [string]$Window, $Region, [int]$Timeout = 300)
    $params = @{ text = $Text; window = $Window; region = ($Region -join ','); timeout = $Timeout }
    if ($Gone) { $params.gone = 'true' }
    if ($All) { $params.all = 'true' }
    if ($Change) { $params.change = 'true' }
    Invoke-DesktopAssistant wait $params -TimeoutSec ($Timeout + 30)
}
function global:batch([Parameter(Mandatory = $true)][string]$File) {
    Invoke-DesktopAssistant batch -RawJson (Get-Content -LiteralPath $File -Raw -Encoding UTF8 -ErrorAction Stop) -TimeoutSec 3700
}

# Only interactive SSH terminals get the cheat sheet; `ssh host 'powershell -Command ...'` stays quiet.
if ($env:SSH_TTY) {
    Write-Host ''
    Write-Host ' [DesktopAssistant 远程控制] 完整说明: Invoke-DesktopAssistant help' -ForegroundColor Cyan
    Write-Host ' shot [-Window w] [-Region x,y,w,h] [-MaxWidth 1280] / shot -ListWindows / focus w' -ForegroundColor Yellow
    Write-Host ' txt [w] [-All] : 读可见文字；el [w] [-Name 文字] [-Type button] : 元素与坐标；url [w]'
    Write-Host ' c x y [-Shot] / c -Name 确定 [-Type button] / dc / rc / m x y / sc -240'
    Write-Host ' paste ''文字'' [-Keys ctrl+shift+v] / t ''逐字'' / k ctrl+m,m / waitfor -Text 完成 [-Gone] / batch -File steps.json'
    Write-Host ''
}
#endregion
