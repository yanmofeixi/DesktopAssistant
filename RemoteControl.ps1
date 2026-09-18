#region DesktopAssistant 远程控制快捷命令与 SSH 欢迎提示
function global:Invoke-DesktopAssistant {
    param([string]$Endpoint, [hashtable]$Query = @{}, [string]$Text, [switch]$Post)
    $pairs = foreach ($entry in $Query.GetEnumerator()) {
        [Uri]::EscapeDataString([string]$entry.Key) + '=' + [Uri]::EscapeDataString([string]$entry.Value)
    }
    $uri = 'http://127.0.0.1:18888/' + $Endpoint
    if ($pairs) { $uri += '?' + ($pairs -join '&') }
    $request = @{ Uri = $uri; TimeoutSec = 120; ErrorAction = 'Stop' }
    if ($Post) {
        $request.Method = 'POST'
        $request.ContentType = 'text/plain; charset=utf-8'
        $request.Body = [Text.Encoding]::UTF8.GetBytes($Text)
    }
    if ($Endpoint -eq 'type') {
        $request.TimeoutSec = [int][Math]::Max(120, [Math]::Ceiling($Text.Length * 0.1) + 30)
    }
    Invoke-RestMethod @request
}
function global:shot {
    [CmdletBinding(DefaultParameterSetName = 'Capture')]
    param(
        [Parameter(Position = 0, ParameterSetName = 'Capture')][string]$Window,
        [Parameter(ParameterSetName = 'List', Mandatory = $true)][switch]$ListWindows
    )
    if ($ListWindows) { return (Invoke-DesktopAssistant windows).windows }
    $query = @{ save = 'true' }
    if ($Window) { $query.window = $Window }
    Invoke-DesktopAssistant screenshot $query
}
function global:c($x, $y, $btn = 'left') { Invoke-DesktopAssistant click @{ x=$x; y=$y; button=$btn } }
function global:dc($x, $y) { Invoke-DesktopAssistant click @{ x=$x; y=$y; double='true' } }
function global:rc($x, $y) { Invoke-DesktopAssistant click @{ x=$x; y=$y; button='right' } }
function global:m($x, $y) { Invoke-DesktopAssistant move @{ x=$x; y=$y } }
function global:paste {
    [CmdletBinding(DefaultParameterSetName = 'Text')]
    param(
        [Parameter(Position = 0, Mandatory = $true, ParameterSetName = 'Text')]
        [ValidateNotNullOrEmpty()][string]$Text,
        [Parameter(Mandatory = $true, ParameterSetName = 'File')][string]$File
    )
    if ($PSCmdlet.ParameterSetName -eq 'File') {
        $Text = Get-Content -LiteralPath $File -Raw -Encoding UTF8 -ErrorAction Stop
    }
    Invoke-DesktopAssistant paste -Text $Text -Post
}
function global:t([string]$text) { Invoke-DesktopAssistant type -Text $text -Post }
function global:k([string]$combo) { Invoke-DesktopAssistant key -Text $combo -Post }
# sc is a built-in alias for Set-Content; remove it so the scroll function is callable.
Remove-Item Alias:sc -Force -ErrorAction SilentlyContinue
function global:sc($delta) { Invoke-DesktopAssistant scroll @{ delta=$delta } }

if ($env:SSH_CLIENT -or $env:SSH_CONNECTION -or $env:SSH_TTY) {
    Write-Host ''
    Write-Host ' [远程控制速查] (端口 18888)' -ForegroundColor Cyan
    Write-Host ' shot [-Window active|程序名|标题|0x句柄] : 截图；不传参数截全屏' -ForegroundColor Yellow
    Write-Host ' shot -ListWindows : 列出窗口；截图返回 left/top，用于换算桌面坐标'
    Write-Host ' c x y / dc x y / rc x y : 单击 / 双击 / 右键；m x y : 移动'
    Write-Host ' paste ''大段文字'' / paste -File D:\Temp\input.txt : 快速粘贴（默认）'
    Write-Host ' t ''你好世界'' : 逐字模拟按键（仅需要按键事件时）；k ctrl+w : 组合键；sc -240 : 向下滚动'
    Write-Host ''
}
#endregion
