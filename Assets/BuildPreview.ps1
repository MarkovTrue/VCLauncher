<#
    BuildPreview.ps1 - generates the README images.

    Preview.png / Preview.en.png:
      Four shots of the main window with the hotkeys cheat sheet open
      (light/dark x Russian/English), composed over the desktop wallpaper:
      dark under light, offset, each with a soft drop shadow.
      The window is shown by a copy of VCLauncher.au3 with PreviewStates.au3
      appended: same GUI code, the sync status is set without running Sync.exe.

    Compare.png (only with -Only Compare or All; the README image is made by hand):
      video-compare started with the exact command the launcher built,
      in native fit, so the letterboxed release shows its black bars.
      The frame is saved by video-compare itself (F key).

    Nothing touches the mouse or the keyboard focus: windows stay off-screen
    and get input through posted messages.

    The file is UTF-8 with BOM: Windows PowerShell reads Cyrillic paths only with it.

    Run:  powershell -ExecutionPolicy Bypass -File Assets\BuildPreview.ps1 [-Only All|Compare]
#>
param(
    [ValidateSet('All', 'Launcher', 'Compare')]
    [string]$Only      = 'Launcher',
    [string]$Video1    = 'F:\Гнев (2004, UnHardsubs, Пучков) [BDRip 1080p].mkv',
    [string]$Video2    = 'F:\Гнев (2004, Open Matte, Пучков) [WEBRip 1080p].mkv',
    [int]   $OffsetMs  = 8010,                  # Sync.exe result for the pair above
    [string]$ShownFrom = 'Гнев',                # replaced in the file names in the EN readme shots
    [string]$ShownTo   = 'Man on Fire',
    [int]   $CompareAt = 1630,                  # seconds to seek in video-compare
    [string]$ShownDir  = 'D:\Portable\VCLauncher',  # video-compare folder shown in the command field
    [string]$AutoIt    = 'D:\Soft\AutoIt Script\AutoIt\AutoIt3_x64.exe',
    [string]$Wallpaper = '',                    # empty = autodetect in D:\Pictures
    [int]   $OffsetX   = 160,                   # dark shot horizontal offset from light
    [int]   $OffsetY   = 60,                    # dark shot vertical offset from light
    [int]   $Pad       = 42                     # wallpaper margin around the group
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$Root       = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$Au3        = Join-Path $Root 'VCLauncher.au3'
$States     = Join-Path $Root 'Assets\PreviewStates.au3'
$PreviewAu3 = Join-Path $Root '_Preview.au3'   # next to the source, so Include\ and @ScriptDir resolve
$Ini        = Join-Path $Root 'VCLauncher.ini'
$Cache      = Join-Path $Root 'VCLauncher.cache'
$Work       = Join-Path ([System.IO.Path]::GetTempPath()) 'VCLauncherPreview'
$CmdFile    = Join-Path $Work 'command.txt'
$OutPng     = Join-Path $Root 'Assets\Preview.png'     # RU readme: light RU over dark EN
$OutPngEn   = Join-Path $Root 'Assets\Preview.en.png'  # EN readme: light EN over dark RU
$OutCompare = Join-Path $Root 'Assets\Compare.png'

foreach ($f in @($AutoIt, $Au3, $States, $Video1, $Video2)) {
    if (-not (Test-Path -LiteralPath $f)) { throw "Not found: $f" }
}
# Wallpaper file name contains Cyrillic - match by ASCII mask to avoid script-encoding issues.
if ($Only -ne 'Compare' -and -not $Wallpaper) {
    $Wallpaper = (Get-ChildItem -LiteralPath 'D:\Pictures' -Filter '2560*1440 1.png' |
        Select-Object -First 1).FullName
    if (-not $Wallpaper) { throw "Wallpaper not found. Pass -Wallpaper <png path>." }
}
New-Item -ItemType Directory -Path $Work -Force | Out-Null

# --- WinAPI ---
Add-Type -Namespace Win -Name Api -MemberDefinition @'
public delegate bool EnumProc(System.IntPtr h, System.IntPtr l);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool EnumWindows(EnumProc cb, System.IntPtr l);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern int GetClassName(System.IntPtr h, System.Text.StringBuilder s, int n);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool IsWindowVisible(System.IntPtr h);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint pid);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool GetWindowRect(System.IntPtr hWnd, out RECT r);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool GetClientRect(System.IntPtr hWnd, out RECT r);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool ClientToScreen(System.IntPtr hWnd, ref POINT p);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool GetCursorPos(out POINT p);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool SetWindowPos(System.IntPtr hWnd, System.IntPtr after, int x, int y, int cx, int cy, uint flags);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool PostMessage(System.IntPtr hWnd, uint msg, System.IntPtr w, System.IntPtr l);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern uint MapVirtualKey(uint code, uint type);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool PrintWindow(System.IntPtr hWnd, System.IntPtr hdc, uint flags);
[System.Runtime.InteropServices.DllImport("dwmapi.dll")]
public static extern int DwmGetWindowAttribute(System.IntPtr hWnd, int attr, out RECT r, int size);
[System.Runtime.InteropServices.DllImport("kernel32.dll", CharSet=System.Runtime.InteropServices.CharSet.Unicode)]
public static extern bool WritePrivateProfileString(string section, string key, string val, string file);
public struct RECT { public int Left, Top, Right, Bottom; }
public struct POINT { public int X, Y; }
'@

$PW_RENDERFULLCONTENT = 2
$DWMWA_EXTENDED_FRAME = 9
$SWP_MOVE_ONLY        = 0x0001 -bor 0x0004 -bor 0x0010   # NOSIZE | NOZORDER | NOACTIVATE
$OffscreenY           = 10000                            # below every monitor

function Set-Ini([string]$section, [string]$key, [string]$val) {
    [Win.Api]::WritePrivateProfileString($section, $key, $val, $Ini) | Out-Null
}

# Finds the top-level window of a specific process (by PID and class), so a window
# of a previous, still-closing run is never picked up.
function Find-ProcessWindow([int]$targetPid, [string]$class) {
    $script:_found = [IntPtr]::Zero
    $cb = {
        param($h, $l)
        if ([Win.Api]::IsWindowVisible($h)) {
            $wpid = [uint32]0
            [Win.Api]::GetWindowThreadProcessId($h, [ref]$wpid) | Out-Null
            if ($wpid -eq $targetPid) {
                $sb = New-Object System.Text.StringBuilder 256
                [Win.Api]::GetClassName($h, $sb, 256) | Out-Null
                if ($sb.ToString() -eq $class) { $script:_found = $h; return $false }
            }
        }
        return $true
    }
    [Win.Api]::EnumWindows([Win.Api+EnumProc]$cb, [IntPtr]::Zero) | Out-Null
    return $script:_found
}

function Wait-ProcessWindow($proc, [string]$class, [int]$timeoutMs = 15000) {
    for ($t = 0; $t -lt $timeoutMs; $t += 100) {
        $hwnd = Find-ProcessWindow $proc.Id $class
        if ($hwnd -ne [IntPtr]::Zero) { return $hwnd }
        if ($proc.HasExited) { break }
        Start-Sleep -Milliseconds 100
    }
    throw "Window '$class' did not appear."
}

function Stop-Gracefully($proc, [int]$waitMs = 1500) {
    if ($proc.HasExited) { return }
    $proc.CloseMainWindow() | Out-Null
    if (-not $proc.WaitForExit($waitMs)) { $proc.Kill() }
}

# === Main window ===

# Copy of the source with the preview entry point: _MainGUI() at top level becomes
# _Preview(), the window is shown by _PreviewShow() off-screen and without activation.
function Build-PreviewScript {
    $lines = Get-Content -LiteralPath $Au3 -Encoding UTF8
    $out   = New-Object System.Collections.Generic.List[string]
    $entryDone = $false
    $showDone  = $false

    foreach ($line in $lines) {
        if (-not $entryDone -and $line -eq '_MainGUI()') {
            $out.Add('_Preview()'); $entryDone = $true; continue
        }
        if (-not $showDone -and $line -eq "`tGUISetState(@SW_SHOW)") {
            $out.Add("`t_PreviewShow()"); $showDone = $true; continue
        }
        $out.Add($line)
    }
    if (-not $entryDone) { throw 'Entry point _MainGUI() not found in VCLauncher.au3' }
    if (-not $showDone)  { throw 'GUISetState(@SW_SHOW) not found in _MainGUI' }

    $out.Add(''); $out.Add('')
    $out.AddRange([string[]](Get-Content -LiteralPath $States -Encoding UTF8))
    # UTF-8 with BOM: without it AutoIt reads Cyrillic as ANSI
    [System.IO.File]::WriteAllLines($PreviewAu3, $out, (New-Object System.Text.UTF8Encoding($true)))
}

# Starts the preview window and waits until PreviewStates writes the command file.
# $from/$to replace text in the shown file names; empty $from keeps the real names.
function Start-Launcher([string]$theme, [string]$lang, [string]$from = '', [string]$to = '') {
    Set-Ini 'Settings' 'Theme'    $theme
    Set-Ini 'Settings' 'Language' $lang
    Set-Ini 'Settings' 'Hotkeys'  '1'
    Set-Ini 'Settings' 'Fit'      'native'
    Set-Ini 'LastDirs' 'Video1'   $Video1
    Set-Ini 'LastDirs' 'Video2'   $Video2
    [Win.Api]::WritePrivateProfileString($null, $null, $null, $Ini) | Out-Null  # flush ini cache to disk

    Remove-Item -LiteralPath $CmdFile -Force -ErrorAction SilentlyContinue
    $argList = @("`"$PreviewAu3`"", $OffsetMs, "`"$CmdFile`"", "`"$ShownDir`"", "`"$from`"", "`"$to`"")
    $proc = Start-Process -FilePath $AutoIt -ArgumentList $argList -WorkingDirectory $Root -PassThru
    for ($t = 0; $t -lt 20000 -and -not (Test-Path -LiteralPath $CmdFile); $t += 100) {
        if ($proc.HasExited) { throw "Preview script exited with code $($proc.ExitCode)" }
        Start-Sleep -Milliseconds 100
    }
    if (-not (Test-Path -LiteralPath $CmdFile)) { Stop-Gracefully $proc; throw 'Preview window was not ready in 20 s.' }
    return $proc
}

function Capture-Window([IntPtr]$hwnd, [string]$outFile) {
    Start-Sleep -Milliseconds 1500  # let the skin timers finish the last repaint

    # full window rect (incl. invisible borders) - for PrintWindow
    $wr = New-Object Win.Api+RECT
    [Win.Api]::GetWindowRect($hwnd, [ref]$wr) | Out-Null
    $fullW = $wr.Right - $wr.Left
    $fullH = $wr.Bottom - $wr.Top

    $bmp = New-Object System.Drawing.Bitmap($fullW, $fullH, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $hdc = $g.GetHdc()
    [Win.Api]::PrintWindow($hwnd, $hdc, $PW_RENDERFULLCONTENT) | Out-Null
    $g.ReleaseHdc($hdc)
    $g.Dispose()

    # crop DWM invisible border - use the visual bounds
    $ext = New-Object Win.Api+RECT
    $hr  = [Win.Api]::DwmGetWindowAttribute($hwnd, $DWMWA_EXTENDED_FRAME, [ref]$ext, [System.Runtime.InteropServices.Marshal]::SizeOf($ext))
    if ($hr -eq 0) {
        $cx = $ext.Left   - $wr.Left
        $cy = $ext.Top    - $wr.Top
        $cw = $ext.Right  - $ext.Left
        $ch = $ext.Bottom - $ext.Top
        if ($cw -gt 0 -and $ch -gt 0 -and $cx -ge 0 -and $cy -ge 0) {
            $crop = New-Object System.Drawing.Rectangle($cx, $cy, $cw, $ch)
            $sub  = $bmp.Clone($crop, $bmp.PixelFormat)
            $bmp.Dispose()
            $bmp = $sub
        }
    }

    $bmp.Save($outFile, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
}

function Shoot-Launcher([string]$theme, [string]$lang, [string]$outFile, [string]$from = '', [string]$to = '') {
    $proc = Start-Launcher $theme $lang $from $to
    try {
        Capture-Window (Wait-ProcessWindow $proc 'AutoIt v3 GUI') $outFile
    }
    finally {
        Stop-Gracefully $proc
    }
}

function New-RoundedPath([float]$x, [float]$y, [float]$w, [float]$h, [float]$r) {
    $d = $r * 2
    $p = New-Object System.Drawing.Drawing2D.GraphicsPath
    $p.AddArc($x, $y, $d, $d, 180, 90)
    $p.AddArc($x + $w - $d, $y, $d, $d, 270, 90)
    $p.AddArc($x + $w - $d, $y + $h - $d, $d, $d, 0, 90)
    $p.AddArc($x, $y + $h - $d, $d, $d, 90, 90)
    $p.CloseFigure()
    return $p
}

# Soft shadow: stack translucent rounded rects from large to small.
function Draw-Shadow($g, [float]$x, [float]$y, [float]$w, [float]$h, [int]$spread, [int]$dy) {
    for ($i = $spread; $i -ge 1; $i--) {
        $a = [int](90 / $spread)
        if ($a -lt 4) { $a = 4 }
        $col   = [System.Drawing.Color]::FromArgb($a, 0, 0, 0)
        $brush = New-Object System.Drawing.SolidBrush($col)
        $path  = New-RoundedPath ($x - $i) ($y - $i + $dy) ($w + 2*$i) ($h + 2*$i) (10 + $i)
        $g.FillPath($brush, $path)
        $path.Dispose(); $brush.Dispose()
    }
}

function Draw-Shot($g, $img, [float]$x, [float]$y, [float]$r) {
    $clip = New-RoundedPath $x $y $img.Width $img.Height $r
    $g.SetClip($clip)
    $g.DrawImage($img, $x, $y, [float]$img.Width, [float]$img.Height)
    $g.ResetClip()
    # thin border to separate from wallpaper
    $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(60, 255, 255, 255), 1)
    $g.DrawPath($pen, $clip)
    $pen.Dispose(); $clip.Dispose()
}

# Light shot on top-left, dark shot below-right with offset, soft shadows, over the wallpaper.
function Compose-Preview([string]$topLightPng, [string]$bottomDarkPng, [string]$outPng) {
    $light = [System.Drawing.Image]::FromFile($topLightPng)
    $dark  = [System.Drawing.Image]::FromFile($bottomDarkPng)
    $wall  = [System.Drawing.Image]::FromFile($Wallpaper)

    $shotW = [Math]::Max($light.Width,  $dark.Width)
    $shotH = [Math]::Max($light.Height, $dark.Height)
    $canvasW = $shotW + $OffsetX + 2 * $Pad
    $canvasH = $shotH + $OffsetY + 2 * $Pad

    $canvas = New-Object System.Drawing.Bitmap($canvasW, $canvasH, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($canvas)
    $g.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.PixelOffsetMode   = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality

    # wallpaper: cover scale, centered
    $scale = [Math]::Max($canvasW / $wall.Width, $canvasH / $wall.Height)
    $bw = [int]([Math]::Ceiling($wall.Width  * $scale))
    $bh = [int]([Math]::Ceiling($wall.Height * $scale))
    $g.DrawImage($wall, [int](($canvasW - $bw) / 2), [int](($canvasH - $bh) / 2), $bw, $bh)

    $lightX = $Pad;            $lightY = $Pad
    $darkX  = $Pad + $OffsetX; $darkY  = $Pad + $OffsetY
    Draw-Shadow $g $darkX  $darkY  $dark.Width  $dark.Height 16 14
    Draw-Shot   $g $dark   $darkX  $darkY  4
    Draw-Shadow $g $lightX $lightY $light.Width $light.Height 16 14
    Draw-Shot   $g $light  $lightX $lightY 4
    $g.Dispose()

    $tmp = [System.IO.Path]::ChangeExtension($outPng, '.tmp.png')
    $canvas.Save($tmp, [System.Drawing.Imaging.ImageFormat]::Png)
    $canvas.Dispose(); $light.Dispose(); $dark.Dispose(); $wall.Dispose()
    Move-Item -LiteralPath $tmp -Destination $outPng -Force
    Write-Host "Done: $outPng ($canvasW x $canvasH)"
}

# === video-compare ===

# Posts a key press. SDL takes the key from the scan code in lParam, so it must be real;
# arrows and Page Up are extended keys.
function Send-Key([IntPtr]$hwnd, [int]$vk, [bool]$extended = $false) {
    $l = 1 -bor ([Win.Api]::MapVirtualKey($vk, 0) -shl 16)
    if ($extended) { $l = $l -bor (1 -shl 24) }
    [Win.Api]::PostMessage($hwnd, 0x0100, [IntPtr]$vk, [IntPtr]$l) | Out-Null            # WM_KEYDOWN
    Start-Sleep -Milliseconds 40
    $lUp = [int64]$l -bor 3221225472
    [Win.Api]::PostMessage($hwnd, 0x0101, [IntPtr]$vk, [IntPtr]$lUp) | Out-Null          # WM_KEYUP
    Start-Sleep -Milliseconds 80
}

# The split follows the mouse. When the cursor leaves the window SDL reports its real
# position, so the window is put straight below the cursor: SDL then sees x = half width.
# The window stays off-screen and the cursor never enters it. Retried if the mouse moved.
function Set-SplitToCenter([IntPtr]$hwnd) {
    $client = New-Object Win.Api+RECT
    [Win.Api]::GetClientRect($hwnd, [ref]$client) | Out-Null
    $half = [int]($client.Right / 2)

    for ($try = 0; $try -lt 10; $try++) {
        $before = New-Object Win.Api+POINT
        [Win.Api]::GetCursorPos([ref]$before) | Out-Null
        $origin = New-Object Win.Api+POINT
        [Win.Api]::ClientToScreen($hwnd, [ref]$origin) | Out-Null
        $wr = New-Object Win.Api+RECT
        [Win.Api]::GetWindowRect($hwnd, [ref]$wr) | Out-Null

        $x = $wr.Left + ($before.X - $origin.X - $half)
        [Win.Api]::SetWindowPos($hwnd, [IntPtr]::Zero, $x, $OffscreenY, 0, 0, $SWP_MOVE_ONLY) | Out-Null
        Start-Sleep -Milliseconds 300
        # a move inside the window makes SDL track the mouse and report the leave position
        [Win.Api]::PostMessage($hwnd, 0x0200, [IntPtr]::Zero, [IntPtr]([int]($client.Bottom / 2) -shl 16 -bor $half)) | Out-Null
        Start-Sleep -Milliseconds 500

        $after = New-Object Win.Api+POINT
        [Win.Api]::GetCursorPos([ref]$after) | Out-Null
        if ($after.X -eq $before.X) { return }
    }
    Write-Warning 'Mouse kept moving: the split may be off-center.'
}

function Shoot-Compare([string]$outFile) {
    $cmd = [System.IO.File]::ReadAllText($CmdFile).Trim()
    if ($cmd -notmatch '^"([^"]+)"\s+(.+)$') { throw "Unexpected command: $cmd" }
    $exe = $Matches[1]; $argLine = $Matches[2]
    Write-Host "  $cmd"

    $shotDir = Join-Path $Work 'vc'
    Remove-Item -LiteralPath $shotDir -Recurse -Force -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Path $shotDir -Force | Out-Null

    # SDL hint: the new window must not take the focus, typed keys would reach video-compare
    $env:SDL_WINDOW_NO_ACTIVATION_WHEN_SHOWN = '1'
    $proc = Start-Process -FilePath $exe -ArgumentList $argLine -WorkingDirectory $shotDir -PassThru
    try {
        $hwnd = Wait-ProcessWindow $proc 'SDL_app'
        [Win.Api]::SetWindowPos($hwnd, [IntPtr]::Zero, 0, $OffscreenY, 0, 0, $SWP_MOVE_ONLY) | Out-Null
        Start-Sleep -Seconds 2

        # Seek while playing: seeking a paused video shows stale frames on one side
        $rest = $CompareAt
        while ($rest -ge 600) { Send-Key $hwnd 0x21 $true; $rest -= 600 }   # Page Up
        while ($rest -ge 15)  { Send-Key $hwnd 0x26 $true; $rest -= 15 }    # Up
        while ($rest -ge 1)   { Send-Key $hwnd 0x27 $true; $rest -= 1 }     # Right
        Start-Sleep -Seconds 3
        Send-Key $hwnd 0x20                                                 # Space: pause
        Start-Sleep -Seconds 1

        Set-SplitToCenter $hwnd
        Send-Key $hwnd 0x46                                                 # F: save frames and on-screen view

        # F writes both frames and the view; the view has "_osd_" in its name
        $shot = $null
        for ($t = 0; $t -lt 15000 -and -not $shot; $t += 250) {
            Start-Sleep -Milliseconds 250
            $shot = Get-ChildItem -LiteralPath $shotDir -Filter '*_osd_*.png' | Select-Object -First 1
        }
        if (-not $shot) { throw 'video-compare did not save the frame.' }
        Start-Sleep -Milliseconds 500   # file may still be written
        Copy-Item -LiteralPath $shot.FullName -Destination $outFile -Force
        Write-Host "Done: $outFile"

        Send-Key $hwnd 0x1B                                                 # Esc: quit
        if (-not $proc.WaitForExit(3000)) { $proc.Kill() }
    }
    finally {
        if (-not $proc.HasExited) { $proc.Kill() }
    }
}

# === Run ===
$ShotLightRU = Join-Path $Work 'light_ru.png'
$ShotDarkEN  = Join-Path $Work 'dark_en.png'
$ShotLightEN = Join-Path $Work 'light_en.png'
$ShotDarkRU  = Join-Path $Work 'dark_ru.png'

# The preview writes ini (theme, files) and the cache (resolutions): both are restored
$iniBackup   = [System.IO.File]::ReadAllBytes($Ini)
$cacheBackup = [System.IO.File]::ReadAllBytes($Cache)
Build-PreviewScript
try {
    if ($Only -ne 'Compare') {
        # file names follow the readme, not the window language: both shots of a readme match
        Write-Host 'Capturing light/RU...'; Shoot-Launcher 'Light' 'Russian' $ShotLightRU
        Write-Host 'Capturing dark/EN...';  Shoot-Launcher 'Dark'  'English' $ShotDarkEN
        Write-Host 'Capturing light/EN...'; Shoot-Launcher 'Light' 'English' $ShotLightEN $ShownFrom $ShownTo
        Write-Host 'Capturing dark/RU...';  Shoot-Launcher 'Dark'  'Russian' $ShotDarkRU  $ShownFrom $ShownTo
    }
    else {
        # only the command is needed: start the window, take the file, close it
        Stop-Gracefully (Start-Launcher 'Light' 'English')
    }
}
finally {
    Remove-Item -LiteralPath $PreviewAu3 -Force -ErrorAction SilentlyContinue
    [System.IO.File]::WriteAllBytes($Ini, $iniBackup)
    [System.IO.File]::WriteAllBytes($Cache, $cacheBackup)
}

if ($Only -ne 'Compare') {
    Write-Host 'Composing previews...'
    Compose-Preview $ShotLightRU $ShotDarkEN $OutPng
    Compose-Preview $ShotLightEN $ShotDarkRU $OutPngEn
}
if ($Only -ne 'Launcher') {
    Write-Host 'Capturing video-compare...'
    Shoot-Compare $OutCompare
}

Remove-Item -LiteralPath $Work -Recurse -Force -ErrorAction SilentlyContinue
