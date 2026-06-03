<#
    BuildPreview.ps1 - generates the README preview image.

    Captures two GUI screenshots:
      - light theme + Russian language
      - dark theme  + English language
    Then composes them over the desktop wallpaper: dark under light,
    offset, each with a soft drop shadow. Result -> Assets/Preview.png.

    Run:  powershell -ExecutionPolicy Bypass -File BuildPreview.ps1
#>
param(
    [string]$Exe       = '',                  # compiled exe; empty = <root>\VCLauncher.exe
    [string]$AutoIt    = 'D:\Soft\AutoIt Script\AutoIt\AutoIt3.exe',  # fallback if no exe
    [string]$Wallpaper = '',                  # empty = autodetect in D:\Pictures
    [int]   $OffsetX   = 160,                  # dark shot horizontal offset from light
    [int]   $OffsetY   = 60,                    # dark shot vertical offset from light
    [int]   $Pad       = 42                    # wallpaper margin around the group
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

# Repo root: script may live in repo root or in Assets/ — locate by VCLauncher.au3.
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Root      = $ScriptDir
if (-not (Test-Path -LiteralPath (Join-Path $Root 'VCLauncher.au3'))) {
    $Root = Split-Path -Parent $ScriptDir
}
$Au3       = Join-Path $Root 'VCLauncher.au3'
$Ini       = Join-Path $Root 'VCLauncher.ini'
$ShotLightRU = Join-Path $Root '_shot_light_ru.png'
$ShotDarkEN  = Join-Path $Root '_shot_dark_en.png'
$ShotLightEN = Join-Path $Root '_shot_light_en.png'
$ShotDarkRU  = Join-Path $Root '_shot_dark_ru.png'
$OutPng      = Join-Path $Root 'Assets\Preview.png'     # RU readme: light RU over dark EN
$OutPngEn    = Join-Path $Root 'Assets\Preview.en.png'  # EN readme: light EN over dark RU

# Wallpaper file name contains Cyrillic - match by ASCII mask to avoid script-encoding issues.
if (-not $Wallpaper) {
    $Wallpaper = (Get-ChildItem -LiteralPath 'D:\Pictures' -Filter '2560*1440 1.png' |
        Select-Object -First 1).FullName
}
if (-not $Wallpaper -or -not (Test-Path -LiteralPath $Wallpaper)) {
    throw "Wallpaper not found. Pass -Wallpaper <png path>."
}
# Prefer the compiled exe; fall back to running the .au3 via AutoIt3.exe.
if (-not $Exe) { $Exe = Join-Path $Root 'VCLauncher.exe' }
$UseExe = Test-Path -LiteralPath $Exe
if (-not $UseExe) {
    if (-not (Test-Path -LiteralPath $AutoIt)) { throw "No exe and no AutoIt3.exe: $AutoIt" }
    if (-not (Test-Path -LiteralPath $Au3))    { throw "VCLauncher.au3 not found: $Au3" }
}

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
public static extern bool SetForegroundWindow(System.IntPtr hWnd);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool GetWindowRect(System.IntPtr hWnd, out RECT r);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool PrintWindow(System.IntPtr hWnd, System.IntPtr hdc, uint flags);
[System.Runtime.InteropServices.DllImport("dwmapi.dll")]
public static extern int DwmGetWindowAttribute(System.IntPtr hWnd, int attr, out RECT r, int size);
[System.Runtime.InteropServices.DllImport("kernel32.dll", CharSet=System.Runtime.InteropServices.CharSet.Unicode)]
public static extern bool WritePrivateProfileString(string section, string key, string val, string file);
public struct RECT { public int Left, Top, Right, Bottom; }
'@

$AUTOIT_CLASS         = 'AutoIt v3 GUI'
$PW_RENDERFULLCONTENT = 2
$DWMWA_EXTENDED_FRAME = 9

function Set-Config([string]$theme, [string]$lang) {
    [Win.Api]::WritePrivateProfileString('Settings', 'Theme',    $theme, $Ini) | Out-Null
    [Win.Api]::WritePrivateProfileString('Settings', 'Language', $lang,  $Ini) | Out-Null
    [Win.Api]::WritePrivateProfileString($null, $null, $null, $Ini) | Out-Null  # flush ini cache to disk
}

# Finds the GUI window of a specific process (by class + PID), so we never grab a
# leftover window from a previous, still-closing capture. FindWindow proved unreliable here.
function Find-GuiWindow([int]$targetPid) {
    $script:_found = [IntPtr]::Zero
    $cb = {
        param($h, $l)
        if ([Win.Api]::IsWindowVisible($h)) {
            $sb = New-Object System.Text.StringBuilder 256
            [Win.Api]::GetClassName($h, $sb, 256) | Out-Null
            if ($sb.ToString() -eq $AUTOIT_CLASS) {
                $wpid = [uint32]0
                [Win.Api]::GetWindowThreadProcessId($h, [ref]$wpid) | Out-Null
                if ($wpid -eq $targetPid) { $script:_found = $h; return $false }
            }
        }
        return $true
    }
    [Win.Api]::EnumWindows([Win.Api+EnumProc]$cb, [IntPtr]::Zero) | Out-Null
    return $script:_found
}

function Capture-Window([int]$targetPid, [string]$outFile) {
    # wait for this process's GUI window (class+PID — title collides with MsgBox)
    $hwnd = [IntPtr]::Zero
    for ($i = 0; $i -lt 100; $i++) {
        $hwnd = Find-GuiWindow $targetPid
        if ($hwnd -ne [IntPtr]::Zero) { break }
        Start-Sleep -Milliseconds 100
    }
    if ($hwnd -eq [IntPtr]::Zero) { throw "GUI window did not appear." }

    [Win.Api]::SetForegroundWindow($hwnd) | Out-Null
    Start-Sleep -Milliseconds 1800   # let header/theme/info labels render

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

function Shoot([string]$theme, [string]$lang, [string]$outFile) {
    Set-Config $theme $lang
    if ($UseExe) {
        $proc = Start-Process -FilePath $Exe -WorkingDirectory $Root -PassThru
    } else {
        $proc = Start-Process -FilePath $AutoIt -ArgumentList "`"$Au3`"" -WorkingDirectory $Root -PassThru
    }
    try {
        Capture-Window $proc.Id $outFile
    }
    finally {
        if (-not $proc.HasExited) {
            $proc.CloseMainWindow() | Out-Null
            Start-Sleep -Milliseconds 400
            if (-not $proc.HasExited) { $proc.Kill() }
        }
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

# === 1. Screenshots (4 combos) ===
$iniBackup = [System.IO.File]::ReadAllBytes($Ini)
try {
    Write-Host 'Capturing light/RU...'; Shoot 'Light' 'Russian' $ShotLightRU
    Write-Host 'Capturing dark/EN...';  Shoot 'Dark'  'English' $ShotDarkEN
    Write-Host 'Capturing light/EN...'; Shoot 'Light' 'English' $ShotLightEN
    Write-Host 'Capturing dark/RU...';  Shoot 'Dark'  'Russian' $ShotDarkRU
}
finally {
    [System.IO.File]::WriteAllBytes($Ini, $iniBackup)   # restore ini
}

# === 2. Composition ===
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

    New-Item -ItemType Directory -Path (Split-Path $outPng) -Force | Out-Null
    $tmp = [System.IO.Path]::ChangeExtension($outPng, '.tmp.png')
    $canvas.Save($tmp, [System.Drawing.Imaging.ImageFormat]::Png)
    Move-Item -LiteralPath $tmp -Destination $outPng -Force

    $canvas.Dispose(); $light.Dispose(); $dark.Dispose(); $wall.Dispose()
    Write-Host "Done: $outPng ($canvasW x $canvasH)"
}

Write-Host 'Composing previews...'
Compose-Preview $ShotLightRU $ShotDarkEN $OutPng    # RU readme: light RU over dark EN
Compose-Preview $ShotLightEN $ShotDarkRU $OutPngEn  # EN readme: light EN over dark RU

# remove temp screenshots (files unlocked after Dispose)
Remove-Item -LiteralPath $ShotLightRU, $ShotDarkEN, $ShotLightEN, $ShotDarkRU -Force -ErrorAction SilentlyContinue
