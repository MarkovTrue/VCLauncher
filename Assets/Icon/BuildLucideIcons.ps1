<#
    BuildLucideIcons.ps1 - renders Lucide icons (ISC license) to neutral-gray PNG.

    Icon set: Lucide https://lucide.dev  (lucide-static, pinned version below).
    Source SVGs are fetched from unpkg at build time; strokes are rendered in a
    single neutral gray (#909090) so one PNG reads on both light and dark themes.

    Mapping:
      settings       -> Assets/Icon/Settings.png            (settings button)
      arrow-down-up  -> Assets/Icon/Swap.png                (swap button)
      columns-2      -> Assets/CompareIcons/CompareDirect.png   (direct compare)
      rows-2         -> Assets/CompareIcons/CompareVstack.png   (vertical compare)

    Run:  powershell -ExecutionPolicy Bypass -File Assets\Icon\BuildLucideIcons.ps1
#>
param(
    [string]$Version = '1.17.0',
    [int]   $Size    = 64,
    [string]$Gray    = '909090'
)

# RenderTargetBitmap requires an STA thread.
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    & powershell.exe -NoProfile -Sta -ExecutionPolicy Bypass -File $PSCommandPath `
        -Version $Version -Size $Size -Gray $Gray
    exit $LASTEXITCODE
}

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase, System.Xaml

$IconDir    = Split-Path -Parent $MyInvocation.MyCommand.Path        # Assets/Icon
$AssetsDir  = Split-Path -Parent $IconDir                            # Assets
$CompareDir = Join-Path $AssetsDir 'CompareIcons'

# sw = stroke-width override (null = use the SVG's own). Compare icons чуть тоньше.
$map = @(
    @{ name = 'settings';      out = (Join-Path $IconDir    'Settings.png')      ; sw = $null },
    @{ name = 'arrow-down-up'; out = (Join-Path $IconDir    'Swap.png')          ; sw = $null },
    @{ name = 'columns-2';     out = (Join-Path $CompareDir 'CompareDirect.png') ; sw = 1.5  },
    @{ name = 'rows-2';        out = (Join-Path $CompareDir 'CompareVstack.png') ; sw = 1.5  }
)

$rgb   = [System.Windows.Media.Color]::FromRgb(
            [Convert]::ToInt32($Gray.Substring(0,2),16),
            [Convert]::ToInt32($Gray.Substring(2,2),16),
            [Convert]::ToInt32($Gray.Substring(4,2),16))
$brush = New-Object System.Windows.Media.SolidColorBrush($rgb)
$brush.Freeze()

function Get-Attr([string]$attrs, [string]$name) {
    $m = [regex]::Match($attrs, $name + '="([\d.\-]+)"')
    if ($m.Success) { return [double]$m.Groups[1].Value }
    return $null
}

function Render-Svg([string]$svg, [string]$outPng, $swOverride) {
    $vb = [regex]::Match($svg, 'viewBox="0 0 ([\d.]+) ([\d.]+)"')
    $vw = if ($vb.Success) { [double]$vb.Groups[1].Value } else { 24.0 }
    $sw = 2.0
    $m = [regex]::Match($svg, 'stroke-width="([\d.]+)"'); if ($m.Success) { $sw = [double]$m.Groups[1].Value }
    if ($null -ne $swOverride) { $sw = [double]$swOverride }

    $pen = New-Object System.Windows.Media.Pen($brush, $sw)
    $pen.StartLineCap = [System.Windows.Media.PenLineCap]::Round
    $pen.EndLineCap   = [System.Windows.Media.PenLineCap]::Round
    $pen.LineJoin     = [System.Windows.Media.PenLineJoin]::Round

    $geos = New-Object System.Collections.Generic.List[System.Windows.Media.Geometry]
    foreach ($mm in [regex]::Matches($svg, '<path\b[^>]*\bd="([^"]+)"')) {
        $geos.Add([System.Windows.Media.Geometry]::Parse($mm.Groups[1].Value))
    }
    foreach ($mm in [regex]::Matches($svg, '<rect\b([^>]*)>')) {
        $a = $mm.Groups[1].Value
        $x = Get-Attr $a '\bx'; $y = Get-Attr $a '\by'
        $w = Get-Attr $a 'width'; $h = Get-Attr $a 'height'
        $rx = Get-Attr $a '\brx'; if ($null -eq $rx) { $rx = 0 }
        $rect = New-Object System.Windows.Rect($x, $y, $w, $h)
        $geos.Add((New-Object System.Windows.Media.RectangleGeometry($rect, $rx, $rx)))
    }
    foreach ($mm in [regex]::Matches($svg, '<circle\b([^>]*)>')) {
        $a = $mm.Groups[1].Value
        $cx = Get-Attr $a 'cx'; $cy = Get-Attr $a 'cy'; $r = Get-Attr $a '\br'
        $geos.Add((New-Object System.Windows.Media.EllipseGeometry((New-Object System.Windows.Point($cx, $cy)), $r, $r)))
    }
    foreach ($mm in [regex]::Matches($svg, '<line\b([^>]*)>')) {
        $a = $mm.Groups[1].Value
        $p1 = New-Object System.Windows.Point((Get-Attr $a 'x1'), (Get-Attr $a 'y1'))
        $p2 = New-Object System.Windows.Point((Get-Attr $a 'x2'), (Get-Attr $a 'y2'))
        $geos.Add((New-Object System.Windows.Media.LineGeometry($p1, $p2)))
    }

    $dv = New-Object System.Windows.Media.DrawingVisual
    $dc = $dv.RenderOpen()
    $scale = $Size / $vw
    $dc.PushTransform((New-Object System.Windows.Media.ScaleTransform($scale, $scale)))
    foreach ($g in $geos) { $dc.DrawGeometry($null, $pen, $g) }
    $dc.Pop()
    $dc.Close()

    $rtb = New-Object System.Windows.Media.Imaging.RenderTargetBitmap($Size, $Size, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
    $rtb.Render($dv)
    $enc = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
    $enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb))
    $fs = [System.IO.File]::Open($outPng, [System.IO.FileMode]::Create)
    try { $enc.Save($fs) } finally { $fs.Close() }
}

foreach ($it in $map) {
    $url = "https://unpkg.com/lucide-static@$Version/icons/$($it.name).svg"
    $svg = (Invoke-WebRequest -UseBasicParsing -Uri $url).Content
    Render-Svg $svg $it.out $it.sw
    Write-Host ("  {0,-14} -> {1}" -f $it.name, $it.out)
}
Write-Host "Lucide icons rendered ($Size px, #$Gray), version $Version"
