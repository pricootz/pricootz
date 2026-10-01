$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing

$root = Join-Path $PSScriptRoot "..\com.pricop.powerpoint-tools.sdPlugin\imgs\plugin"
New-Item -ItemType Directory -Force -Path $root | Out-Null

function New-RoundedPath([System.Drawing.RectangleF]$rect, [float]$radius) {
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $radius * 2
    $path.AddArc($rect.Left, $rect.Top, $d, $d, 180, 90)
    $path.AddArc($rect.Right-$d, $rect.Top, $d, $d, 270, 90)
    $path.AddArc($rect.Right-$d, $rect.Bottom-$d, $d, $d, 0, 90)
    $path.AddArc($rect.Left, $rect.Bottom-$d, $d, $d, 90, 90)
    $path.CloseFigure()
    return $path
}

function New-PluginIcon([int]$size, [string]$path) {
    $bmp = New-Object System.Drawing.Bitmap $size, $size
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.Clear([System.Drawing.Color]::Transparent)

    $margin = [single]($size * 0.04)
    $rect = New-Object System.Drawing.RectangleF $margin, $margin, ($size - 2*$margin), ($size - 2*$margin)
    $radius = [single]($size * 0.18)

    $bgPath = New-RoundedPath $rect $radius
    $bgBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        $rect,
        [System.Drawing.Color]::FromArgb(24,43,55),
        [System.Drawing.Color]::FromArgb(8,17,22),
        45
    )
    $g.FillPath($bgBrush, $bgPath)

    $border = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(50,32,224,194)), ([single]($size*0.012))
    $g.DrawPath($border, $bgPath)

    # Duotone PowerPoint-style P mark
    $white = [System.Drawing.Color]::FromArgb(248,250,252)
    $teal = [System.Drawing.Color]::FromArgb(32,224,194)

    $penW = [single]($size * 0.055)
    $penWhite = New-Object System.Drawing.Pen $white, $penW
    $penWhite.StartCap = $penWhite.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $penWhite.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round

    $penTeal = New-Object System.Drawing.Pen $teal, $penW
    $penTeal.StartCap = $penTeal.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $penTeal.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round

    $x = [single]($size * 0.34)
    $y1 = [single]($size * 0.28)
    $y2 = [single]($size * 0.72)
    $g.DrawLine($penWhite, $x, $y1, $x, $y2)

    $loopRect = New-Object System.Drawing.RectangleF ([single]($size*0.34)), ([single]($size*0.29)), ([single]($size*0.34)), ([single]($size*0.25))
    $loopPath = New-RoundedPath $loopRect ([single]($size*0.07))
    $g.DrawPath($penTeal, $loopPath)

    $dotBrush = New-Object System.Drawing.SolidBrush $teal
    $g.FillEllipse($dotBrush, [single]($size*0.62), [single]($size*0.63), [single]($size*0.10), [single]($size*0.10))

    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)

    $dotBrush.Dispose()
    $loopPath.Dispose()
    $penTeal.Dispose()
    $penWhite.Dispose()
    $border.Dispose()
    $bgBrush.Dispose()
    $bgPath.Dispose()
    $g.Dispose()
    $bmp.Dispose()
}

New-PluginIcon 256 (Join-Path $root "marketplace.png")
New-PluginIcon 512 (Join-Path $root "marketplace@2x.png")
Write-Host "Duotone plugin icons generated."
