$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing

$root = Join-Path $PSScriptRoot "..\com.pricop.powerpoint-tools.sdPlugin\imgs\plugin"
New-Item -ItemType Directory -Force -Path $root | Out-Null

function New-PluginIcon([int]$size, [string]$path) {
    $bmp = New-Object System.Drawing.Bitmap $size, $size
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.Clear([System.Drawing.Color]::Transparent)

    $accent = [System.Drawing.Color]::FromArgb(19, 183, 166)
    $bg = New-Object System.Drawing.SolidBrush $accent
    $margin = [int]($size * 0.04)
    $radius = [int]($size * 0.18)

    $rect = New-Object System.Drawing.Rectangle $margin, $margin, ($size - 2*$margin), ($size - 2*$margin)
    $pathShape = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = 2 * $radius
    $pathShape.AddArc($rect.Left, $rect.Top, $d, $d, 180, 90)
    $pathShape.AddArc($rect.Right-$d, $rect.Top, $d, $d, 270, 90)
    $pathShape.AddArc($rect.Right-$d, $rect.Bottom-$d, $d, $d, 0, 90)
    $pathShape.AddArc($rect.Left, $rect.Bottom-$d, $d, $d, 90, 90)
    $pathShape.CloseFigure()
    $g.FillPath($bg, $pathShape)

    $fontSize = [single]($size * 0.47)
    $font = New-Object System.Drawing.Font("Segoe UI", $fontSize, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment = [System.Drawing.StringAlignment]::Center
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
    $textRect = New-Object System.Drawing.RectangleF 0, (-$size*0.025), $size, $size
    $g.DrawString("P", $font, [System.Drawing.Brushes]::White, $textRect, $sf)

    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)

    $sf.Dispose()
    $font.Dispose()
    $pathShape.Dispose()
    $bg.Dispose()
    $g.Dispose()
    $bmp.Dispose()
}

New-PluginIcon 256 (Join-Path $root "marketplace.png")
New-PluginIcon 512 (Join-Path $root "marketplace@2x.png")
Write-Host "Stream Deck plugin icons generated."
