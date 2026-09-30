$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$assets = Join-Path $root "installer\assets"
New-Item -ItemType Directory -Force -Path $assets | Out-Null

Add-Type -AssemblyName System.Drawing
$bmp = New-Object System.Drawing.Bitmap 64,64
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.Clear([System.Drawing.Color]::Transparent)
$accent = [System.Drawing.Color]::FromArgb(19,183,166)
$brush = New-Object System.Drawing.SolidBrush $accent
$g.FillEllipse($brush,4,4,56,56)
$font = New-Object System.Drawing.Font("Segoe UI Semibold",28,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Pixel)
$sf = New-Object System.Drawing.StringFormat
$sf.Alignment = [System.Drawing.StringAlignment]::Center
$sf.LineAlignment = [System.Drawing.StringAlignment]::Center
$g.DrawString("P",$font,[System.Drawing.Brushes]::White,(New-Object System.Drawing.RectangleF(4,2,56,58)),$sf)

$pngPath = Join-Path $assets "app.png"
$icoPath = Join-Path $assets "app.ico"
$bmp.Save($pngPath,[System.Drawing.Imaging.ImageFormat]::Png)
$hIcon = $bmp.GetHicon()
$ico = [System.Drawing.Icon]::FromHandle($hIcon)
$fs = [IO.File]::Open($icoPath,[IO.FileMode]::Create)
$ico.Save($fs)
$fs.Close()
$ico.Dispose()
$font.Dispose()
$brush.Dispose()
$g.Dispose()
$bmp.Dispose()
Write-Host "Assets generated."
