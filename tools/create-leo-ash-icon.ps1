$ErrorActionPreference = "Stop"

$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$assetDir = Join-Path $root "KeywordOcr.App\Assets"
New-Item -ItemType Directory -Force -Path $assetDir | Out-Null

Add-Type -AssemblyName System.Drawing

function New-PointArray($points) {
    $arr = New-Object "System.Drawing.PointF[]" $points.Count
    for ($i = 0; $i -lt $points.Count; $i++) {
        $arr[$i] = New-Object System.Drawing.PointF $points[$i][0], $points[$i][1]
    }
    return $arr
}

function New-LeoAshPngBytes([int]$size) {
    $bmp = New-Object System.Drawing.Bitmap $size, $size, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $g.Clear([System.Drawing.Color]::Transparent)

    $scale = $size / 256.0
    function S([float]$value) { return [float]($value * $scale) }

    $black = [System.Drawing.Color]::FromArgb(255, 24, 23, 22)
    $beige = [System.Drawing.Color]::FromArgb(255, 205, 196, 185)
    $line = [System.Drawing.Color]::FromArgb(255, 50, 48, 45)
    $white = [System.Drawing.Color]::FromArgb(255, 250, 249, 247)

    $beigeBrush = New-Object System.Drawing.SolidBrush $beige
    $blackBrush = New-Object System.Drawing.SolidBrush $black
    $whiteBrush = New-Object System.Drawing.SolidBrush $white
    $eyeBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 35, 34, 32))
    $lightEyeBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 245, 244, 240))
    $thinPen = New-Object System.Drawing.Pen $line, (S 2)
    $arcPen = New-Object System.Drawing.Pen $line, (S 4)
    $lightPen = New-Object System.Drawing.Pen $white, (S 2)

    $g.DrawArc($arcPen, (S 22), (S 32), (S 212), (S 145), 196, 148)

    $leftEar1 = New-PointArray @(@((S 78),(S 84)), @((S 91),(S 51)), @((S 103),(S 85)))
    $leftEar2 = New-PointArray @(@((S 119),(S 83)), @((S 135),(S 54)), @((S 139),(S 91)))
    $rightEar1 = New-PointArray @(@((S 132),(S 87)), @((S 145),(S 58)), @((S 158),(S 88)))
    $rightEar2 = New-PointArray @(@((S 174),(S 87)), @((S 191),(S 61)), @((S 194),(S 96)))

    $g.FillPolygon($beigeBrush, $leftEar1)
    $g.FillPolygon($beigeBrush, $leftEar2)
    $g.FillEllipse($beigeBrush, (S 58), (S 116), (S 82), (S 56))
    $g.FillEllipse($whiteBrush, (S 72), (S 72), (S 68), (S 58))
    $g.FillEllipse($beigeBrush, (S 70), (S 72), (S 34), (S 48))
    $g.DrawPolygon($thinPen, $leftEar1)
    $g.DrawPolygon($thinPen, $leftEar2)
    $g.DrawEllipse($thinPen, (S 72), (S 72), (S 68), (S 58))

    $g.FillPolygon($blackBrush, $rightEar1)
    $g.FillPolygon($blackBrush, $rightEar2)
    $g.FillEllipse($blackBrush, (S 116), (S 116), (S 84), (S 58))
    $g.FillEllipse($blackBrush, (S 124), (S 75), (S 70), (S 56))
    $g.DrawPolygon($thinPen, $rightEar1)
    $g.DrawPolygon($thinPen, $rightEar2)

    $g.FillEllipse($eyeBrush, (S 92), (S 97), (S 5), (S 4))
    $g.FillEllipse($eyeBrush, (S 118), (S 97), (S 5), (S 4))
    $g.DrawArc($thinPen, (S 101), (S 101), (S 14), (S 10), 20, 140)
    $g.DrawLine($thinPen, (S 108), (S 106), (S 108), (S 111))

    $g.FillEllipse($lightEyeBrush, (S 145), (S 98), (S 5), (S 4))
    $g.FillEllipse($lightEyeBrush, (S 171), (S 98), (S 5), (S 4))
    $g.DrawArc($lightPen, (S 154), (S 103), (S 13), (S 9), 20, 140)
    $g.DrawLine($lightPen, (S 160), (S 107), (S 160), (S 112))

    $fontSize = [Math]::Max(8, [int](22 * $scale))
    $font = New-Object System.Drawing.Font "Segoe UI", $fontSize, ([System.Drawing.FontStyle]::Regular), ([System.Drawing.GraphicsUnit]::Pixel)
    $format = New-Object System.Drawing.StringFormat
    $format.Alignment = [System.Drawing.StringAlignment]::Center
    $textBrush = New-Object System.Drawing.SolidBrush $line
    $g.DrawString("leo & ash", $font, $textBrush, (New-Object System.Drawing.RectangleF (S 0), (S 188), (S 256), (S 42)), $format)

    $ms = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $bytes = $ms.ToArray()

    $g.Dispose()
    $bmp.Dispose()
    $ms.Dispose()
    return $bytes
}

$pngPath = Join-Path $assetDir "leo-ash-icon.png"
[System.IO.File]::WriteAllBytes($pngPath, (New-LeoAshPngBytes 512))

$sizes = @(256, 128, 64, 48, 32, 16)
$images = @()
foreach ($size in $sizes) {
    $images += ,@($size, (New-LeoAshPngBytes $size))
}

$icoPath = Join-Path $assetDir "leo-ash.ico"
$fs = [System.IO.File]::Create($icoPath)
$bw = New-Object System.IO.BinaryWriter $fs
$bw.Write([UInt16]0)
$bw.Write([UInt16]1)
$bw.Write([UInt16]$images.Count)

$offset = 6 + (16 * $images.Count)
foreach ($image in $images) {
    $size = [int]$image[0]
    $bytes = [byte[]]$image[1]
    $dimension = if ($size -eq 256) { 0 } else { $size }
    $bw.Write([byte]$dimension)
    $bw.Write([byte]$dimension)
    $bw.Write([byte]0)
    $bw.Write([byte]0)
    $bw.Write([UInt16]1)
    $bw.Write([UInt16]32)
    $bw.Write([UInt32]$bytes.Length)
    $bw.Write([UInt32]$offset)
    $offset += $bytes.Length
}

foreach ($image in $images) {
    $bw.Write([byte[]]$image[1])
}

$bw.Close()
$fs.Close()

Write-Host "Icon: $icoPath"
Write-Host "Preview: $pngPath"
