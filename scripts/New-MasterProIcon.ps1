param(
    [string]$SourcePath = (Join-Path $PSScriptRoot "..\installer\assets\app_icon.png"),
    [string]$OutputPath = (Join-Path $PSScriptRoot "..\installer\assets\MasterPRO.ico")
)

$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Drawing

function Get-AlphaBounds {
    param([System.Drawing.Bitmap]$Bitmap)

    $minX = $Bitmap.Width
    $minY = $Bitmap.Height
    $maxX = -1
    $maxY = -1

    for ($y = 0; $y -lt $Bitmap.Height; $y++) {
        for ($x = 0; $x -lt $Bitmap.Width; $x++) {
            if ($Bitmap.GetPixel($x, $y).A -gt 0) {
                if ($x -lt $minX) { $minX = $x }
                if ($y -lt $minY) { $minY = $y }
                if ($x -gt $maxX) { $maxX = $x }
                if ($y -gt $maxY) { $maxY = $y }
            }
        }
    }

    if ($maxX -lt 0 -or $maxY -lt 0) {
        return [System.Drawing.Rectangle]::FromLTRB(0, 0, $Bitmap.Width, $Bitmap.Height)
    }

    return [System.Drawing.Rectangle]::FromLTRB($minX, $minY, $maxX + 1, $maxY + 1)
}

function New-IconFrameBytes {
    param(
        [System.Drawing.Bitmap]$SourceBitmap,
        [System.Drawing.Rectangle]$CropBounds,
        [int]$Size
    )

    $frame = New-Object System.Drawing.Bitmap $Size, $Size, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    try {
        $graphics = [System.Drawing.Graphics]::FromImage($frame)
        try {
            $graphics.Clear([System.Drawing.Color]::Transparent)
            $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
            $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
            $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias

            $contentSize = [Math]::Round($Size * 0.90)
            $offset = [Math]::Round(($Size - $contentSize) / 2.0)
            $destRect = New-Object System.Drawing.Rectangle $offset, $offset, $contentSize, $contentSize
            $graphics.DrawImage($SourceBitmap, $destRect, $CropBounds, [System.Drawing.GraphicsUnit]::Pixel)
        }
        finally {
            $graphics.Dispose()
        }

        $memory = New-Object System.IO.MemoryStream
        try {
            $frame.Save($memory, [System.Drawing.Imaging.ImageFormat]::Png)
            return ,$memory.ToArray()
        }
        finally {
            $memory.Dispose()
        }
    }
    finally {
        $frame.Dispose()
    }
}

$resolvedSourcePath = [System.IO.Path]::GetFullPath($SourcePath)
$resolvedOutputPath = [System.IO.Path]::GetFullPath($OutputPath)

if (-not (Test-Path -LiteralPath $resolvedSourcePath)) {
    throw "Source image not found: $resolvedSourcePath"
}

$outputDirectory = Split-Path -Path $resolvedOutputPath -Parent
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory | Out-Null
}

$sizes = @(16, 20, 24, 32, 40, 48, 64, 128, 256)
$sourceBitmap = [System.Drawing.Bitmap]::FromFile($resolvedSourcePath)

try {
    $cropBounds = Get-AlphaBounds -Bitmap $sourceBitmap
    $frames = foreach ($size in $sizes) {
        $frameBytes = [byte[]](New-IconFrameBytes -SourceBitmap $sourceBitmap -CropBounds $cropBounds -Size $size)
        [PSCustomObject]@{
            Size = $size
            Bytes = $frameBytes
        }
    }
}
finally {
    $sourceBitmap.Dispose()
}

$fileStream = [System.IO.File]::Create($resolvedOutputPath)
$writer = New-Object System.IO.BinaryWriter $fileStream

try {
    $writer.Write([UInt16]0)
    $writer.Write([UInt16]1)
    $writer.Write([UInt16]$frames.Count)

    $imageOffset = 6 + (16 * $frames.Count)
    foreach ($frame in $frames) {
        $dimension = if ($frame.Size -ge 256) { 0 } else { [byte]$frame.Size }
        $writer.Write([byte]$dimension)
        $writer.Write([byte]$dimension)
        $writer.Write([byte]0)
        $writer.Write([byte]0)
        $writer.Write([UInt16]1)
        $writer.Write([UInt16]32)
        $writer.Write([UInt32]$frame.Bytes.Length)
        $writer.Write([UInt32]$imageOffset)
        $imageOffset += $frame.Bytes.Length
    }

    foreach ($frame in $frames) {
        $writer.Write([byte[]]$frame.Bytes)
    }
}
finally {
    $writer.Dispose()
    $fileStream.Dispose()
}

Write-Output "Created icon: $resolvedOutputPath"
Write-Output ("Frames: " + (($sizes | ForEach-Object { "${_}x${_}" }) -join ", "))
