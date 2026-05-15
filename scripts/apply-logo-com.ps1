param(
  [Parameter(Mandatory = $true)]
  [string]$PayloadPath
)

$ErrorActionPreference = "Stop"
$utf8NoBom = New-Object System.Text.UTF8Encoding -ArgumentList $false
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom

function Write-JsonResult($Value) {
  $Value | ConvertTo-Json -Depth 6 -Compress
}

function Set-StampMetadata($Shape, [string]$Name, [string]$LogoId, [string]$PlacementId) {
  $Shape.Name = $Name
  $Shape.AlternativeText = "PPT Logo Stamp|$PlacementId|$LogoId"
  try {
    $Shape.Tags.Add("brandStamp", "true")
    $Shape.Tags.Add("logoId", $LogoId)
    $Shape.Tags.Add("placementId", $PlacementId)
  } catch {
    # Tags are best-effort; shape name remains the primary marker.
  }
}

function Set-ShapePlacement($Shape, [single]$Left, [single]$Top, [single]$Width, [single]$Height) {
  try {
    $Shape.LockAspectRatio = 0
  } catch {
    # Some shape types may not expose LockAspectRatio consistently.
  }

  $Shape.Width = $Width
  $Shape.Height = $Height
  $Shape.Left = $Left
  $Shape.Top = $Top
}

function Test-ExistingStampForPlacement($Shape, [string]$PlacementId, [string]$TemporaryShapePrefix) {
  if ($Shape.Name -like "$TemporaryShapePrefix*") {
    return $false
  }

  $parts = [string]$Shape.Name -split "\|"
  if ($parts.Length -lt 3 -or $parts[0] -ne "BrandLogoStamp") {
    return $false
  }

  return $parts[1] -eq $PlacementId
}

$payload = Get-Content -LiteralPath $PayloadPath -Raw -Encoding UTF8 | ConvertFrom-Json
$dataUrl = [string]$payload.asset.data
$commaIndex = $dataUrl.IndexOf(",")
$base64 = if ($commaIndex -ge 0) { $dataUrl.Substring($commaIndex + 1) } else { $dataUrl }

$extension = switch ([string]$payload.asset.mimeType) {
  "image/jpeg" { ".jpg" }
  "image/webp" { ".webp" }
  default { ".png" }
}

$logoPath = Join-Path ([System.IO.Path]::GetTempPath()) ("brand-logo-stamp-" + [Guid]::NewGuid().ToString("N") + $extension)
[System.IO.File]::WriteAllBytes($logoPath, [Convert]::FromBase64String($base64))

try {
  try {
    $powerPoint = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
  } catch {
    throw "PowerPoint is not running. Open the target presentation first."
  }

  $presentation = $powerPoint.ActivePresentation
  if ($null -eq $presentation) {
    throw "No active PowerPoint presentation was found."
  }

  $placement = $payload.placement
  $placementId = [string]$placement.id
  $logoId = [string]$payload.asset.id
  $shapeName = "BrandLogoStamp|$placementId|$logoId"
  $shapePrefix = "BrandLogoStamp|$placementId|"
  $temporaryPlacementId = "temp-" + [Guid]::NewGuid().ToString("N")
  $temporaryShapeName = "BrandLogoStamp|$temporaryPlacementId|$logoId"
  $temporaryShapePrefix = "BrandLogoStamp|$temporaryPlacementId|"

  $left = [single]$placement.left
  $top = [single]$placement.top
  $width = [single]$placement.width
  $height = [single]$placement.height

  $activeSlideIndex = $null
  try {
    $activeSlideIndex = $powerPoint.ActiveWindow.View.Slide.SlideIndex
  } catch {
    $activeSlideIndex = $null
  }

  $slideCount = $presentation.Slides.Count
  try {
    for ($slideIndex = 1; $slideIndex -le $slideCount; $slideIndex++) {
      $slide = $presentation.Slides.Item($slideIndex)
      $inserted = $slide.Shapes.AddPicture($logoPath, 0, -1, $left, $top, -1, -1)
      Set-ShapePlacement $inserted $left $top $width $height
      Set-StampMetadata $inserted $temporaryShapeName $logoId $temporaryPlacementId
    }
  } catch {
    for ($slideIndex = 1; $slideIndex -le $slideCount; $slideIndex++) {
      $slide = $presentation.Slides.Item($slideIndex)
      for ($shapeIndex = $slide.Shapes.Count; $shapeIndex -ge 1; $shapeIndex--) {
        $shape = $slide.Shapes.Item($shapeIndex)
        if ($shape.Name -like "$temporaryShapePrefix*") {
          $shape.Delete()
        }
      }
    }
    throw
  }

  for ($slideIndex = 1; $slideIndex -le $slideCount; $slideIndex++) {
    $slide = $presentation.Slides.Item($slideIndex)

    for ($shapeIndex = $slide.Shapes.Count; $shapeIndex -ge 1; $shapeIndex--) {
      $shape = $slide.Shapes.Item($shapeIndex)
      if (Test-ExistingStampForPlacement $shape $placementId $temporaryShapePrefix) {
        $shape.Delete()
      }
    }

    for ($shapeIndex = $slide.Shapes.Count; $shapeIndex -ge 1; $shapeIndex--) {
      $shape = $slide.Shapes.Item($shapeIndex)
      if ($shape.Name -like "$temporaryShapePrefix*") {
        Set-ShapePlacement $shape $left $top $width $height
        Set-StampMetadata $shape $shapeName $logoId $placementId
      }
    }
  }

  Start-Sleep -Milliseconds 120

  for ($slideIndex = 1; $slideIndex -le $slideCount; $slideIndex++) {
    $slide = $presentation.Slides.Item($slideIndex)
    for ($shapeIndex = 1; $shapeIndex -le $slide.Shapes.Count; $shapeIndex++) {
      $shape = $slide.Shapes.Item($shapeIndex)
      if ($shape.Name -eq $shapeName) {
        Set-ShapePlacement $shape $left $top $width $height
      }
    }
  }

  if ($null -ne $activeSlideIndex) {
    try {
      $powerPoint.ActiveWindow.View.GotoSlide($activeSlideIndex)
    } catch {
      # Restoring view is best-effort.
    }
  }

  Write-JsonResult @{
    ok = $true
    slideCount = $slideCount
  }
} finally {
  if (Test-Path -LiteralPath $logoPath) {
    Remove-Item -LiteralPath $logoPath -Force -ErrorAction SilentlyContinue
  }
}
