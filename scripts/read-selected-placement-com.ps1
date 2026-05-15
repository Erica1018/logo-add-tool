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

function New-PlacementId {
  "placement-" + [Guid]::NewGuid().ToString("N")
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

$payload = Get-Content -LiteralPath $PayloadPath -Raw -Encoding UTF8 | ConvertFrom-Json
$activeLogoId = [string]$payload.activeLogoId

try {
  $powerPoint = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
} catch {
  throw "PowerPoint is not running. Open the target presentation first."
}

try {
  $shapeRange = $powerPoint.ActiveWindow.Selection.ShapeRange
} catch {
  throw "Select exactly one logo shape on the current slide first."
}

if ($shapeRange.Count -ne 1) {
  throw "Select exactly one logo shape on the current slide first."
}

$shape = $shapeRange.Item(1)
$placementId = $null
$logoId = $null

$parts = [string]$shape.Name -split "\|"
if ($parts.Length -ge 3 -and $parts[0] -eq "BrandLogoStamp") {
  $placementId = $parts[1]
  $logoId = $parts[2]
}

if ([string]::IsNullOrWhiteSpace($placementId)) {
  $placementId = New-PlacementId
}
if ([string]::IsNullOrWhiteSpace($logoId)) {
  $logoId = $activeLogoId
}
if ([string]::IsNullOrWhiteSpace($logoId)) {
  throw "The selected shape is not a logo inserted by this add-in. Insert a logo from the task pane first."
}

$shapeName = "BrandLogoStamp|$placementId|$logoId"
Set-StampMetadata $shape $shapeName $logoId $placementId

Write-JsonResult @{
  id = $placementId
  logoId = $logoId
  left = [math]::Round([double]$shape.Left, 2)
  top = [math]::Round([double]$shape.Top, 2)
  width = [math]::Round([double]$shape.Width, 2)
  height = [math]::Round([double]$shape.Height, 2)
  scope = "allSlides"
  createdAt = (Get-Date).ToUniversalTime().ToString("o")
}
