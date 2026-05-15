$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$catalogPath = Join-Path $root "catalog"
$manifestPath = Join-Path $root "manifest.xml"
$shareName = "LogoAddToolCatalog"
$sharePath = "\\localhost\$shareName"
$catalogId = "{8D74D909-117D-4D68-A4F7-35F0F041AD87}"
$registryPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$catalogId"

Set-Location $root

Write-Host "Logo Add Tool installer"
Write-Host "1/4 Checking Node.js..."
if ($null -eq (Get-Command node -ErrorAction SilentlyContinue)) {
  throw "Node.js was not found. Install Node.js 20 or later first: https://nodejs.org/"
}

Write-Host "2/4 Installing trusted localhost certificate..."
npx --yes office-addin-dev-certs install

Write-Host "3/4 Preparing manifest catalog..."
New-Item -ItemType Directory -Force -Path $catalogPath | Out-Null
Copy-Item -LiteralPath $manifestPath -Destination (Join-Path $catalogPath "manifest.xml") -Force

Write-Host "4/4 Trying to register Windows trusted add-in catalog..."
$shareReady = $false
if ($null -ne (Get-Command Get-SmbShare -ErrorAction SilentlyContinue)) {
  $existingShare = Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue
  if ($null -ne $existingShare) {
    $shareReady = $true
  } else {
    try {
      New-SmbShare -Name $shareName -Path $catalogPath -ReadAccess "Everyone" | Out-Null
      $shareReady = $true
    } catch {
      Write-Warning "Could not create SMB share automatically. Run this script as Administrator, or follow docs/user-installation.md."
      Write-Warning $_.Exception.Message
    }
  }
}

if ($shareReady) {
  New-Item -Path $registryPath -Force | Out-Null
  New-ItemProperty -Path $registryPath -Name "Id" -Value $catalogId -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $registryPath -Name "Url" -Value $sharePath -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $registryPath -Name "Flags" -Value 1 -PropertyType DWord -Force | Out-Null
  Write-Host "Trusted catalog registered: $sharePath"
} else {
  Write-Host "Manifest copied to: $catalogPath"
  Write-Host "Trusted catalog was not registered automatically."
}

Write-Host ""
Write-Host "Install step finished."
Write-Host "Next: run start-windows.bat, restart PowerPoint, then add Logo Add Tool from Home > Add-ins."
