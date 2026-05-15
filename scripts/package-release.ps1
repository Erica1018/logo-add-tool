param(
  [string]$Version = (Get-Content -Encoding UTF8 "$PSScriptRoot\..\VERSION" | Select-Object -First 1)
)

$ErrorActionPreference = "Stop"

$root = Resolve-Path "$PSScriptRoot\.."
$releaseRoot = Join-Path $root "release"
$versionLabel = if ($Version.StartsWith("v")) { $Version } else { "v$Version" }
$packageName = "logo-add-tool-$versionLabel"
$packageDir = Join-Path $releaseRoot $packageName
$zipPath = Join-Path $releaseRoot "$packageName.zip"

if (-not (Test-Path (Join-Path $root "dist"))) {
  throw "dist directory was not found. Run npm run build first."
}

if (Test-Path $packageDir) {
  Remove-Item -LiteralPath $packageDir -Recurse -Force
}

if (Test-Path $zipPath) {
  Remove-Item -LiteralPath $zipPath -Force
}

New-Item -ItemType Directory -Path $packageDir | Out-Null

$rootFiles = @(
  "README.md",
  "CHANGELOG.md",
  "VERSION",
  "LICENSE",
  "manifest.xml",
  "install-windows.ps1",
  "start-windows.bat",
  "install-mac.command",
  "start-mac.command"
)

foreach ($file in $rootFiles) {
  Copy-Item -LiteralPath (Join-Path $root $file) -Destination $packageDir
}

foreach ($dir in @("dist", "docs", "scripts", "server")) {
  Copy-Item -LiteralPath (Join-Path $root $dir) -Destination $packageDir -Recurse
}

Compress-Archive -LiteralPath $packageDir -DestinationPath $zipPath -Force

Write-Host "Created $zipPath"
