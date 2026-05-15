#!/bin/zsh
set -e

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
WEF_DIR="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"

cd "$ROOT_DIR"

echo "Logo Add Tool installer"
echo "1/3 Checking Node.js..."
if ! command -v node >/dev/null 2>&1; then
  echo "Node.js was not found. Install Node.js 20 or later first: https://nodejs.org/"
  exit 1
fi

echo "2/3 Installing trusted localhost certificate..."
npx --yes office-addin-dev-certs install

echo "3/3 Copying manifest to PowerPoint sideload folder..."
mkdir -p "$WEF_DIR"
cp "$ROOT_DIR/manifest.xml" "$WEF_DIR/logo-add-tool.xml"

echo ""
echo "Install step finished."
echo "Next: run start-mac.command, restart PowerPoint, then open Home > Add-ins > Logo Add Tool."
