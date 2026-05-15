#!/bin/zsh
set -e

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT_DIR"

if ! command -v node >/dev/null 2>&1; then
  echo "Node.js was not found. Install Node.js 20 or later first: https://nodejs.org/"
  exit 1
fi

if [ ! -f "$HOME/.office-addin-dev-certs/localhost.crt" ]; then
  echo "Missing trusted localhost certificate."
  echo "Run install-mac.command first."
  exit 1
fi

node server/local-server.mjs
