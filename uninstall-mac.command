#!/bin/zsh
set -e

WEF_DIR="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
MANIFEST_TARGET="$WEF_DIR/logo-add-tool.xml"
LEGACY_MANIFEST_TARGET="$WEF_DIR/logo-add-tool-mac-portable.xml"

echo "移除 Logo 添加工具 Mac 免安装版"
echo ""

REMOVED=0

if [ -f "$MANIFEST_TARGET" ]; then
  rm "$MANIFEST_TARGET"
  echo "已移除：$MANIFEST_TARGET"
  REMOVED=1
fi

if [ -f "$LEGACY_MANIFEST_TARGET" ]; then
  rm "$LEGACY_MANIFEST_TARGET"
  echo "已移除旧版：$LEGACY_MANIFEST_TARGET"
  REMOVED=1
fi

if [ "$REMOVED" -eq 0 ]; then
  echo "没有找到已安装的 manifest。"
fi

echo ""
echo "如果 PowerPoint 菜单里仍然显示旧插件，请关闭 PowerPoint 后重新打开。"

osascript -e 'display dialog "Logo 添加工具已移除。请重启 PowerPoint。" buttons {"好"} default button "好"' >/dev/null 2>&1 || true
