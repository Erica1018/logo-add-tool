#!/bin/zsh
set -e

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
WEF_DIR="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
MANIFEST_SOURCE="$ROOT_DIR/manifest.xml"
MANIFEST_TARGET="$WEF_DIR/logo-add-tool.xml"
LEGACY_MANIFEST_TARGET="$WEF_DIR/logo-add-tool-mac-portable.xml"

echo "Logo 添加工具 Mac 免安装版"
echo ""

if [ ! -f "$MANIFEST_SOURCE" ]; then
  echo "没有找到 manifest.xml。请确认压缩包已经完整解压。"
  exit 1
fi

if grep -q "__TASKPANE_ORIGIN__" "$MANIFEST_SOURCE"; then
  echo "manifest.xml 还没有写入正式 HTTPS 地址。"
  echo "请先由交付方生成最终用户包，再把这个包发给 Mac 用户。"
  exit 1
fi

if grep -E -q "example.com|your-domain|yourdomain|\\.invalid|\\.test" "$MANIFEST_SOURCE"; then
  echo "manifest.xml 仍然是示例 HTTPS 地址。"
  echo "请先由交付方生成最终用户包，再把这个包发给 Mac 用户。"
  exit 1
fi

mkdir -p "$WEF_DIR"
rm -f "$LEGACY_MANIFEST_TARGET"
cp "$MANIFEST_SOURCE" "$MANIFEST_TARGET"

echo "已把插件加入 PowerPoint 加载项目录："
echo "$MANIFEST_TARGET"
echo ""
echo "下一步："
echo "1. 关闭并重新打开 PowerPoint。"
echo "2. 打开一个 PPT。"
echo "3. 进入“开始 > 加载项”，选择“Logo 添加工具”。"
echo ""
echo "这个版本不需要安装 Node.js，不需要本地服务，也不需要管理员密码。"

osascript -e 'display dialog "Logo 添加工具已加入 PowerPoint。请重启 PowerPoint 后，在“开始 > 加载项”里打开它。" buttons {"好"} default button "好"' >/dev/null 2>&1 || true
