# Logo 添加工具技术说明

## 版本

当前版本：`0.1.4`

## 项目定位

Logo 添加工具是一个 PowerPoint Task Pane Web Add-in。用户从侧边栏导入 Logo，插入当前页后用 PowerPoint 原生画布拖拽和缩放，再把当前位置批量应用到整份演示文稿。

同一个仓库维护两条平台路径：

- Windows：Office.js 侧边栏 + 本地 PowerPoint COM helper。
- Mac：Office.js 侧边栏 + HTTPS 静态托管 + 当前用户 sideload manifest。

## 运行形态

### Windows 本地增强版

Windows 版使用本地 HTTPS 开发/运行服务：

- `manifest.xml` 指向 `https://localhost:3001/index.html`。
- `server/local-server.mjs` 或 Vite dev server 提供侧边栏。
- Vite dev server 暴露 PowerPoint COM helper API：
  - `POST /api/powerpoint/read-selected-placement`
  - `POST /api/powerpoint/apply-all`

对应脚本：

- `scripts/read-selected-placement-com.ps1`
- `scripts/apply-logo-com.ps1`

### Mac 免安装版

Mac 版不要求用户安装 Node、本地证书或系统软件：

- `dist/` 部署到 GitHub Pages 或其他 HTTPS 静态托管。
- `manifest.template.xml` 在打包时写入正式 URL。
- `install-mac.command` 把最终 `manifest.xml` 复制到：

```text
~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef
```

这是 PowerPoint for Mac 官方 sideload manifest 目录，属于当前用户目录，不需要管理员权限。

## 核心数据模型

```ts
interface LogoAsset {
  id: string;
  name: string;
  mimeType: string;
  data: string;
  intrinsicWidth: number;
  intrinsicHeight: number;
  createdAt: string;
}

interface LogoPlacement {
  id: string;
  logoId: string;
  left: number;
  top: number;
  width: number;
  height: number;
  scope: "allSlides" | "selectedSlides" | "currentLayout" | "currentMaster";
  createdAt: string;
}
```

`LogoAsset` 和 `LogoPlacement` 保存在插件本地 IndexedDB：

- 数据库：`brand-logo-stamp`
- Logo 表：`logoAssets`
- 位置表：`logoPlacements`

IndexedDB 连接如果被 PowerPoint WebView 关闭，会自动重开并重试一次，避免 Logo 库页面因为 transient connection closing 报错失效。

## PowerPoint 标记策略

插件生成的 Logo shape 以 shape name 作为主标记：

```text
BrandLogoStamp|<placementId>|<logoId>
```

tags 是可选增强：

- `brandStamp=true`
- `logoId=<logoId>`
- `placementId=<placementId>`

Mac PowerPoint 的 tags 在部分版本上不稳定，因此 shape name 是唯一必须依赖的标记。批量更新和删除以 `placementId` 为边界，只替换当前这一组 Logo，不删除其他 Logo 或其他位置。

## Windows 批量写入

Windows 桌面版优先使用 PowerPoint COM helper。批量写入流程：

1. 从 payload 解出 Logo data URL，写入临时图片文件。
2. 连接当前运行中的 `PowerPoint.Application`。
3. 读取当前活动演示文稿。
4. 先用临时 `placementId` 插入每一页。
5. 对新图形显式设置 `Left`、`Top`、`Width`、`Height`。
6. 删除同一 `placementId` 的旧图形。
7. 把临时图形改名为正式插件标记。
8. 短暂等待后再次校正位置和尺寸，避免 PowerPoint 首次解析图片时覆盖尺寸。

## Mac 批量写入

Mac 使用 Office.js。能力分两层：

- 如果当前 PowerPoint 开放 `ShapeCollection.addPicture`，直接按坐标批量写入。
- 如果没有开放 preview 图片 API，使用 `setSelectedDataAsync` 兼容模式逐页写入，并在每页写入后重新标记和校正尺寸。

Mac 兼容模式会短暂切换幻灯片，这是 Office.js 限制造成的交互代价。为了降低 `GeneralException` 风险，超大 Logo 会先在浏览器端等比压缩，再写入 PowerPoint。

## 打包和发布

构建检查：

```powershell
npm run typecheck
npm test
npm run build
```

生成 Mac 免安装交付包：

```powershell
$env:LOGO_ADD_TOOL_URL="https://erica1018.github.io/logo-add-tool"
npm run package:mac
```

打包结果位于：

```text
release/logo-add-tool-mac-portable-<version>.zip
```

## 目录说明

```text
dist/                         构建后的侧边栏静态资源
docs/                         用户和技术文档
manifest.xml                  本地开发/Windows 旁加载 manifest
manifest.template.xml         Mac 交付包 manifest 模板
tools/render-manifest.mjs     渲染正式 manifest
tools/package-mac-release.mjs 生成 Mac 免安装 zip
scripts/apply-logo-com.ps1    Windows 批量应用 helper
scripts/read-selected-*.ps1   Windows 读取选中 Logo 位置 helper
server/local-server.mjs       Windows 本地运行服务
install-windows.ps1           Windows 第一次安装脚本
start-windows.bat             Windows 启动本地服务
install-mac.command           Mac 免安装 manifest 写入脚本
uninstall-mac.command         Mac manifest 移除脚本
```
