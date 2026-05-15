# Logo 添加工具技术说明

## 版本

当前版本：`0.1.0-mvp`

## 目标

Logo 添加工具是一个 PowerPoint Task Pane Web Add-in。用户从侧边栏导入 Logo，插入当前页后用 PowerPoint 原生画布拖拽和缩放，再把当前位置批量应用到整份演示文稿。

## 运行形态

第一版采用本地旁加载模式：

- `manifest.xml` 声明 PowerPoint 插件和功能区按钮。
- `dist/` 是 React 侧边栏静态资源。
- `server/local-server.mjs` 提供本地 HTTPS 服务，默认监听 `https://localhost:3001`。
- Windows 下，本地服务还提供 PowerPoint COM helper API。

Office Web Add-in 需要 HTTPS 页面来源。压缩包不能只靠复制文件完成运行，必须启动本地 HTTPS 服务。

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

## PowerPoint 标记策略

每个由插件生成的 Logo shape 都写入名称和 tags：

```text
BrandLogoStamp|<placementId>|<logoId>
```

tags：

- `brandStamp=true`
- `logoId=<logoId>`
- `placementId=<placementId>`

批量更新和删除以 `placementId` 为边界，只替换当前这一组 Logo，不删除其他 Logo 或其他位置。

## Windows 批量写入

Windows 桌面版使用 PowerPoint COM，入口是：

- `POST /api/powerpoint/read-selected-placement`
- `POST /api/powerpoint/apply-all`

对应脚本：

- `scripts/read-selected-placement-com.ps1`
- `scripts/apply-logo-com.ps1`

批量写入流程：

1. 从 payload 解出 Logo data URL，写入临时图片文件。
2. 连接当前运行中的 `PowerPoint.Application`。
3. 读取当前活动演示文稿。
4. 先用临时 `placementId` 插入每一页。
5. 对新图形显式设置 `Left`、`Top`、`Width`、`Height`。
6. 删除同一 `placementId` 的旧图形。
7. 把临时图形改名为正式插件标记。
8. 短暂等待后再次校正位置和尺寸，避免 PowerPoint 首次解析图片时覆盖尺寸。

## Mac 支持范围

Mac 使用 Office.js 路径。第一版不提供 Mac 原生 helper：

- 插入当前页可使用 Office.js。
- 批量应用依赖 PowerPoint Office.js 图片插入能力。
- 如果当前 PowerPoint 版本没有开放相关 API，界面会显示限制。

## 目录说明

```text
dist/                         构建后的侧边栏静态资源
docs/                         用户和技术文档
manifest.xml                  Office Add-in manifest
scripts/apply-logo-com.ps1    Windows 批量应用 helper
scripts/read-selected-*.ps1   Windows 读取选中 Logo 位置 helper
server/local-server.mjs       压缩包运行时使用的本地 HTTPS 服务
install-windows.ps1           Windows 第一次安装脚本
start-windows.bat             Windows 启动本地服务
install-mac.command           Mac 第一次安装脚本
start-mac.command             Mac 启动本地服务
```

## 构建和验证

```powershell
npm run typecheck
npm test
npm run build
npx office-addin-manifest validate manifest.xml
```

## 打包

打包时需要包含：

- `dist/`
- `docs/`
- `manifest.xml`
- `scripts/`
- `server/`
- `install-windows.ps1`
- `start-windows.bat`
- `install-mac.command`
- `start-mac.command`
- `README.md`
- `CHANGELOG.md`
- `VERSION`

不要包含：

- `node_modules/`
- `.git/`
- 开发日志
- 截图临时文件
