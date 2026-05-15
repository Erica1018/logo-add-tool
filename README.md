# Logo 添加工具

版本：`0.1.4`

这是一个 PowerPoint 侧边栏插件，用来把 Logo 添加到当前 PPT 的每一页。用户可以导入 Logo、插入当前页、在 PowerPoint 页面里直接拖拽缩放，然后把这个位置应用到整份 PPT。

这个仓库统一维护两个交付形态：

- Windows 本地增强版：保留本地 PowerPoint COM helper，批量应用更稳定。
- Mac 免安装版：不安装 Node.js，不启动本地服务，不需要管理员密码，通过用户目录旁加载 manifest。

## Windows 用户使用指南

Windows 版适合在自己的电脑上长期使用，需要第一次运行安装脚本，并在使用时启动本地服务。

### 第一次安装

1. 下载或解压项目到固定位置，例如：

```text
C:\LogoAddTool
```

2. 进入这个文件夹。
3. 右键 `install-windows.ps1`，选择“使用 PowerShell 运行”。

如果 Windows 不允许直接运行脚本，可以在这个文件夹里打开 PowerShell，然后运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\install-windows.ps1
```

脚本会准备本地 HTTPS 证书、PowerPoint manifest 和 Windows 批量写入 helper。

### 每次使用

1. 双击 `start-windows.bat`。
2. 保持弹出的窗口打开。
3. 打开或重启 PowerPoint。
4. 在 PowerPoint 顶部菜单里点击 `添加 Logo`，打开侧边栏。

### Windows 看不到插件怎么办

如果 PowerPoint 顶部没有 `添加 Logo`：

1. 打开 PowerPoint。
2. 进入 `文件 > 选项 > 信任中心 > 信任中心设置`。
3. 进入 `受信任的加载项目录`。
4. 添加目录：

```text
\\localhost\LogoAddToolCatalog
```

5. 勾选“显示在菜单中”。
6. 重启 PowerPoint。
7. 进入 `开始 > 加载项`，从共享目录里添加 `Logo 添加工具`。

## Mac 用户使用指南

Mac 版是给普通用户的免安装包。用户只需要解压、双击安装脚本、重启 PowerPoint。

### 安装

1. 下载 GitHub Release 里的 `logo-add-tool-mac-portable-*.zip`。
2. 解压这个 zip。
3. 双击 `install-mac.command`。
4. 关闭并重新打开 PowerPoint。
5. 打开一个 PPT。
6. 进入 `开始 > 加载项`，选择 `Logo 添加工具`。

如果 macOS 提示无法打开脚本：右键点 `install-mac.command`，选择“打开”。

### 卸载

双击 `uninstall-mac.command`，然后重启 PowerPoint。

## 插件使用方法

1. 点击 `导入 Logo`。
2. 选择一个或多个 PNG、JPEG 或 WebP 文件。
3. 选中 Logo，点击 `插入当前页`。
4. 在 PowerPoint 页面里拖拽、缩放 Logo。
5. 点击 `使用当前选中 Logo 的位置`。
6. 点击 `应用全部`。

如果要添加第二个 Logo，重复同样步骤即可。不同 Logo 和不同位置不会互相删除。

## 开发和打包

本地开发：

```powershell
npm install
npm run dev
```

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

生成的 zip 会放在 `release/` 目录。这个 zip 才是可以发给 Mac 用户的包。

## 维护说明

Mac 免安装版已经并入本仓库，不需要单独维护另一个功能项目。后续建议：

- 主仓库 `logo-add-tool`：维护源码、文档、打包脚本、Windows 和 Mac 适配逻辑。
- GitHub Pages：托管 Mac 免安装版侧边栏静态文件。
- GitHub Releases：发布最终用户 zip。

旧的 `logo-add-tool-mac` 可以保留为历史部署仓库，不建议继续作为功能分支开发。
