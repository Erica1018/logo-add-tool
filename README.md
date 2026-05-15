# Logo 添加工具

版本：`0.1.0-mvp`

这是一个 PowerPoint 侧边栏插件，用来把 Logo 添加到整份 PPT 的固定位置。

## 使用前准备

先安装 Node.js 20 或更高版本：

https://nodejs.org/

安装完成后，打开命令行输入：

```bash
node -v
```

能看到版本号就可以继续。

## Windows 用户使用指南

### 第一次安装

1. 把压缩包解压到一个固定位置，例如：

```text
C:\LogoAddTool
```

2. 进入解压后的文件夹。
3. 右键 `install-windows.ps1`，选择“使用 PowerShell 运行”。

如果 Windows 不允许直接运行脚本，可以在这个文件夹里打开 PowerShell，然后运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\install-windows.ps1
```

安装脚本会自动完成：

- 安装本地 HTTPS 证书。
- 准备 PowerPoint 插件清单。
- 尝试把插件目录加入 PowerPoint 受信任加载项目录。

如果脚本提示无法自动注册受信任目录，通常是因为没有管理员权限。此时可以先继续下面的“每次使用”，如果 PowerPoint 里看不到插件，再按“Windows 看不到插件怎么办”处理。

### 每次使用

1. 双击 `start-windows.bat`。
2. 保持弹出的黑色窗口打开。
3. 打开或重启 PowerPoint。
4. 在 PowerPoint 顶部菜单中找到“添加 Logo”，打开侧边栏。

### Windows 看不到插件怎么办

如果 PowerPoint 顶部没有“添加 Logo”：

1. 打开 PowerPoint。
2. 点击 `文件 > 选项 > 信任中心 > 信任中心设置`。
3. 进入 `受信任的加载项目录`。
4. 添加目录：

```text
\\localhost\LogoAddToolCatalog
```

5. 勾选“显示在菜单中”。
6. 关闭并重新打开 PowerPoint。
7. 进入 `开始 > 加载项`，从共享文件夹里添加 `Logo 添加工具`。

## Mac 用户使用指南

### 第一次安装

1. 把压缩包解压到一个固定位置，例如：

```text
~/Applications/LogoAddTool
```

2. 打开终端，进入解压后的文件夹。
3. 运行：

```bash
sh install-mac.command
```

安装脚本会自动完成：

- 安装本地 HTTPS 证书。
- 把 `manifest.xml` 放到 PowerPoint 的加载目录。

### 每次使用

1. 打开终端，进入解压后的文件夹。
2. 运行：

```bash
sh start-mac.command
```

3. 保持终端窗口打开。
4. 打开或重启 PowerPoint。
5. 进入 `开始 > 加载项`，选择 `Logo 添加工具`。

## 插件使用方法

1. 点击“导入 Logo”。
2. 可以一次选择多个 PNG、JPEG 或 WebP 文件。
3. 选择一个 Logo，点击“插入当前页”。
4. 在 PowerPoint 页面里直接拖拽、缩放这个 Logo。
5. 点击“使用当前选中 Logo 的位置”。
6. 点击“应用全部”。

如果要添加第二个 Logo，重复同样步骤即可。不同 Logo 或不同位置不会互相删除。

## 重要说明

使用插件时，`start-windows.bat` 或 `start-mac.command` 打开的窗口必须保持打开。这个窗口提供本地 HTTPS 服务，PowerPoint 需要通过它加载侧边栏。

如果 PowerPoint 显示加载失败，请先确认浏览器能打开：

```text
https://localhost:3001/index.html
```

## 技术文档

普通用户不需要阅读下面文档。后续开发或维护时再看：

- `docs/technical-architecture.md`
- `docs/user-installation.md`
- `CHANGELOG.md`

## 开源参与

这个仓库同时维护 Windows 和 Mac 的 PowerPoint 插件方案，主分支保存通用源码和平台适配逻辑。旧的 `logo-add-tool-mac` 可以保留为历史版本或部署仓库，不建议作为长期 Mac 功能分支继续开发。

欢迎提交 Issue 和 Pull Request。参与开发前请先阅读：

- `CONTRIBUTING.md`
- `docs/technical-architecture.md`
