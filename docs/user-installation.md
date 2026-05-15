# Logo 添加工具用户安装指南

## 先说明

这是一个 PowerPoint Web Add-in，不是传统 `.ppam` 或 VSTO 插件。Office 要求这类插件从 HTTPS 页面加载，所以压缩包里包含一个本地 HTTPS 服务。使用时需要保持启动窗口打开。

官方依据：

- Windows 旁加载：Microsoft 建议通过受信任加载项目录加载 manifest，同时 Web 应用本身仍然要部署到 HTTPS 服务。参考 [Sideload Office Add-ins from a network share](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)。
- Mac 旁加载：PowerPoint manifest 可以复制到 `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`。参考 [Sideload Office Add-ins on Mac](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac)。

## 安装前准备

需要安装 Node.js 20 或更高版本：

https://nodejs.org/

安装后，打开命令行运行下面命令，能看到版本号即可：

```bash
node -v
```

## Windows 安装

推荐把压缩包解压到一个固定路径，例如：

```text
C:\LogoAddTool
```

### 第一次安装

1. 右键 `install-windows.ps1`，选择“使用 PowerShell 运行”。
2. 如果 Windows 阻止脚本运行，用 PowerShell 进入目录后执行：

```powershell
powershell -ExecutionPolicy Bypass -File .\install-windows.ps1
```

安装脚本会做三件事：

- 安装并信任 `https://localhost:3001` 使用的本地证书。
- 准备 `manifest.xml`。
- 尝试把 manifest 目录注册为 PowerPoint 的受信任加载项目录。

如果它提示无法自动创建受信任目录，通常是因为没有管理员权限。此时可以先继续下一步，必要时按 PowerPoint 的“加载项/共享文件夹”流程手动添加。

### 每次使用

1. 双击 `start-windows.bat`，保持窗口打开。
2. 打开或重启 PowerPoint。
3. 在 PowerPoint 顶部菜单里打开“添加 Logo”。

如果顶部还看不到按钮：

1. 打开 PowerPoint。
2. 进入 `开始 > 加载项`。
3. 在共享文件夹或高级加载项里选择 `Logo 添加工具`。
4. 添加后，顶部按钮会出现在功能区中。

## Mac 安装

推荐把压缩包解压到固定路径，例如：

```text
~/Applications/LogoAddTool
```

### 第一次安装

打开终端，进入解压目录，运行：

```bash
sh install-mac.command
```

安装脚本会做两件事：

- 安装并信任 `https://localhost:3001` 使用的本地证书。
- 把 `manifest.xml` 复制到 PowerPoint 的 `wef` 旁加载目录。

### 每次使用

1. 运行：

```bash
sh start-mac.command
```

2. 保持终端窗口打开。
3. 打开或重启 PowerPoint。
4. 进入 `开始 > 加载项`，选择 `Logo 添加工具`。

## 使用流程

1. 点击“导入 Logo”，可以一次选择多个 PNG、JPEG 或 WebP 文件。
2. 选择一个 Logo，点击“插入当前页”。
3. 在 PowerPoint 页面中拖拽、缩放 Logo。
4. 点击“使用当前选中 Logo 的位置”。
5. 点击“应用全部”。

如果要添加第二个 Logo，重复同样流程即可。不同 Logo 或不同位置不会互相删除。

## 常见问题

### PowerPoint 显示加载失败

确认本地服务窗口还开着，并且能访问：

```text
https://localhost:3001/index.html
```

### 顶部按钮名称还是旧的

Office 会缓存 manifest。请关闭 PowerPoint 后重新打开；如果仍然没有更新，需要清理 Office Web Add-in 缓存或重新旁加载 manifest。

### Mac 上批量应用不可用

第一版 Mac 端依赖 PowerPoint Office.js 能力。如果当前 Mac PowerPoint 没有开放稳定图片批量插入 API，插件会显示限制。Windows 桌面版会使用 PowerPoint COM helper，因此能力更完整。
