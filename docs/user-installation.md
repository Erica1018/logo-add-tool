# Logo 添加工具用户安装指南

## Windows

Windows 版需要运行本地服务，因为它要保留 PowerPoint COM helper 来稳定批量写入 Logo。

### 第一次安装

1. 解压项目到固定目录，例如 `C:\LogoAddTool`。
2. 右键 `install-windows.ps1`，选择“使用 PowerShell 运行”。
3. 如果被执行策略拦住，在该目录打开 PowerShell 后运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\install-windows.ps1
```

### 每次使用

1. 双击 `start-windows.bat`。
2. 保持窗口打开。
3. 重启 PowerPoint。
4. 在 PowerPoint 顶部菜单或 `开始 > 加载项` 中打开 `Logo 添加工具`。

如果 PowerPoint 看不到插件，请在 `文件 > 选项 > 信任中心 > 信任中心设置 > 受信任的加载项目录` 中添加：

```text
\\localhost\LogoAddToolCatalog
```

并勾选“显示在菜单中”。

## Mac

Mac 版使用免安装交付包，不需要 Node、本地 HTTPS 服务或管理员密码。

### 安装

1. 下载 `logo-add-tool-mac-portable-*.zip`。
2. 解压。
3. 双击 `install-mac.command`。
4. 重启 PowerPoint。
5. 在 `开始 > 加载项` 中打开 `Logo 添加工具`。

如果 macOS 提示无法打开脚本，请右键 `install-mac.command`，选择“打开”。

### 卸载

双击 `uninstall-mac.command`，然后重启 PowerPoint。

## 使用流程

1. 点击 `导入 Logo`。
2. 选择一个或多个 PNG、JPEG 或 WebP 文件。
3. 选中 Logo，点击 `插入当前页`。
4. 在 PowerPoint 页面中拖拽和缩放 Logo。
5. 点击 `使用当前选中 Logo 的位置`。
6. 点击 `应用全部`。

第二个 Logo 重复同样流程即可。插件会按不同 `placementId` 管理不同 Logo，不会因为应用第二个 Logo 而删除第一个 Logo。
