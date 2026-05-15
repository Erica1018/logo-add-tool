# Mac 免安装交付说明

## 目标

Mac 用户只需要解压压缩包并运行一次 `install-mac.command`，就能在 PowerPoint 的侧边栏里打开 `Logo 添加工具`。用户侧不安装 Node.js，不安装系统证书，不启动本地服务，不需要管理员密码。

## 架构

```text
PowerPoint for Mac
  -> Office Add-in manifest
  -> HTTPS task pane
  -> React / Office.js
  -> PowerPoint 文档 API
```

交付包里的 `install-mac.command` 只做一件事：把 `manifest.xml` 复制到当前用户目录：

```text
~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef
```

这是微软官方的 Mac sideload 目录，属于用户目录，不需要管理员权限。

## 为什么不能直接加载本地 HTML

PowerPoint Office Add-in 的 task pane 需要通过 HTTPS 加载。为了去掉 Node、本地 HTTPS 证书和管理员弹窗，Mac 免安装版把侧边栏静态资源托管到正式 HTTPS 地址。

## 打包流程

1. 构建静态页面：

```bash
npm run build
```

2. 把 `dist/` 上传到正式 HTTPS 地址。
3. 生成 manifest 和最终交付包：

```bash
LOGO_ADD_TOOL_URL=https://erica1018.github.io/logo-add-tool npm run package:mac
```

4. 输出文件：

```text
release/logo-add-tool-mac-portable-0.1.4.zip
```

## 当前能力边界

- 保留侧边栏 Logo 库。
- 保留导入 Logo。
- 保留插入当前页。
- 保留在 PowerPoint 画布中拖拽、缩放。
- 保留读取当前选中 Logo 位置。
- 保留应用到整份 PPT；如果 Mac 版 PowerPoint 没有开放 preview 图片 API，会使用兼容模式逐页写入。
- 超大 Logo 会在写入 PowerPoint 前自动等比压缩，减少 Mac Office API 的 `GeneralException`。
- 不包含 Windows COM helper；Windows helper 仍保留在同一个主仓库里。
- 不包含本地 Node/Vite/localhost 服务。
- 暂不做图形锁定。Mac 端后续如果要做强锁定，需要单独设计 Mac 原生 Helper App。

## 风险点

Mac 免安装版依赖 PowerPoint for Mac 的 Office.js 能力。没有 preview `addPicture` API 时会使用兼容模式逐页写入；如果连基础图片插入或形状读取 API 也不可用，仍需要升级到 Microsoft 365 PowerPoint 最新版。
