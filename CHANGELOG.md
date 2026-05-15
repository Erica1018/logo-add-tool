# 更新记录

## 0.1.4

把 Mac 免安装交付版并入主仓库，作为 `logo-add-tool` 的正式发布形态。

- 新增 `npm run package:mac`，可生成 Mac 最终用户 zip。
- 新增 `manifest.template.xml` 和 `tools/render-manifest.mjs`，用于把正式 HTTPS 托管地址写入交付包 manifest。
- Mac 安装脚本改为只写入当前用户 PowerPoint sideload 目录，不需要 Node.js、本地 HTTPS 服务或管理员权限。
- Mac Office.js 路径加入兼容模式：缺少 preview 图片 API 时逐页写入 Logo。
- 插入前自动压缩超大 Logo，降低 Mac PowerPoint `GeneralException` 风险。
- Logo 库 IndexedDB 增加连接关闭后的自动重试，避免 `database connection is closing` 报错卡住页面。
- 主项目构建不再依赖本地 Office 开发证书，方便 GitHub Pages 和 Release 流程。

## 0.1.0-mvp

首个可用版本，面向本地旁加载和小范围试用。

- 侧边栏产品名改为“Logo 添加工具”。
- 支持导入 PNG、JPEG、WebP Logo，Logo 库保存在插件本地 IndexedDB。
- 支持插入当前页，然后在 PowerPoint 画布中直接拖拽和缩放。
- 支持读取当前选中 Logo 的位置和尺寸。
- 支持把当前 Logo 位置应用到整份 PPT。
- Windows 桌面版使用本机 PowerPoint COM helper 做稳定批量写入。
- 支持同一份 PPT 中保留多个 Logo 或多个 Logo 位置。
- 隐藏环境诊断和历史位置列表，只在错误时显示诊断信息。
- 添加本地 HTTPS 运行服务和 Windows/Mac 安装脚本。
