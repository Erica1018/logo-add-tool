# 更新记录

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
