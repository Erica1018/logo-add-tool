# 参与贡献

感谢你参与 Logo 添加工具。这个项目的目标是做一个简单、稳定、能在 PowerPoint 侧边栏使用的 Logo 添加插件。

## 本地开发

先安装 Node.js 20 或更高版本，然后运行：

```bash
npm install
npm run dev
```

开发服务器默认运行在：

```text
https://localhost:3001
```

PowerPoint 旁加载请参考 `README.md` 中的 Windows 或 Mac 使用指南。

## 提交前检查

提交 PR 前请至少运行：

```bash
npm run typecheck
npm test
npm run build
```

如果改动涉及 `manifest.xml`，请同时验证 Office 清单：

```bash
npx office-addin-manifest validate manifest.xml
```

## 适合贡献的方向

- 修复 Windows 或 Mac PowerPoint 里的兼容性问题。
- 改进 Logo 库导入、命名、删除和本地存储体验。
- 改进批量应用位置的稳定性。
- 补充安装教程、故障排查和截图。
- 探索后续的 Logo 清理、去背景、母版/版式锁定能力。

## 提交 PR

请让 PR 尽量小而清楚，并写明：

- 你解决的问题是什么。
- 你改了哪些主要文件。
- 你在哪些 PowerPoint 版本或系统上验证过。
- 如果没有验证某个平台，请直接说明。

## 报告问题

提交 Issue 时，请尽量提供：

- Windows 或 Mac 版本。
- PowerPoint 版本。
- 插件版本。
- 复现步骤。
- 错误截图或侧边栏错误文案。
