# Windows 原生桥接协议

当前 MVP 仍是 Office.js 侧边栏。为了让 Windows 桌面版的“应用全部”稳定可用，开发服务器额外提供了一个本机 COM helper：

- 侧边栏请求 `POST /api/powerpoint/apply-all`。
- Vite 中间件把 Logo 和 `LogoPlacement` 写入临时 JSON。
- `scripts/apply-logo-com.ps1` 通过 PowerPoint COM 连接当前活动演示文稿，并把 Logo 插入每一页。
- 脚本先插入临时 Logo，确认所有页都成功后再删除旧 Logo 并改名为正式标记，避免失败时把现有 Logo 清空。

下面是后续 VSTO/WebView2 原生宿主的正式桥接协议。

共享的 React 侧边栏可以由未来的 Windows VSTO/COM 插件通过 WebView2 承载。
该原生宿主需要注入：

```ts
window.BrandLogoStampNative = { enabled: true, platform: "windows" };
```

并实现 `window.chrome.webview.postMessage` 的请求/响应处理。

## 请求格式

```ts
{
  id: string;
  type:
    | "capabilities:get"
    | "logo:insert-current"
    | "logo:read-selected-placement"
    | "logo:apply-all"
    | "logo:update-placement"
    | "logo:remove-placement";
  payload?: unknown;
}
```

## 响应格式

```ts
{
  id: string;
  ok: boolean;
  payload?: unknown;
  error?: string;
}
```

## 原生 PowerPoint 宿主职责

- 使用 PowerPoint COM 图形 API 插入 Logo。
- 尽可能把同一份元数据写入图形名称、tags 或自定义数据。
- 第二阶段在可用时使用 PowerPoint 原生 `Locked` 能力。
- 返回与 Office.js 适配器一致的 `LogoPlacement` 数据模型。
