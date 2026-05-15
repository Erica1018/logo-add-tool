/// <reference types="vite/client" />

interface NativeBridgeEnvelope<TPayload = unknown> {
  id: string;
  type: string;
  payload?: TPayload;
}

interface NativeBridgeResponse<TPayload = unknown> {
  id: string;
  ok: boolean;
  payload?: TPayload;
  error?: string;
}

interface Window {
  BrandLogoStampNative?: {
    enabled: true;
    platform: "windows";
  };
  chrome?: {
    webview?: {
      postMessage(message: NativeBridgeEnvelope): void;
      addEventListener(
        type: "message",
        listener: (event: MessageEvent<NativeBridgeResponse>) => void,
      ): void;
      removeEventListener(
        type: "message",
        listener: (event: MessageEvent<NativeBridgeResponse>) => void,
      ): void;
    };
  };
}
