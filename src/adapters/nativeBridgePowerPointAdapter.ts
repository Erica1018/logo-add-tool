import type {
  InsertedLogo,
  LogoAsset,
  LogoPlacement,
  NativeBridgeRequestMap,
  PowerPointAdapter,
  PowerPointCapabilities,
} from "../domain/types";

type NativeRequestType = keyof NativeBridgeRequestMap;

export class NativeBridgePowerPointAdapter implements PowerPointAdapter {
  async insertLogoOnCurrentSlide(asset: LogoAsset): Promise<InsertedLogo> {
    return this.request("logo:insert-current", { asset });
  }

  async readSelectedLogoPlacement(activeLogoId?: string): Promise<LogoPlacement> {
    return this.request("logo:read-selected-placement", { activeLogoId });
  }

  async applyPlacementToAllSlides(asset: LogoAsset, placement: LogoPlacement): Promise<void> {
    await this.request("logo:apply-all", { asset, placement });
  }

  async updatePlacement(placement: LogoPlacement): Promise<void> {
    await this.request("logo:update-placement", { placement });
  }

  async removePlacement(placementId: string): Promise<void> {
    await this.request("logo:remove-placement", { placementId });
  }

  async getCapabilities(): Promise<PowerPointCapabilities> {
    return this.request("capabilities:get", undefined);
  }

  private request<TType extends NativeRequestType, TResult>(
    type: TType,
    payload: NativeBridgeRequestMap[TType],
  ): Promise<TResult> {
    const webview = window.chrome?.webview;
    if (!webview) {
      return Promise.reject(new Error("Windows 原生桥接不可用。"));
    }

    const id = crypto.randomUUID();
    return new Promise((resolve, reject) => {
      const timeout = window.setTimeout(() => {
        webview.removeEventListener("message", onMessage);
        reject(new Error(`原生桥接请求超时：${type}`));
      }, 20000);

      const onMessage = (event: MessageEvent<NativeBridgeResponse<unknown>>) => {
        if (event.data?.id !== id) {
          return;
        }

        window.clearTimeout(timeout);
        webview.removeEventListener("message", onMessage);

        if (event.data.ok) {
          resolve(event.data.payload as TResult);
        } else {
          reject(new Error(event.data.error ?? `原生桥接请求失败：${type}`));
        }
      };

      webview.addEventListener("message", onMessage);
      webview.postMessage({ id, type, payload });
    });
  }
}
