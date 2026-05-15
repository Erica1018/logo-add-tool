import type { PowerPointAdapter } from "../domain/types";
import { MockPowerPointAdapter } from "./mockPowerPointAdapter";
import { NativeBridgePowerPointAdapter } from "./nativeBridgePowerPointAdapter";
import { OfficeJsPowerPointAdapter } from "./officeJsPowerPointAdapter";

export async function createPowerPointAdapter(): Promise<PowerPointAdapter> {
  if (window.BrandLogoStampNative?.enabled && window.chrome?.webview) {
    return new NativeBridgePowerPointAdapter();
  }

  if (typeof Office !== "undefined") {
    try {
      const readyInfo = await Office.onReady();
      const isPowerPoint =
        Office.context?.host === Office.HostType.PowerPoint ||
        readyInfo.host === Office.HostType.PowerPoint ||
        typeof PowerPoint !== "undefined";

      if (isPowerPoint) {
        return new OfficeJsPowerPointAdapter();
      }
    } catch {
      // Fall through to mock mode for browser-only development.
    }
  }

  return new MockPowerPointAdapter();
}
