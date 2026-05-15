import { createId } from "../domain/id";
import { defaultLogoSizeInPoints } from "../domain/image";
import type {
  InsertedLogo,
  LogoAsset,
  LogoPlacement,
  PowerPointAdapter,
  PowerPointCapabilities,
} from "../domain/types";

export class MockPowerPointAdapter implements PowerPointAdapter {
  private lastPlacement: LogoPlacement | null = null;

  async insertLogoOnCurrentSlide(asset: LogoAsset): Promise<InsertedLogo> {
    const size = defaultLogoSizeInPoints(asset);
    const placement: LogoPlacement = {
      id: createId("placement"),
      logoId: asset.id,
      left: 36,
      top: 36,
      width: size.width,
      height: size.height,
      scope: "allSlides",
      createdAt: new Date().toISOString(),
    };

    this.lastPlacement = placement;
    return { shapeId: createId("mockShape"), placement };
  }

  async readSelectedLogoPlacement(activeLogoId?: string): Promise<LogoPlacement> {
    if (!this.lastPlacement) {
      throw new Error("模拟模式下没有选中的 Logo。请先插入一个 Logo。");
    }

    return {
      ...this.lastPlacement,
      logoId: activeLogoId ?? this.lastPlacement.logoId,
      left: this.lastPlacement.left + 8,
      top: this.lastPlacement.top + 6,
      width: this.lastPlacement.width,
      height: this.lastPlacement.height,
    };
  }

  async applyPlacementToAllSlides(_asset: LogoAsset, placement: LogoPlacement): Promise<void> {
    this.lastPlacement = placement;
  }

  async updatePlacement(placement: LogoPlacement): Promise<void> {
    this.lastPlacement = placement;
  }

  async removePlacement(placementId: string): Promise<void> {
    if (this.lastPlacement?.id === placementId) {
      this.lastPlacement = null;
    }
  }

  async getCapabilities(): Promise<PowerPointCapabilities> {
    return {
      host: "mock",
      platform: "unknown",
      canInsertOnCurrentSlide: true,
      canReadSelectedShapePosition: true,
      canApplyToAllSlides: true,
      canRemoveStampedLogos: true,
      canUseMasterLayout: false,
      canLockShapes: false,
      warnings: ["当前不在 PowerPoint 中运行，操作会以浏览器开发模拟模式执行。"],
    };
  }
}
