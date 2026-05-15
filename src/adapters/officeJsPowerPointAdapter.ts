import { createId } from "../domain/id";
import { dataUrlToBase64, defaultLogoSizeInPoints, roundPoint } from "../domain/image";
import { createStampShapeName, isStampShapeName, parseStampShapeName } from "../domain/stampMetadata";
import type {
  InsertedLogo,
  LogoAsset,
  LogoPlacement,
  PowerPointAdapter,
  PowerPointCapabilities,
  StampMetadata,
} from "../domain/types";

const TAG_BRAND_STAMP = "brandStamp";
const TAG_LOGO_ID = "logoId";
const TAG_PLACEMENT_ID = "placementId";

type ShapeCollectionWithPreviewPicture = PowerPoint.ShapeCollection & {
  addPicture(
    base64EncodedImage: string,
    options?: Pick<LogoPlacement, "left" | "top" | "width" | "height">,
  ): PowerPoint.Shape;
};

type SlideShapeSnapshot = Map<string, Set<string>>;

export class OfficeJsPowerPointAdapter implements PowerPointAdapter {
  async insertLogoOnCurrentSlide(asset: LogoAsset): Promise<InsertedLogo> {
    await this.ensureOfficeReady();
    return PowerPoint.run(async (context) => {
      const selectedSlides = context.presentation.getSelectedSlides();
      const slide = selectedSlides.getItemAt(0);
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

      const base64Image = dataUrlToBase64(asset.data);

      if (!this.hasPreviewAddPicture(slide.shapes)) {
        const beforeShapeIds = await this.getDeckShapeIdsBySlide();
        slide.load("id");
        await context.sync();
        await this.insertImageWithDocumentApi(base64Image);
        await this.markInsertedShapeSince(
          slide.id,
          beforeShapeIds,
          { brandStamp: true, logoId: asset.id, placementId: placement.id },
          placement,
        );
        return { shapeId: "office-selected-image", placement };
      }

      const picture = this.addPicture(slide.shapes, base64Image, placement);
      this.markShape(picture, { brandStamp: true, logoId: asset.id, placementId: placement.id });
      picture.load("id");
      await context.sync();

      slide.setSelectedShapes([picture.id]);
      await context.sync();

      return { shapeId: picture.id, placement };
    });
  }

  async readSelectedLogoPlacement(activeLogoId?: string): Promise<LogoPlacement> {
    await this.ensureOfficeReady();
    if (this.detectPlatform() === "windows") {
      return this.readSelectedPlacementWithLocalComHelper(activeLogoId);
    }

    return PowerPoint.run(async (context) => {
      const selectedShapes = context.presentation.getSelectedShapes();
      const count = selectedShapes.getCount();
      selectedShapes.load("items/id,items/name,items/left,items/top,items/width,items/height");
      await context.sync();

      if (count.value !== 1) {
        throw new Error("请在当前幻灯片中只选择一个 Logo 图形。");
      }

      const shape = selectedShapes.items[0];
      const metadata = parseStampShapeName(shape.name);
      if (!metadata && !activeLogoId) {
        throw new Error("当前选中的图形不是插件插入的 Logo。请先从侧边栏插入 Logo。");
      }
      const placementId = metadata?.placementId ?? createId("placement");
      const logoId = metadata?.logoId ?? activeLogoId ?? "";

      if (!metadata && logoId) {
        this.markShape(shape, { brandStamp: true, logoId, placementId });
        await context.sync();
      }

      return {
        id: placementId,
        logoId,
        left: roundPoint(shape.left),
        top: roundPoint(shape.top),
        width: roundPoint(shape.width),
        height: roundPoint(shape.height),
        scope: "allSlides",
        createdAt: new Date().toISOString(),
      };
    });
  }

  async applyPlacementToAllSlides(asset: LogoAsset, placement: LogoPlacement): Promise<void> {
    await this.ensureOfficeReady();
    if (!(await this.canUsePreviewAddPicture())) {
      if (this.detectPlatform() !== "windows") {
        throw new Error("当前 PowerPoint 版本暂不支持批量插入图片。Windows 版本会使用本机 PowerPoint COM 兜底。");
      }
      await this.applyPlacementWithLocalComHelper(asset, placement);
      return;
    }

    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/id,items/shapes/items/id,items/shapes/items/name");
      await context.sync();

      for (const slide of slides.items) {
        for (const shape of slide.shapes.items) {
          if (isStampShapeName(shape.name, placement.id)) {
            shape.delete();
          }
        }

        const picture = this.addPicture(slide.shapes, dataUrlToBase64(asset.data), {
          left: placement.left,
          top: placement.top,
          width: placement.width,
          height: placement.height,
        });
        this.markShape(picture, { brandStamp: true, logoId: asset.id, placementId: placement.id });
      }

      await context.sync();
    });
  }

  async updatePlacement(placement: LogoPlacement): Promise<void> {
    await this.ensureOfficeReady();
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/shapes/items/id,items/shapes/items/name,items/shapes/items/left,items/shapes/items/top,items/shapes/items/width,items/shapes/items/height");
      await context.sync();

      for (const slide of slides.items) {
        for (const shape of slide.shapes.items) {
          if (isStampShapeName(shape.name, placement.id)) {
            shape.left = placement.left;
            shape.top = placement.top;
            shape.width = placement.width;
            shape.height = placement.height;
          }
        }
      }

      await context.sync();
    });
  }

  async removePlacement(placementId: string): Promise<void> {
    await this.ensureOfficeReady();
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/shapes/items/id,items/shapes/items/name");
      await context.sync();

      for (const slide of slides.items) {
        for (const shape of slide.shapes.items) {
          if (isStampShapeName(shape.name, placementId)) {
            shape.delete();
          }
        }
      }

      await context.sync();
    });
  }

  async getCapabilities(): Promise<PowerPointCapabilities> {
    await this.ensureOfficeReady();
    const platform = this.detectPlatform();
    const isPowerPoint = Office.context?.host === Office.HostType.PowerPoint;
    const warnings: string[] = [];

    if (!isPowerPoint) {
      warnings.push("请在 PowerPoint 内打开此插件。");
    }

    warnings.push("当前环境会优先使用 PowerPoint 图片 API；如果缺少预览能力，Windows 会改用本机 PowerPoint COM 批量写入。");
    warnings.push("Office.js 目前没有稳定暴露图形锁定能力，后续阶段会优先探索母版/版式模式。");

    return {
      host: "office-js",
      platform,
      canInsertOnCurrentSlide: isPowerPoint,
      canReadSelectedShapePosition: isPowerPoint,
      canApplyToAllSlides: isPowerPoint,
      canRemoveStampedLogos: isPowerPoint,
      canUseMasterLayout: false,
      canLockShapes: false,
      warnings,
    };
  }

  private markShape(shape: PowerPoint.Shape, metadata: StampMetadata): void {
    shape.name = createStampShapeName(metadata);
    shape.tags.add(TAG_BRAND_STAMP, "true");
    shape.tags.add(TAG_LOGO_ID, metadata.logoId);
    shape.tags.add(TAG_PLACEMENT_ID, metadata.placementId);
    shape.altTextTitle = "品牌 Logo 固定标识";
    shape.altTextDescription = "由 Logo 添加工具放置的 Logo。";
  }

  private addPicture(
    shapes: PowerPoint.ShapeCollection,
    base64Image: string,
    options: Pick<LogoPlacement, "left" | "top" | "width" | "height">,
  ): PowerPoint.Shape {
    if (!this.hasPreviewAddPicture(shapes)) {
      throw new Error("当前 PowerPoint 版本暂不支持按坐标批量插入图片。请先使用“插入当前页”验证当前页插入。");
    }

    return (shapes as ShapeCollectionWithPreviewPicture).addPicture(base64Image, options);
  }

  private hasPreviewAddPicture(shapes: PowerPoint.ShapeCollection): boolean {
    return typeof (shapes as Partial<ShapeCollectionWithPreviewPicture>).addPicture === "function";
  }

  private async canUsePreviewAddPicture(): Promise<boolean> {
    return PowerPoint.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(0);
      await context.sync();
      return this.hasPreviewAddPicture(slide.shapes);
    });
  }

  private async applyPlacementWithLocalComHelper(asset: LogoAsset, placement: LogoPlacement): Promise<void> {
    const response = await fetch("/api/powerpoint/apply-all", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ asset, placement }),
    });

    const rawBody = await response.text();
    let body: { error?: string } | null = null;
    try {
      body = rawBody ? JSON.parse(rawBody) : null;
    } catch {
      body = null;
    }

    if (!response.ok) {
      throw new Error(this.formatLocalComError(body?.error ?? rawBody, "Windows 本地 PowerPoint 批量写入失败。"));
    }
  }

  private async readSelectedPlacementWithLocalComHelper(activeLogoId?: string): Promise<LogoPlacement> {
    const response = await fetch("/api/powerpoint/read-selected-placement", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ activeLogoId }),
    });

    const rawBody = await response.text();
    let body: (LogoPlacement & { error?: string }) | null = null;
    try {
      body = rawBody ? JSON.parse(rawBody) : null;
    } catch {
      body = null;
    }

    if (!response.ok || !body) {
      throw new Error(this.formatLocalComError(body?.error ?? rawBody, "Windows 本地 PowerPoint 坐标读取失败。"));
    }

    return {
      id: body.id,
      logoId: body.logoId,
      left: roundPoint(body.left),
      top: roundPoint(body.top),
      width: roundPoint(body.width),
      height: roundPoint(body.height),
      scope: "allSlides",
      createdAt: body.createdAt ?? new Date().toISOString(),
    };
  }

  private formatLocalComError(message: string | undefined, fallback: string): string {
    if (!message) {
      return fallback;
    }
    if (message.includes("PowerPoint is not running")) {
      return "没有找到正在运行的 PowerPoint。请先打开目标 PPT。";
    }
    if (message.includes("No active PowerPoint presentation")) {
      return "没有找到当前打开的演示文稿。";
    }
    if (message.includes("Select exactly one logo shape")) {
      return "请先在当前幻灯片中只选中一个 Logo 图形。";
    }
    if (message.includes("selected shape is not a logo inserted by this add-in")) {
      return "当前选中的图形不是插件插入的 Logo。请先从侧边栏插入 Logo。";
    }
    return message;
  }

  private async getDeckShapeIdsBySlide(): Promise<SlideShapeSnapshot> {
    return PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/id,items/shapes/items/id");
      await context.sync();
      return new Map(
        slides.items.map((slide) => [
          slide.id,
          new Set(slide.shapes.items.map((shape) => shape.id)),
        ]),
      );
    });
  }

  private async markInsertedShapeSince(
    preferredSlideId: string,
    beforeShapeIds: SlideShapeSnapshot,
    metadata: StampMetadata,
    placement?: Pick<LogoPlacement, "left" | "top" | "width" | "height">,
  ): Promise<string | null> {
    return PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/id,items/shapes/items/id");
      await context.sync();

      const inserted: Array<{ slide: PowerPoint.Slide; shape: PowerPoint.Shape }> = [];
      for (const slide of slides.items) {
        const previousIds = beforeShapeIds.get(slide.id) ?? new Set<string>();
        for (const shape of slide.shapes.items) {
          if (!previousIds.has(shape.id)) {
            inserted.push({ slide, shape });
          }
        }
      }

      const target =
        inserted.filter((item) => item.slide.id === preferredSlideId).at(-1) ??
        inserted.at(-1);
      if (!target) {
        return null;
      }

      const { slide, shape } = target;
      this.markShape(shape, metadata);
      if (placement) {
        shape.left = placement.left;
        shape.top = placement.top;
        shape.width = placement.width;
        shape.height = placement.height;
      }
      slide.setSelectedShapes([shape.id]);
      await context.sync();
      return slide.id;
    });
  }

  private insertImageWithDocumentApi(base64Image: string): Promise<void> {
    return new Promise((resolve, reject) => {
      Office.context.document.setSelectedDataAsync(
        base64Image,
        { coercionType: Office.CoercionType.Image },
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            reject(new Error(asyncResult.error.message));
            return;
          }

          resolve();
        },
      );
    });
  }

  private async ensureOfficeReady(): Promise<void> {
    if (typeof Office === "undefined") {
      throw new Error("PowerPoint Office.js API 不可用。");
    }

    await Office.onReady();

    if (typeof PowerPoint === "undefined") {
      throw new Error("PowerPoint Office.js API 不可用。");
    }
  }

  private detectPlatform(): PowerPointCapabilities["platform"] {
    const platform = Office.context?.platform;
    if (platform === Office.PlatformType.Mac) {
      return "mac";
    }
    if (platform === Office.PlatformType.PC) {
      return "windows";
    }
    if (platform === Office.PlatformType.OfficeOnline) {
      return "web";
    }
    return "unknown";
  }
}
