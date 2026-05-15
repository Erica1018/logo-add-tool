import { createId } from "../domain/id";
import { dataUrlToBase64, defaultLogoSizeInPoints, prepareLogoDataForPowerPoint, roundPoint } from "../domain/image";
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
const POWERPOINT_HOST_SETTLE_DELAY_MS = 80;

type ShapeCollectionWithPreviewPicture = PowerPoint.ShapeCollection & {
  addPicture(
    base64EncodedImage: string,
    options?: Pick<LogoPlacement, "left" | "top" | "width" | "height">,
  ): PowerPoint.Shape;
};

type ImageInsertionOptions = Office.SetSelectedDataOptions & {
  imageLeft?: number;
  imageTop?: number;
  imageWidth?: number;
  imageHeight?: number;
};

type SlideInfo = {
  id: string;
  index: number;
};

type SlideShapeSnapshot = Map<string, Set<string>>;

type CurrentSlideShapeSnapshot = {
  slideId: string;
  shapeIds: Set<string>;
};

type InsertedShapeRef = {
  slideId: string;
  shapeId: string;
};

export class OfficeJsPowerPointAdapter implements PowerPointAdapter {
  async insertLogoOnCurrentSlide(asset: LogoAsset): Promise<InsertedLogo> {
    await this.ensureOfficeReady();
    const insertedShapeRef: { current: InsertedShapeRef | null } = { current: null };
    const insertedMetadata: { current: StampMetadata | null } = { current: null };

    const insertedLogo = await PowerPoint.run(async (context) => {
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

      const imageData = await prepareLogoDataForPowerPoint(asset);
      const base64Image = dataUrlToBase64(imageData);
      const metadata: StampMetadata = { brandStamp: true, logoId: asset.id, placementId: placement.id };
      insertedMetadata.current = metadata;

      if (!this.hasPreviewAddPicture(slide.shapes)) {
        const beforeShapeIds = await this.getDeckShapeIdsBySlide();
        slide.load("id");
        await context.sync();
        await this.insertImageWithDocumentApi(base64Image);
        const insertedShape = await this.markInsertedShapeSince(
          slide.id,
          beforeShapeIds,
          metadata,
          placement,
        );
        if (insertedShape) {
          insertedShapeRef.current = insertedShape;
        }
        return { shapeId: "office-selected-image", placement };
      }

      const picture = this.addPicture(slide.shapes, base64Image, placement);
      this.markShape(picture, metadata);
      picture.load("id");
      slide.load("id");
      await context.sync();

      insertedShapeRef.current = { slideId: slide.id, shapeId: picture.id };

      return { shapeId: picture.id, placement };
    });

    if (insertedShapeRef.current && insertedMetadata.current) {
      await this.tryMarkShapeTags(insertedShapeRef.current.slideId, insertedShapeRef.current.shapeId, insertedMetadata.current);
      await this.trySelectShape(insertedShapeRef.current.slideId, insertedShapeRef.current.shapeId);
    }

    return insertedLogo;
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
    const imageData = await prepareLogoDataForPowerPoint(asset);
    if (!(await this.canUsePreviewAddPicture())) {
      if (this.detectPlatform() === "windows") {
        await this.applyPlacementWithLocalComHelper(asset, placement);
        return;
      }

      await this.applyPlacementWithSelectedDataFallback(asset, placement, imageData);
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

        const picture = this.addPicture(slide.shapes, dataUrlToBase64(imageData), {
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
    const canUseBatchPictures = isPowerPoint ? await this.canUsePreviewAddPicture() : false;
    const warnings: string[] = [];

    if (!isPowerPoint) {
      warnings.push("请在 PowerPoint 内打开此插件。");
    }

    if (isPowerPoint && !canUseBatchPictures) {
      if (platform === "windows") {
        warnings.push("当前 Windows PowerPoint 会使用本地 COM helper 批量写入 Logo。");
      } else {
        warnings.push("当前 PowerPoint 会使用兼容模式逐页写入 Logo，过程中可能短暂切换幻灯片。");
      }
    }

    if (platform === "mac") {
      warnings.push("Mac 免安装版不需要 Node、本地 HTTPS 服务或管理员权限；图形锁定会在后续 Helper 版中继续实现。");
    }

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
  }

  private markShapeTags(shape: PowerPoint.Shape, metadata: StampMetadata): void {
    shape.tags.add(TAG_BRAND_STAMP, "true");
    shape.tags.add(TAG_LOGO_ID, metadata.logoId);
    shape.tags.add(TAG_PLACEMENT_ID, metadata.placementId);
  }

  private async tryMarkShapeTags(slideId: string, shapeId: string, metadata: StampMetadata): Promise<void> {
    try {
      await PowerPoint.run(async (context) => {
        const shape = context.presentation.slides.getItem(slideId).shapes.getItem(shapeId);
        this.markShapeTags(shape, metadata);
        await context.sync();
      });
    } catch {
      // Shape name is the canonical marker. Tags are best-effort on Mac Office builds.
    }
  }

  private async trySelectShape(slideId: string, shapeId: string): Promise<void> {
    try {
      await PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getItem(slideId);
        slide.setSelectedShapes([shapeId]);
        await context.sync();
      });
    } catch {
      // Selection is only a convenience. It should not make insertion/apply fail.
    }
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
    try {
      return await PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getItemAt(0);
        await context.sync();
        return this.hasPreviewAddPicture(slide.shapes);
      });
    } catch {
      return false;
    }
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

  private async applyPlacementWithSelectedDataFallback(
    asset: LogoAsset,
    placement: LogoPlacement,
    imageData: string,
  ): Promise<void> {
    const base64Image = dataUrlToBase64(imageData);
    const slides = await this.getSlideInfos();
    if (slides.length === 0) {
      throw new Error("当前演示文稿没有可写入的幻灯片。");
    }

    const originalSlideId = await this.getCurrentSlideId().catch(() => null);
    const originalSlideOrdinal = Math.max(0, slides.findIndex((slide) => slide.id === originalSlideId));
    const temporaryPlacementId = createId("placement_tmp");
    const temporaryMetadata: StampMetadata = {
      brandStamp: true,
      logoId: asset.id,
      placementId: temporaryPlacementId,
    };
    const finalMetadata: StampMetadata = {
      brandStamp: true,
      logoId: asset.id,
      placementId: placement.id,
    };

    const insertedShapeIds: InsertedShapeRef[] = [];
    try {
      await this.goToSlideOrdinal(0);
      for (let slideIndex = 0; slideIndex < slides.length; slideIndex += 1) {
        const beforeSnapshot = await this.getCurrentSlideShapeSnapshot();
        await this.insertImageWithDocumentApi(base64Image, placement);
        await this.waitForPowerPointHostToSettle();
        const shapeId = await this.markInsertedShapeOnSlideSince(
          beforeSnapshot.slideId,
          beforeSnapshot.shapeIds,
          temporaryMetadata,
          placement,
        );

        if (!shapeId) {
          throw new Error("PowerPoint 已插入图片，但插件没有找到新 Logo 图形。请重试。");
        }

        insertedShapeIds.push({ slideId: beforeSnapshot.slideId, shapeId });
        if (slideIndex < slides.length - 1) {
          await this.goToRelativeSlide(Office.Index.Next);
        }
      }

      await this.commitTemporaryPlacement(placement.id, temporaryPlacementId, finalMetadata, insertedShapeIds);
    } catch (error) {
      await this.removePlacement(temporaryPlacementId).catch(() => undefined);
      throw error;
    } finally {
      await this.goToSlideOrdinal(originalSlideOrdinal).catch(() => undefined);
    }
  }

  private async commitTemporaryPlacement(
    finalPlacementId: string,
    temporaryPlacementId: string,
    finalMetadata: StampMetadata,
    insertedShapeIds: InsertedShapeRef[],
  ): Promise<void> {
    const insertedShapeKeys = new Set(insertedShapeIds.map((item) => `${item.slideId}:${item.shapeId}`));
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/id,items/shapes/items/id,items/shapes/items/name");
      await context.sync();

      for (const slide of slides.items) {
        for (const shape of slide.shapes.items) {
          const shapeKey = `${slide.id}:${shape.id}`;
          if (isStampShapeName(shape.name, finalPlacementId)) {
            shape.delete();
            continue;
          }

          if (insertedShapeKeys.has(shapeKey) && isStampShapeName(shape.name, temporaryPlacementId)) {
            this.markShape(shape, finalMetadata);
          }
        }
      }

      await context.sync();
    });
  }

  private async getSlideInfos(): Promise<SlideInfo[]> {
    return PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/id,items/index");
      await context.sync();
      return slides.items
        .map((slide) => ({ id: slide.id, index: slide.index }))
        .sort((a, b) => a.index - b.index);
    });
  }

  private async getCurrentSlideId(): Promise<string> {
    return PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      slide.load("id");
      await context.sync();
      return slide.id;
    });
  }

  private async getCurrentSlideShapeSnapshot(): Promise<CurrentSlideShapeSnapshot> {
    return PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      slide.load("id");
      slide.load("shapes/items/id");
      await context.sync();
      return {
        slideId: slide.id,
        shapeIds: new Set(slide.shapes.items.map((shape) => shape.id)),
      };
    });
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
  ): Promise<InsertedShapeRef | null> {
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
      await context.sync();
      return { slideId: slide.id, shapeId: shape.id };
    });
  }

  private async markInsertedShapeOnSlideSince(
    slideId: string,
    beforeShapeIds: Set<string>,
    metadata: StampMetadata,
    placement: Pick<LogoPlacement, "left" | "top" | "width" | "height">,
  ): Promise<string | null> {
    return PowerPoint.run(async (context) => {
      const slide = context.presentation.slides.getItem(slideId);
      slide.load("shapes/items/id");
      await context.sync();

      const target = slide.shapes.items.filter((shape) => !beforeShapeIds.has(shape.id)).at(-1);
      if (!target) {
        return null;
      }

      this.markShape(target, metadata);
      target.left = placement.left;
      target.top = placement.top;
      target.width = placement.width;
      target.height = placement.height;
      await context.sync();
      return target.id;
    });
  }

  private insertImageWithDocumentApi(
    base64Image: string,
    placement?: Pick<LogoPlacement, "left" | "top" | "width" | "height">,
  ): Promise<void> {
    return new Promise((resolve, reject) => {
      const options: ImageInsertionOptions = { coercionType: Office.CoercionType.Image };
      if (placement) {
        options.imageLeft = placement.left;
        options.imageTop = placement.top;
        options.imageWidth = placement.width;
        options.imageHeight = placement.height;
      }

      Office.context.document.setSelectedDataAsync(
        base64Image,
        options,
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

  private async goToSlideOrdinal(ordinal: number): Promise<void> {
    await this.goToRelativeSlide(Office.Index.First);
    for (let index = 0; index < ordinal; index += 1) {
      await this.goToRelativeSlide(Office.Index.Next);
    }
  }

  private goToRelativeSlide(target: Office.Index): Promise<void> {
    return new Promise((resolve, reject) => {
      Office.context.document.goToByIdAsync(
        target,
        Office.GoToType.Index,
        { selectionMode: Office.SelectionMode.None },
        async (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            reject(new Error(asyncResult.error.message));
            return;
          }

          await this.waitForPowerPointHostToSettle();
          resolve();
        },
      );
    });
  }

  private waitForPowerPointHostToSettle(): Promise<void> {
    return new Promise((resolve) => {
      window.setTimeout(resolve, POWERPOINT_HOST_SETTLE_DELAY_MS);
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
