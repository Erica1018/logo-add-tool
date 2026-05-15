export type PlacementScope = "allSlides" | "selectedSlides" | "currentLayout" | "currentMaster";

export interface LogoAsset {
  id: string;
  name: string;
  mimeType: string;
  data: string;
  intrinsicWidth: number;
  intrinsicHeight: number;
  createdAt: string;
}

export interface LogoPlacement {
  id: string;
  logoId: string;
  left: number;
  top: number;
  width: number;
  height: number;
  scope: PlacementScope;
  createdAt: string;
}

export interface StampMetadata {
  brandStamp: true;
  logoId: string;
  placementId: string;
}

export interface PowerPointCapabilities {
  host: "office-js" | "native-windows" | "mock";
  platform: "windows" | "mac" | "web" | "unknown";
  canInsertOnCurrentSlide: boolean;
  canReadSelectedShapePosition: boolean;
  canApplyToAllSlides: boolean;
  canRemoveStampedLogos: boolean;
  canUseMasterLayout: boolean;
  canLockShapes: boolean;
  warnings: string[];
}

export interface InsertedLogo {
  shapeId: string;
  placement: LogoPlacement;
}

export interface PowerPointAdapter {
  insertLogoOnCurrentSlide(asset: LogoAsset): Promise<InsertedLogo>;
  readSelectedLogoPlacement(activeLogoId?: string): Promise<LogoPlacement>;
  applyPlacementToAllSlides(asset: LogoAsset, placement: LogoPlacement): Promise<void>;
  updatePlacement(placement: LogoPlacement): Promise<void>;
  removePlacement(placementId: string): Promise<void>;
  getCapabilities(): Promise<PowerPointCapabilities>;
}

export interface NativeBridgeRequestMap {
  "capabilities:get": undefined;
  "logo:insert-current": { asset: LogoAsset };
  "logo:read-selected-placement": { activeLogoId?: string };
  "logo:apply-all": { asset: LogoAsset; placement: LogoPlacement };
  "logo:update-placement": { placement: LogoPlacement };
  "logo:remove-placement": { placementId: string };
}
