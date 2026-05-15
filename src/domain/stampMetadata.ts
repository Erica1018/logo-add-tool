import type { StampMetadata } from "./types";

const SHAPE_NAME_PREFIX = "BrandLogoStamp";

export function createStampShapeName(metadata: StampMetadata): string {
  return `${SHAPE_NAME_PREFIX}|${metadata.placementId}|${metadata.logoId}`;
}

export function parseStampShapeName(name: string): StampMetadata | null {
  const [prefix, placementId, logoId] = name.split("|");
  if (prefix !== SHAPE_NAME_PREFIX || !placementId || !logoId) {
    return null;
  }

  return {
    brandStamp: true,
    logoId,
    placementId,
  };
}

export function isStampShapeName(name: string, placementId?: string): boolean {
  const metadata = parseStampShapeName(name);
  if (!metadata) {
    return false;
  }
  return placementId ? metadata.placementId === placementId : true;
}
