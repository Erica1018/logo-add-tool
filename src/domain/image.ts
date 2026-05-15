import { createId } from "./id";
import type { LogoAsset } from "./types";

const SUPPORTED_MIME_TYPES = new Set(["image/png", "image/jpeg", "image/webp"]);
const POWERPOINT_IMAGE_MAX_DIMENSION = 1200;

export function isSupportedLogoFile(file: File): boolean {
  return SUPPORTED_MIME_TYPES.has(file.type);
}

export function dataUrlToBase64(dataUrl: string): string {
  const markerIndex = dataUrl.indexOf(",");
  return markerIndex >= 0 ? dataUrl.slice(markerIndex + 1) : dataUrl;
}

export function isTransparentPng(asset: LogoAsset): boolean {
  return asset.mimeType === "image/png" && asset.data.startsWith("data:image/png");
}

export function readFileAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(reader.error ?? new Error("无法读取文件。"));
    reader.onload = () => resolve(String(reader.result));
    reader.readAsDataURL(file);
  });
}

export function loadImageDimensions(dataUrl: string): Promise<{ width: number; height: number }> {
  return loadImageElement(dataUrl).then((image) => ({
    width: image.naturalWidth,
    height: image.naturalHeight,
  }));
}

function loadImageElement(dataUrl: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("无法读取图片。"));
    image.src = dataUrl;
  });
}

export async function createLogoAssetFromFile(file: File): Promise<LogoAsset> {
  if (!isSupportedLogoFile(file)) {
    throw new Error("当前版本仅支持 PNG、JPEG 和 WebP Logo 文件。");
  }

  const data = await readFileAsDataUrl(file);
  const dimensions = await loadImageDimensions(data);

  return {
    id: createId("logo"),
    name: file.name,
    mimeType: file.type,
    data,
    intrinsicWidth: dimensions.width,
    intrinsicHeight: dimensions.height,
    createdAt: new Date().toISOString(),
  };
}

export function defaultLogoSizeInPoints(asset: LogoAsset): { width: number; height: number } {
  const maxWidth = 108;
  const maxHeight = 54;
  const ratio = asset.intrinsicWidth > 0 ? asset.intrinsicHeight / asset.intrinsicWidth : 0.4;
  let width = maxWidth;
  let height = Math.max(12, width * ratio);

  if (height > maxHeight) {
    height = maxHeight;
    width = Math.max(12, height / Math.max(ratio, 0.1));
  }

  return { width: roundPoint(width), height: roundPoint(height) };
}

export function roundPoint(value: number): number {
  return Math.round(value * 100) / 100;
}

export async function prepareLogoDataForPowerPoint(asset: LogoAsset): Promise<string> {
  if (!shouldDownsampleForPowerPoint(asset)) {
    return asset.data;
  }

  if (typeof document === "undefined") {
    return asset.data;
  }

  const image = await loadImageElement(asset.data);
  const scale = Math.min(
    1,
    POWERPOINT_IMAGE_MAX_DIMENSION / Math.max(image.naturalWidth, image.naturalHeight),
  );

  if (scale >= 1) {
    return asset.data;
  }

  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(image.naturalWidth * scale));
  canvas.height = Math.max(1, Math.round(image.naturalHeight * scale));

  const context = canvas.getContext("2d");
  if (!context) {
    return asset.data;
  }

  context.clearRect(0, 0, canvas.width, canvas.height);
  context.drawImage(image, 0, 0, canvas.width, canvas.height);

  return canvas.toDataURL("image/png");
}

export function shouldDownsampleForPowerPoint(asset: Pick<LogoAsset, "intrinsicWidth" | "intrinsicHeight">): boolean {
  return Math.max(asset.intrinsicWidth, asset.intrinsicHeight) > POWERPOINT_IMAGE_MAX_DIMENSION;
}
