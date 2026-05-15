import { createId } from "./id";
import type { LogoAsset } from "./types";

const SUPPORTED_MIME_TYPES = new Set(["image/png", "image/jpeg", "image/webp"]);

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
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve({ width: image.naturalWidth, height: image.naturalHeight });
    image.onerror = () => reject(new Error("无法读取图片尺寸。"));
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
