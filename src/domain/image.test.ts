import { describe, expect, it } from "vitest";
import {
  dataUrlToBase64,
  defaultLogoSizeInPoints,
  isSupportedLogoFile,
  roundPoint,
  shouldDownsampleForPowerPoint,
} from "./image";
import type { LogoAsset } from "./types";

describe("image helpers", () => {
  it("extracts base64 payload from data URLs", () => {
    expect(dataUrlToBase64("data:image/png;base64,abc123")).toBe("abc123");
    expect(dataUrlToBase64("abc123")).toBe("abc123");
  });

  it("accepts only MVP logo image types", () => {
    expect(isSupportedLogoFile(new File(["x"], "logo.png", { type: "image/png" }))).toBe(true);
    expect(isSupportedLogoFile(new File(["x"], "logo.jpg", { type: "image/jpeg" }))).toBe(true);
    expect(isSupportedLogoFile(new File(["x"], "logo.svg", { type: "image/svg+xml" }))).toBe(false);
  });

  it("keeps default logo size within MVP bounds", () => {
    const asset: LogoAsset = {
      id: "logo_1",
      name: "wide.png",
      mimeType: "image/png",
      data: "data:image/png;base64,abc123",
      intrinsicWidth: 1200,
      intrinsicHeight: 300,
      createdAt: "2026-05-13T00:00:00.000Z",
    };

    expect(defaultLogoSizeInPoints(asset)).toEqual({ width: 108, height: 27 });
  });

  it("rounds PowerPoint point values consistently", () => {
    expect(roundPoint(12.345)).toBe(12.35);
    expect(roundPoint(12.344)).toBe(12.34);
  });

  it("flags oversized images for PowerPoint downsampling", () => {
    expect(shouldDownsampleForPowerPoint({ intrinsicWidth: 3508, intrinsicHeight: 2481 })).toBe(true);
    expect(shouldDownsampleForPowerPoint({ intrinsicWidth: 1024, intrinsicHeight: 512 })).toBe(false);
  });
});
