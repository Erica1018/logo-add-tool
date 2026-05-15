import { describe, expect, it } from "vitest";
import { createStampShapeName, isStampShapeName, parseStampShapeName } from "./stampMetadata";

describe("stamp metadata", () => {
  it("round-trips metadata through the PowerPoint shape name", () => {
    const name = createStampShapeName({
      brandStamp: true,
      logoId: "logo_abc",
      placementId: "placement_xyz",
    });

    expect(parseStampShapeName(name)).toEqual({
      brandStamp: true,
      logoId: "logo_abc",
      placementId: "placement_xyz",
    });
  });

  it("detects placement-specific stamp names", () => {
    const name = "BrandLogoStamp|placement_1|logo_1";
    expect(isStampShapeName(name)).toBe(true);
    expect(isStampShapeName(name, "placement_1")).toBe(true);
    expect(isStampShapeName(name, "placement_2")).toBe(false);
  });

  it("rejects unrelated shape names", () => {
    expect(parseStampShapeName("Rectangle 1")).toBeNull();
  });
});
