import { describe, expect, it, vi } from "vitest";
import { OfficeJsPowerPointAdapter } from "./officeJsPowerPointAdapter";
import type { StampMetadata } from "../domain/types";

describe("OfficeJsPowerPointAdapter", () => {
  it("keeps the core shape marker independent from PowerPoint tags", () => {
    const adapter = new OfficeJsPowerPointAdapter() as unknown as {
      markShape(shape: PowerPoint.Shape, metadata: StampMetadata): void;
    };
    const addTag = vi.fn(() => {
      throw new Error("Mac PowerPoint tags are unavailable");
    });
    const shape = {
      name: "",
      tags: { add: addTag },
    } as unknown as PowerPoint.Shape;

    adapter.markShape(shape, {
      brandStamp: true,
      logoId: "logo_1",
      placementId: "placement_1",
    });

    expect(shape.name).toBe("BrandLogoStamp|placement_1|logo_1");
    expect(addTag).not.toHaveBeenCalled();
  });
});
