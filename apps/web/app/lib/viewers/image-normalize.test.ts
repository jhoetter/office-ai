import { describe, expect, it } from "vitest";
import { normalizeImageForViewer } from "./image-normalize";

describe("normalizeImageForViewer", () => {
  it("passes browser-renderable images through unchanged", async () => {
    const bytes = new Uint8Array([1, 2, 3]);
    const result = await normalizeImageForViewer(bytes, "pixel.png");
    expect(result).toMatchObject({
      mediaType: "image/png",
      extension: "png",
      browserRenderable: true,
      diagnostics: [],
    });
    expect(result.displayBytes).toBe(bytes);
  });

  it("normalizes HEIC/TIFF-class codecs into a metadata SVG preview", async () => {
    const result = await normalizeImageForViewer(new Uint8Array([1, 2, 3, 4]), "photo.heic");
    expect(result).toMatchObject({
      mediaType: "image/heic",
      extension: "heic",
      browserRenderable: false,
    });
    expect(result.normalizedPreviewSvg).toContain("HEIC");
    expect(new TextDecoder().decode(result.displayBytes)).toContain("Normalized preview");
    expect(result.diagnostics[0]).toContain("preserved as the working artifact");
  });
});
