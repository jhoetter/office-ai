import { describe, expect, it } from "vitest";
import { PdfAgent } from "@officeai/pdf";
import { setMetadata } from "@officeai/pdf-edit";
import { isPdfBytes, loadFixture } from "./helpers.js";

/**
 * `metadata-rich.pdf` ships every Info-dict field set; the
 * `simple-text-1page.pdf` baseline has none. Both must:
 *
 *   - parse cleanly via PdfAgent
 *   - survive a no-op exportFile round-trip with their metadata intact
 *   - accept a `pdf-edit` setMetadata patch and re-emerge with the
 *     patch persisted on the next parse.
 */

describe("PDF roundtrip — metadata", () => {
  it("metadata-rich.pdf round-trips every Info-dict field", async () => {
    const bytes = await loadFixture("metadata-rich.pdf");
    const agent = await PdfAgent.fromBuffer(bytes);
    const md = agent.getSnapshot().root.metadata;

    expect(md.title).toBe("Metadata-Rich Fixture");
    expect(md.author).toBe("Office AI Night Shift");
    expect(md.subject).toBe("Roundtrip metadata coverage");
    expect(md.keywords).toBeDefined();
    expect(md.keywords ?? "").toMatch(/pdf/);
    expect(md.creator).toBe("officeai/build-fixtures");
    expect(md.producer).toBe("officeai/pdf-lib");

    const exported = await agent.exportFile();
    expect(isPdfBytes(exported)).toBe(true);
    const reparsed = await PdfAgent.fromBuffer(exported);
    const md2 = reparsed.getSnapshot().root.metadata;
    expect(md2.title).toBe(md.title);
    expect(md2.author).toBe(md.author);
    expect(md2.subject).toBe(md.subject);
    expect(md2.creator).toBe(md.creator);
    expect(md2.producer).toBe(md.producer);
  });

  it("simple-text-1page.pdf has no Info-dict title before the patch", async () => {
    const bytes = await loadFixture("simple-text-1page.pdf");
    const agent = await PdfAgent.fromBuffer(bytes);
    expect(agent.getSnapshot().root.metadata.title).toBeUndefined();
  });

  it("set-metadata via pdf-edit persists across re-parse", async () => {
    const bytes = await loadFixture("simple-text-1page.pdf");
    const patched = await setMetadata(bytes, {
      title: "Patched Title",
      author: "Roundtrip Bot",
      subject: "Set-metadata persistence",
    });
    expect(isPdfBytes(patched)).toBe(true);
    const agent = await PdfAgent.fromBuffer(patched);
    const md = agent.getSnapshot().root.metadata;
    expect(md.title).toBe("Patched Title");
    expect(md.author).toBe("Roundtrip Bot");
    expect(md.subject).toBe("Set-metadata persistence");
  });

  it("pdf:set-metadata command-bus mutation persists through exportFile", async () => {
    const bytes = await loadFixture("simple-text-1page.pdf");
    const agent = await PdfAgent.fromBuffer(bytes);
    await agent.applyCommand({
      type: "pdf:set-metadata",
      payload: { title: "Bus Title", author: "Bus Author" },
    });
    const exported = await agent.exportFile();
    const reparsed = await PdfAgent.fromBuffer(exported);
    expect(reparsed.getSnapshot().root.metadata.title).toBe("Bus Title");
    expect(reparsed.getSnapshot().root.metadata.author).toBe("Bus Author");
  });
});
