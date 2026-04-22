import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { ooxml } from "@officeai/core";
import { PptxAgent } from "../agent/agent.js";
import { parsePptx } from "../parser/parse.js";
import type { MediaShape } from "../model/types.js";

const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);

async function loadAgent(name: string): Promise<PptxAgent> {
  const buf = await readFile(join(FIXTURES_DIR.pathname, name));
  return PptxAgent.fromBuffer(buf);
}

// Tiny "fake" MP4 / MP3 bodies. The serializer doesn't decode the
// bytes — it just routes them to `ppt/media/`. Two distinct payloads
// per `mediaType` so we can exercise SHA-256 dedup vs. distinct parts.
const MP4_A = new Uint8Array([
  0x00, 0x00, 0x00, 0x18, 0x66, 0x74, 0x79, 0x70, 0x6d, 0x70, 0x34, 0x32, 0x00, 0x00, 0x00, 0x01, 0x6d, 0x70,
  0x34, 0x32, 0x69, 0x73, 0x6f, 0x6d,
]);
const MP4_B = new Uint8Array([
  0x00, 0x00, 0x00, 0x18, 0x66, 0x74, 0x79, 0x70, 0x6d, 0x70, 0x34, 0x32, 0x00, 0x00, 0x00, 0x02, 0x6d, 0x70,
  0x34, 0x32, 0x69, 0x73, 0x6f, 0x6d,
]);
const MP3_A = new Uint8Array([0xff, 0xfb, 0x90, 0x44, 0x00, 0x00, 0x00, 0x00]);

describe("P1: insert-media (video)", () => {
  it("adds a typed MediaShape, registers media + poster rels and content type", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const m = await agent.applyCommand({
      type: "pptx:insert-media",
      payload: {
        slideIndex: 0,
        mediaType: "video",
        contentType: "video/mp4",
        bytes: MP4_A,
        position: { xEmu: 1_000_000, yEmu: 2_000_000 },
        size: { cxEmu: 5_000_000, cyEmu: 3_000_000 },
        name: "Demo video",
      },
      source: "system",
    });
    expect(m.status).toBe("approved");

    const slide = agent.getSnapshot().root.slides[0];
    const media = slide.shapes.find((s): s is MediaShape => s.kind === "media");
    expect(media).toBeDefined();
    if (!media) return;
    expect(media.mediaType).toBe("video");
    expect(media.name).toBe("Demo video");
    expect(media.position).toEqual({ xEmu: 1_000_000, yEmu: 2_000_000 });
    expect(media.size).toEqual({ cxEmu: 5_000_000, cyEmu: 3_000_000 });
    expect(media.mediaPath).toMatch(/^ppt\/media\/media\d+\.mp4$/);
    expect(media.posterPath).toMatch(/^ppt\/media\/image\d+\.png$/);
    expect(media.mediaRelId).toMatch(/^rId\d+$/);
    expect(media.posterRelId).toMatch(/^rId\d+$/);

    const out = await agent.exportFile();
    const c = await ooxml.OoxmlContainer.load(out);
    expect(c.has(media.mediaPath)).toBe(true);
    expect(c.has(media.posterPath!)).toBe(true);

    // The slide XML now references the video file via a:videoFile
    // and our supplied position/size made it through serialization.
    const slideXml = c.readText(slide.partPath);
    expect(slideXml).toContain("a:videoFile");
    expect(slideXml).toContain('contentType="video/mp4"');
    expect(slideXml).toContain('x="1000000"');
    expect(slideXml).toContain('cx="5000000"');

    // Content types contains the video/mp4 default.
    const ctXml = c.readText("[Content_Types].xml");
    expect(ctXml).toContain('Extension="mp4"');
    expect(ctXml).toContain('ContentType="video/mp4"');

    // Round-trip: re-parse and verify the MediaShape survives.
    const reparsed = await parsePptx(out);
    const remedia = reparsed.root.slides[0].shapes.find((s): s is MediaShape => s.kind === "media");
    expect(remedia).toBeDefined();
    expect(remedia!.mediaType).toBe("video");
    expect(remedia!.mediaPath).toBe(media.mediaPath);
    expect(remedia!.posterPath).toBe(media.posterPath);
    expect(remedia!.position).toEqual(media.position);
    expect(remedia!.size).toEqual(media.size);
  });

  it("dedups identical bytes via SHA-256 (same media path used twice)", async () => {
    const agent = await loadAgent("01-blank.pptx");
    await agent.applyCommand({
      type: "pptx:insert-media",
      payload: {
        slideIndex: 0,
        mediaType: "video",
        contentType: "video/mp4",
        bytes: MP4_A,
        position: { xEmu: 0, yEmu: 0 },
        size: { cxEmu: 1_000_000, cyEmu: 1_000_000 },
      },
      source: "system",
    });
    await agent.applyCommand({
      type: "pptx:insert-media",
      payload: {
        slideIndex: 0,
        mediaType: "video",
        contentType: "video/mp4",
        bytes: MP4_A,
        position: { xEmu: 2_000_000, yEmu: 2_000_000 },
        size: { cxEmu: 1_000_000, cyEmu: 1_000_000 },
      },
      source: "system",
    });
    const medias = agent
      .getSnapshot()
      .root.slides[0].shapes.filter((s): s is MediaShape => s.kind === "media");
    expect(medias.length).toBe(2);
    expect(medias[0].mediaPath).toBe(medias[1].mediaPath);
    expect(medias[0].mediaRelId).toBe(medias[1].mediaRelId);
    // Posters are also deduped — same transparent placeholder.
    expect(medias[0].posterPath).toBe(medias[1].posterPath);
    expect(medias[0].posterRelId).toBe(medias[1].posterRelId);
  });

  it("creates distinct media parts for distinct bytes", async () => {
    const agent = await loadAgent("01-blank.pptx");
    await agent.applyCommand({
      type: "pptx:insert-media",
      payload: {
        slideIndex: 0,
        mediaType: "video",
        contentType: "video/mp4",
        bytes: MP4_A,
        position: { xEmu: 0, yEmu: 0 },
        size: { cxEmu: 1_000_000, cyEmu: 1_000_000 },
      },
      source: "system",
    });
    await agent.applyCommand({
      type: "pptx:insert-media",
      payload: {
        slideIndex: 0,
        mediaType: "video",
        contentType: "video/mp4",
        bytes: MP4_B,
        position: { xEmu: 0, yEmu: 0 },
        size: { cxEmu: 1_000_000, cyEmu: 1_000_000 },
      },
      source: "system",
    });
    const medias = agent
      .getSnapshot()
      .root.slides[0].shapes.filter((s): s is MediaShape => s.kind === "media");
    expect(medias.length).toBe(2);
    expect(medias[0].mediaPath).not.toBe(medias[1].mediaPath);
  });
});

describe("P1: insert-media (audio)", () => {
  it("emits an a:audioFile reference and registers an audio/mpeg default", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const m = await agent.applyCommand({
      type: "pptx:insert-media",
      payload: {
        slideIndex: 0,
        mediaType: "audio",
        contentType: "audio/mpeg",
        bytes: MP3_A,
        position: { xEmu: 0, yEmu: 0 },
        size: { cxEmu: 600_000, cyEmu: 600_000 },
      },
      source: "system",
    });
    expect(m.status).toBe("approved");

    const slide = agent.getSnapshot().root.slides[0];
    const media = slide.shapes.find((s): s is MediaShape => s.kind === "media")!;
    expect(media.mediaType).toBe("audio");
    expect(media.mediaPath).toMatch(/^ppt\/media\/media\d+\.mp3$/);

    const out = await agent.exportFile();
    const c = await ooxml.OoxmlContainer.load(out);
    const slideXml = c.readText(slide.partPath);
    expect(slideXml).toContain("a:audioFile");
    expect(slideXml).toContain('contentType="audio/mpeg"');
    const ctXml = c.readText("[Content_Types].xml");
    expect(ctXml).toContain('Extension="mp3"');
    expect(ctXml).toContain('ContentType="audio/mpeg"');
  });
});

describe("P1: insert-media (validation)", () => {
  it("rejects unsupported MIME for the chosen mediaType", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const m = await agent.applyCommand({
      type: "pptx:insert-media",
      payload: {
        slideIndex: 0,
        mediaType: "video",
        contentType: "video/avi",
        bytes: MP4_A,
        position: { xEmu: 0, yEmu: 0 },
        size: { cxEmu: 100_000, cyEmu: 100_000 },
      },
      source: "system",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-payload");
  });

  it("rejects empty bytes and zero-sized rects", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const empty = await agent.applyCommand({
      type: "pptx:insert-media",
      payload: {
        slideIndex: 0,
        mediaType: "video",
        contentType: "video/mp4",
        bytes: new Uint8Array(),
        position: { xEmu: 0, yEmu: 0 },
        size: { cxEmu: 100_000, cyEmu: 100_000 },
      },
      source: "system",
    });
    expect(empty.status).toBe("rejected");
    expect(empty.rejection?.code).toBe("invalid-payload");

    const zero = await agent.applyCommand({
      type: "pptx:insert-media",
      payload: {
        slideIndex: 0,
        mediaType: "video",
        contentType: "video/mp4",
        bytes: MP4_A,
        position: { xEmu: 0, yEmu: 0 },
        size: { cxEmu: 0, cyEmu: 100_000 },
      },
      source: "system",
    });
    expect(zero.status).toBe("rejected");
    expect(zero.rejection?.code).toBe("invalid-payload");
  });
});
