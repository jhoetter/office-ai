import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { DocxAgent } from "../agent/agent.js";
import { parseDocx } from "../parser/parse.js";
import { buildBlankDocxBuffer } from "../agent/empty.js";

describe("docx:create-header-footer-part", () => {
  it("mints a fresh header part on a blank document and round-trips", async () => {
    const blank = await buildBlankDocxBuffer();
    const agent = await DocxAgent.fromBuffer(blank, { idMinter: deterministicIdMinter() });

    expect(agent.getSnapshot().root.headersAndFooters).toHaveLength(0);

    await agent.applyCommand({
      type: "docx:create-header-footer-part",
      payload: { slot: "header" },
    });
    const snap = agent.getSnapshot();

    const parts = snap.root.headersAndFooters;
    expect(parts).toHaveLength(1);
    expect(parts[0].kind).toBe("header");
    expect(parts[0].partPath).toBe("word/header1.xml");
    expect(parts[0].body).toHaveLength(1);

    // Section gained a default headerReference resolving to the new rel.
    const lastBlock = snap.root.body[snap.root.body.length - 1];
    if (lastBlock.kind !== "section-break") {
      throw new Error("expected trailing section break");
    }
    expect(lastBlock.properties.headerRefs).toHaveLength(1);
    const ref = lastBlock.properties.headerRefs[0];
    expect(ref.type).toBe("default");
    const docRels = snap.root.relationships.get("word/document.xml") ?? [];
    expect(docRels.find((r) => r.id === ref.relationshipId)?.target).toBe("header1.xml");
    // `raw` is dropped so the serializer rebuilds <w:sectPr>.
    expect(lastBlock.raw).toBeUndefined();

    // Round-trip: bytes parse back into the same shape.
    const bytes = await agent.exportFile();
    const reparsed = await parseDocx(new Uint8Array(bytes), { idMinter: deterministicIdMinter() });
    expect(reparsed.root.headersAndFooters).toHaveLength(1);
    expect(reparsed.root.headersAndFooters[0].kind).toBe("header");
    expect(reparsed.root.headersAndFooters[0].partPath).toBe("word/header1.xml");
  });

  it("mints a fresh footer part with a unique relId and partPath", async () => {
    const blank = await buildBlankDocxBuffer();
    const agent = await DocxAgent.fromBuffer(blank, { idMinter: deterministicIdMinter() });

    await agent.applyCommand({
      type: "docx:create-header-footer-part",
      payload: { slot: "header" },
    });
    await agent.applyCommand({
      type: "docx:create-header-footer-part",
      payload: { slot: "footer" },
    });
    const snap = agent.getSnapshot();

    const paths = snap.root.headersAndFooters.map((p) => p.partPath);
    expect(paths).toEqual(["word/header1.xml", "word/footer1.xml"]);
    const docRels = snap.root.relationships.get("word/document.xml") ?? [];
    const ids = docRels.map((r) => r.id);
    expect(new Set(ids).size).toBe(ids.length);
  });

  it("is idempotent when the requested slot already exists", async () => {
    const blank = await buildBlankDocxBuffer();
    const agent = await DocxAgent.fromBuffer(blank, { idMinter: deterministicIdMinter() });

    await agent.applyCommand({
      type: "docx:create-header-footer-part",
      payload: { slot: "header" },
    });
    const revAfterFirst = agent.getSnapshot().revision;
    await agent.applyCommand({
      type: "docx:create-header-footer-part",
      payload: { slot: "header" },
    });
    expect(agent.getSnapshot().revision).toBe(revAfterFirst);
    expect(agent.getSnapshot().root.headersAndFooters).toHaveLength(1);
  });
});
