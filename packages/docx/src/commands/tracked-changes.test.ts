import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { DocxAgent } from "../agent/agent.js";
import { parseDocx } from "../parser/parse.js";
import { paragraphPlainText } from "./helpers.js";
import { makeSyntheticDocx, DEFAULT_DOC_ROOT_ATTRS } from "../test-utils/synthetic.js";
import type { RevisionWrapper } from "../model/types.js";

/**
 * Build a document.xml that wraps content in `<w:ins>` and `<w:del>`
 * wrappers. The shape mirrors what Word emits: insertions wrap one or more
 * `<w:r>` elements, deletions wrap `<w:r>` whose text leaves are
 * `<w:delText>`.
 */
function trackedChangesXml(): string {
  const ins = `<w:ins w:id="100" w:author="Alice" w:date="2026-04-17T10:00:00Z"><w:r><w:t xml:space="preserve">INSERTED </w:t></w:r></w:ins>`;
  const del = `<w:del w:id="200" w:author="Alice" w:date="2026-04-17T10:01:00Z"><w:r><w:delText xml:space="preserve">REMOVED </w:delText></w:r></w:del>`;
  const plain = `<w:r><w:t xml:space="preserve">tail.</w:t></w:r>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body><w:p>${ins}${del}${plain}</w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
}

async function loadAgent(): Promise<DocxAgent> {
  const buf = await makeSyntheticDocx({ documentXml: trackedChangesXml() });
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

function findRevisions(snap: { root: { body: ReadonlyArray<unknown> } }): RevisionWrapper[] {
  const out: RevisionWrapper[] = [];
  for (const b of snap.root.body) {
    const block = b as { kind: string; children?: ReadonlyArray<{ kind: string; revisionId?: string }> };
    if (block.kind !== "paragraph") continue;
    for (const c of block.children ?? []) {
      if (c.kind === "revision") out.push(c as unknown as RevisionWrapper);
    }
  }
  return out;
}

describe("docx tracked-change resolution", () => {
  it("accept-change on a w:ins folds the inserted runs into the paragraph", async () => {
    const agent = await loadAgent();
    expect(findRevisions(agent.getSnapshot())).toHaveLength(2);
    const m = await agent.applyCommand({
      type: "docx:accept-change",
      payload: { revisionId: "100" },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const snap = agent.getSnapshot();
    expect(snap.dirty.body).toBe(true);
    const revs = findRevisions(snap);
    expect(revs.find((r) => r.revisionId === "100")).toBeUndefined();
    // The previously-wrapped run survives at the paragraph level.
    const p0 = snap.root.body[0];
    if (p0.kind !== "paragraph") throw new Error();
    const text = paragraphPlainText(p0);
    expect(text).toBe("INSERTED REMOVED tail.");
  });

  it("accept-change on a w:del removes the deleted runs entirely", async () => {
    const agent = await loadAgent();
    const m = await agent.applyCommand({
      type: "docx:accept-change",
      payload: { revisionId: "200" },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const snap = agent.getSnapshot();
    const revs = findRevisions(snap);
    expect(revs.find((r) => r.revisionId === "200")).toBeUndefined();
    const p0 = snap.root.body[0];
    if (p0.kind !== "paragraph") throw new Error();
    expect(paragraphPlainText(p0)).toBe("INSERTED tail.");
  });

  it("reject-change on a w:ins drops the inserted runs entirely", async () => {
    const agent = await loadAgent();
    const m = await agent.applyCommand({
      type: "docx:reject-change",
      payload: { revisionId: "100" },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const snap = agent.getSnapshot();
    const revs = findRevisions(snap);
    expect(revs.find((r) => r.revisionId === "100")).toBeUndefined();
    const p0 = snap.root.body[0];
    if (p0.kind !== "paragraph") throw new Error();
    expect(paragraphPlainText(p0)).toBe("REMOVED tail.");
  });

  it("reject-change on a w:del unwraps the deletion (text returns)", async () => {
    const agent = await loadAgent();
    const m = await agent.applyCommand({
      type: "docx:reject-change",
      payload: { revisionId: "200" },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const snap = agent.getSnapshot();
    const revs = findRevisions(snap);
    expect(revs.find((r) => r.revisionId === "200")).toBeUndefined();
    const p0 = snap.root.body[0];
    if (p0.kind !== "paragraph") throw new Error();
    expect(paragraphPlainText(p0)).toBe("INSERTED REMOVED tail.");
  });

  it("rejects unknown revision ids with code unknown-revision", async () => {
    const agent = await loadAgent();
    const m1 = await agent.applyCommand({
      type: "docx:accept-change",
      payload: { revisionId: "999" },
      source: "human",
    });
    expect(m1.status).toBe("rejected");
    expect(m1.rejection?.code).toBe("unknown-revision");

    const m2 = await agent.applyCommand({
      type: "docx:reject-change",
      payload: { revisionId: "" },
      source: "human",
    });
    expect(m2.status).toBe("rejected");
    expect(m2.rejection?.code).toBe("unknown-revision");
  });

  it("once a revision is resolved a second accept on the same id is unknown-revision (idempotence in the loud sense)", async () => {
    const agent = await loadAgent();
    const m1 = await agent.applyCommand({
      type: "docx:accept-change",
      payload: { revisionId: "100" },
      source: "human",
    });
    expect(m1.status).toBe("approved");
    const m2 = await agent.applyCommand({
      type: "docx:accept-change",
      payload: { revisionId: "100" },
      source: "human",
    });
    expect(m2.status).toBe("rejected");
    expect(m2.rejection?.code).toBe("unknown-revision");
  });

  it("round-trip: serialize → re-parse leaves no RevisionWrapper for the resolved id", async () => {
    const agent = await loadAgent();
    await agent.applyCommand({
      type: "docx:accept-change",
      payload: { revisionId: "100" },
      source: "human",
    });
    const buf = await agent.exportFile();
    const reparsed = await parseDocx(buf);
    const revs = findRevisions(reparsed);
    expect(revs.find((r) => r.revisionId === "100")).toBeUndefined();
    // The other revision is still present.
    expect(revs.find((r) => r.revisionId === "200")).toBeTruthy();
  });

  it("round-trip: model equality through parse(serialize(s)) — text matches the in-memory snapshot", async () => {
    const agent = await loadAgent();
    await agent.applyCommand({
      type: "docx:reject-change",
      payload: { revisionId: "200" },
      source: "human",
    });
    const expected = paragraphPlainText(
      (() => {
        const p = agent.getSnapshot().root.body[0];
        if (p.kind !== "paragraph") throw new Error();
        return p;
      })()
    );
    const buf = await agent.exportFile();
    const reparsed = await parseDocx(buf);
    const p0 = reparsed.root.body[0];
    if (p0.kind !== "paragraph") throw new Error();
    expect(paragraphPlainText(p0)).toBe(expected);
  });
});
