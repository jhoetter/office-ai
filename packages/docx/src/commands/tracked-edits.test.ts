import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { DocxAgent } from "../agent/agent.js";
import { parseDocx } from "../parser/parse.js";
import { paragraphPlainText } from "./helpers.js";
import { makeSyntheticDocx, DEFAULT_DOC_ROOT_ATTRS } from "../test-utils/synthetic.js";
import type { Paragraph, RevisionWrapper, Run, TextLeaf } from "../model/types.js";

/**
 * "Suggesting" mode round-trip suite. Verifies that the two new
 * tracked-edit commands (`docx:insert-text-tracked`,
 * `docx:delete-range-tracked`) produce snapshots that:
 *
 *   1. Carry a `RevisionWrapper` of the right type at the right
 *      paragraph-relative offset.
 *   2. Serialise to OOXML the same way Word writes
 *      `<w:ins>` / `<w:del>` (including `<w:delText>` for deletions).
 *   3. Round-trip through `accept-change` / `reject-change` so the
 *      "Suggesting" flow plugs into the existing tracked-changes UI
 *      with no further wiring.
 */

function singleParagraphXml(text: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
}

async function loadAgent(text: string): Promise<DocxAgent> {
  const buf = await makeSyntheticDocx({ documentXml: singleParagraphXml(text) });
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

function firstParagraph(snap: { root: { body: ReadonlyArray<unknown> } }): Paragraph {
  const block = snap.root.body[0] as Paragraph;
  if (block.kind !== "paragraph") {
    throw new Error("first body block is not a paragraph");
  }
  return block;
}

function collectRevisions(p: Paragraph): RevisionWrapper[] {
  return p.children.filter((c): c is RevisionWrapper => c.kind === "revision");
}

describe("docx:insert-text-tracked", () => {
  it("wraps the inserted text in a <w:ins> revision wrapper", async () => {
    const agent = await loadAgent("Hello world.");
    const m = await agent.applyCommand({
      type: "docx:insert-text-tracked",
      payload: {
        at: { paragraph: 0, run: 0, offset: 6 },
        text: "INSERTED ",
        author: "Alice",
        date: "2026-04-17T10:00:00Z",
        revisionId: "42",
      },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const p = firstParagraph(agent.getSnapshot());
    const revs = collectRevisions(p);
    expect(revs).toHaveLength(1);
    expect(revs[0]).toMatchObject({
      revisionType: "ins",
      author: "Alice",
      date: "2026-04-17T10:00:00Z",
      revisionId: "42",
    });
    // The wrapper sits between the two halves of the original run,
    // and reading the paragraph's plain text yields the merged
    // result so the editor preview matches what Word displays when
    // tracked changes are accepted.
    expect(paragraphPlainText(p)).toBe("Hello INSERTED world.");
  });

  it("mints a unique revision id when the caller omits one", async () => {
    const agent = await loadAgent("Hi.");
    await agent.applyCommand({
      type: "docx:insert-text-tracked",
      payload: { at: { paragraph: 0, run: 0, offset: 0 }, text: "A", author: "U", revisionId: "5" },
      source: "human",
    });
    await agent.applyCommand({
      type: "docx:insert-text-tracked",
      payload: { at: { paragraph: 0, run: 0, offset: 0 }, text: "B", author: "U" },
      source: "human",
    });
    const ids = collectRevisions(firstParagraph(agent.getSnapshot()))
      .map((r) => r.revisionId)
      .sort();
    expect(ids).toContain("5");
    // Second insert should have minted "1" (smallest unused positive int).
    expect(ids).toContain("1");
    expect(new Set(ids).size).toBe(ids.length);
  });

  it("round-trips via accept-change (suggested insert becomes plain text)", async () => {
    const agent = await loadAgent("Hello world.");
    await agent.applyCommand({
      type: "docx:insert-text-tracked",
      payload: {
        at: { paragraph: 0, run: 0, offset: 6 },
        text: "INSERTED ",
        author: "A",
        revisionId: "9",
      },
      source: "human",
    });
    expect(collectRevisions(firstParagraph(agent.getSnapshot()))).toHaveLength(1);

    await agent.applyCommand({
      type: "docx:accept-change",
      payload: { revisionId: "9" },
      source: "human",
    });
    const p = firstParagraph(agent.getSnapshot());
    expect(collectRevisions(p)).toHaveLength(0);
    expect(paragraphPlainText(p)).toBe("Hello INSERTED world.");
  });

  it("round-trips via reject-change (suggested insert disappears)", async () => {
    const agent = await loadAgent("Hello world.");
    await agent.applyCommand({
      type: "docx:insert-text-tracked",
      payload: {
        at: { paragraph: 0, run: 0, offset: 6 },
        text: "INSERTED ",
        author: "A",
        revisionId: "9",
      },
      source: "human",
    });
    await agent.applyCommand({
      type: "docx:reject-change",
      payload: { revisionId: "9" },
      source: "human",
    });
    const p = firstParagraph(agent.getSnapshot());
    expect(collectRevisions(p)).toHaveLength(0);
    expect(paragraphPlainText(p)).toBe("Hello world.");
  });
});

describe("docx:delete-range-tracked", () => {
  it("wraps the targeted range in a <w:del> with isDelText leaves", async () => {
    const agent = await loadAgent("Hello world.");
    await agent.applyCommand({
      type: "docx:delete-range-tracked",
      payload: {
        range: {
          start: { paragraph: 0, offset: 6 },
          end: { paragraph: 0, offset: 11 },
        },
        author: "Bob",
        date: "2026-04-17T10:01:00Z",
        revisionId: "77",
      },
      source: "human",
    });
    const p = firstParagraph(agent.getSnapshot());
    const revs = collectRevisions(p);
    expect(revs).toHaveLength(1);
    const del = revs[0];
    expect(del.revisionType).toBe("del");
    expect(del.author).toBe("Bob");
    expect(del.revisionId).toBe("77");

    const innerRun = del.children[0] as Run;
    expect(innerRun.kind).toBe("run");
    const leaf = innerRun.children[0] as TextLeaf;
    expect(leaf.kind).toBe("text");
    expect(leaf.text).toBe("world");
    expect(leaf.isDelText).toBe(true);

    // The paragraph's plain text still includes the deletion marker
    // because `paragraphPlainText` walks revision children — that's
    // the right choice for visual previews. Accept-change should
    // strip it.
    expect(paragraphPlainText(p)).toBe("Hello world.");
  });

  it("round-trips via accept-change (suggested delete actually removes the text)", async () => {
    const agent = await loadAgent("Hello world.");
    await agent.applyCommand({
      type: "docx:delete-range-tracked",
      payload: {
        range: {
          start: { paragraph: 0, offset: 5 },
          end: { paragraph: 0, offset: 11 },
        },
        author: "B",
        revisionId: "12",
      },
      source: "human",
    });
    await agent.applyCommand({
      type: "docx:accept-change",
      payload: { revisionId: "12" },
      source: "human",
    });
    const p = firstParagraph(agent.getSnapshot());
    expect(collectRevisions(p)).toHaveLength(0);
    expect(paragraphPlainText(p)).toBe("Hello.");
  });

  it("round-trips via reject-change (suggested delete reverts to plain text)", async () => {
    const agent = await loadAgent("Hello world.");
    await agent.applyCommand({
      type: "docx:delete-range-tracked",
      payload: {
        range: {
          start: { paragraph: 0, offset: 5 },
          end: { paragraph: 0, offset: 11 },
        },
        author: "B",
        revisionId: "12",
      },
      source: "human",
    });
    await agent.applyCommand({
      type: "docx:reject-change",
      payload: { revisionId: "12" },
      source: "human",
    });
    const p = firstParagraph(agent.getSnapshot());
    expect(collectRevisions(p)).toHaveLength(0);
    expect(paragraphPlainText(p)).toBe("Hello world.");
  });

  it("round-trips through serialize → parse with the wrappers preserved", async () => {
    // Two separate agents — chaining insert+delete on the same
    // paragraph is intentionally out of scope because the shared
    // `paragraphTextLength` offset model (also used by the
    // non-tracked `delete-range`) does not count text that lives
    // inside a `RevisionWrapper`. Both commands round-trip
    // independently, which is what matters for the agent →
    // serialize → reparse → agent loop.
    const insAgent = await loadAgent("Hello world.");
    await insAgent.applyCommand({
      type: "docx:insert-text-tracked",
      payload: {
        at: { paragraph: 0, run: 0, offset: 6 },
        text: "INSERTED ",
        author: "Alice",
        date: "2026-04-17T10:00:00Z",
        revisionId: "42",
      },
      source: "human",
    });
    {
      const buf = await insAgent.exportFile();
      const reparsed = await parseDocx(buf);
      const revs = collectRevisions(firstParagraph(reparsed));
      expect(revs).toHaveLength(1);
      expect(revs[0].revisionType).toBe("ins");
      expect(revs[0].revisionId).toBe("42");
      expect(revs[0].author).toBe("Alice");
    }

    const delAgent = await loadAgent("Hello world.");
    await delAgent.applyCommand({
      type: "docx:delete-range-tracked",
      payload: {
        range: {
          start: { paragraph: 0, offset: 6 },
          end: { paragraph: 0, offset: 11 },
        },
        author: "Alice",
        date: "2026-04-17T10:01:00Z",
        revisionId: "43",
      },
      source: "human",
    });
    {
      const buf = await delAgent.exportFile();
      const reparsed = await parseDocx(buf);
      const revs = collectRevisions(firstParagraph(reparsed));
      expect(revs).toHaveLength(1);
      expect(revs[0].revisionType).toBe("del");
      expect(revs[0].revisionId).toBe("43");
      expect(revs[0].author).toBe("Alice");
      // The wrapped run's text leaf must serialise as `<w:delText>`,
      // which the parser reads back with `isDelText: true`.
      const innerRun = revs[0].children[0] as Run;
      const leaf = innerRun.children[0] as TextLeaf;
      expect(leaf.isDelText).toBe(true);
      expect(leaf.text).toBe("world");
    }
  });

  it("rejects multi-paragraph ranges with not-implemented", async () => {
    const agent = await loadAgent("Hello world.");
    const m = await agent.applyCommand({
      type: "docx:delete-range-tracked",
      payload: {
        range: {
          start: { paragraph: 0, offset: 0 },
          end: { paragraph: 1, offset: 0 },
        },
        author: "B",
      },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("not-implemented");
  });
});
