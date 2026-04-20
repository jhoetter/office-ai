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

  it("mints a unique revision id when the caller omits one and there is no neighbour to coalesce with", async () => {
    // Two inserts by *different* authors so they cannot coalesce
    // into a single `<w:ins>` wrapper. The second call omits a
    // revisionId, so the handler must mint the smallest unused
    // positive integer ("1" — author "U" already pinned "5").
    const agent = await loadAgent("Hi.");
    await agent.applyCommand({
      type: "docx:insert-text-tracked",
      payload: { at: { paragraph: 0, run: 0, offset: 0 }, text: "A", author: "U", revisionId: "5" },
      source: "human",
    });
    await agent.applyCommand({
      type: "docx:insert-text-tracked",
      payload: { at: { paragraph: 0, run: 0, offset: 0 }, text: "B", author: "V" },
      source: "human",
    });
    const ids = collectRevisions(firstParagraph(agent.getSnapshot()))
      .map((r) => r.revisionId)
      .sort();
    expect(ids).toContain("5");
    expect(ids).toContain("1");
    expect(new Set(ids).size).toBe(ids.length);
  });

  it("coalesces consecutive same-author inserts at the same boundary into one wrapper (Word-style typing)", async () => {
    // Simulates the PM funnel: every keystroke is its own
    // transaction → its own `docx:insert-text-tracked` command, no
    // explicit revisionId. Word merges these into a single revision
    // (one balloon, one accept/reject) and so should we.
    const agent = await loadAgent("Hi.");
    const author = "Alice";
    for (let i = 0; i < "test".length; i++) {
      await agent.applyCommand({
        type: "docx:insert-text-tracked",
        payload: { at: { paragraph: 0, offset: i }, text: "test"[i], author },
        source: "human",
      });
    }
    const revs = collectRevisions(firstParagraph(agent.getSnapshot()));
    expect(revs).toHaveLength(1);
    expect(revs[0].revisionType).toBe("ins");
    expect(revs[0].author).toBe(author);
    // The merged wrapper carries the full word in its inner run.
    const innerRun = revs[0].children[0] as Run;
    const fullText = innerRun.children.reduce(
      (acc, leaf) => (leaf.kind === "text" ? acc + leaf.text : acc),
      ""
    );
    expect(fullText).toBe("test");
    expect(paragraphPlainText(firstParagraph(agent.getSnapshot()))).toBe("testHi.");
  });

  it("does not coalesce across different authors", async () => {
    const agent = await loadAgent("Hi.");
    await agent.applyCommand({
      type: "docx:insert-text-tracked",
      payload: { at: { paragraph: 0, offset: 0 }, text: "A", author: "Alice" },
      source: "human",
    });
    await agent.applyCommand({
      type: "docx:insert-text-tracked",
      payload: { at: { paragraph: 0, offset: 1 }, text: "B", author: "Bob" },
      source: "human",
    });
    const revs = collectRevisions(firstParagraph(agent.getSnapshot()));
    expect(revs).toHaveLength(2);
    expect(new Set(revs.map((r) => r.author))).toEqual(new Set(["Alice", "Bob"]));
  });

  it("honours an explicit revisionId by always creating a fresh wrapper (no coalescing)", async () => {
    // Programmatic callers (agent rewrites, tests, etc.) that pin a
    // revisionId are explicitly addressing a distinct wrapper;
    // coalescing would silently merge them and break round-trip.
    const agent = await loadAgent("Hi.");
    await agent.applyCommand({
      type: "docx:insert-text-tracked",
      payload: { at: { paragraph: 0, offset: 0 }, text: "A", author: "U", revisionId: "10" },
      source: "human",
    });
    await agent.applyCommand({
      type: "docx:insert-text-tracked",
      payload: { at: { paragraph: 0, offset: 1 }, text: "B", author: "U", revisionId: "11" },
      source: "human",
    });
    const revs = collectRevisions(firstParagraph(agent.getSnapshot()));
    expect(revs).toHaveLength(2);
    expect(revs.map((r) => r.revisionId).sort()).toEqual(["10", "11"]);
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

  it("coalesces consecutive same-author backspaces into a single <w:del> wrapper", async () => {
    // Simulates the PM funnel for backspacing one character at a
    // time. Each backspace becomes its own
    // `docx:delete-range-tracked` command without a pinned
    // revisionId. Word merges these into a single revision (one
    // balloon, one accept/reject) and so should we.
    //
    // Initial paragraph "Hello world." has visible length 12;
    // paragraph offsets count wrapper text, so even after the first
    // wrap the offset of "d" in "world" stays at 10 because the
    // wrapper for "." still contributes one visible character.
    const agent = await loadAgent("Hello world.");
    const author = "Alice";
    await agent.applyCommand({
      type: "docx:delete-range-tracked",
      payload: {
        range: { start: { paragraph: 0, offset: 11 }, end: { paragraph: 0, offset: 12 } },
        author,
      },
      source: "human",
    });
    await agent.applyCommand({
      type: "docx:delete-range-tracked",
      payload: {
        range: { start: { paragraph: 0, offset: 10 }, end: { paragraph: 0, offset: 11 } },
        author,
      },
      source: "human",
    });
    const revs = collectRevisions(firstParagraph(agent.getSnapshot()));
    expect(revs).toHaveLength(1);
    expect(revs[0].revisionType).toBe("del");
    expect(revs[0].author).toBe(author);
    // Inner runs concatenate to the document-order text "d." — the
    // exact contiguous range the user removed.
    const merged = revs[0].children
      .filter((c): c is Run => c.kind === "run")
      .flatMap((r) => r.children)
      .filter((c): c is TextLeaf => c.kind === "text")
      .map((t) => t.text)
      .join("");
    expect(merged).toBe("d.");
    // Accept the merged wrapper and confirm the underlying range is
    // gone in one shot — proving the wrapper still round-trips
    // through the Accept handler unchanged.
    await agent.applyCommand({
      type: "docx:accept-change",
      payload: { revisionId: revs[0].revisionId },
      source: "human",
    });
    expect(paragraphPlainText(firstParagraph(agent.getSnapshot()))).toBe("Hello worl");
  });

  it("does not coalesce deletions across different authors", async () => {
    const agent = await loadAgent("Hello world.");
    await agent.applyCommand({
      type: "docx:delete-range-tracked",
      payload: {
        range: { start: { paragraph: 0, offset: 11 }, end: { paragraph: 0, offset: 12 } },
        author: "Alice",
      },
      source: "human",
    });
    await agent.applyCommand({
      type: "docx:delete-range-tracked",
      payload: {
        range: { start: { paragraph: 0, offset: 10 }, end: { paragraph: 0, offset: 11 } },
        author: "Bob",
      },
      source: "human",
    });
    const revs = collectRevisions(firstParagraph(agent.getSnapshot()));
    expect(revs).toHaveLength(2);
    expect(new Set(revs.map((r) => r.author))).toEqual(new Set(["Alice", "Bob"]));
  });

  it("does not coalesce deletions separated by visible content", async () => {
    // Two single-char deletes at opposite ends of the paragraph by
    // the same author. There's plenty of unwrapped text between them
    // so the coalescer must keep them as two separate balloons.
    const agent = await loadAgent("Hello world.");
    await agent.applyCommand({
      type: "docx:delete-range-tracked",
      payload: {
        range: { start: { paragraph: 0, offset: 11 }, end: { paragraph: 0, offset: 12 } },
        author: "Alice",
      },
      source: "human",
    });
    await agent.applyCommand({
      type: "docx:delete-range-tracked",
      payload: {
        range: { start: { paragraph: 0, offset: 0 }, end: { paragraph: 0, offset: 1 } },
        author: "Alice",
      },
      source: "human",
    });
    const revs = collectRevisions(firstParagraph(agent.getSnapshot()));
    expect(revs).toHaveLength(2);
  });

  it("honours an explicit revisionId by always creating a fresh <w:del> wrapper (no coalescing)", async () => {
    const agent = await loadAgent("Hello world.");
    await agent.applyCommand({
      type: "docx:delete-range-tracked",
      payload: {
        range: { start: { paragraph: 0, offset: 11 }, end: { paragraph: 0, offset: 12 } },
        author: "Alice",
        revisionId: "100",
      },
      source: "human",
    });
    await agent.applyCommand({
      type: "docx:delete-range-tracked",
      payload: {
        range: { start: { paragraph: 0, offset: 10 }, end: { paragraph: 0, offset: 11 } },
        author: "Alice",
        revisionId: "101",
      },
      source: "human",
    });
    const revs = collectRevisions(firstParagraph(agent.getSnapshot()));
    expect(revs).toHaveLength(2);
    expect(revs.map((r) => r.revisionId).sort()).toEqual(["100", "101"]);
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
