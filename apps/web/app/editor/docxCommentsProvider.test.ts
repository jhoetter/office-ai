import { describe, expect, it } from "vitest";
import { DocxAgent } from "@officeai/docx";
import { deterministicIdMinter } from "@officeai/core";
import { createDocxCommentsProvider } from "./docxCommentsProvider";

/**
 * Realtime identity (peer id + colour) is *not* persisted into OOXML —
 * the comments provider keeps an in-memory map and merges it back into
 * the `CommentBody` snapshot returned by `threads()`. These tests pin
 * that contract so the comments sidebar can keep showing the author
 * swatch + stable id even though the underlying `DocxAgent` only
 * remembers the human-readable display name.
 */
describe("createDocxCommentsProvider — author identity stamping", () => {
  async function loadProvider() {
    const agent = await DocxAgent.empty({ idMinter: deterministicIdMinter() });
    // `empty()` gives us a single empty paragraph; insert a few chars so
    // the docx-range anchor below resolves to a real run.
    await agent.applyCommand({
      type: "docx:insert-text",
      payload: {
        position: { paragraph: 0, run: 0, offset: 0 },
        text: "hello world",
      },
      source: "human",
    });
    return { agent, provider: createDocxCommentsProvider({ agent }) };
  }

  const range = {
    start: { paragraph: 0, run: 0, offset: 0 },
    end: { paragraph: 0, run: 0, offset: 5 },
  } as const;

  it("hydrates authorId / authorColor on the next threads() snapshot", async () => {
    const { provider } = await loadProvider();
    const id = await provider.add({
      author: "faithful-dog",
      text: "first take",
      anchor: { kind: "docx-range", paragraphIndex: 0, range },
      authorId: "peer-42",
      authorColor: "#ff8800",
    });
    expect(id).toBeTruthy();

    const threads = provider.threads();
    expect(threads).toHaveLength(1);
    const root = threads[0]!.parent;
    expect(root.author).toBe("faithful-dog");
    expect(root.authorId).toBe("peer-42");
    expect(root.authorColor).toBe("#ff8800");
  });

  it("stamps replies with their own identity (independent of the parent)", async () => {
    const { provider } = await loadProvider();
    const parentId = await provider.add({
      author: "faithful-dog",
      text: "first take",
      anchor: { kind: "docx-range", paragraphIndex: 0, range },
      authorId: "peer-42",
      authorColor: "#ff8800",
    });
    await provider.reply({
      parentId,
      author: "curious-otter",
      text: "+1",
      authorId: "peer-99",
      authorColor: "#00aaff",
    });

    const threads = provider.threads();
    expect(threads).toHaveLength(1);
    const replies = threads[0]!.replies;
    expect(replies).toHaveLength(1);
    expect(replies[0]!.author).toBe("curious-otter");
    expect(replies[0]!.authorId).toBe("peer-99");
    expect(replies[0]!.authorColor).toBe("#00aaff");
  });

  it("leaves identity fields undefined when no realtime identity is supplied", async () => {
    const { provider } = await loadProvider();
    await provider.add({
      author: "Anonymous",
      text: "no peer id here",
      anchor: { kind: "docx-range", paragraphIndex: 0, range },
    });
    const threads = provider.threads();
    expect(threads).toHaveLength(1);
    expect(threads[0]!.parent.authorId).toBeUndefined();
    expect(threads[0]!.parent.authorColor).toBeUndefined();
  });
});
