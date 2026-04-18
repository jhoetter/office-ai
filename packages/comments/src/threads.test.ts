import { describe, expect, it } from "vitest";
import { countOpenThreads, groupThreads } from "./threads.js";
import type { CommentBody } from "./types.js";

function c(id: string, parentId?: string, opts?: Partial<CommentBody>): CommentBody {
  return {
    id,
    author: opts?.author ?? "alice",
    text: opts?.text ?? `body of ${id}`,
    anchor: opts?.anchor ?? { kind: "none" },
    ...(parentId ? { parentId } : {}),
    ...(opts?.resolved !== undefined ? { resolved: opts.resolved } : {}),
  };
}

describe("groupThreads", () => {
  it("preserves the input order of top-level comments", () => {
    const out = groupThreads([c("a"), c("b"), c("c")]);
    expect(out.map((t) => t.parent.id)).toEqual(["a", "b", "c"]);
  });

  it("nests replies under their parent in input order", () => {
    const out = groupThreads([c("a"), c("a1", "a"), c("b"), c("a2", "a")]);
    expect(out).toHaveLength(2);
    expect(out[0]!.parent.id).toBe("a");
    expect(out[0]!.replies.map((r) => r.id)).toEqual(["a1", "a2"]);
  });

  it("surfaces orphan replies as their own top-level threads", () => {
    const out = groupThreads([c("a"), c("ghost1", "missing")]);
    expect(out.map((t) => t.parent.id)).toEqual(["a", "ghost1"]);
  });

  it("counts only unresolved parents as open", () => {
    const out = groupThreads([c("a", undefined, { resolved: true }), c("b")]);
    expect(countOpenThreads(out)).toBe(1);
  });
});
