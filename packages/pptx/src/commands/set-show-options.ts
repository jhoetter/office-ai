import type { CommandHandler } from "@officeai/core";
import type { OpaqueXml, PptxPresentation, PptxSnapshot } from "../model/types.js";
import { ATTR_PREFIX } from "../parser/xml-helpers.js";
import { buildDiff, evolveSnapshot, makeError } from "./helpers.js";
import type { SetShowOptionsPayload } from "./payloads.js";

/**
 * Configure presentation-wide slideshow options. Mirrors PowerPoint's
 * "Set Up Slide Show" dialog by patching the `<p:showPr>` element on
 * `ppt/presentation.xml`.
 *
 * Implementation strategy: locate `<p:showPr>` inside
 * `presentationOpaqueTail`. If absent, append a new entry just before
 * `<p:extLst>` (or at the end of the tail). Replace its attributes and
 * the showType child element (`<p:browse/>` / `<p:kiosk/>`).
 *
 * The presentation serializer rebuilds the tail verbatim (except for
 * `<p:sldIdLst>` and `<p:sldSz>`), so the new opaque entry round-trips
 * correctly without further serializer changes.
 */
export const setShowOptionsHandler: CommandHandler<SetShowOptionsPayload, PptxSnapshot> = {
  type: "pptx:set-show-options",
  apply(snapshot, payload) {
    const tail = snapshot.root.presentationOpaqueTail;
    const existingIdx = tail.findIndex((o) => o.tag === "p:showPr");

    if (payload.clear) {
      if (existingIdx === -1) {
        return { next: snapshot, diff: buildDiff(snapshot.revision, snapshot.revision) };
      }
      const nextTail: OpaqueXml[] = [...tail];
      nextTail.splice(existingIdx, 1);
      return commit(snapshot, nextTail, "cleared");
    }

    if (
      payload.showType === undefined &&
      payload.loop === undefined &&
      payload.showNarration === undefined &&
      payload.showAnimation === undefined &&
      payload.useTimings === undefined
    ) {
      throw makeError(
        "invalid-payload",
        "set-show-options: must supply at least one of showType/loop/showNarration/showAnimation/useTimings, or set clear:true",
      );
    }

    const existing = existingIdx === -1 ? null : tail[existingIdx];
    const next = buildShowPr(existing, payload);

    if (existing && shallowEqual(existing, next)) {
      return { next: snapshot, diff: buildDiff(snapshot.revision, snapshot.revision) };
    }

    const nextTail: OpaqueXml[] = [...tail];
    if (existingIdx === -1) {
      const insertAt = preferredInsertIndex(tail);
      nextTail.splice(insertAt, 0, next);
    } else {
      nextTail[existingIdx] = next;
    }
    return commit(snapshot, nextTail, summarise(payload));
  },
};

function commit(
  snapshot: PptxSnapshot,
  nextTail: OpaqueXml[],
  summary: string,
): { next: PptxSnapshot; diff: ReturnType<typeof buildDiff> } {
  const root: PptxPresentation = { ...snapshot.root, presentationOpaqueTail: nextTail };
  const evolved = evolveSnapshot(snapshot, root, { presentation: true });
  return {
    next: evolved,
    diff: buildDiff(snapshot.revision, evolved.revision, {
      kind: "node-updated",
      nodeId: snapshot.root.id,
      path: ["presentation", "showPr"],
      field: "showPr",
      summary,
    }),
  };
}

function buildShowPr(prev: OpaqueXml | null, payload: SetShowOptionsPayload): OpaqueXml {
  const prevAttrs = prev?.attrs ?? {};
  const attrs: Record<string, string> = { ...prevAttrs };

  applyBoolAttr(attrs, "loop", payload.loop);
  applyBoolAttr(attrs, "showNarration", payload.showNarration);
  applyBoolAttr(attrs, "showAnimation", payload.showAnimation);
  applyBoolAttr(attrs, "useTimings", payload.useTimings);

  // Drop any prior showType child elements; rebuild from showType.
  const subtree: unknown[] = [];
  if (prev) {
    for (const c of prev.subtree) {
      if (!c || typeof c !== "object" || Array.isArray(c)) {
        subtree.push(c);
        continue;
      }
      const obj = c as Record<string, unknown>;
      const keys = Object.keys(obj).filter((k) => k !== ":@");
      const tag = keys[0];
      if (tag === "p:browse" || tag === "p:kiosk" || tag === "p:present") continue;
      subtree.push(c);
    }
  }
  const effectiveShowType = payload.showType ?? inferShowType(prev);
  if (effectiveShowType === "browse") {
    subtree.unshift({ "p:browse": [] });
  } else if (effectiveShowType === "kiosk") {
    subtree.unshift({ "p:kiosk": [] });
  }
  // "presenter" emits no showType child element (PowerPoint default).

  const rawAttrs: Record<string, string> = {};
  for (const [k, v] of Object.entries(attrs)) {
    rawAttrs[`${ATTR_PREFIX}${k}`] = v;
  }

  return {
    tag: "p:showPr",
    attrs,
    rawAttrs,
    subtree,
  };
}

function inferShowType(prev: OpaqueXml | null): "presenter" | "browse" | "kiosk" {
  if (!prev) return "presenter";
  for (const c of prev.subtree) {
    if (!c || typeof c !== "object" || Array.isArray(c)) continue;
    const obj = c as Record<string, unknown>;
    const keys = Object.keys(obj).filter((k) => k !== ":@");
    if (keys[0] === "p:browse") return "browse";
    if (keys[0] === "p:kiosk") return "kiosk";
  }
  return "presenter";
}

function applyBoolAttr(
  attrs: Record<string, string>,
  name: string,
  value: boolean | undefined,
): void {
  if (value === undefined) return;
  if (value) {
    attrs[name] = "1";
  } else {
    delete attrs[name];
  }
}

function preferredInsertIndex(tail: ReadonlyArray<OpaqueXml>): number {
  // p:showPr should appear before p:extLst per ECMA-376 schema order.
  const extIdx = tail.findIndex((o) => o.tag === "p:extLst");
  if (extIdx >= 0) return extIdx;
  return tail.length;
}

function shallowEqual(a: OpaqueXml, b: OpaqueXml): boolean {
  if (a.tag !== b.tag) return false;
  const ak = Object.keys(a.attrs);
  const bk = Object.keys(b.attrs);
  if (ak.length !== bk.length) return false;
  for (const k of ak) {
    if (a.attrs[k] !== b.attrs[k]) return false;
  }
  // Compare subtree by structural JSON; tail items here are tiny.
  return JSON.stringify(a.subtree) === JSON.stringify(b.subtree);
}

function summarise(p: SetShowOptionsPayload): string {
  const parts: string[] = [];
  if (p.showType) parts.push(p.showType);
  if (p.loop !== undefined) parts.push(`loop=${p.loop}`);
  if (p.showNarration !== undefined) parts.push(`narr=${p.showNarration}`);
  if (p.showAnimation !== undefined) parts.push(`anim=${p.showAnimation}`);
  if (p.useTimings !== undefined) parts.push(`timings=${p.useTimings}`);
  return parts.join(",");
}
