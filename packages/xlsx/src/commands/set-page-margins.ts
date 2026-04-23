import { CommandError, type CommandHandler } from "@officeai/core";
import type { Sheet, XlsxSnapshot } from "../model/types.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type { SetPageMarginsPayload } from "./payloads.js";
import { resolveSheet } from "./validation.js";

/**
 * `xlsx:set-page-margins` — overwrite the worksheet's `<pageMargins>`
 * element. Mirrors Page Layout → Margins.
 *
 * Three presets (`normal` / `wide` / `narrow`) match Excel's built-in
 * picker. Individual `*In` overrides take precedence after the preset
 * is applied. Pass `clear: true` to drop the element entirely.
 *
 * The `<pageMargins>` element is required to have all six attributes,
 * so the handler always emits a complete element when present (Excel
 * tolerates missing attrs but PowerPoint-style strict consumers don't).
 */
export const setPageMarginsHandler: CommandHandler<SetPageMarginsPayload, XlsxSnapshot> = {
  type: "xlsx:set-page-margins",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);

    if (payload.clear) {
      if (!sheet.pageMarginsXml) {
        return { next: snapshot, diff: buildDiff(snapshot.revision, snapshot.revision, []) };
      }
      const next: Sheet = { ...sheet, pageMarginsXml: undefined };
      return commit(snapshot, sheet, next, "cleared");
    }

    const current = parseAttrs(sheet.pageMarginsXml) ?? defaults();
    const merged = { ...current };
    if (payload.preset) {
      const preset = PRESETS[payload.preset];
      Object.assign(merged, preset);
    }
    if (payload.leftIn !== undefined) merged.left = clampMargin(payload.leftIn, "leftIn");
    if (payload.rightIn !== undefined) merged.right = clampMargin(payload.rightIn, "rightIn");
    if (payload.topIn !== undefined) merged.top = clampMargin(payload.topIn, "topIn");
    if (payload.bottomIn !== undefined) merged.bottom = clampMargin(payload.bottomIn, "bottomIn");
    if (payload.headerIn !== undefined) merged.header = clampMargin(payload.headerIn, "headerIn");
    if (payload.footerIn !== undefined) merged.footer = clampMargin(payload.footerIn, "footerIn");

    const xml = buildElement(merged);
    if (xml === sheet.pageMarginsXml) {
      return { next: snapshot, diff: buildDiff(snapshot.revision, snapshot.revision, []) };
    }
    const next: Sheet = { ...sheet, pageMarginsXml: xml };
    return commit(snapshot, sheet, next, payload.preset ?? "patched");
  },
};

interface Margins {
  left: number;
  right: number;
  top: number;
  bottom: number;
  header: number;
  footer: number;
}

const PRESETS: Record<string, Margins> = {
  normal: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
  wide: { left: 1.0, right: 1.0, top: 1.0, bottom: 1.0, header: 0.5, footer: 0.5 },
  narrow: { left: 0.25, right: 0.25, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
};

function defaults(): Margins {
  return { ...PRESETS.normal };
}

function clampMargin(v: number, name: string): number {
  if (!Number.isFinite(v) || v < 0) {
    throw new CommandError(
      "invalid-payload",
      `set-page-margins: ${name} must be a non-negative number (got ${v}).`
    );
  }
  return v;
}

function parseAttrs(xml: string | undefined): Margins | null {
  if (!xml) return null;
  const out: Partial<Margins> = {};
  const attrRe = /\b([A-Za-z]+)\s*=\s*"([^"]*)"/g;
  let m: RegExpExecArray | null;
  while ((m = attrRe.exec(xml)) !== null) {
    const key = m[1] as keyof Margins;
    const num = Number(m[2]);
    if (!Number.isFinite(num)) continue;
    out[key] = num;
  }
  return {
    left: out.left ?? 0.7,
    right: out.right ?? 0.7,
    top: out.top ?? 0.75,
    bottom: out.bottom ?? 0.75,
    header: out.header ?? 0.3,
    footer: out.footer ?? 0.3,
  };
}

function buildElement(m: Margins): string {
  const fmt = (v: number) => formatNum(v);
  return `<pageMargins left="${fmt(m.left)}" right="${fmt(m.right)}" top="${fmt(m.top)}" bottom="${fmt(m.bottom)}" header="${fmt(m.header)}" footer="${fmt(m.footer)}"/>`;
}

function formatNum(v: number): string {
  // Excel writes inches as plain decimals without trailing zeros.
  if (Number.isInteger(v)) return String(v);
  return Number(v.toFixed(6)).toString();
}

function commit(
  snapshot: XlsxSnapshot,
  before: Sheet,
  after: Sheet,
  summary: string
): { next: XlsxSnapshot; diff: ReturnType<typeof buildDiff> } {
  const root = replaceSheet(snapshot.root, after);
  const evolved = evolveSnapshot(snapshot, root, { sheets: [before.partPath] });
  return {
    next: evolved,
    diff: buildDiff(snapshot.revision, evolved.revision, [
      {
        kind: "node-updated",
        nodeId: before.id,
        path: ["sheets", before.index, "pageMargins"],
        field: "pageMargins",
        summary,
      },
    ]),
  };
}
