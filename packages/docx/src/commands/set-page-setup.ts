import { CommandError, type CommandHandler } from "@officeai/core";
import type {
  BlockNode,
  DocxSnapshot,
  PageMargins,
  PageSize,
  SectionBreak,
  SectionProperties,
} from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import type { SetPageSetupPayload } from "./payloads.js";

/**
 * B3 — Page Setup.
 *
 * Mutates the typed `pgSz` / `pgMar` of the section that owns the
 * paragraph at `paragraphIndex`. Drops the section's `raw` cache so
 * the serializer rebuilds `<w:sectPr>` from the typed projection;
 * untouched sections still round-trip byte-identical.
 *
 * The handler is intentionally permissive: any subset of the
 * `pgSz` / `pgMar` fields may be supplied; omitted fields keep
 * their current value. Numeric fields are validated against
 * Word's legal twip range so a pasted-in value never produces an
 * invalid OOXML document.
 */
export const setPageSetupHandler: CommandHandler<SetPageSetupPayload, DocxSnapshot> = {
  type: "docx:set-page-setup",
  apply(snapshot, payload) {
    const { paragraphIndex, pgSz: pgSzPatch, pgMar: pgMarPatch } = payload;
    if (!Number.isInteger(paragraphIndex) || paragraphIndex < 0) {
      throw new CommandError(
        "invalid-payload",
        `paragraphIndex must be a non-negative integer (got ${paragraphIndex})`
      );
    }
    if (!pgSzPatch && !pgMarPatch) {
      throw new CommandError("invalid-payload", "set-page-setup requires at least one of pgSz / pgMar");
    }

    const located = findOwningSection(snapshot, paragraphIndex);
    if (!located) {
      throw new CommandError(
        "unknown-target",
        `no section found at or after paragraph ${paragraphIndex} (body has ${snapshot.root.body.length} blocks)`
      );
    }

    const currentPgSz: PageSize = located.section.properties.pgSz ?? {
      w: 12240,
      h: 15840,
    };
    const currentPgMar: PageMargins = located.section.properties.pgMar ?? {
      top: 1440,
      right: 1440,
      bottom: 1440,
      left: 1440,
      header: 720,
      footer: 720,
    };

    const nextPgSz: PageSize = pgSzPatch
      ? validatePgSz({
          w: pgSzPatch.w ?? currentPgSz.w,
          h: pgSzPatch.h ?? currentPgSz.h,
          orient: pgSzPatch.orient ?? currentPgSz.orient,
        })
      : currentPgSz;

    const nextPgMar: PageMargins = pgMarPatch
      ? validatePgMar(
          {
            top: pgMarPatch.top ?? currentPgMar.top,
            right: pgMarPatch.right ?? currentPgMar.right,
            bottom: pgMarPatch.bottom ?? currentPgMar.bottom,
            left: pgMarPatch.left ?? currentPgMar.left,
            header: pgMarPatch.header ?? currentPgMar.header,
            footer: pgMarPatch.footer ?? currentPgMar.footer,
            gutter: pgMarPatch.gutter ?? currentPgMar.gutter,
          },
          nextPgSz
        )
      : currentPgMar;

    if (pageSizeEqual(currentPgSz, nextPgSz) && pageMarginsEqual(currentPgMar, nextPgMar)) {
      return {
        next: snapshot,
        diff: buildDiff(snapshot.revision, snapshot.revision, {
          kind: "node-updated",
          nodeId: located.section.id,
          path: ["body", located.index],
          field: "pageSetup",
          summary: "no-op (geometry unchanged)",
        }),
      };
    }

    const nextProps: SectionProperties = {
      ...located.section.properties,
      pgSz: nextPgSz,
      pgMar: nextPgMar,
    };
    const updated: SectionBreak = {
      ...located.section,
      properties: nextProps,
      raw: undefined,
    };

    const newBody: BlockNode[] = snapshot.root.body.slice();
    newBody[located.index] = updated;
    const next = evolveSnapshot(snapshot, { ...snapshot.root, body: newBody }, { body: true });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: located.section.id,
        path: ["body", located.index],
        field: "pageSetup",
        summary: summary(currentPgSz, nextPgSz, currentPgMar, nextPgMar),
      }),
    };
  },
};

interface LocatedSection {
  readonly index: number;
  readonly section: SectionBreak;
}

function findOwningSection(snapshot: DocxSnapshot, paragraphIndex: number): LocatedSection | null {
  const body = snapshot.root.body;
  for (let i = paragraphIndex; i < body.length; i++) {
    const block = body[i];
    if (block.kind === "section-break") return { index: i, section: block };
  }
  for (let i = body.length - 1; i >= 0; i--) {
    const block = body[i];
    if (block.kind === "section-break") return { index: i, section: block };
  }
  return null;
}

// Word allows page sizes between roughly 1/8" and 22" — anything
// outside that range is almost always a paste error from a different
// units system, so we hard-clamp.
const MIN_PAGE_DIM = 180; // 1/8"
const MAX_PAGE_DIM = 31680; // 22"

function validatePgSz(s: PageSize): PageSize {
  const w = clamp(s.w, MIN_PAGE_DIM, MAX_PAGE_DIM);
  const h = clamp(s.h, MIN_PAGE_DIM, MAX_PAGE_DIM);
  return { w, h, orient: s.orient };
}

function validatePgMar(m: PageMargins, page: PageSize): PageMargins {
  const top = clamp(m.top, 0, page.h - 360);
  const bottom = clamp(m.bottom, 0, page.h - top - 360);
  const left = clamp(m.left, 0, page.w - 360);
  const right = clamp(m.right, 0, page.w - left - 360);
  const header = clamp(m.header, 0, page.h);
  const footer = clamp(m.footer, 0, page.h);
  const gutter = m.gutter !== undefined ? clamp(m.gutter, 0, page.w) : undefined;
  return { top, right, bottom, left, header, footer, gutter };
}

function clamp(n: number, lo: number, hi: number): number {
  if (!Number.isFinite(n)) return lo;
  if (n < lo) return lo;
  if (n > hi) return hi;
  return Math.round(n);
}

function pageSizeEqual(a: PageSize, b: PageSize): boolean {
  return a.w === b.w && a.h === b.h && (a.orient ?? null) === (b.orient ?? null);
}

function pageMarginsEqual(a: PageMargins, b: PageMargins): boolean {
  return (
    a.top === b.top &&
    a.right === b.right &&
    a.bottom === b.bottom &&
    a.left === b.left &&
    a.header === b.header &&
    a.footer === b.footer &&
    (a.gutter ?? undefined) === (b.gutter ?? undefined)
  );
}

function summary(prevSz: PageSize, nextSz: PageSize, prevMar: PageMargins, nextMar: PageMargins): string {
  const parts: string[] = [];
  if (!pageSizeEqual(prevSz, nextSz)) {
    parts.push(`pgSz: ${prevSz.w}×${prevSz.h} → ${nextSz.w}×${nextSz.h}`);
  }
  if (!pageMarginsEqual(prevMar, nextMar)) {
    parts.push(
      `pgMar: ${prevMar.top}/${prevMar.right}/${prevMar.bottom}/${prevMar.left} → ${nextMar.top}/${nextMar.right}/${nextMar.bottom}/${nextMar.left}`
    );
  }
  return parts.join(" · ");
}
