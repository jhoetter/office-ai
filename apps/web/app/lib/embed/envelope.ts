import type { XlsxClipboardSnapshot } from "@officeai/xlsx";

/**
 * Cross-format embed envelope — the structured handoff we paint
 * onto the system clipboard whenever a copy in one editor should
 * survive a paste in a different editor with full semantic fidelity
 * (e.g. an XLSX range pasted into a DOCX as a real `<w:tbl>`,
 * not as a TSV string).
 *
 * Why a custom MIME type rather than relying on `text/html`?
 *   - The HTML fallback is what Word/PowerPoint *outside* our app
 *     consume, and the existing XLSX clipboard already paints it.
 *   - But our DOCX/PPTX editors operate on a typed snapshot, not
 *     on HTML — going via `text/html` would mean re-parsing the
 *     same data we just emitted, losing formulas, style ids, merge
 *     metadata, chart kind, etc.
 *   - The envelope keeps the structured data verbatim so the paste
 *     handler can dispatch a typed `docx:insert-table` /
 *     `pptx:replace-text` etc. directly.
 *
 * Wire format: a single JSON object written to the
 * `application/x-officeai-embed+json` clipboard MIME type, alongside
 * (not instead of) `text/html` and `text/plain` for external apps.
 *
 * Always on: an earlier iteration gated this behind
 * `NEXT_PUBLIC_OAI_EMBED`, but the user explicitly asked for the
 * cross-format embed to ship default-on (no env-var configs).
 * Receivers tolerate a missing envelope (they fall back to the
 * existing `text/html` / `text/plain` paths) so there's no harm in
 * always painting it onto the clipboard.
 */
export const EMBED_MIME = "application/x-officeai-embed+json";

/** Schema version. Bump on breaking shape changes; readers tolerate. */
export const EMBED_VERSION = 1;

/**
 * Discriminated union of payloads we know how to emit/consume.
 * Keep the variant names short and stable; they're persisted on
 * the system clipboard which can outlive the user's session.
 */
export type OfficeAIEmbedPayload =
  | XlsxRangeEmbed
  | XlsxChartImageEmbed
  | DocxTableEmbed;

/**
 * A copied XLSX range. The raw `XlsxClipboardSnapshot` is preserved
 * verbatim so a same-format paste back into XLSX still goes through
 * the lossless fingerprint path; cross-format pastes downgrade to
 * the row-major `cells` matrix and ignore formulas/styles.
 */
export interface XlsxRangeEmbed {
  readonly kind: "xlsx-range";
  readonly snapshot: XlsxClipboardSnapshot;
  /**
   * Human-readable origin (e.g. `"Sheet1!A1:C5"`). Only used for
   * toast messages and undo labels — never round-tripped back into
   * the snapshot.
   */
  readonly originLabel: string;
}

/**
 * A copied XLSX chart, rendered to PNG bytes (base64) so that a
 * paste into DOCX/PPTX can insert a picture even though the typed
 * chart model lives only in XLSX.
 *
 * `width` / `height` are in CSS pixels at the time of capture so
 * the receiver can pick a reasonable default canvas size.
 */
export interface XlsxChartImageEmbed {
  readonly kind: "xlsx-chart-image";
  readonly png: string;
  readonly width: number;
  readonly height: number;
  readonly chartKind: "column" | "bar" | "line" | "pie";
  readonly title?: string;
}

/**
 * D5 — a copied DOCX table, projected onto a 2D string matrix so a
 * paste into XLSX lands as a real range (cells with values) rather
 * than a pasted-as-text block of TSV. Each row's length matches the
 * widest row; sparse cells are empty strings.
 *
 * Only emitted when the entire ProseMirror selection sits inside a
 * single `<w:tbl>` — partial selections fall through to the default
 * PM serialiser (text / HTML) so cross-app pastes keep working.
 */
export interface DocxTableEmbed {
  readonly kind: "docx-table";
  readonly cells: ReadonlyArray<ReadonlyArray<string>>;
  /**
   * Human-readable origin (e.g. `"Document table"` or, when the
   * caller has a better label, the surrounding heading text). Only
   * used for toast messages and undo labels — never re-applied to
   * the pasted range.
   */
  readonly originLabel?: string;
}

export interface OfficeAIEmbedEnvelope {
  readonly type: "officeai/embed";
  readonly version: number;
  readonly source: "xlsx" | "docx" | "pptx";
  readonly createdAt: string;
  readonly payload: OfficeAIEmbedPayload;
}

export function makeEnvelope(
  source: OfficeAIEmbedEnvelope["source"],
  payload: OfficeAIEmbedPayload
): OfficeAIEmbedEnvelope {
  return {
    type: "officeai/embed",
    version: EMBED_VERSION,
    source,
    createdAt: new Date().toISOString(),
    payload,
  };
}

export function serializeEnvelope(env: OfficeAIEmbedEnvelope): string {
  return JSON.stringify(env);
}

/**
 * Parse a string off the system clipboard back into an envelope.
 * Returns `null` for anything that isn't a recognisable envelope —
 * unknown JSON, wrong `type`, malformed payload, etc. We never
 * throw because the caller tries the embed parser as one of
 * several MIME readers and any throw would mask the rest.
 */
export function parseEnvelope(raw: string | null | undefined): OfficeAIEmbedEnvelope | null {
  if (!raw) return null;
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    return null;
  }
  if (!isEnvelope(parsed)) return null;
  return parsed;
}

function isEnvelope(v: unknown): v is OfficeAIEmbedEnvelope {
  if (!v || typeof v !== "object") return false;
  const o = v as Record<string, unknown>;
  if (o.type !== "officeai/embed") return false;
  if (typeof o.version !== "number") return false;
  if (o.source !== "xlsx" && o.source !== "docx" && o.source !== "pptx") return false;
  const p = o.payload;
  if (!p || typeof p !== "object") return false;
  const pk = (p as { kind?: unknown }).kind;
  if (pk !== "xlsx-range" && pk !== "xlsx-chart-image" && pk !== "docx-table") return false;
  return true;
}
