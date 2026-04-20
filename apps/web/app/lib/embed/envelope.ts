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
 * The whole feature is gated on `NEXT_PUBLIC_OAI_EMBED` — see
 * `isEmbedEnabled()`. When disabled, copy paths skip the extra MIME
 * and paste paths ignore it; the shipping `text/html` fallback
 * keeps working unchanged.
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
  | XlsxChartImageEmbed;

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

export interface OfficeAIEmbedEnvelope {
  readonly type: "officeai/embed";
  readonly version: number;
  readonly source: "xlsx" | "docx" | "pptx";
  readonly createdAt: string;
  readonly payload: OfficeAIEmbedPayload;
}

/**
 * Truthy when the cross-format embed feature flag is on. The flag
 * is a string `"1"` to avoid the classic `"false"` truthiness foot-
 * gun. We deliberately default to OFF in production builds so the
 * extra clipboard MIME never ships before we've finished QA.
 */
export function isEmbedEnabled(): boolean {
  const flag = (process.env.NEXT_PUBLIC_OAI_EMBED ?? "").trim();
  return flag === "1" || flag === "true";
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
  if (pk !== "xlsx-range" && pk !== "xlsx-chart-image") return false;
  return true;
}
