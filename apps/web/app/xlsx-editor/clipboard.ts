"use client";

import {
  parseExternalClipboard,
  parseFingerprintHtml,
  parseHtmlTable as parseHtmlTablePure,
  sniffDelimiter,
  snapshotToTsv,
  type XlsxClipboardSnapshot,
} from "@officeai/xlsx";
import { EMBED_MIME, makeEnvelope, serializeEnvelope } from "@/lib/embed/envelope";

/**
 * Web-side system-clipboard bridge for the XLSX editor.
 *
 * Outbound (`Cmd+C` / `Cmd+X`):
 *   - `marshalClipboard(snapshot)` produces `{ tsv, html }`.
 *   - The HTML carries the full snapshot as a JSON fingerprint on a
 *     `data-xlsx-fingerprint` attribute on the root `<table>` so a
 *     paste back into our own app round-trips formulas + styles +
 *     merges with full fidelity.
 *   - The TSV is the universal fallback (Excel, Sheets, plain text
 *     editors).
 *   - `writeToSystemClipboard` writes both MIME types in one
 *     `ClipboardItem` call.
 *
 * Inbound (`Cmd+V`):
 *   - `readFromSystemClipboard()` returns a `XlsxClipboardSnapshot`.
 *   - Priority: our own HTML fingerprint > generic HTML table >
 *     TSV (text/plain with tabs) > CSV (text/plain with `,` `;` `|`).
 */

const FINGERPRINT_ATTR = "data-xlsx-fingerprint";

export interface MarshalResult {
  readonly tsv: string;
  readonly html: string;
  /**
   * Structured embed envelope for cross-format pastes (XLSX → DOCX
   * table, XLSX → PPTX table). Always emitted; readers in DOCX/PPTX
   * still tolerate its absence and fall back to the HTML/TSV channels
   * when an external app pasted in.
   */
  readonly embed: string;
}

/**
 * Build the TSV + HTML (+ embed) payloads for a snapshot.
 *
 * Outputs are pure strings so the caller decides which channels to
 * paint onto the clipboard (synchronous `event.clipboardData.setData`
 * vs async `ClipboardItem` write).
 */
export function marshalClipboard(snapshot: XlsxClipboardSnapshot): MarshalResult {
  const tsv = snapshotToTsv(snapshot);
  const html = buildHtmlTable(snapshot);
  const env = makeEnvelope("xlsx", {
    kind: "xlsx-range",
    snapshot,
    originLabel: `${snapshot.origin.sheet}!${snapshot.origin.range}`,
  });
  return { tsv, html, embed: serializeEnvelope(env) };
}

function buildHtmlTable(snapshot: XlsxClipboardSnapshot): string {
  const fp = encodeURIComponent(JSON.stringify(snapshot));
  const rows: string[] = [];
  for (let r = 0; r < snapshot.height; r++) {
    const row = snapshot.cells[r] ?? [];
    const cols: string[] = [];
    for (let c = 0; c < snapshot.width; c++) {
      const cell = row[c];
      cols.push(`<td>${cell ? escapeHtml(displayValue(cell)) : ""}</td>`);
    }
    rows.push(`<tr>${cols.join("")}</tr>`);
  }
  return `<table ${FINGERPRINT_ATTR}="${fp}"><tbody>${rows.join("")}</tbody></table>`;
}

function displayValue(cell: { value: unknown; formula?: string }): string {
  if (cell.formula) return `=${cell.formula}`;
  const v = cell.value;
  if (v === null || v === undefined) return "";
  if (typeof v === "string") return v;
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (typeof v === "object" && v && "code" in v) return String((v as { code: string }).code);
  return String(v);
}

function escapeHtml(s: string): string {
  return s.replace(/[&<>"']/g, (ch) => {
    switch (ch) {
      case "&":
        return "&amp;";
      case "<":
        return "&lt;";
      case ">":
        return "&gt;";
      case '"':
        return "&quot;";
      case "'":
        return "&#39;";
      default:
        return ch;
    }
  });
}

/**
 * Write the TSV + HTML pair to the system clipboard. Falls back to
 * plain `writeText(tsv)` on browsers / contexts that don't allow the
 * `ClipboardItem` API.
 */
export async function writeToSystemClipboard(payload: MarshalResult): Promise<void> {
  if (typeof navigator === "undefined") return;
  const clip = navigator.clipboard;
  if (!clip) return;
  try {
    if (typeof window !== "undefined" && "ClipboardItem" in window) {
      const parts: Record<string, Blob> = {
        "text/plain": new Blob([payload.tsv], { type: "text/plain" }),
        "text/html": new Blob([payload.html], { type: "text/html" }),
        [EMBED_MIME]: new Blob([payload.embed], { type: EMBED_MIME }),
      };
      const item = new (window as unknown as { ClipboardItem: typeof ClipboardItem }).ClipboardItem(parts);
      await clip.write([item]);
      return;
    }
  } catch {
    // Fall through to writeText.
  }
  await clip.writeText(payload.tsv);
}

/**
 * Read the system clipboard and return the best-fit
 * {@link XlsxClipboardSnapshot}.
 *
 * Returns `null` if the clipboard is empty / unreadable. Throws only
 * for truly unrecoverable errors (no `navigator.clipboard`, etc.).
 *
 * All actual parsing is delegated to `parseExternalClipboard` in
 * `@officeai/xlsx` so the same logic exercises the package test
 * suite (fingerprint round-trip, Excel HTML, Sheets HTML, CSV, TSV).
 */
export async function readFromSystemClipboard(): Promise<XlsxClipboardSnapshot | null> {
  if (typeof navigator === "undefined") return null;
  const clip = navigator.clipboard;
  if (!clip) return null;

  try {
    if (typeof clip.read === "function") {
      const items = await clip.read();
      let html: string | null = null;
      let text: string | null = null;
      for (const it of items) {
        if (!html && it.types.includes("text/html")) {
          html = await (await it.getType("text/html")).text();
        }
        if (!text && it.types.includes("text/plain")) {
          text = await (await it.getType("text/plain")).text();
        }
      }
      const out = parseExternalClipboard({ html, text });
      if (out) return out;
    }
  } catch {
    // Fall through to readText.
  }

  try {
    const text = await clip.readText();
    if (!text) return null;
    return parseExternalClipboard({ text });
  } catch {
    return null;
  }
}

/**
 * Parse a clipboard payload that the caller has already extracted
 * from an event (e.g. the React `onPaste` handler). Convenient for
 * `paste`-event-driven flows that don't go through the async
 * `navigator.clipboard.read()` permission dance.
 */
export function parseClipboardPayload(opts: {
  readonly html?: string | null;
  readonly text?: string | null;
}): XlsxClipboardSnapshot | null {
  return parseExternalClipboard(opts);
}

/**
 * Re-exports of the headless parsers so existing callers keep
 * working through this module's public surface.
 */
export const parseFingerprint = parseFingerprintHtml;
export const parseHtmlTable = parseHtmlTablePure;
export { sniffDelimiter };
