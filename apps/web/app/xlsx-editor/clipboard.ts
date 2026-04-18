"use client";

import {
  delimitedToSnapshot,
  sniffDelimiter,
  snapshotToTsv,
  tsvToSnapshot,
  type XlsxClipboardSnapshot,
} from "@officeai/xlsx";

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
}

/** Build the TSV + HTML payloads for a snapshot. */
export function marshalClipboard(snapshot: XlsxClipboardSnapshot): MarshalResult {
  const tsv = snapshotToTsv(snapshot);
  const html = buildHtmlTable(snapshot);
  return { tsv, html };
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
      const item = new (window as unknown as { ClipboardItem: typeof ClipboardItem }).ClipboardItem({
        "text/plain": new Blob([payload.tsv], { type: "text/plain" }),
        "text/html": new Blob([payload.html], { type: "text/html" }),
      });
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
 */
export async function readFromSystemClipboard(): Promise<XlsxClipboardSnapshot | null> {
  if (typeof navigator === "undefined") return null;
  const clip = navigator.clipboard;
  if (!clip) return null;

  // Prefer the rich `read()` API for HTML detection. Fallback to plain
  // text on browsers that gate it behind a permission we don't have.
  try {
    if (typeof clip.read === "function") {
      const items = await clip.read();
      for (const it of items) {
        if (it.types.includes("text/html")) {
          const blob = await it.getType("text/html");
          const html = await blob.text();
          const fromFingerprint = parseFingerprint(html);
          if (fromFingerprint) return fromFingerprint;
          const fromTable = parseHtmlTable(html);
          if (fromTable) return fromTable;
        }
      }
      // Then fall through to text/plain.
      for (const it of items) {
        if (it.types.includes("text/plain")) {
          const blob = await it.getType("text/plain");
          const text = await blob.text();
          return parsePlain(text);
        }
      }
    }
  } catch {
    // Fall through to readText.
  }

  try {
    const text = await clip.readText();
    if (!text) return null;
    return parsePlain(text);
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
  if (opts.html) {
    const fromFingerprint = parseFingerprint(opts.html);
    if (fromFingerprint) return fromFingerprint;
    const fromTable = parseHtmlTable(opts.html);
    if (fromTable) return fromTable;
  }
  if (opts.text) return parsePlain(opts.text);
  return null;
}

function parsePlain(text: string): XlsxClipboardSnapshot {
  const delim = sniffDelimiter(text);
  if (delim === "\t") return tsvToSnapshot(text);
  return delimitedToSnapshot(text, delim);
}

/**
 * Look for a `data-xlsx-fingerprint` attribute on the first `<table>`
 * element. The attribute carries the URI-encoded JSON of the
 * snapshot so we can recover values + formulas + styleIds + merges
 * without lossy HTML→snapshot inference.
 */
export function parseFingerprint(html: string): XlsxClipboardSnapshot | null {
  const re = new RegExp(`<table[^>]*\\s${FINGERPRINT_ATTR}="([^"]+)"`, "i");
  const m = re.exec(html);
  if (!m) return null;
  try {
    const decoded = decodeURIComponent(m[1]!);
    const snap = JSON.parse(decoded) as unknown;
    if (
      snap &&
      typeof snap === "object" &&
      "cells" in snap &&
      Array.isArray((snap as { cells: unknown[] }).cells)
    ) {
      return snap as XlsxClipboardSnapshot;
    }
  } catch {
    // fall through
  }
  return null;
}

/**
 * Generic `<table>` parser for external HTML (Excel-on-web, Google
 * Sheets, Word). Walks `<tr>`/`<td>`/`<th>` and emits a snapshot
 * with values only — no formulas, styles, or merges (we don't try
 * to reconstruct those from external HTML).
 */
export function parseHtmlTable(html: string): XlsxClipboardSnapshot | null {
  if (!html) return null;
  // Prefer DOMParser when available (browser + jsdom in tests).
  if (typeof DOMParser !== "undefined") {
    const doc = new DOMParser().parseFromString(html, "text/html");
    const table = doc.querySelector("table");
    if (!table) return null;
    const rows = Array.from(table.querySelectorAll("tr"));
    if (rows.length === 0) return null;
    const matrix: Array<Array<string>> = [];
    let width = 0;
    for (const row of rows) {
      const cells = Array.from(row.querySelectorAll("td,th"));
      const r = cells.map((c) => c.textContent ?? "");
      matrix.push(r);
      if (r.length > width) width = r.length;
    }
    const cells: Array<Array<{ value: number | string | boolean | null } | null>> = [];
    for (const row of matrix) {
      const out: Array<{ value: number | string | boolean | null } | null> = [];
      for (let c = 0; c < width; c++) {
        const raw = (row[c] ?? "").trim();
        if (raw === "") {
          out.push(null);
          continue;
        }
        if (/^-?\d+(?:\.\d+)?$/.test(raw)) {
          const n = Number(raw);
          if (Number.isFinite(n)) {
            out.push({ value: n });
            continue;
          }
        }
        const lower = raw.toLowerCase();
        if (lower === "true") {
          out.push({ value: true });
          continue;
        }
        if (lower === "false") {
          out.push({ value: false });
          continue;
        }
        out.push({ value: raw });
      }
      cells.push(out);
    }
    return {
      origin: { sheet: "", range: "" },
      width,
      height: cells.length,
      cells,
      merges: [],
    };
  }
  return null;
}

/**
 * Re-export for tests that don't want to reach into `@officeai/xlsx`.
 */
export { sniffDelimiter };
