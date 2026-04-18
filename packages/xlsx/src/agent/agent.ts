import { CommandBus, type Command, type CommandLite, type DocumentDiff, type Mutation } from "@officeai/core";
import { allXlsxHandlers } from "../commands/index.js";
import {
  cellKey,
  formatA1,
  formatRange,
  parseA1,
  parseRange,
  type CellAddress,
  type CellRange,
} from "../model/refs.js";
import type { Cell, Sheet, XlsxSnapshot } from "../model/types.js";
import { parseXlsx, type ParseOptions } from "../parser/parse.js";
import { serializeXlsx } from "../serializer/serialize.js";
import { diffXlsxSnapshots } from "./diff.js";

export interface XlsxAgentOptions extends ParseOptions {
  readonly sessionId?: string;
}

/**
 * Range request for read-side projection. Two shapes:
 *   - { kind: "xlsx-range", sheet, range }     A1 range like "A1:C10"
 *   - { kind: "xlsx-sheet", sheet }            entire sheet (sparse)
 */
export type XlsxRangeRequest =
  | {
      readonly kind: "xlsx-range";
      readonly sheet: string;
      readonly range: string;
    }
  | {
      readonly kind: "xlsx-sheet";
      readonly sheet: string;
    };

export interface XlsxRangeSnapshot {
  readonly sheet: string;
  readonly range: string;
  /**
   * Sparse list of populated cells in the range. Order: row-major
   * (top-to-bottom, then left-to-right). Caller can rebuild a dense
   * matrix via the returned `(rows, cols)` for sizing.
   */
  readonly cells: ReadonlyArray<{
    readonly ref: string;
    readonly row: number;
    readonly col: number;
    readonly value: Cell["value"];
    readonly formula?: string;
  }>;
  readonly rows: number;
  readonly cols: number;
}

export interface XlsxSearchSpec {
  readonly query: string;
  readonly caseSensitive?: boolean;
  readonly regex?: boolean;
  /** Restrict to a single sheet name. Default: search every worksheet. */
  readonly sheet?: string;
}

export interface XlsxSearchResult {
  readonly sheet: string;
  readonly ref: string;
  readonly row: number;
  readonly col: number;
  readonly value: string;
  readonly match: string;
  readonly start: number;
  readonly end: number;
}

/**
 * Headless XLSX agent. Implements the same `DocumentAgent` shape as
 * `DocxAgent` (see spec/shared/agent-api.md): read / write /
 * diff & review / I/O / subscriptions. All command application
 * goes through the shared `CommandBus`; agent-sourced commands are
 * staged as `pending` for human review.
 */
export class XlsxAgent {
  private bus: CommandBus<XlsxSnapshot>;
  private readonly opts: XlsxAgentOptions;

  private constructor(initial: XlsxSnapshot, opts: XlsxAgentOptions) {
    this.opts = opts;
    this.bus = new CommandBus<XlsxSnapshot>(initial, {
      ...(opts.sessionId ? { sessionId: opts.sessionId } : {}),
      ...(opts.idMinter ? { mintNodeId: opts.idMinter } : {}),
    });
    this.bus.registerAll(allXlsxHandlers);
  }

  static async fromBuffer(buffer: ArrayBuffer | Uint8Array, opts: XlsxAgentOptions = {}): Promise<XlsxAgent> {
    const snap = await parseXlsx(buffer, opts);
    return new XlsxAgent(snap, opts);
  }

  /** Replace the in-memory document. Drops all pending mutations. */
  async importFile(buffer: ArrayBuffer | Uint8Array): Promise<void> {
    const snap = await parseXlsx(buffer, this.opts);
    this.bus = new CommandBus<XlsxSnapshot>(snap, {
      ...(this.opts.sessionId ? { sessionId: this.opts.sessionId } : {}),
      ...(this.opts.idMinter ? { mintNodeId: this.opts.idMinter } : {}),
    });
    this.bus.registerAll(allXlsxHandlers);
  }

  // ── Read ───────────────────────────────────────────────────────────────
  getSnapshot(): XlsxSnapshot {
    return this.bus.getSnapshot();
  }

  getApprovedSnapshot(): XlsxSnapshot {
    return this.bus.getApproved();
  }

  /** List sheet names in tab order. */
  listSheets(): ReadonlyArray<{ name: string; index: number; kind: Sheet["kind"]; state: Sheet["state"] }> {
    return this.getSnapshot().root.sheets.map((s) => ({
      name: s.name,
      index: s.index,
      kind: s.kind,
      state: s.state,
    }));
  }

  /**
   * Markdown-table preview of one or more sheets. Useful for LLM
   * surfacing. Each sheet renders only its populated rectangle (the
   * bounding box of its sparse cell map). Empty sheets render as
   * "(empty)".
   */
  toMarkdown(opts: { sheet?: string; maxRows?: number; maxCols?: number } = {}): string {
    const snap = this.getSnapshot();
    const sheets = opts.sheet
      ? snap.root.sheets.filter((s) => s.name === opts.sheet)
      : snap.root.sheets.filter((s) => s.kind === "worksheet");
    const out: string[] = [];
    for (const sheet of sheets) {
      out.push(`## ${sheet.name}`);
      out.push(sheetToMarkdown(sheet, opts.maxRows ?? 50, opts.maxCols ?? 20));
      out.push("");
    }
    return out.join("\n").trimEnd();
  }

  /** Sparse projection of a range or whole sheet. */
  getRange(req: XlsxRangeRequest): XlsxRangeSnapshot {
    const snap = this.getSnapshot();
    const sheet = snap.root.sheets.find((s) => s.name === req.sheet);
    if (!sheet) {
      throw new Error(`unknown sheet "${req.sheet}"`);
    }
    if (req.kind === "xlsx-sheet") {
      const bbox = sheetBoundingBox(sheet);
      return projectRange(sheet, bbox);
    }
    const range = parseRange(req.range);
    return projectRange(sheet, range);
  }

  search(spec: XlsxSearchSpec): XlsxSearchResult[] {
    const snap = this.getSnapshot();
    const flags = spec.caseSensitive ? "g" : "gi";
    const pattern = spec.regex ? new RegExp(spec.query, flags) : new RegExp(escapeRegex(spec.query), flags);
    const out: XlsxSearchResult[] = [];
    const sheets = spec.sheet
      ? snap.root.sheets.filter((s) => s.name === spec.sheet)
      : snap.root.sheets.filter((s) => s.kind === "worksheet");

    for (const sheet of sheets) {
      for (const cell of sheet.cells.values()) {
        const text = cellValueToString(cell.value);
        if (text === "") continue;
        let m: RegExpExecArray | null;
        while ((m = pattern.exec(text)) !== null) {
          out.push({
            sheet: sheet.name,
            ref: formatA1({ row: cell.row, col: cell.col }),
            row: cell.row,
            col: cell.col,
            value: text,
            match: m[0],
            start: m.index,
            end: m.index + m[0].length,
          });
          if (m[0].length === 0) pattern.lastIndex++;
        }
      }
    }
    return out;
  }

  // ── Write ──────────────────────────────────────────────────────────────
  async applyCommand(command: Command | CommandLite): Promise<Mutation<XlsxSnapshot>> {
    return this.bus.dispatch(command);
  }

  async applyCommands(
    commands: ReadonlyArray<Command | CommandLite>
  ): Promise<ReadonlyArray<Mutation<XlsxSnapshot>>> {
    return this.bus.dispatchAll(commands);
  }

  // ── Diff & Review ──────────────────────────────────────────────────────
  /**
   * Structural diff between two snapshots. See `agent/diff.ts`.
   */
  getDiff(from: XlsxSnapshot, to: XlsxSnapshot): DocumentDiff {
    return diffXlsxSnapshots(from, to);
  }

  getPendingMutations(): ReadonlyArray<Mutation<XlsxSnapshot>> {
    return this.bus.getPending();
  }

  approveMutation(id: string): void {
    this.bus.approveMutation(id);
  }

  rejectMutation(id: string): void {
    this.bus.rejectMutation(id);
  }

  rollback(toRevision: number): void {
    this.bus.rollback(toRevision);
  }

  // ── I/O ────────────────────────────────────────────────────────────────
  async exportFile(): Promise<ArrayBuffer> {
    return serializeXlsx(this.getSnapshot());
  }

  // ── Subscriptions ──────────────────────────────────────────────────────
  subscribe(listener: (snapshot: XlsxSnapshot, mutation: Mutation<XlsxSnapshot>) => void): () => void {
    return this.bus.subscribe(listener);
  }
}

function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function sheetBoundingBox(sheet: Sheet): CellRange {
  let minRow = Infinity;
  let minCol = Infinity;
  let maxRow = -1;
  let maxCol = -1;
  for (const c of sheet.cells.values()) {
    if (c.row < minRow) minRow = c.row;
    if (c.col < minCol) minCol = c.col;
    if (c.row > maxRow) maxRow = c.row;
    if (c.col > maxCol) maxCol = c.col;
  }
  if (maxRow === -1) {
    return { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } };
  }
  return { start: { row: minRow, col: minCol }, end: { row: maxRow, col: maxCol } };
}

function projectRange(sheet: Sheet, range: CellRange): XlsxRangeSnapshot {
  const cells: Array<{
    ref: string;
    row: number;
    col: number;
    value: Cell["value"];
    formula?: string;
  }> = [];
  for (let r = range.start.row; r <= range.end.row; r++) {
    for (let c = range.start.col; c <= range.end.col; c++) {
      const cell = sheet.cells.get(cellKey(r, c));
      if (!cell) continue;
      const ref = formatA1({ row: r, col: c });
      cells.push({
        ref,
        row: r,
        col: c,
        value: cell.value,
        ...(cell.formula ? { formula: cell.formula.text } : {}),
      });
    }
  }
  return {
    sheet: sheet.name,
    range: formatRange(range),
    cells,
    rows: range.end.row - range.start.row + 1,
    cols: range.end.col - range.start.col + 1,
  };
}

function sheetToMarkdown(sheet: Sheet, maxRows: number, maxCols: number): string {
  if (sheet.cells.size === 0) return "_(empty)_";
  const bbox = sheetBoundingBox(sheet);
  const rowEnd = Math.min(bbox.end.row, bbox.start.row + maxRows - 1);
  const colEnd = Math.min(bbox.end.col, bbox.start.col + maxCols - 1);

  const header: string[] = [""];
  for (let c = bbox.start.col; c <= colEnd; c++) {
    header.push(formatA1({ row: bbox.start.row, col: c }).replace(/[0-9]+$/, ""));
  }
  const lines: string[] = [`| ${header.join(" | ")} |`, `| ${header.map(() => "---").join(" | ")} |`];

  for (let r = bbox.start.row; r <= rowEnd; r++) {
    const row: string[] = [String(r + 1)];
    for (let c = bbox.start.col; c <= colEnd; c++) {
      const cell = sheet.cells.get(cellKey(r, c));
      row.push(cell ? cellValueToString(cell.value) : "");
    }
    lines.push(`| ${row.join(" | ")} |`);
  }

  if (bbox.end.row > rowEnd || bbox.end.col > colEnd) {
    lines.push(`_(truncated; full bounding box ${formatA1(bbox.start)}:${formatA1(bbox.end)})_`);
  }
  return lines.join("\n");
}

function cellValueToString(v: Cell["value"]): string {
  if (v === null || v === undefined) return "";
  if (typeof v === "string") return v;
  if (typeof v === "object" && "kind" in v) return v.code;
  return String(v);
}

export type { CellAddress };
export { parseA1 };
