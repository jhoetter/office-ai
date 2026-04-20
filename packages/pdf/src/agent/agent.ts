import { CommandBus, type Command, type CommandLite, type DocumentDiff, type Mutation } from "@officeai/core";
import { allPdfHandlers } from "../commands/index.js";
import type { PdfRect, PdfSnapshot } from "../model/types.js";
import { parsePdf, type PdfParseOptions } from "../parser/parse.js";
import { serializePdf } from "../serializer/serialize.js";
import { snapshotToMarkdown } from "./markdown.js";
import { findInStructuredPage } from "../text/search.js";

export interface PdfAgentOptions extends PdfParseOptions {
  readonly sessionId?: string;
}

export interface PdfRangeRequest {
  readonly kind: "pdf-pages";
  readonly start: number;
  readonly end: number;
}

export interface PdfRangeSnapshot {
  readonly pages: ReadonlyArray<{
    readonly pageNumber: number;
    readonly id: string;
    readonly width: number;
    readonly height: number;
    readonly rotation: number;
    readonly hasTextLayer: boolean;
    readonly text: string;
  }>;
}

export interface PdfSearchSpec {
  readonly query: string;
  readonly caseSensitive?: boolean;
  readonly regex?: boolean;
  readonly pageRange?: { readonly start: number; readonly end: number };
}

export interface PdfSearchResult {
  readonly pageNumber: number;
  readonly start: number;
  readonly end: number;
  readonly preview: string;
  readonly match: string;
  /**
   * Per-line bounding boxes of the matched glyphs, in PDF
   * user-space (origin bottom-left). Empty for matches that
   * couldn't be projected back onto the structured layout (e.g.
   * regex hits that span column / paragraph boundaries) — callers
   * should treat empty as "fall back to a page-level pulse".
   */
  readonly rects: ReadonlyArray<PdfRect>;
}

const ESCAPE_RE = /[.*+?^${}()|[\]\\]/g;
const escapeRegex = (s: string): string => s.replace(ESCAPE_RE, "\\$&");
const PREVIEW_RADIUS = 40;

const buildPreview = (text: string, start: number, end: number): string => {
  const a = Math.max(0, start - PREVIEW_RADIUS);
  const b = Math.min(text.length, end + PREVIEW_RADIUS);
  return (a > 0 ? "…" : "") + text.slice(a, b) + (b < text.length ? "…" : "");
};

/**
 * Headless PDF agent — implements the same shape as PptxAgent / XlsxAgent.
 * Holds the original PDF buffer so incremental save can produce minimal
 * deltas. Has zero DOM/React imports.
 */
export class PdfAgent {
  private readonly bus: CommandBus<PdfSnapshot>;
  private readonly originalBuffer: Uint8Array;

  private constructor(initial: PdfSnapshot, originalBuffer: Uint8Array, opts: PdfAgentOptions) {
    this.originalBuffer = originalBuffer;
    this.bus = new CommandBus<PdfSnapshot>(initial, {
      ...(opts.sessionId ? { sessionId: opts.sessionId } : {}),
      ...(opts.idMinter ? { mintNodeId: opts.idMinter } : {}),
    });
    this.bus.registerAll(allPdfHandlers);
  }

  static async fromBuffer(buffer: ArrayBuffer | Uint8Array, opts: PdfAgentOptions = {}): Promise<PdfAgent> {
    const source = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
    // PDF.js's worker pipeline transfers the underlying ArrayBuffer
    // out of the calling thread which detaches the original Uint8Array
    // and corrupts our pristine "for incremental save" copy. Defensive
    // clone-then-pass keeps both copies intact.
    const pristine = new Uint8Array(source.byteLength);
    pristine.set(source);
    const forParse = new Uint8Array(source.byteLength);
    forParse.set(source);
    const snap = await parsePdf(forParse, opts);
    return new PdfAgent(snap, pristine, opts);
  }

  /**
   * Produce a brand-new agent from an in-memory empty PDF (one blank
   * Letter-sized page). Use as the entry point for `office-agent pdf
   * create --out`.
   */
  static async empty(opts: PdfAgentOptions = {}): Promise<PdfAgent> {
    const { PDFDocument } = await import("pdf-lib");
    const pdf = await PDFDocument.create();
    pdf.addPage([612, 792]);
    const bytes = await pdf.save();
    return PdfAgent.fromBuffer(bytes, opts);
  }

  // ── Read ──────────────────────────────────────────────────────────────
  getSnapshot(): PdfSnapshot {
    return this.bus.getSnapshot();
  }

  getApprovedSnapshot(): PdfSnapshot {
    return this.bus.getApproved();
  }

  toMarkdown(): string {
    return snapshotToMarkdown(this.getSnapshot());
  }

  getRange(req: PdfRangeRequest): PdfRangeSnapshot {
    const pages = this.getSnapshot().root.pages;
    const start = Math.max(1, req.start);
    const end = Math.min(pages.length, req.end);
    const out: Array<PdfRangeSnapshot["pages"][number]> = [];
    for (let i = start; i <= end; i++) {
      const p = pages[i - 1];
      out.push({
        pageNumber: p.pageNumber,
        id: p.id,
        width: p.width,
        height: p.height,
        rotation: p.rotation,
        hasTextLayer: p.hasTextLayer,
        text: p.text,
      });
    }
    return { pages: out };
  }

  search(spec: PdfSearchSpec): ReadonlyArray<PdfSearchResult> {
    const flags = spec.caseSensitive ? "g" : "gi";
    const pattern = spec.regex ? spec.query : escapeRegex(spec.query);
    let re: RegExp;
    try {
      re = new RegExp(pattern, flags);
    } catch (err) {
      throw new Error(`pdf:search invalid pattern: ${(err as Error).message}`, { cause: err });
    }
    const out: PdfSearchResult[] = [];
    const pages = this.getSnapshot().root.pages;
    const startPage = Math.max(1, spec.pageRange?.start ?? 1);
    const endPage = Math.min(pages.length, spec.pageRange?.end ?? pages.length);
    for (let i = startPage; i <= endPage; i++) {
      const page = pages[i - 1];
      // Glyph-precise hits via the structured page (per-line
      // bbox quads). Falls back transparently to the legacy
      // string-only search when the page has no structured layer
      // (e.g. fully scanned PDF).
      const structuredHits =
        page.structured.blocks.length > 0 ? findInStructuredPage(page.structured, re, page.text) : null;
      if (structuredHits) {
        for (const h of structuredHits) {
          out.push({
            pageNumber: i,
            start: h.start,
            end: h.end,
            preview: buildPreview(page.text, h.start, h.end),
            match: h.match,
            rects: h.rects,
          });
        }
        continue;
      }
      re.lastIndex = 0;
      let m: RegExpExecArray | null;
      while ((m = re.exec(page.text)) !== null) {
        const matchStart = m.index;
        const matchEnd = m.index + m[0].length;
        out.push({
          pageNumber: i,
          start: matchStart,
          end: matchEnd,
          preview: buildPreview(page.text, matchStart, matchEnd),
          match: m[0],
          rects: [],
        });
        if (m[0].length === 0) re.lastIndex++;
      }
    }
    return out;
  }

  // ── Mutate ────────────────────────────────────────────────────────────
  async applyCommand(command: Command | CommandLite): Promise<Mutation<PdfSnapshot>> {
    return this.bus.dispatch(command);
  }

  async applyCommands(
    commands: ReadonlyArray<Command | CommandLite>
  ): Promise<ReadonlyArray<Mutation<PdfSnapshot>>> {
    return this.bus.dispatchAll(commands);
  }

  // ── Diff & Review ─────────────────────────────────────────────────────
  getDiff(from: PdfSnapshot, to: PdfSnapshot): DocumentDiff {
    return {
      format: "pdf",
      fromRevision: from.revision,
      toRevision: to.revision,
      changes: [
        {
          kind: "node-updated",
          nodeId: "root",
          path: ["root"],
          field: "snapshot",
          summary: `revision ${from.revision} → ${to.revision}`,
        },
      ],
    };
  }

  getPendingMutations(): ReadonlyArray<Mutation<PdfSnapshot>> {
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

  canUndo(): boolean {
    return this.bus.canUndo();
  }

  canRedo(): boolean {
    return this.bus.canRedo();
  }

  undo(): Mutation<PdfSnapshot> | null {
    return this.bus.undo();
  }

  redo(): Mutation<PdfSnapshot> | null {
    return this.bus.redo();
  }

  // ── I/O ───────────────────────────────────────────────────────────────
  async exportFile(): Promise<Uint8Array> {
    return serializePdf(this.getSnapshot(), this.originalBuffer);
  }

  /** Access the original PDF bytes (for diffing and audit purposes). */
  originalBytes(): Uint8Array {
    return this.originalBuffer;
  }

  // ── Subscriptions ─────────────────────────────────────────────────────
  subscribe(listener: (snapshot: PdfSnapshot, mutation: Mutation<PdfSnapshot>) => void): () => void {
    return this.bus.subscribe(listener);
  }
}
