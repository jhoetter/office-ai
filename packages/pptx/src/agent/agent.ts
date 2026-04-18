import {
  CommandBus,
  type Command,
  type CommandLite,
  type DocumentDiff,
  type Mutation,
} from "@officeai/core";
import { allPptxHandlers } from "../commands/index.js";
import type { PptxSnapshot, Shape, Slide, TextShape } from "../model/types.js";
import { parsePptx, type ParseOptions } from "../parser/parse.js";
import { serializePptx } from "../serializer/serialize.js";
import { paragraphText, snapshotToMarkdown } from "./markdown.js";

export interface PptxAgentOptions extends ParseOptions {
  readonly sessionId?: string;
}

export interface PptxRangeRequest {
  readonly kind: "pptx-slides";
  readonly start: number;
  readonly end: number;
}

export interface PptxRangeSnapshot {
  readonly slides: ReadonlyArray<{
    readonly index: number;
    readonly id: string;
    readonly partPath: string;
    readonly slideId: number;
    readonly shapeCount: number;
    readonly text: string;
  }>;
}

export interface PptxSearchSpec {
  query: string;
  caseSensitive?: boolean;
  regex?: boolean;
}

export interface PptxSearchResult {
  slideIndex: number;
  shapeId: string;
  paragraphIndex: number;
  preview: string;
  match: string;
  start: number;
  end: number;
  /** Present when the match is inside a typed TableShape cell. */
  tableCell?: { row: number; column: number };
}

/**
 * Headless PPTX agent — implements the shared `DocumentAgent` shape from
 * spec/shared/agent-api.md. Built on top of the headless parser, command
 * bus, and serializer; has zero DOM/React imports.
 */
export class PptxAgent {
  private readonly bus: CommandBus<PptxSnapshot>;

  private constructor(initial: PptxSnapshot, opts: PptxAgentOptions) {
    this.bus = new CommandBus<PptxSnapshot>(initial, {
      ...(opts.sessionId ? { sessionId: opts.sessionId } : {}),
      ...(opts.idMinter ? { mintNodeId: opts.idMinter } : {}),
    });
    this.bus.registerAll(allPptxHandlers);
  }

  static async fromBuffer(
    buffer: ArrayBuffer | Uint8Array,
    opts: PptxAgentOptions = {}
  ): Promise<PptxAgent> {
    const snap = await parsePptx(buffer, opts);
    return new PptxAgent(snap, opts);
  }

  // ── Read ───────────────────────────────────────────────────────────────
  getSnapshot(): PptxSnapshot {
    return this.bus.getSnapshot();
  }

  getApprovedSnapshot(): PptxSnapshot {
    return this.bus.getApproved();
  }

  toMarkdown(): string {
    return snapshotToMarkdown(this.getSnapshot());
  }

  getRange(req: PptxRangeRequest): PptxRangeSnapshot {
    const slides = this.getSnapshot().root.slides;
    const start = Math.max(0, req.start);
    const end = Math.min(slides.length, req.end);
    const out: Array<PptxRangeSnapshot["slides"][number]> = [];
    for (let i = start; i < end; i++) {
      const s = slides[i];
      out.push({
        index: i,
        id: s.id,
        partPath: s.partPath,
        slideId: s.slideId,
        shapeCount: s.shapes.length,
        text: slidePlainText(s),
      });
    }
    return { slides: out };
  }

  search(spec: PptxSearchSpec): PptxSearchResult[] {
    const slides = this.getSnapshot().root.slides;
    const out: PptxSearchResult[] = [];
    const flags = spec.caseSensitive ? "g" : "gi";
    const pattern = spec.regex
      ? new RegExp(spec.query, flags)
      : new RegExp(escapeRegex(spec.query), flags);
    for (let si = 0; si < slides.length; si++) {
      walkShapes(slides[si].shapes, (shape) => {
        if (shape.kind === "text") {
          for (let pi = 0; pi < shape.txBody.paragraphs.length; pi++) {
            const text = paragraphText(shape.txBody.paragraphs[pi]);
            let m: RegExpExecArray | null;
            while ((m = pattern.exec(text)) !== null) {
              out.push({
                slideIndex: si,
                shapeId: shape.id,
                paragraphIndex: pi,
                start: m.index,
                end: m.index + m[0].length,
                match: m[0],
                preview: snippet(text, m.index, m.index + m[0].length),
              });
              if (m[0].length === 0) pattern.lastIndex++;
            }
          }
          return;
        }
        if (shape.kind === "table") {
          for (let ri = 0; ri < shape.rows.length; ri++) {
            const row = shape.rows[ri];
            for (let ci = 0; ci < row.cells.length; ci++) {
              const cell = row.cells[ci];
              for (let pi = 0; pi < cell.txBody.paragraphs.length; pi++) {
                const text = paragraphText(cell.txBody.paragraphs[pi]);
                let m: RegExpExecArray | null;
                while ((m = pattern.exec(text)) !== null) {
                  out.push({
                    slideIndex: si,
                    shapeId: shape.id,
                    paragraphIndex: pi,
                    start: m.index,
                    end: m.index + m[0].length,
                    match: m[0],
                    preview: snippet(text, m.index, m.index + m[0].length),
                    tableCell: { row: ri, column: ci },
                  });
                  if (m[0].length === 0) pattern.lastIndex++;
                }
              }
            }
          }
        }
      });
    }
    return out;
  }

  // ── Write ──────────────────────────────────────────────────────────────
  async applyCommand(command: Command | CommandLite): Promise<Mutation<PptxSnapshot>> {
    return this.bus.dispatch(command);
  }

  async applyCommands(
    commands: ReadonlyArray<Command | CommandLite>
  ): Promise<ReadonlyArray<Mutation<PptxSnapshot>>> {
    return this.bus.dispatchAll(commands);
  }

  // ── Diff & Review ──────────────────────────────────────────────────────
  getDiff(from: PptxSnapshot, to: PptxSnapshot): DocumentDiff {
    return {
      format: "pptx",
      fromRevision: from.revision,
      toRevision: to.revision,
      changes: [
        {
          kind: "node-updated",
          nodeId: from.root.id,
          path: ["root"],
          field: "snapshot",
          summary: `revision ${from.revision} → ${to.revision}`,
        },
      ],
    };
  }

  getPendingMutations(): ReadonlyArray<Mutation<PptxSnapshot>> {
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
    return serializePptx(this.getSnapshot());
  }

  // ── Subscriptions ──────────────────────────────────────────────────────
  subscribe(
    listener: (snapshot: PptxSnapshot, mutation: Mutation<PptxSnapshot>) => void
  ): () => void {
    return this.bus.subscribe(listener);
  }
}

function slidePlainText(slide: Slide): string {
  const out: string[] = [];
  walkShapes(slide.shapes, (shape) => {
    if (shape.kind === "text") {
      for (const p of shape.txBody.paragraphs) {
        const t = paragraphText(p);
        if (t.length > 0) out.push(t);
      }
      return;
    }
    if (shape.kind === "table") {
      for (const row of shape.rows) {
        for (const cell of row.cells) {
          for (const p of cell.txBody.paragraphs) {
            const t = paragraphText(p);
            if (t.length > 0) out.push(t);
          }
        }
      }
    }
  });
  return out.join("\n");
}

function walkShapes(shapes: ReadonlyArray<Shape>, visit: (s: TextShape | Shape) => void): void {
  for (const s of shapes) {
    visit(s);
    if (s.kind === "group") walkShapes(s.children, visit);
  }
}

function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function snippet(text: string, start: number, end: number, span = 40): string {
  const a = Math.max(0, start - span);
  const b = Math.min(text.length, end + span);
  return (a > 0 ? "…" : "") + text.slice(a, b) + (b < text.length ? "…" : "");
}
