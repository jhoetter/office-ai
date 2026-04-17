import {
  CommandBus,
  type Command,
  type CommandLite,
  type DocumentDiff,
  type Mutation,
} from "@officeai/core";
import { allDocxHandlers } from "../commands/index.js";
import type { DocxSnapshot, DocxPosition } from "../model/types.js";
import { paragraphPlainText } from "../commands/helpers.js";
import { parseDocx, type ParseOptions } from "../parser/parse.js";
import { serializeDocx } from "../serializer/serialize.js";
import { snapshotToMarkdown } from "./markdown.js";

export interface DocxAgentOptions extends ParseOptions {
  readonly sessionId?: string;
}

export interface DocxRangeRequest {
  readonly kind: "docx-paragraphs";
  readonly start: number;
  readonly end: number;
}

export interface DocxRangeSnapshot {
  readonly paragraphs: ReadonlyArray<{
    readonly index: number;
    readonly id: string;
    readonly styleId?: string;
    readonly text: string;
  }>;
}

export interface DocxSearchSpec {
  query: string;
  caseSensitive?: boolean;
  regex?: boolean;
}

export interface DocxSearchResult {
  paragraphIndex: number;
  preview: string;
  match: string;
  start: number;
  end: number;
}

/**
 * Headless DOCX agent. Implements the shared `DocumentAgent` shape
 * described in spec/shared/agent-api.md.
 */
export class DocxAgent {
  private readonly bus: CommandBus<DocxSnapshot>;

  private constructor(initial: DocxSnapshot, opts: DocxAgentOptions) {
    this.bus = new CommandBus<DocxSnapshot>(initial, {
      ...(opts.sessionId ? { sessionId: opts.sessionId } : {}),
      ...(opts.idMinter ? { mintNodeId: opts.idMinter } : {}),
    });
    this.bus.registerAll(allDocxHandlers);
  }

  static async fromBuffer(buffer: ArrayBuffer | Uint8Array, opts: DocxAgentOptions = {}): Promise<DocxAgent> {
    const snap = await parseDocx(buffer, opts);
    return new DocxAgent(snap, opts);
  }

  // ── Read ───────────────────────────────────────────────────────────────
  getSnapshot(): DocxSnapshot {
    return this.bus.getSnapshot();
  }

  getApprovedSnapshot(): DocxSnapshot {
    return this.bus.getApproved();
  }

  toMarkdown(): string {
    return snapshotToMarkdown(this.getSnapshot());
  }

  getRange(req: DocxRangeRequest): DocxRangeSnapshot {
    const body = this.getSnapshot().root.body;
    const start = Math.max(0, req.start);
    const end = Math.min(body.length, req.end);
    const paragraphs: Array<DocxRangeSnapshot["paragraphs"][number]> = [];
    for (let i = start; i < end; i++) {
      const block = body[i];
      if (block.kind !== "paragraph") continue;
      paragraphs.push({
        index: i,
        id: block.id,
        ...(block.properties.styleId ? { styleId: block.properties.styleId } : {}),
        text: paragraphPlainText(block),
      });
    }
    return { paragraphs };
  }

  search(spec: DocxSearchSpec): DocxSearchResult[] {
    const body = this.getSnapshot().root.body;
    const out: DocxSearchResult[] = [];
    const flags = spec.caseSensitive ? "g" : "gi";
    const pattern = spec.regex
      ? new RegExp(spec.query, flags)
      : new RegExp(escapeRegex(spec.query), flags);
    for (let i = 0; i < body.length; i++) {
      const block = body[i];
      if (block.kind !== "paragraph") continue;
      const text = paragraphPlainText(block);
      let m: RegExpExecArray | null;
      while ((m = pattern.exec(text)) !== null) {
        out.push({
          paragraphIndex: i,
          start: m.index,
          end: m.index + m[0].length,
          match: m[0],
          preview: previewSnippet(text, m.index, m.index + m[0].length),
        });
        if (m[0].length === 0) pattern.lastIndex++;
      }
    }
    return out;
  }

  // ── Write ──────────────────────────────────────────────────────────────
  async applyCommand(command: Command | CommandLite): Promise<Mutation<DocxSnapshot>> {
    return this.bus.dispatch(command);
  }

  async applyCommands(commands: ReadonlyArray<Command | CommandLite>): Promise<ReadonlyArray<Mutation<DocxSnapshot>>> {
    return this.bus.dispatchAll(commands);
  }

  // ── Diff & Review ──────────────────────────────────────────────────────
  getDiff(from: DocxSnapshot, to: DocxSnapshot): DocumentDiff {
    return {
      format: "docx",
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

  getPendingMutations(): ReadonlyArray<Mutation<DocxSnapshot>> {
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
    return serializeDocx(this.getSnapshot());
  }

  // ── Subscriptions ──────────────────────────────────────────────────────
  subscribe(listener: (snapshot: DocxSnapshot, mutation: Mutation<DocxSnapshot>) => void): () => void {
    return this.bus.subscribe(listener);
  }
}

function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function previewSnippet(text: string, start: number, end: number, span = 40): string {
  const a = Math.max(0, start - span);
  const b = Math.min(text.length, end + span);
  return (a > 0 ? "…" : "") + text.slice(a, b) + (b < text.length ? "…" : "");
}

export type { DocxPosition };
