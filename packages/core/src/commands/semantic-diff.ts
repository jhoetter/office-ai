import type { DiffChange, DiffPath, DocumentDiff, DocumentFormat } from "../types/document.js";
import type { CommandDiagnostic } from "./lifecycle.js";

export type SemanticDiffRisk = "low" | "medium" | "high" | "unknown";

export interface SemanticDiffAnchor {
  readonly id: string;
  readonly label: string;
  readonly path: DiffPath;
}

export interface SemanticDiffChange {
  readonly kind: DiffChange["kind"] | "opaque";
  readonly summary: string;
  readonly risk: SemanticDiffRisk;
  readonly anchor?: SemanticDiffAnchor;
  readonly before?: unknown;
  readonly after?: unknown;
  readonly field?: string;
}

export interface SemanticDiffSummary {
  readonly text: string;
  readonly changeCount: number;
  readonly risk: SemanticDiffRisk;
}

export interface SemanticDiff {
  readonly schema: "office-ai/semantic-diff@1";
  readonly format: DocumentFormat;
  readonly fromRevision: number;
  readonly toRevision: number;
  readonly summary: SemanticDiffSummary;
  readonly anchors: ReadonlyArray<SemanticDiffAnchor>;
  readonly changes: ReadonlyArray<SemanticDiffChange>;
  readonly diagnostics: ReadonlyArray<CommandDiagnostic>;
  readonly fallback: boolean;
  readonly truncated: boolean;
}

export interface NormalizeDocumentDiffOptions {
  readonly maxChanges?: number;
  readonly operation?: string;
}

export function normalizeDocumentDiff(
  diff: DocumentDiff,
  opts: NormalizeDocumentDiffOptions = {}
): SemanticDiff {
  const maxChanges = Math.max(1, opts.maxChanges ?? 25);
  const fallback = diff.changes.length === 0 && diff.fromRevision !== diff.toRevision;
  const sourceChanges = fallback ? [opaqueChange(diff, opts.operation)] : diff.changes;
  const visible = sourceChanges.slice(0, maxChanges);
  const changes = visible.map((change) => semanticChange(diff.format, change));
  const anchors = uniqueAnchors(changes);
  const truncated = sourceChanges.length > visible.length;
  const diagnostics: CommandDiagnostic[] = [];
  if (fallback) {
    diagnostics.push({
      level: "warning",
      code: "semantic-diff-fallback",
      message: "Only an opaque revision change is available for this mutation.",
    });
  }
  if (truncated) {
    diagnostics.push({
      level: "warning",
      code: "semantic-diff-truncated",
      message: `Semantic diff includes ${visible.length} of ${sourceChanges.length} changes.`,
    });
  }
  const risk = aggregateRisk(changes.map((change) => change.risk));
  return {
    schema: "office-ai/semantic-diff@1",
    format: diff.format,
    fromRevision: diff.fromRevision,
    toRevision: diff.toRevision,
    summary: {
      text: summarize(diff.format, sourceChanges.length, risk, opts.operation, fallback),
      changeCount: sourceChanges.length,
      risk,
    },
    anchors,
    changes,
    diagnostics,
    fallback,
    truncated,
  };
}

function opaqueChange(diff: DocumentDiff, operation?: string): DiffChange {
  return {
    kind: "node-updated",
    nodeId: "document",
    path: ["document"],
    field: "revision",
    summary: `${operation ?? diff.format} changed revision ${diff.fromRevision} -> ${diff.toRevision}.`,
    meta: { opaque: true },
  };
}

function semanticChange(format: DocumentFormat, change: DiffChange): SemanticDiffChange {
  const anchor = anchorFor(format, change);
  return {
    kind: change.meta?.opaque === true ? "opaque" : change.kind,
    summary: change.summary,
    risk: riskFor(change),
    ...(anchor ? { anchor } : {}),
    ...("field" in change ? { field: change.field } : {}),
    ...("meta" in change && change.meta && "before" in change.meta ? { before: change.meta.before } : {}),
    ...("meta" in change && change.meta && "after" in change.meta ? { after: change.meta.after } : {}),
  };
}

function anchorFor(format: DocumentFormat, change: DiffChange): SemanticDiffAnchor | undefined {
  const path = "path" in change ? change.path : "to" in change ? change.to : undefined;
  if (!path || path.length === 0) return undefined;
  const id = "nodeId" in change ? change.nodeId : path.join("/");
  return {
    id,
    label: `${format.toUpperCase()} ${path.map(String).join(" > ")}`,
    path,
  };
}

function riskFor(change: DiffChange): SemanticDiffRisk {
  if (change.meta?.opaque === true) return "unknown";
  if (change.kind === "node-deleted") return "high";
  if (change.kind === "node-moved" || change.kind === "part-added") return "medium";
  return "low";
}

function aggregateRisk(risks: ReadonlyArray<SemanticDiffRisk>): SemanticDiffRisk {
  if (risks.includes("unknown")) return "unknown";
  if (risks.includes("high")) return "high";
  if (risks.includes("medium")) return "medium";
  return "low";
}

function uniqueAnchors(changes: ReadonlyArray<SemanticDiffChange>): ReadonlyArray<SemanticDiffAnchor> {
  const seen = new Set<string>();
  const anchors: SemanticDiffAnchor[] = [];
  for (const change of changes) {
    if (!change.anchor || seen.has(change.anchor.id)) continue;
    seen.add(change.anchor.id);
    anchors.push(change.anchor);
  }
  return anchors;
}

function summarize(
  format: DocumentFormat,
  changeCount: number,
  risk: SemanticDiffRisk,
  operation?: string,
  fallback?: boolean
): string {
  const op = operation ? `${operation}: ` : "";
  if (fallback) return `${op}opaque ${format.toUpperCase()} change; review raw diagnostics.`;
  if (changeCount === 0) return `${op}no ${format.toUpperCase()} changes.`;
  const noun = changeCount === 1 ? "change" : "changes";
  return `${op}${changeCount} ${format.toUpperCase()} ${noun}; ${risk} review risk.`;
}
