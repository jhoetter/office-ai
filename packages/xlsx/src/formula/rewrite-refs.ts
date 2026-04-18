import type { AstNode, LiteralNode } from "./ast.js";
import { ErrorKinds, type CellError } from "./errors.js";
import { parse } from "./parser.js";
import { isRefError, type CellRef, type RangeRef } from "./references.js";
import { serializeAst } from "./serialize-ast.js";
import { err } from "./values.js";

export type AdjustFn = (ref: CellRef | RangeRef) => CellRef | RangeRef | CellError;

export interface RewriteResult {
  /** New formula text WITHOUT a leading `=`. */
  readonly text: string;
  /** True iff at least one ref became `#REF!` during the rewrite. */
  readonly hasRefError: boolean;
  /** True iff any ref's identity changed (or became an error). */
  readonly changed: boolean;
}

/**
 * Re-parse `formulaText` against `anchor`, walk the AST replacing
 * every `ref` / `range` node by feeding its target through `adjust`,
 * and re-emit canonical formula text.
 *
 * Per `EC-R2`: when `adjust` returns `#REF!` (the cell or range was
 * inside a deletion band), the node is rewritten to a literal
 * `#REF!` token — Excel's canonical "deleted target" rendering.
 *
 * Used by the four structural reshape commands. Cross-sheet refs are
 * unaffected by an adjust function that scopes itself to a single
 * sheet.
 */
export function rewriteFormulaRefs(formulaText: string, anchor: CellRef, adjust: AdjustFn): RewriteResult {
  const parsed = parse(formulaText, { anchor });
  const acc = { changed: false, hasRefError: false };
  const next = rewriteNode(parsed.ast, adjust, acc);
  const text = serializeAst(next, anchor.sheet);
  return { text, changed: acc.changed, hasRefError: acc.hasRefError };
}

interface Acc {
  changed: boolean;
  hasRefError: boolean;
}

function rewriteNode(node: AstNode, adjust: AdjustFn, acc: Acc): AstNode {
  switch (node.kind) {
    case "lit":
    case "name":
      return node;
    case "ref": {
      const next = adjust(node.ref);
      if (isRefError(next)) {
        acc.changed = true;
        acc.hasRefError = true;
        return refErrorLiteral(node.start, node.end);
      }
      if (!isCellRefShape(next)) {
        // adjust returned a RangeRef where a CellRef was expected —
        // this should never happen in our adjust functions; treat
        // defensively as no-op.
        return node;
      }
      if (cellRefEqual(node.ref, next)) return node;
      acc.changed = true;
      return { kind: "ref", ref: next, start: node.start, end: node.end };
    }
    case "range": {
      const next = adjust(node.ref);
      if (isRefError(next)) {
        acc.changed = true;
        acc.hasRefError = true;
        return refErrorLiteral(node.start, node.end);
      }
      if (isCellRefShape(next)) {
        return node;
      }
      if (rangeRefEqual(node.ref, next)) return node;
      acc.changed = true;
      return { kind: "range", ref: next, start: node.start, end: node.end };
    }
    case "binary": {
      const left = rewriteNode(node.left, adjust, acc);
      const right = rewriteNode(node.right, adjust, acc);
      if (left === node.left && right === node.right) return node;
      return { ...node, left, right };
    }
    case "unary": {
      const operand = rewriteNode(node.operand, adjust, acc);
      if (operand === node.operand) return node;
      return { ...node, operand };
    }
    case "pct": {
      const operand = rewriteNode(node.operand, adjust, acc);
      if (operand === node.operand) return node;
      return { ...node, operand };
    }
    case "call": {
      let any = false;
      const args = node.args.map((a) => {
        const out = rewriteNode(a, adjust, acc);
        if (out !== a) any = true;
        return out;
      });
      if (!any) return node;
      return { ...node, args };
    }
    case "array": {
      let any = false;
      const rows = node.rows.map((row) =>
        row.map((cell) => {
          const out = rewriteNode(cell, adjust, acc);
          if (out !== cell) any = true;
          return out;
        })
      );
      if (!any) return node;
      return { ...node, rows };
    }
  }
}

function refErrorLiteral(start: number, end: number): LiteralNode {
  return { kind: "lit", value: err({ kind: ErrorKinds.REF }), start, end };
}

function isCellRefShape(r: CellRef | RangeRef): r is CellRef {
  return "row" in r;
}

function cellRefEqual(a: CellRef, b: CellRef): boolean {
  return a.sheet === b.sheet && a.row === b.row && a.col === b.col && a.abs === b.abs;
}

function rangeRefEqual(a: RangeRef, b: RangeRef): boolean {
  return (
    a.sheet === b.sheet &&
    a.r0 === b.r0 &&
    a.c0 === b.c0 &&
    a.r1 === b.r1 &&
    a.c1 === b.c1 &&
    a.abs0 === b.abs0 &&
    a.abs1 === b.abs1
  );
}
