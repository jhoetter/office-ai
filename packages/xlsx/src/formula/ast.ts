import type { CellRef, RangeRef } from "./references.js";
import type { Value } from "./values.js";

/**
 * AST node types — discriminated union, post-order walked by the
 * evaluator. Every node carries a source span for diagnostics.
 *
 * Spec: `spec/xlsx/formula-engine.md` §9.
 */
export type AstNode =
  | LiteralNode
  | RefNode
  | RangeRefNode
  | NameNode
  | BinaryNode
  | UnaryNode
  | PercentNode
  | CallNode
  | ArrayLitNode;

export interface NodeBase {
  readonly start: number;
  readonly end: number;
}

export interface LiteralNode extends NodeBase {
  readonly kind: "lit";
  readonly value: Value;
}

export interface RefNode extends NodeBase {
  readonly kind: "ref";
  readonly ref: CellRef;
}

export interface RangeRefNode extends NodeBase {
  readonly kind: "range";
  readonly ref: RangeRef;
}

export interface NameNode extends NodeBase {
  readonly kind: "name";
  readonly name: string;
}

export interface BinaryNode extends NodeBase {
  readonly kind: "binary";
  readonly op: BinaryOp;
  readonly left: AstNode;
  readonly right: AstNode;
}

export interface UnaryNode extends NodeBase {
  readonly kind: "unary";
  readonly op: UnaryOp;
  readonly operand: AstNode;
}

export interface PercentNode extends NodeBase {
  readonly kind: "pct";
  readonly operand: AstNode;
}

export interface CallNode extends NodeBase {
  readonly kind: "call";
  readonly name: string; //  upper-cased function name
  readonly args: ReadonlyArray<AstNode>;
}

export interface ArrayLitNode extends NodeBase {
  readonly kind: "array";
  readonly rows: ReadonlyArray<ReadonlyArray<AstNode>>;
}

export type BinaryOp = "+" | "-" | "*" | "/" | "^" | "&" | "=" | "<>" | "<" | ">" | "<=" | ">=";
export type UnaryOp = "+" | "-";

/** Wrapper around an AST + source text + collected dependencies. */
export interface Formula {
  readonly text: string;
  readonly ast: AstNode;
  readonly anchor: CellRef;
  readonly dependencies: ReadonlyArray<CellRef | RangeRef>;
  readonly volatile: boolean;
}

/** Volatile function names (force-dirty every recalc). */
export const VOLATILE_FUNCTIONS: ReadonlySet<string> = new Set([
  "RAND",
  "RANDBETWEEN",
  "NOW",
  "TODAY",
  "INDIRECT",
  "OFFSET",
]);

/**
 * Walk an AST post-order, collecting every CellRef / RangeRef that
 * appears as a `ref` / `range` node. De-duplicated by serialised key,
 * preserved in encounter order. Per `formula-engine.md` §12.3.
 */
export function collectDependencies(ast: AstNode): ReadonlyArray<CellRef | RangeRef> {
  const out: Array<CellRef | RangeRef> = [];
  const seen = new Set<string>();
  walk(ast, (n) => {
    if (n.kind === "ref") {
      const k = `c|${n.ref.sheet}|${n.ref.row}|${n.ref.col}`;
      if (!seen.has(k)) {
        seen.add(k);
        out.push(n.ref);
      }
    } else if (n.kind === "range") {
      const k = `r|${n.ref.sheet}|${n.ref.r0}|${n.ref.c0}|${n.ref.r1}|${n.ref.c1}`;
      if (!seen.has(k)) {
        seen.add(k);
        out.push(n.ref);
      }
    }
  });
  return out;
}

/** Returns true iff any CallNode below `ast` is in `VOLATILE_FUNCTIONS`. */
export function containsVolatile(ast: AstNode): boolean {
  let found = false;
  walk(ast, (n) => {
    if (n.kind === "call" && VOLATILE_FUNCTIONS.has(n.name)) found = true;
  });
  return found;
}

function walk(node: AstNode, visit: (n: AstNode) => void): void {
  visit(node);
  switch (node.kind) {
    case "lit":
    case "ref":
    case "range":
    case "name":
      return;
    case "binary":
      walk(node.left, visit);
      walk(node.right, visit);
      return;
    case "unary":
    case "pct":
      walk(node.operand, visit);
      return;
    case "call":
      for (const a of node.args) walk(a, visit);
      return;
    case "array":
      for (const row of node.rows) for (const c of row) walk(c, visit);
      return;
  }
}

export class FormulaParseError extends Error {
  readonly code: FormulaParseErrorCode;
  readonly span: { start: number; end: number };
  readonly hint?: string;

  constructor(
    code: FormulaParseErrorCode,
    message: string,
    span: { start: number; end: number },
    hint?: string
  ) {
    super(message);
    this.name = "FormulaParseError";
    this.code = code;
    this.span = span;
    if (hint !== undefined) this.hint = hint;
  }
}

export type FormulaParseErrorCode =
  | "empty-formula"
  | "unexpected-token"
  | "unexpected-eof"
  | "unbalanced-paren"
  | "intersection-operator-not-supported"
  | "implicit-intersection-not-supported"
  | "structured-ref-not-supported"
  | "external-ref-not-supported"
  | "3d-ref-not-supported"
  | "malformed-expression"
  | "invalid-ref"
  | "invalid-number"
  | "unterminated-string";
