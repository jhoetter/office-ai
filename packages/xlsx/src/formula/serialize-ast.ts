import type { AstNode, BinaryOp, LiteralNode } from "./ast.js";
import { serializeCellRef, serializeRangeRef } from "./references.js";
import type { Value } from "./values.js";

/**
 * Emit canonical Excel formula text from an AST. Used by
 * `rewriteFormulaRefs` after a structural shift mutates `ref` /
 * `range` nodes so the rewritten cell carries fresh formula text.
 *
 * Spec: `spec/xlsx/agent-commands.md` §§5–8 (the structural commands
 * delegate to this for the "re-emit formula" step).
 *
 * The output is **semantically** equivalent to the parsed AST. We
 * track the surrounding operator precedence so binary children get
 * parens iff dropping them would re-associate the expression on a
 * round-trip (e.g. `(1+2)*3` keeps its parens; `1+2*3` does not).
 * Whitespace is dropped and function names are uppercased — both
 * match Excel's own canonicalisation on edit.
 */
export function serializeAst(node: AstNode, anchorSheet?: string): string {
  return emit(node, anchorSheet, 0);
}

const PRECEDENCE: Record<BinaryOp, number> = {
  "^": 8,
  "*": 7,
  "/": 7,
  "+": 6,
  "-": 6,
  "&": 5,
  "=": 4,
  "<>": 4,
  "<": 4,
  ">": 4,
  "<=": 4,
  ">=": 4,
};

const RIGHT_ASSOC: ReadonlySet<BinaryOp> = new Set<BinaryOp>(["^"]);

function emit(node: AstNode, anchorSheet: string | undefined, parentPrec: number): string {
  switch (node.kind) {
    case "lit":
      return emitLiteral(node);
    case "ref":
      return serializeCellRef(node.ref, anchorSheet ? { sheet: anchorSheet } : undefined);
    case "range":
      return serializeRangeRef(node.ref, anchorSheet ? { sheet: anchorSheet } : undefined);
    case "name":
      return node.name;
    case "binary": {
      const myPrec = PRECEDENCE[node.op];
      const rightAssoc = RIGHT_ASSOC.has(node.op);
      const leftMin = rightAssoc ? myPrec + 1 : myPrec;
      const rightMin = rightAssoc ? myPrec : myPrec + 1;
      const left = emit(node.left, anchorSheet, leftMin);
      const right = emit(node.right, anchorSheet, rightMin);
      const text = `${left}${node.op}${right}`;
      return parentPrec > myPrec ? `(${text})` : text;
    }
    case "unary": {
      // Unary `-` binds tighter than `^` (Excel quirk: `-2^2 == 4`).
      const inner = emit(node.operand, anchorSheet, 9);
      return `${node.op}${inner}`;
    }
    case "pct": {
      const inner = emit(node.operand, anchorSheet, 9);
      return `${inner}%`;
    }
    case "call": {
      const args = node.args.map((a) => emit(a, anchorSheet, 0)).join(",");
      return `${node.name}(${args})`;
    }
    case "array": {
      const rows = node.rows.map((row) => row.map((c) => emit(c, anchorSheet, 0)).join(",")).join(";");
      return `{${rows}}`;
    }
  }
}

function emitLiteral(node: LiteralNode): string {
  const v: Value = node.value;
  switch (v.kind) {
    case "n":
      return formatNumber(v.v);
    case "s":
      return `"${v.v.replace(/"/g, '""')}"`;
    case "b":
      return v.v ? "TRUE" : "FALSE";
    case "e":
      return v.v.kind;
    case "r":
      // Range values are runtime-only — they never appear as literal
      // AST nodes from a parse. Emit as an empty array literal so the
      // serializer is total over the union without throwing.
      return "{}";
  }
}

function formatNumber(n: number): string {
  if (Number.isInteger(n)) return n.toString();
  return Number(n.toPrecision(15)).toString();
}
