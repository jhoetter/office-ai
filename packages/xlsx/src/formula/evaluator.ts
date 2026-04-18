import type { AstNode, BinaryOp } from "./ast.js";
import { ErrorKinds } from "./errors.js";
import type { EvalContext, LazyEntry, LazyEvalAccess } from "./function-registry.js";
import {
  add,
  concat,
  div,
  eq,
  err,
  gt,
  gte,
  lt,
  lte,
  mul,
  neg,
  neq,
  pct,
  pow,
  rangeVal,
  sub,
  type Range2D,
  type Value,
} from "./values.js";

/**
 * Tree-walking evaluator. Single post-order pass over the AST.
 *
 * Spec: `spec/xlsx/formula-engine.md` §13.
 *
 * The evaluator is small on purpose: every binary op delegates to a
 * helper in `values.ts` (so error short-circuit and coercion live in
 * exactly one place) and every function call delegates to the
 * registry. The only complexity here is the lazy-arg dispatch path.
 */
export function evaluate(node: AstNode, ctx: EvalContext): Value {
  switch (node.kind) {
    case "lit":
      return node.value;
    case "ref":
      return ctx.getCell(node.ref);
    case "range":
      return rangeVal(ctx.getRange(node.ref));
    case "name": {
      const resolved = ctx.resolveName(node.name);
      if (!resolved) return err(ErrorKinds.NAME);
      return "row" in resolved ? ctx.getCell(resolved) : rangeVal(ctx.getRange(resolved));
    }
    case "binary": {
      const l = evaluate(node.left, ctx);
      const r = evaluate(node.right, ctx);
      return applyBinary(node.op, l, r);
    }
    case "unary": {
      const o = evaluate(node.operand, ctx);
      return node.op === "-" ? neg(o) : o;
    }
    case "pct":
      return pct(evaluate(node.operand, ctx));
    case "call": {
      const impl = ctx.registry.get(node.name);
      if (!impl) return err(ErrorKinds.NAME);
      if (!impl.arity.accepts(node.args.length)) return err(ErrorKinds.NA);
      if (impl.lazyArgs) return invokeLazy(impl, node.args, ctx);
      const args = node.args.map((a) => evaluate(a, ctx));
      return impl.fn(args, ctx);
    }
    case "array":
      return rangeVal(evaluateArrayLiteral(node.rows, ctx));
  }
}

function evaluateArrayLiteral(rows: ReadonlyArray<ReadonlyArray<AstNode>>, ctx: EvalContext): Range2D {
  return rows.map((row) => row.map((c) => evaluate(c, ctx)));
}

function invokeLazy(impl: LazyEntry, args: ReadonlyArray<AstNode>, ctx: EvalContext): Value {
  const lazy: LazyEvalAccess = {
    evaluate: (node) => evaluate(node, ctx),
    ctx,
  };
  return impl.fn(args, lazy);
}

const BINARY_OPS: Readonly<Record<BinaryOp, (a: Value, b: Value) => Value>> = {
  "+": add,
  "-": sub,
  "*": mul,
  "/": div,
  "^": pow,
  "&": concat,
  "=": eq,
  "<>": neq,
  "<": lt,
  ">": gt,
  "<=": lte,
  ">=": gte,
};

function applyBinary(op: BinaryOp, a: Value, b: Value): Value {
  return BINARY_OPS[op](a, b);
}
