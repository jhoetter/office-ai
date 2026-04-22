import { ErrorKinds } from "../errors.js";
import { arity, type EagerFn, type MutableFunctionRegistry } from "../function-registry.js";
import { err } from "../values.js";

/**
 * Pivot-aware function stubs (`GETPIVOTDATA`, `CUBEMEMBER`, `CUBEVALUE`,
 * `CUBESET`).
 *
 * Spec: `spec/xlsx/pivot-tables.md` §"Phase 4 — formula integration".
 *
 * These functions are *registered* — meaning the parser stops emitting
 * `#NAME?` for the bare token and the evaluator can wire them up via
 * the standard registry — but they always evaluate to `#NAME?` until
 * the pivot model + cache resolver lands. The stub serves three jobs:
 *
 *  1. Roundtrip safety: a workbook authored in Excel that contains
 *     `=GETPIVOTDATA(...)` survives reparse/serialize without the
 *     parser tripping on an unknown function during validation.
 *  2. Discoverability: the function list surfaced to the agent
 *     mentions every pivot helper from day one, so prompts can be
 *     written against a stable surface area.
 *  3. Forward compatibility: when Phase-4 (pivot evaluation) lands,
 *     swapping the body is a one-file change with no parser plumbing
 *     work required.
 *
 * The chosen error kind is `#NAME?` rather than `#REF!` or `#N/A` so a
 * grep for `#NAME?` in fixtures still spots genuinely missing
 * functions; pivot-evaluation gaps will be tracked separately as
 * "implemented stub returns #NAME?" instead.
 */
export function registerPivotStubs(reg: MutableFunctionRegistry): void {
  // GETPIVOTDATA(data_field, pivot_table, [field1, item1], …)
  // Min arity is 2. Excel allows up to 254 arguments.
  reg.register({ name: "GETPIVOTDATA", arity: arity(2, 254), fn: stubFn });
  // CUBE-* family (OLAP). All accept 1+ args; we keep the upper bound
  // permissive so stubbed calls don't blow up on long member lists.
  reg.register({ name: "CUBEMEMBER", arity: arity(2, 4), fn: stubFn });
  reg.register({ name: "CUBEVALUE", arity: arity(1, 254), fn: stubFn });
  reg.register({ name: "CUBESET", arity: arity(2, 5), fn: stubFn });
  reg.register({ name: "CUBEMEMBERPROPERTY", arity: arity(3, 3), fn: stubFn });
  reg.register({ name: "CUBESETCOUNT", arity: arity(1, 1), fn: stubFn });
  reg.register({ name: "CUBERANKEDMEMBER", arity: arity(3, 4), fn: stubFn });
  reg.register({ name: "CUBEKPIMEMBER", arity: arity(3, 4), fn: stubFn });
}

const stubFn: EagerFn = () => err(ErrorKinds.NAME);
