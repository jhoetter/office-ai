import { ErrorKinds } from "../errors.js";
import { arity, type EagerFn, type MutableFunctionRegistry } from "../function-registry.js";
import { Blank, bool, err, isError, num, toNumber, type Value } from "../values.js";

/**
 * Info functions (P0).
 *
 * Spec: `spec/xlsx/formula-engine.md` §16.7.
 *
 * Two design points worth flagging up-front:
 *
 * 1. `IS*` predicates intentionally do NOT propagate errors. They are
 *    *type inspectors* — `=ISERROR(1/0)` must return `TRUE` and
 *    `=ISNUMBER(1/0)` must return `FALSE`. Only `ISODD` / `ISEVEN`
 *    propagate, since they actually consume the underlying number.
 * 2. `ISBLANK` returns `TRUE` only when the engine receives the
 *    exported `Blank` sentinel reference (object identity). At the
 *    `Value` level there's no other way to distinguish a deliberately
 *    typed `0` from an empty cell — Excel's distinction lives in the
 *    cell-data layer, not the formula runtime. Hosts that want full
 *    parity must surface empty cells via the `Blank` singleton from
 *    `getCell`.
 */
export function registerInfo(reg: MutableFunctionRegistry): void {
  reg.register({ name: "ISBLANK", arity: arity(1, 1), fn: isBlankFn });
  reg.register({ name: "ISNUMBER", arity: arity(1, 1), fn: isNumberFn });
  reg.register({ name: "ISTEXT", arity: arity(1, 1), fn: isTextFn });
  reg.register({ name: "ISERROR", arity: arity(1, 1), fn: isErrorFn });
  reg.register({ name: "ISNA", arity: arity(1, 1), fn: isNaFn });
  reg.register({ name: "ISODD", arity: arity(1, 1), fn: isOddFn });
  reg.register({ name: "ISEVEN", arity: arity(1, 1), fn: isEvenFn });
  reg.register({ name: "TYPE", arity: arity(1, 1), fn: typeFn });
  reg.register({ name: "N", arity: arity(1, 1), fn: nFn });
  reg.register({ name: "NA", arity: arity(0, 0), fn: naFn });
}

const isBlankFn: EagerFn = (args) => bool(args[0] === Blank);

const isNumberFn: EagerFn = (args) => bool(args[0].kind === "n");

const isTextFn: EagerFn = (args) => bool(args[0].kind === "s");

const isErrorFn: EagerFn = (args) => bool(args[0].kind === "e");

const isNaFn: EagerFn = (args) => {
  const v = args[0];
  return bool(v.kind === "e" && v.v.kind === ErrorKinds.NA);
};

const isOddFn: EagerFn = (args) => {
  const v = args[0];
  if (v.kind === "e") return v;
  const n = toNumber(v);
  if (isError(n)) return n;
  const t = Math.trunc(n.v);
  return bool(t % 2 !== 0);
};

const isEvenFn: EagerFn = (args) => {
  const v = args[0];
  if (v.kind === "e") return v;
  const n = toNumber(v);
  if (isError(n)) return n;
  const t = Math.trunc(n.v);
  return bool(t % 2 === 0);
};

const typeFn: EagerFn = (args) => {
  const v: Value = args[0];
  switch (v.kind) {
    case "n":
      return num(1);
    case "s":
      return num(2);
    case "b":
      return num(4);
    case "e":
      return num(16);
    case "r":
      return num(64);
  }
};

const nFn: EagerFn = (args) => {
  const v: Value = args[0];
  switch (v.kind) {
    case "n":
      return v;
    case "b":
      return num(v.v ? 1 : 0);
    case "s":
      return num(0);
    case "e":
      return v;
    case "r": {
      // Implicit-intersection-style collapse: a 1×1 range is treated
      // as its scalar; anything larger yields `#VALUE!` per the
      // existing `toNumber` semantics in `values.ts`.
      const n = toNumber(v);
      if (isError(n)) return n;
      return n;
    }
  }
};

const naFn: EagerFn = () => err(ErrorKinds.NA);
