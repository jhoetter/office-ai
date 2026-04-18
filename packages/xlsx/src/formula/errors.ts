/**
 * Excel error values — single source of truth for the formula engine.
 *
 * The seven canonical Excel errors (`#DIV/0!`, `#NAME?`, `#VALUE!`,
 * `#NUM!`, `#N/A`, `#REF!`, `#NULL!`) are fully supported runtime
 * values. The four "modern" errors (`#SPILL!`, `#CALC!`, `#CYCLE!`,
 * `#GETTING_DATA`) are recognised as literal tokens so they round-trip
 * but the P0 engine never internally generates them.
 *
 * Spec: `spec/xlsx/formula-engine.md` §6.
 */

export type CellErrorKind =
  | "#DIV/0!"
  | "#NAME?"
  | "#VALUE!"
  | "#NUM!"
  | "#N/A"
  | "#REF!"
  | "#NULL!"
  | "#SPILL!"
  | "#CALC!"
  | "#CYCLE!"
  | "#GETTING_DATA";

/** Named constants for the canonical Excel error kinds. */
export const ErrorKinds = {
  DIV0: "#DIV/0!",
  NAME: "#NAME?",
  VALUE: "#VALUE!",
  NUM: "#NUM!",
  NA: "#N/A",
  REF: "#REF!",
  NULL: "#NULL!",
  SPILL: "#SPILL!",
  CALC: "#CALC!",
  CYCLE: "#CYCLE!",
  GETTING_DATA: "#GETTING_DATA",
} as const satisfies Record<string, CellErrorKind>;

export interface CellError {
  readonly kind: CellErrorKind;
  readonly meta?: Readonly<Record<string, unknown>>;
}

const ALL_KINDS: ReadonlyArray<CellErrorKind> = [
  ErrorKinds.DIV0,
  ErrorKinds.NAME,
  ErrorKinds.VALUE,
  ErrorKinds.NUM,
  ErrorKinds.NA,
  ErrorKinds.REF,
  ErrorKinds.NULL,
  ErrorKinds.SPILL,
  ErrorKinds.CALC,
  ErrorKinds.CYCLE,
  ErrorKinds.GETTING_DATA,
];

/**
 * Interned singleton per error kind (no `meta`). Cheap propagation:
 * `add(#DIV/0!, x)` returns the same `Errors[#DIV/0!]` object.
 */
export const Errors: Readonly<Record<CellErrorKind, CellError>> = Object.freeze(
  Object.fromEntries(ALL_KINDS.map((k) => [k, Object.freeze({ kind: k })])) as Record<
    CellErrorKind,
    CellError
  >
);

const KIND_LOOKUP = new Map<string, CellErrorKind>(ALL_KINDS.map((k) => [k, k]));

/** Recognise an Excel error literal token. Returns `undefined` if no match. */
export function parseErrorLiteral(text: string): CellError | undefined {
  const k = KIND_LOOKUP.get(text);
  return k === undefined ? undefined : Errors[k];
}

/** `#REF!` enriched with a structured cycle path (for `EC-F1`). */
export function refWithCycle(cycle: ReadonlyArray<string>): CellError {
  return { kind: ErrorKinds.REF, meta: { cycle: [...cycle] } };
}

/** `#REF!` enriched with the now-deleted target (for `EC-R2` / `EC-F4`). */
export function refWithDeletedTarget(deletedRef: string): CellError {
  return { kind: ErrorKinds.REF, meta: { deletedRef } };
}
