import { createFormulaEngine, type EngineHost, type FormulaEngine } from "./recalc.js";
import { createRegistry, type FunctionRegistry } from "./function-registry.js";
import { registerAllFunctions } from "./functions/index.js";
import { type CellRef, type RangeRef } from "./references.js";
import { Blank, bool, err, num, str, type Range2D, type Value } from "./values.js";
import { ErrorKinds, type CellErrorKind } from "./errors.js";
import { cellKey } from "../model/refs.js";
import type { CellErrorValue, CellValue, Sheet, XlsxWorkbook } from "../model/types.js";

/**
 * Workbook ↔ formula-engine adapter.
 *
 * Spec: `spec/xlsx/agent-commands.md` §2 + `formula-engine.md` §15.
 *
 * Commands that touch formulas (`xlsx:set-cell-formula`,
 * `xlsx:set-cell-value` when overwriting a formula, …) build a
 * fresh engine over the current `XlsxWorkbook` via
 * `bindEngineToWorkbook`. The engine reads raw cell values through
 * the host (which maps the typed `Cell.value` union to the engine's
 * `Value` discriminated union), evaluates the dependency graph, and
 * exposes the post-recalc cached values back to the command.
 *
 * The function registry is a module-level singleton: there's no
 * per-engine state on the registry, and re-registering the P0
 * library on every command would dominate the recalc cost on small
 * workbooks.
 */

let sharedRegistry: FunctionRegistry | undefined;

function getRegistry(): FunctionRegistry {
  if (!sharedRegistry) {
    const reg = createRegistry();
    registerAllFunctions(reg);
    sharedRegistry = reg;
  }
  return sharedRegistry;
}

/** Build an engine + host bound to a workbook snapshot. */
export function bindEngineToWorkbook(workbook: XlsxWorkbook): {
  engine: FormulaEngine;
  host: WorkbookHost;
} {
  const host = new WorkbookHost(workbook);
  const engine = createFormulaEngine({ registry: getRegistry(), host });
  host.engine = engine;
  return { engine, host };
}

/**
 * `EngineHost` implementation that reads cell values from a workbook
 * snapshot's typed `Sheet.cells` map (and falls through to the engine's
 * formula cache for cells that have been added as formulas during the
 * current recalc cycle).
 */
export class WorkbookHost implements EngineHost {
  engine: FormulaEngine | null = null;
  private sheets = new Map<string, Sheet>();

  constructor(workbook: XlsxWorkbook) {
    for (const s of workbook.sheets) this.sheets.set(s.name, s);
  }

  readCell(ref: CellRef): Value {
    // Engine cache wins for formula cells (precedents are computed first).
    const cached = this.engine?.getCachedValue(ref);
    if (cached) return cached;
    const sheet = this.sheets.get(ref.sheet);
    if (!sheet) return err({ kind: ErrorKinds.REF });
    const cell = sheet.cells.get(cellKey(ref.row, ref.col));
    if (!cell) return Blank;
    return toEngineValue(cell.value);
  }

  readRange(ref: RangeRef): Range2D {
    const out: Value[][] = [];
    const sheet = this.sheets.get(ref.sheet);
    for (let r = ref.r0; r <= ref.r1; r++) {
      const row: Value[] = [];
      for (let c = ref.c0; c <= ref.c1; c++) {
        const cellRef: CellRef = { sheet: ref.sheet, row: r, col: c, abs: 0 };
        const cached = this.engine?.getCachedValue(cellRef);
        if (cached) {
          row.push(cached);
          continue;
        }
        if (!sheet) {
          row.push(err({ kind: ErrorKinds.REF }));
          continue;
        }
        const cell = sheet.cells.get(cellKey(r, c));
        row.push(cell ? toEngineValue(cell.value) : Blank);
      }
      out.push(row);
    }
    return out;
  }

  /**
   * Seed the engine with every formula cell in the workbook so the
   * dependency graph reflects the steady-state.
   *
   * Returns the list of `{ ref, formula }` pairs that failed to parse;
   * those cells stay literal-only and the parse-error surfaces at
   * import time per `EC-F3`.
   */
  seedFormulas(engine: FormulaEngine): Array<{ ref: CellRef; error: Error }> {
    const failures: Array<{ ref: CellRef; error: Error }> = [];
    for (const [name, sheet] of this.sheets) {
      for (const [, cell] of sheet.cells) {
        if (!cell.formula) continue;
        const ref: CellRef = { sheet: name, row: cell.row, col: cell.col, abs: 0 };
        try {
          const parsed = engine.parse(cell.formula.text, ref);
          engine.addCell(ref, parsed, toEngineValue(cell.value));
        } catch (e) {
          failures.push({ ref, error: e as Error });
        }
      }
    }
    return failures;
  }
}

/** Convert the model's `CellValue` union to an engine `Value`. */
export function toEngineValue(v: CellValue | undefined): Value {
  if (v === null || v === undefined) return Blank;
  if (typeof v === "number") return num(v);
  if (typeof v === "string") return str(v);
  if (typeof v === "boolean") return bool(v);
  if (v.kind === "error") return err({ kind: errCodeToKind(v.code) });
  return Blank;
}

/** Convert an engine `Value` back to the model's `CellValue` for storage. */
export function fromEngineValue(v: Value): CellValue {
  switch (v.kind) {
    case "n":
      return v.v;
    case "s":
      return v.v;
    case "b":
      return v.v;
    case "e":
      return { kind: "error", code: kindToErrCode(v.v.kind) };
    case "r":
      // 1×1 → scalar; otherwise stash first cell (Excel collapses arrays at storage).
      if (v.v.length === 1 && v.v[0].length === 1) return fromEngineValue(v.v[0][0]);
      return fromEngineValue(v.v[0][0]);
  }
}

/** Reverse map of model error codes ↔ engine `CellErrorKind`. */
function errCodeToKind(code: CellErrorValue["code"]): CellErrorKind {
  switch (code) {
    case "#REF!":
      return ErrorKinds.REF;
    case "#VALUE!":
      return ErrorKinds.VALUE;
    case "#DIV/0!":
      return ErrorKinds.DIV0;
    case "#NAME?":
      return ErrorKinds.NAME;
    case "#N/A":
      return ErrorKinds.NA;
    case "#NULL!":
      return ErrorKinds.NULL;
    case "#NUM!":
      return ErrorKinds.NUM;
    case "#GETTING_DATA":
      return ErrorKinds.GETTING_DATA;
    case "#SPILL!":
      return ErrorKinds.SPILL;
  }
}

function kindToErrCode(kind: CellErrorKind): CellErrorValue["code"] {
  switch (kind) {
    case ErrorKinds.REF:
      return "#REF!";
    case ErrorKinds.VALUE:
      return "#VALUE!";
    case ErrorKinds.DIV0:
      return "#DIV/0!";
    case ErrorKinds.NAME:
      return "#NAME?";
    case ErrorKinds.NA:
      return "#N/A";
    case ErrorKinds.NULL:
      return "#NULL!";
    case ErrorKinds.NUM:
      return "#NUM!";
    case ErrorKinds.GETTING_DATA:
      return "#GETTING_DATA";
    case ErrorKinds.SPILL:
      return "#SPILL!";
    case ErrorKinds.CALC:
      // No model code for #CALC! — collapse to #VALUE! at the boundary.
      return "#VALUE!";
    case ErrorKinds.CYCLE:
      // Cycle errors surface as #REF! (with meta.cycle) per spec §6.
      return "#REF!";
  }
}
