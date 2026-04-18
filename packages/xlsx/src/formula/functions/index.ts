import type { MutableFunctionRegistry } from "../function-registry.js";
import { registerInfo } from "./info.js";
import { registerLogic } from "./logic.js";
import { registerLookup } from "./lookup.js";
import { registerMath } from "./math.js";
import { registerText } from "./text.js";

/**
 * Bulk-register every P0 function across all categories.
 *
 * Spec: `spec/xlsx/formula-engine.md` §16. Each category lives in
 * its own file under `functions/` and exposes a `registerXxx(reg)`
 * helper; this aggregator lets the workbook layer wire the entire
 * library in one call.
 *
 * Date / Finance categories are P1 (per spec §16.5 / §16.6); they
 * are intentionally absent here so unknown calls fall through to
 * the evaluator's `#NAME?` path.
 */
export function registerAllFunctions(reg: MutableFunctionRegistry): void {
  registerMath(reg);
  registerLogic(reg);
  registerInfo(reg);
  registerLookup(reg);
  registerText(reg);
}

export { registerInfo, registerLogic, registerLookup, registerMath, registerText };
