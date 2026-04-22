import { createRegistry, type FunctionEntry, type MutableFunctionRegistry } from "./function-registry.js";
import { registerInfo } from "./functions/info.js";
import { registerLogic } from "./functions/logic.js";
import { registerLookup } from "./functions/lookup.js";
import { registerMath } from "./functions/math.js";
import { registerPivotStubs } from "./functions/pivot.js";
import { registerText } from "./functions/text.js";

export type FunctionCategory = "math" | "logic" | "info" | "lookup" | "text" | "pivot";

export interface RegisteredFunctionInfo {
  readonly name: string;
  readonly arity: { readonly min: number; readonly max: number };
  readonly category: FunctionCategory;
  readonly volatile: boolean;
}

/**
 * Public catalogue of every P0 formula function: name, arity, and the
 * category bucket it lives under.
 *
 * Used by the web editor's formula autocomplete (`FormulaSuggest.tsx`)
 * to surface a "type-ahead" of function names without forcing
 * consumers to import the per-category register helpers themselves.
 *
 * The data is built by registering each category into a throwaway
 * registry wrapped in a category-tagging proxy. This keeps the
 * single-source-of-truth in `functions/` (no parallel constant to
 * maintain) and lets new functions automatically appear in the
 * autocomplete the moment they are wired up.
 */
let cached: ReadonlyArray<RegisteredFunctionInfo> | undefined;

export function listRegisteredFunctions(): ReadonlyArray<RegisteredFunctionInfo> {
  if (cached) return cached;
  const out: RegisteredFunctionInfo[] = [];
  const collect = (category: FunctionCategory, register: (r: MutableFunctionRegistry) => void) => {
    const target = createRegistry();
    const proxy: MutableFunctionRegistry = {
      register(entry: FunctionEntry): void {
        target.register(entry);
        out.push({
          name: entry.name.toUpperCase(),
          arity: { min: entry.arity.min, max: entry.arity.max },
          category,
          volatile: !!entry.volatile,
        });
      },
      get: target.get.bind(target),
      has: target.has.bind(target),
      entries: target.entries.bind(target),
      volatileNames: target.volatileNames.bind(target),
    };
    register(proxy);
  };

  collect("math", registerMath);
  collect("logic", registerLogic);
  collect("info", registerInfo);
  collect("lookup", registerLookup);
  collect("text", registerText);
  // Surface pivot/CUBE stubs in the autocomplete so users can discover
  // them; today they always evaluate to `#NAME?` (see
  // `functions/pivot.ts` and `spec/xlsx/pivot-tables.md`).
  collect("pivot", registerPivotStubs);

  // Sort alphabetically by name for stable autocomplete ordering.
  out.sort((a, b) => (a.name < b.name ? -1 : a.name > b.name ? 1 : 0));
  cached = out;
  return cached;
}
