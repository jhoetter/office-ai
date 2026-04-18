import {
  flattenCellXf,
  type CellFormatPatch,
  type EffectiveStyle,
  type Sheet,
  type StyleTable,
  type XlsxAgent,
} from "@officeai/xlsx";
import {
  collapse,
  MIXED,
  normalizeColor,
  type ActiveTextFormat,
  type MaybeMixed,
  type TextFormat,
  type TextFormatProvider,
} from "@officeai/text-formatting";
import {
  formatSelection,
  normalizeSelection,
  type Selection,
} from "./selection";

/**
 * XLSX adapter for the shared TextFormatBar.
 *
 * Reads the per-cell effective style for every cell in the active
 * range and collapses each attribute into a `MaybeMixed` so the
 * shared toolbar can show MIXED placeholders. Patches are dispatched
 * as `xlsx:set-cell-format` with the canonical RRGGBB hex / point
 * units the command handler already speaks.
 *
 * Highlight is reported as `"fill-fallback"` so the shared
 * HighlightPicker maps onto a solid cell fill (`fill.color` +
 * `pattern: "solid"`). XLSX has no character-run highlight; this
 * matches what users expect when they press the highlight chip on
 * a spreadsheet selection.
 */
export interface XlsxProviderDeps {
  readonly agentRef: { readonly current: XlsxAgent | null };
  readonly selectionRef: { readonly current: Selection | null };
  readonly sheetRef: { readonly current: Sheet | null };
  readonly stylesRef: { readonly current: StyleTable | null };
  readonly pushToast: (kind: "info" | "warn" | "error", text: string) => void;
}

const CAPABILITIES = {
  highlight: "fill-fallback",
  underlineVariants: false,
  fontFamily: true,
  fontSize: true,
  strike: true,
} as const;

export function createXlsxFormatProvider(deps: XlsxProviderDeps): TextFormatProvider {
  return {
    capabilities: CAPABILITIES,
    hasSelection: () => deps.selectionRef.current != null,
    getActive: () => readActive(deps.sheetRef.current, deps.stylesRef.current, deps.selectionRef.current),
    apply: (patch) => {
      void dispatchPatch(deps, patch);
    },
  };
}

/**
 * Pure render-path helper that mirrors `provider.getActive()`.
 *
 * Exported so the editor can compute `ActiveTextFormat` from values
 * already in scope during render (`activeSheet`, `selection`,
 * `styles`) without bouncing through the provider's ref-reads.
 */
export function computeXlsxActive(
  sheet: Sheet | null,
  styles: StyleTable | null,
  selection: Selection | null
): ActiveTextFormat {
  return readActive(sheet, styles, selection);
}

const EMPTY: ActiveTextFormat = {
  bold: undefined,
  italic: undefined,
  underline: undefined,
  strike: undefined,
  fontFamily: undefined,
  fontSizePt: undefined,
  color: undefined,
  highlight: undefined,
};

function readActive(
  sheet: Sheet | null,
  styles: StyleTable | null,
  selection: Selection | null
): ActiveTextFormat {
  if (!sheet || !styles || !selection) return EMPTY;

  const effs = collectEffective(sheet, styles, selection);
  if (effs.length === 0) return EMPTY;

  const bold = collapse(effs.map((e) => normBool(e.font.bold)));
  const italic = collapse(effs.map((e) => normBool(e.font.italic)));
  const underline = collapse(effs.map((e) => normUnderline(e.font.underline)));
  const strike = collapse(effs.map((e) => normBool(e.font.strike)));
  const fontFamily = collapse(effs.map((e) => e.font.name));
  const fontSizePt = collapse(effs.map((e) => e.font.size));
  const color = collapse(effs.map((e) => normalizeColor(e.font.color?.rgb)));
  const highlight = collapse(effs.map((e) => normalizeFillColor(e)));

  return { bold, italic, underline, strike, fontFamily, fontSizePt, color, highlight };
}

function collectEffective(
  sheet: Sheet,
  styles: StyleTable,
  selection: Selection
): EffectiveStyle[] {
  const n = normalizeSelection(selection);
  const out: EffectiveStyle[] = [];
  for (let r = n.r0; r <= n.r1; r++) {
    for (let c = n.c0; c <= n.c1; c++) {
      const cell = sheet.cells.get(`${r}:${c}`);
      out.push(flattenCellXf(styles, cell?.styleId));
    }
  }
  return out;
}

function normBool(value: boolean | undefined): boolean {
  return value === true;
}

function normUnderline(value: string | true | undefined): boolean {
  return value === true || (typeof value === "string" && value.length > 0 && value !== "none");
}

function normalizeFillColor(eff: EffectiveStyle): string | undefined {
  if (eff.fill.kind !== "pattern") return undefined;
  if (eff.fill.patternType !== "solid") return undefined;
  return normalizeColor(eff.fill.fgColor?.rgb);
}

async function dispatchPatch(deps: XlsxProviderDeps, patch: TextFormat): Promise<void> {
  const agent = deps.agentRef.current;
  const sheet = deps.sheetRef.current;
  const selection = deps.selectionRef.current;
  if (!agent || !sheet || !selection) {
    deps.pushToast("info", "Select some cells first.");
    return;
  }
  const cellPatch = canonicalToCellFormat(patch);
  if (!cellPatch.font && !cellPatch.fill) return;
  try {
    await agent.applyCommand({
      type: "xlsx:set-cell-format",
      payload: {
        sheet: sheet.name,
        range: formatSelection(selection),
        format: cellPatch,
      },
      source: "human",
    });
  } catch (err) {
    deps.pushToast("error", err instanceof Error ? err.message : String(err));
  }
}

function canonicalToCellFormat(patch: TextFormat): CellFormatPatch {
  const out: { font?: NonNullable<CellFormatPatch["font"]>; fill?: NonNullable<CellFormatPatch["fill"]> } = {};
  const font: Record<string, unknown> = {};
  if (patch.bold !== undefined) font.bold = patch.bold;
  if (patch.italic !== undefined) font.italic = patch.italic;
  if (patch.underline !== undefined) {
    font.underline = patch.underline === false ? false : true;
  }
  if (patch.strike !== undefined) font.strike = patch.strike;
  if (patch.fontFamily !== undefined && patch.fontFamily !== "") {
    font.fontFamily = patch.fontFamily;
  }
  if (patch.fontSizePt !== undefined) font.size = patch.fontSizePt;
  if (patch.color !== undefined) {
    const norm = normalizeColor(patch.color);
    font.color = norm ?? "000000";
  }
  if (Object.keys(font).length > 0) {
    out.font = font as NonNullable<CellFormatPatch["font"]>;
  }

  if (patch.highlight !== undefined) {
    if (patch.highlight === "") {
      out.fill = { pattern: "none" } as NonNullable<CellFormatPatch["fill"]>;
    } else {
      const norm = normalizeColor(patch.highlight);
      if (norm) {
        out.fill = { color: norm, pattern: "solid" } as NonNullable<CellFormatPatch["fill"]>;
      }
    }
  }

  return out as CellFormatPatch;
}

// Re-export so callers needn't re-import the sentinel just to interpret active values.
export { MIXED };
export type { MaybeMixed };
