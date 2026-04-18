import type { EditorView } from "prosemirror-view";
import type { DocxAgent, DocxSnapshot, MountResult, TextFormat as DocxTextFormat } from "@officeai/docx";
import { NotImplementedError } from "@officeai/core";
import {
  MIXED,
  halfPointsToPt,
  highlightByDocxName,
  nearestHighlight,
  normalizeColor,
  ptToHalfPoints,
  type ActiveTextFormat,
  type MaybeMixed,
  type TextFormat,
  type TextFormatProvider,
} from "@officeai/text-formatting";
import {
  activeMarks as computeActiveMarks,
  activeRunAttr,
  pmSelectionToRange,
  MIXED as DOCX_MIXED,
  type MaybeMixed as DocxMaybeMixed,
} from "@/lib/format-helpers";

/**
 * DOCX adapter for the shared TextFormatBar.
 *
 * Reads run-level formatting through the existing PM-aware helpers
 * (which already understand the DOCX style cascade) and translates
 * canonical patches back into `docx:format-range` payloads with
 * native units (half-points) and DOCX highlight enum names.
 *
 * The adapter is built around `MutableRefObject` accessors rather
 * than React-state values so the React Compiler doesn't flag it as
 * "passing a ref to a function during render". This mirrors the
 * pattern already used by the editor's per-action callbacks
 * (`mountRef.current?.view`, `agentRef.current`).
 */

export interface DocxProviderDeps {
  mountRef: { readonly current: MountResult | null };
  agentRef: { readonly current: DocxAgent | null };
  /** Pushes a toast — used for "select some text first" / errors. */
  pushToast: (kind: "info" | "warn" | "error", text: string) => void;
}

const CAPABILITIES = {
  highlight: "native",
  underlineVariants: false,
  fontFamily: true,
  fontSize: true,
  strike: true,
} as const;

/** Translate the local (DOCX) MIXED sentinel into the shared sentinel. */
function bridge<T>(value: DocxMaybeMixed<T>): MaybeMixed<T> {
  if (value === DOCX_MIXED) return MIXED;
  return value as MaybeMixed<T>;
}

export function createDocxFormatProvider(deps: DocxProviderDeps): TextFormatProvider {
  return {
    capabilities: CAPABILITIES,
    hasSelection: () => {
      const view = deps.mountRef.current?.view ?? null;
      return !!view && !view.state.selection.empty;
    },
    getActive: () => {
      const view = deps.mountRef.current?.view ?? null;
      const snapshot = deps.agentRef.current?.getSnapshot() ?? null;
      return readActive(view, snapshot);
    },
    apply: (patch) => {
      void dispatchPatch(deps, patch);
    },
  };
}

/**
 * Eagerly compute the ActiveTextFormat for the current PM view.
 * Exported separately so the editor can call it during render
 * without going through `provider.getActive()` — the render path
 * already has `view` and `snapshot` in scope.
 */
export function computeDocxActive(
  view: EditorView | null,
  snapshot: DocxSnapshot | null
): ActiveTextFormat {
  return readActive(view, snapshot);
}

function readActive(view: EditorView | null, snapshot: DocxSnapshot | null): ActiveTextFormat {
  if (!view) return EMPTY;

  const state = view.state;
  const marks = computeActiveMarks(state);

  const fontSizeHalf = activeRunAttr<number>(
    state,
    "font_size",
    "halfPoints",
    snapshot,
    (rPr) => rPr.fontSize
  );
  const fontFamily = activeRunAttr<string>(
    state,
    "font_family",
    "family",
    snapshot,
    (rPr) => rPr.fontFamily
  );
  const colorRaw = activeRunAttr<string>(
    state,
    "color",
    "rgb",
    snapshot,
    (rPr) => rPr.color
  );
  const highlightRaw = activeRunAttr<string>(
    state,
    "highlight",
    "name",
    snapshot,
    (rPr) => rPr.highlight
  );

  return {
    bold: marks.has("bold"),
    italic: marks.has("italic"),
    underline: marks.has("underline"),
    strike: marks.has("strikethrough"),
    fontFamily: bridge(fontFamily),
    fontSizePt: mapFontSize(bridge(fontSizeHalf)),
    color: mapColor(bridge(colorRaw)),
    highlight: mapDocxHighlight(bridge(highlightRaw)),
  };
}

function mapFontSize(value: MaybeMixed<number>): MaybeMixed<number> {
  if (value === MIXED || value === undefined) return value;
  return halfPointsToPt(value);
}

function mapColor(value: MaybeMixed<string>): MaybeMixed<string> {
  if (value === MIXED || value === undefined) return value;
  return normalizeColor(value) ?? undefined;
}

function mapDocxHighlight(value: MaybeMixed<string>): MaybeMixed<string> {
  if (value === MIXED || value === undefined) return value;
  const swatch = highlightByDocxName(value);
  return swatch ? swatch.hex : undefined;
}

async function dispatchPatch(deps: DocxProviderDeps, patch: TextFormat): Promise<void> {
  const view = deps.mountRef.current?.view ?? null;
  const agent = deps.agentRef.current;
  if (!view || !agent) return;
  if (view.state.selection.empty) {
    deps.pushToast("info", "Select some text first.");
    return;
  }
  const docxFormat = canonicalToDocx(patch);
  if (Object.keys(docxFormat).length === 0) return;
  const range = pmSelectionToRange(view.state);
  try {
    await agent.applyCommand({
      type: "docx:format-range",
      payload: { range, format: docxFormat },
      source: "human",
    });
  } catch (err) {
    if (err instanceof NotImplementedError) {
      deps.pushToast("warn", "Not yet supported in this build.");
      return;
    }
    deps.pushToast("error", err instanceof Error ? err.message : String(err));
  }
}

function canonicalToDocx(patch: TextFormat): DocxTextFormat {
  const out: DocxTextFormat = {};
  if (patch.bold !== undefined) out.bold = patch.bold;
  if (patch.italic !== undefined) out.italic = patch.italic;
  if (patch.underline !== undefined) {
    out.underline = patch.underline === false ? false : true;
  }
  if (patch.strike !== undefined) out.strike = patch.strike;
  if (patch.fontFamily !== undefined) out.fontFamily = patch.fontFamily;
  if (patch.fontSizePt !== undefined) out.fontSize = ptToHalfPoints(patch.fontSizePt);
  if (patch.color !== undefined) {
    const norm = normalizeColor(patch.color);
    out.color = norm ?? "";
  }
  if (patch.highlight !== undefined) {
    if (patch.highlight === "") {
      out.highlight = "";
    } else {
      const norm = normalizeColor(patch.highlight);
      out.highlight = norm ? nearestHighlight(norm).docxName : "";
    }
  }
  return out;
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
