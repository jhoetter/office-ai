"use client";

import type { ReactNode } from "react";
import { useCallback, useMemo } from "react";
import {
  AlignLeft,
  AlignCenter,
  AlignRight,
  Merge,
  Split,
  Undo2,
  Redo2,
  TableProperties,
  MessageSquarePlus,
  Filter,
  Image as ImageIcon,
  Snowflake,
  ChevronDown,
  Square,
  SquareDashed,
  Grid3x3,
  Paintbrush,
  BarChart3,
  Sigma,
  Pencil,
  FileText,
  Frame,
  Printer,
  PrinterCheck,
  Repeat,
  Calculator,
  Eye,
  EyeOff,
  Lock,
  ZoomIn,
  Ruler,
  Tag,
  Percent,
  DollarSign,
  ArrowUpAZ,
  ArrowDownZA,
  TrendingUp,
  Heading,
  Lightbulb,
  CopyMinus,
  Group,
  Ungroup,
  Sigma as SigmaIcon,
  Target,
  Wand2,
  Filter as FilterIcon,
} from "@officeai/ui/sonaloop-icons";
import { useRef, useState } from "react";
import { TextFormatBar, cn } from "@officeai/ui";
import type { ActiveTextFormat, TextFormatProvider } from "@officeai/text-formatting";
import type { CellFormatPatch, EffectiveStyle, StyleTable } from "@officeai/xlsx";
import { flattenCellXf } from "@officeai/xlsx";
import { NUMBER_FORMAT_PRESETS, type NumberFormatPresetId } from "./styles";
import type { Selection } from "./selection";
import { Ribbon, ToolbarMenu, type RibbonCatalogue } from "../lib/shell";

export interface ToolbarProps {
  readonly disabled: boolean;
  /**
   * The active anchor cell's effective style, used to flip toggle
   * states (Bold pressed, etc.) on the alignment buttons that aren't
   * yet routed through the shared TextFormatBar.
   */
  readonly anchorStyleId: number | undefined;
  readonly styles: StyleTable;
  readonly selection: Selection | null;
  /**
   * Dispatch a `xlsx:set-cell-format` over the current selection. The
   * parent owns the agent + sheet name so this stays a pure
   * presentational component.
   */
  readonly onApply: (patch: CellFormatPatch) => void;
  /**
   * Shared text-formatting plumbing. Bold / italic / underline /
   * strike / font family / font size / font color / cell highlight
   * are all routed through this so the spreadsheet ribbon shares
   * pickers + MIXED-state UX with the document and slide editors.
   */
  readonly textFormatProvider: TextFormatProvider;
  readonly textFormatActive: ActiveTextFormat;
  /** True iff the active selection range exactly matches an existing merge. */
  readonly canUnmerge: boolean;
  /** True iff the active selection covers ≥ 2 cells (eligible for merge). */
  readonly canMerge: boolean;
  readonly onMerge: () => void;
  readonly onUnmerge: () => void;
  /**
   * P13 Undo/Redo. The bus exposes `canUndo()`/`canRedo()` so the
   * parent flips these every render. When the bus surface isn't
   * wired yet (early P13 commits) the parent passes `false` +
   * no-op callbacks.
   */
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  readonly onUndo: () => void;
  readonly onRedo: () => void;
  /** Open the Text-to-Columns popover for the active selection. */
  readonly onTextToColumns: () => void;
  /** Disabled when no single-column selection is available. */
  readonly canTextToColumns: boolean;
  /**
   * Focus the comments composer for the active cell. Disabled when
   * there is no selection (the composer needs an A1 anchor).
   */
  readonly onAddComment: () => void;
  /**
   * Toggle the AutoFilter band on the active sheet. The toolbar
   * stays presentational; the parent computes the range to use
   * (current selection or auto-detected used range) and dispatches
   * `xlsx:set-auto-filter` with `range: null` to remove.
   */
  readonly onToggleFilter: () => void;
  /** True iff the active sheet currently has an AutoFilter applied. */
  readonly filterActive: boolean;
  /**
   * Open the OS file picker for an image (PNG/JPEG/GIF). The selected
   * file is anchored at the active selection's anchor cell on the
   * active sheet via `xlsx:add-image`.
   */
  readonly onInsertImage: () => void;
  /**
   * Apply a freeze configuration to the active sheet. `rows` ≥ 0,
   * `cols` ≥ 0, and `(0, 0)` clears any existing freeze. The toolbar
   * surfaces the four standard Excel presets (top row, first column,
   * panes-at-selection, unfreeze) plus the current state, so the user
   * can tell at a glance whether the sheet is frozen and where.
   */
  readonly onFreeze: (rows: number, cols: number) => void;
  /** Currently active freeze, if any. */
  readonly freeze: { rows: number; cols: number } | undefined;
  /**
   * Anchor row/col for "Freeze panes at selection". When the user
   * dispatches this action, the freeze split lands above + left of
   * this position (matching Excel's behaviour). Pass `null` when
   * there's no selection so the menu entry is disabled.
   */
  readonly freezeAnchor: { row: number; col: number } | null;
  /**
   * C6 — Borders splitter. Dispatched preset is applied to the active
   * selection. The parent computes per-cell sub-range dispatches when
   * needed (outside / inside borders).
   */
  readonly onApplyBorderPreset: (preset: BorderPreset) => void;
  /**
   * C6 — "More borders…" entry opens the Format Cells dialog on the
   * Border tab so the user can pick a thick line, a colour, etc.
   */
  readonly onOpenMoreBorders: () => void;
  /**
   * C8 — Format Painter. Single-click activates one-shot mode;
   * double-click pins (sticky) until the user clicks again or hits
   * Esc. The button reflects the active state via aria-pressed.
   */
  readonly onActivateFormatPainter: (sticky: boolean) => void;
  readonly formatPainterActive: boolean;
  /**
   * Open the Insert Chart dialog. Disabled when no selection exists.
   */
  readonly onOpenInsertChart: () => void;
  /**
   * Open the Insert PivotTable dialog. Disabled when no selection
   * exists.
   */
  readonly onOpenInsertPivot: () => void;
  /**
   * The currently selected chart id on the active sheet, or `null`.
   * Drives the contextual "Diagrammtools" tab.
   */
  readonly selectedChartId: string | null;
  /**
   * Open the chart editor for the currently selected chart.
   */
  readonly onEditSelectedChart: () => void;

  // ── Page Layout tab ──────────────────────────────────────────────
  /** Open the Excel-style Page Setup dialog (initial tab optional). */
  readonly onOpenPageSetup: (tab?: "page" | "margins" | "sheet") => void;
  /** Apply a margins preset directly (Page Layout → Margins splitter). */
  readonly onApplyMarginsPreset: (preset: "normal" | "wide" | "narrow") => void;
  /** Apply an orientation preset directly (Page Layout → Orientation splitter). */
  readonly onApplyOrientation: (orientation: "portrait" | "landscape") => void;
  /** Apply a paper-size preset directly (Page Layout → Size splitter). */
  readonly onApplyPaperSize: (paperSize: number) => void;
  /** Set / clear / extend the worksheet print area. */
  readonly onPrintArea: (mode: "set" | "clear" | "add") => void;
  /** Toggle "Print gridlines" / "Print headings" — mirrors the Sheet Options group. */
  readonly onTogglePrintFlag: (flag: "gridLines" | "headings", value: boolean) => void;
  /** Live read-throughs of the print options, so the toggle UIs reflect truth. */
  readonly printGridLines: boolean;
  readonly printHeadings: boolean;

  // ── Formulas tab ─────────────────────────────────────────────────
  /** Open the Name Manager dialog. */
  readonly onOpenNameManager: () => void;
  /** Set Excel's calculation mode. */
  readonly onSetCalcMode: (mode: "auto" | "autoNoTable" | "manual") => void;
  /** Toggle "Calculate before saving". */
  readonly onSetCalcOnSave: (value: boolean) => void;
  /** Toggle "Show formulas" on the active sheet. */
  readonly onToggleShowFormulas: () => void;
  /** Live calc mode read-through. */
  readonly calcMode: "auto" | "autoNoTable" | "manual";
  readonly calcOnSave: boolean;
  /** Whether the active sheet currently has Show Formulas enabled. */
  readonly showFormulas: boolean;

  // ── Review tab ───────────────────────────────────────────────────
  /** Open the Protect Sheet dialog. */
  readonly onOpenProtectSheet: () => void;
  /** Open the Protect Workbook dialog. */
  readonly onOpenProtectWorkbook: () => void;
  /** Whether the active sheet currently has any protection. */
  readonly sheetProtected: boolean;
  /** Whether the workbook currently has any protection. */
  readonly workbookProtected: boolean;

  // ── 9b: Editing / quick-format / sort additions ─────────────────
  /**
   * Excel's classic AutoSum splitter. Compose `=SUM(range)` (or the
   * matching aggregate) and dispatch `xlsx:set-cell-formula` to land
   * it in the active anchor cell. The parent owns range detection
   * (Excel's "walk up while non-empty" heuristic) and falls back to
   * the current selection.
   */
  readonly onAutoSum: (kind: "sum" | "average" | "count" | "max" | "min") => void;
  /**
   * Apply one of the named number-format presets in a single click
   * (no select-from-dropdown needed) — % / $ / "comma". Composes
   * `xlsx:set-cell-format` with the preset's code. Disabled when no
   * selection exists.
   */
  readonly onQuickNumberFormat: (preset: "percent" | "currency-usd" | "comma") => void;
  /**
   * Bump the active selection's number-format up or down one
   * decimal place. The parent inspects the *current* preset and
   * either appends `.0` / `.00` … or strips one trailing `0`.
   * Mirrors Excel's "Increase Decimal" / "Decrease Decimal" buttons.
   */
  readonly onAdjustDecimals: (delta: 1 | -1) => void;
  /**
   * Sort the active range A→Z (asc) / Z→A (desc) on its first
   * column. Disabled when the selection covers fewer than two rows
   * (nothing to sort) or no selection exists.
   */
  readonly onSort: (direction: "asc" | "desc") => void;
  readonly canSort: boolean;
  /**
   * Open the "Remove Duplicates" dialog (Data ▸ Remove Duplicates).
   * The dialog inspects the active selection and lets the user pick
   * key columns before dispatching `xlsx:remove-duplicates`.
   * Disabled when no selection exists.
   */
  readonly onOpenRemoveDuplicates: () => void;
  readonly canRemoveDuplicates: boolean;
  /**
   * Hide / unhide the row(s) or column(s) covered by the active
   * selection — Excel-parity quick-actions composing
   * `xlsx:set-row-height` (height=0 hides) and the column equivalent.
   * Surfaced via the row/column header context menu in
   * `XlsxEditor.tsx`; toolbar-side we keep them invisible.
   * (Reserved for future direct-toolbar surfacing.)
   */
  readonly onHideRows?: () => void;
  readonly onUnhideRows?: () => void;
  readonly onHideColumns?: () => void;
  readonly onUnhideColumns?: () => void;

  // ── View tab (extended) ──────────────────────────────────────────
  /** Set the worksheet view mode. */
  readonly onSetSheetView: (view: "normal" | "pageBreakPreview" | "pageLayout") => void;
  /** Toggle one of the boolean flags in the View tab. */
  readonly onToggleViewFlag: (
    flag: "showGridLines" | "showRowColHeaders" | "showRuler" | "rightToLeft",
    value: boolean
  ) => void;
  /** Open the Zoom dialog. */
  readonly onOpenZoom: () => void;
  /** Live read-throughs for view toggles, so the buttons reflect truth. */
  readonly viewMode: "normal" | "pageBreakPreview" | "pageLayout";
  readonly showGridLinesView: boolean;
  readonly showRowColHeadersView: boolean;
  readonly showRulerView: boolean;
  readonly rightToLeft: boolean;
  readonly zoomScale: number;
}

/**
 * C6 — Border presets surfaced in the toolbar splitter. Mirrors
 * Excel's "Home > Borders" dropdown. The icon-side of the splitter
 * re-applies the last used preset; the chevron opens the menu.
 */
export type BorderPreset =
  | "all"
  | "outside"
  | "thick-outside"
  | "top"
  | "bottom"
  | "left"
  | "right"
  | "top-bottom"
  | "top-thick-bottom"
  | "none";

/**
 * Excel-style formatting toolbar.
 *
 * Layout (left → right):
 *   [Undo / Redo] | [shared TextFormatBar — bold/italic/.../color/highlight] |
 *   [Align L/C/R] | [Merge / Unmerge] | [Text-to-Columns] | [Number format]
 *
 * The text-formatting cluster (font family + size, bold/italic/
 * underline/strike, font color, cell highlight) is now the shared
 * `<TextFormatBar />` from `@officeai/ui`, driven by an
 * `xlsxFormatProvider`. The XLSX-specific bits below — alignment,
 * merge/unmerge, structural row/column edits, undo/redo,
 * text-to-columns, number format — stay local because they have no
 * counterpart in DOCX or PPTX.
 */
export function Toolbar(props: ToolbarProps): ReactNode {
  const { disabled, anchorStyleId, styles, selection, onApply, textFormatProvider, textFormatActive } = props;

  // C6 — Last-used border preset, so the icon-side click re-applies
  // it (Excel parity). Defaults to "all" since that's what the icon
  // visually depicts at first launch.
  const [lastBorder, setLastBorder] = useState<BorderPreset>("all");

  const effective: EffectiveStyle = useMemo(
    () => flattenCellXf(styles, anchorStyleId),
    [styles, anchorStyleId]
  );

  const apply = useCallback(
    (patch: CellFormatPatch) => {
      if (!selection) return;
      onApply(patch);
    },
    [onApply, selection]
  );

  const onAlign = useCallback(
    (h: "left" | "center" | "right") => {
      apply({ alignment: { horizontal: h } });
    },
    [apply]
  );

  const onNumberFormat = useCallback(
    (id: NumberFormatPresetId) => {
      const code = NUMBER_FORMAT_PRESETS.find((p) => p.id === id)?.code;
      if (!code) return;
      apply({ numberFormat: code });
    },
    [apply]
  );

  const ctx: XlsxRibbonCtx = {
    props,
    effective,
    lastBorder,
    setLastBorder,
    onAlign,
    onNumberFormat,
  };

  const catalogue = useMemo<RibbonCatalogue<XlsxRibbonCtx>>(
    () => buildXlsxRibbonCatalogue({ disabled, textFormatProvider, textFormatActive }),
    [disabled, textFormatProvider, textFormatActive]
  );

  return <Ribbon ariaLabel="Spreadsheet toolbar" testId="xlsx-toolbar" catalogue={catalogue} ctx={ctx} />;
}

interface XlsxRibbonCtx {
  readonly props: ToolbarProps;
  readonly effective: EffectiveStyle;
  readonly lastBorder: BorderPreset;
  readonly setLastBorder: (preset: BorderPreset) => void;
  readonly onAlign: (h: "left" | "center" | "right") => void;
  readonly onNumberFormat: (id: NumberFormatPresetId) => void;
}

interface XlsxRibbonOptions {
  readonly disabled: boolean;
  readonly textFormatProvider: TextFormatProvider;
  readonly textFormatActive: ActiveTextFormat;
}

/**
 * XLSX subset-pragmatic catalogue. Tab names mirror Excel DE:
 * Start / Einfügen / Daten / Ansicht. Diagrammtools auto-activates
 * when a chart is selected on the sheet.
 */
function buildXlsxRibbonCatalogue(opts: XlsxRibbonOptions): RibbonCatalogue<XlsxRibbonCtx> {
  const { disabled, textFormatProvider, textFormatActive } = opts;
  return {
    defaultTabId: "start",
    tabs: [
      {
        id: "start",
        label: "Start",
        groups: [
          {
            id: "undo",
            label: "Rückgängig",
            render: ({ props }) => (
              <>
                <ActionBtn
                  icon={<Undo2 size={14} />}
                  label="Undo (⌘Z)"
                  testId="action-undo"
                  disabled={!props.canUndo}
                  onClick={props.onUndo}
                />
                <ActionBtn
                  icon={<Redo2 size={14} />}
                  label="Redo (⇧⌘Z)"
                  testId="action-redo"
                  disabled={!props.canRedo}
                  onClick={props.onRedo}
                />
              </>
            ),
          },
          {
            id: "clipboard",
            label: "Zwischenablage",
            render: ({ props }) => (
              <button
                type="button"
                data-testid="action-format-painter"
                title="Format Painter (double-click for sticky)"
                aria-label="Format Painter"
                aria-pressed={props.formatPainterActive}
                disabled={disabled}
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => props.onActivateFormatPainter(false)}
                onDoubleClick={() => props.onActivateFormatPainter(true)}
                className={cn(
                  "inline-flex h-7 w-7 items-center justify-center rounded text-foreground hover:bg-hover disabled:opacity-40",
                  props.formatPainterActive && "bg-accent-soft text-accent"
                )}
              >
                <Paintbrush size={14} />
              </button>
            ),
          },
          {
            id: "font",
            label: "Schriftart",
            render: () => (
              <TextFormatBar
                provider={textFormatProvider}
                active={textFormatActive}
                disabled={disabled}
                testIdPrefix="format"
              />
            ),
          },
          {
            id: "alignment",
            label: "Ausrichtung",
            render: ({ effective, onAlign, props }) => (
              <>
                <ToggleBtn
                  icon={<AlignLeft size={14} />}
                  label="Align left"
                  testId="format-align-left"
                  active={effective.alignment?.horizontal === "left"}
                  disabled={disabled}
                  onClick={() => onAlign("left")}
                />
                <ToggleBtn
                  icon={<AlignCenter size={14} />}
                  label="Align center"
                  testId="format-align-center"
                  active={effective.alignment?.horizontal === "center"}
                  disabled={disabled}
                  onClick={() => onAlign("center")}
                />
                <ToggleBtn
                  icon={<AlignRight size={14} />}
                  label="Align right"
                  testId="format-align-right"
                  active={effective.alignment?.horizontal === "right"}
                  disabled={disabled}
                  onClick={() => onAlign("right")}
                />
                <ActionBtn
                  icon={<Merge size={14} />}
                  label="Merge cells"
                  testId="format-merge"
                  disabled={disabled || !props.canMerge}
                  onClick={props.onMerge}
                />
                <ActionBtn
                  icon={<Split size={14} />}
                  label="Unmerge cells"
                  testId="format-unmerge"
                  disabled={disabled || !props.canUnmerge}
                  onClick={props.onUnmerge}
                />
              </>
            ),
          },
          {
            id: "borders",
            label: "Rahmen",
            render: ({ lastBorder, setLastBorder, props }) => (
              <BordersMenu
                disabled={disabled}
                last={lastBorder}
                onApply={(preset) => {
                  if (preset !== "none") setLastBorder(preset);
                  props.onApplyBorderPreset(preset);
                }}
                onOpenMore={props.onOpenMoreBorders}
              />
            ),
          },
          {
            id: "number",
            label: "Zahl",
            render: ({ onNumberFormat, props }) => (
              <div className="flex items-center gap-1">
                <select
                  data-testid="format-number"
                  aria-label="Number format"
                  disabled={disabled}
                  defaultValue=""
                  onChange={(e) => {
                    const v = e.target.value;
                    if (v) onNumberFormat(v as NumberFormatPresetId);
                    e.currentTarget.value = "";
                  }}
                  className="h-7 rounded border border-divider bg-background px-1 text-xs text-foreground disabled:opacity-50"
                >
                  <option value="" disabled>
                    Format…
                  </option>
                  {NUMBER_FORMAT_PRESETS.map((p) => (
                    <option key={p.id} value={p.id}>
                      {p.label}
                    </option>
                  ))}
                </select>
                {/*
                 * Excel's Home → Number quick-action cluster. These
                 * compose `xlsx:set-cell-format` with the same number
                 * presets as the dropdown but in one click.
                 */}
                <ActionBtn
                  icon={<DollarSign size={14} />}
                  label="Currency ($)"
                  testId="format-quick-currency"
                  disabled={disabled || !props.selection}
                  onClick={() => props.onQuickNumberFormat("currency-usd")}
                />
                <ActionBtn
                  icon={<Percent size={14} />}
                  label="Percent"
                  testId="format-quick-percent"
                  disabled={disabled || !props.selection}
                  onClick={() => props.onQuickNumberFormat("percent")}
                />
                <button
                  type="button"
                  data-testid="format-quick-comma"
                  title="Comma style (1,234.56)"
                  aria-label="Comma style"
                  onMouseDown={(e) => e.preventDefault()}
                  disabled={disabled || !props.selection}
                  onClick={() => props.onQuickNumberFormat("comma")}
                  className="inline-flex h-7 min-w-7 items-center justify-center rounded px-1 text-[11px] font-semibold text-foreground hover:bg-hover disabled:opacity-40"
                >
                  ,
                </button>
                <button
                  type="button"
                  data-testid="format-decimal-up"
                  title="Increase decimal"
                  aria-label="Increase decimal"
                  onMouseDown={(e) => e.preventDefault()}
                  disabled={disabled || !props.selection}
                  onClick={() => props.onAdjustDecimals(+1)}
                  className="inline-flex h-7 min-w-7 items-center justify-center rounded px-1 text-[10px] font-mono text-foreground hover:bg-hover disabled:opacity-40"
                >
                  .0←
                </button>
                <button
                  type="button"
                  data-testid="format-decimal-down"
                  title="Decrease decimal"
                  aria-label="Decrease decimal"
                  onMouseDown={(e) => e.preventDefault()}
                  disabled={disabled || !props.selection}
                  onClick={() => props.onAdjustDecimals(-1)}
                  className="inline-flex h-7 min-w-7 items-center justify-center rounded px-1 text-[10px] font-mono text-foreground hover:bg-hover disabled:opacity-40"
                >
                  ←.0
                </button>
              </div>
            ),
          },
          {
            id: "editing",
            label: "Bearbeiten",
            render: ({ props }) => (
              <AutoSumSplit
                disabled={disabled || !props.selection}
                onPick={(kind) => props.onAutoSum(kind)}
              />
            ),
          },
        ],
      },
      {
        id: "insert",
        label: "Einfügen",
        groups: [
          {
            id: "tables",
            label: "Tabellen",
            render: ({ props }) => (
              <ActionBtn
                icon={<Sigma size={14} />}
                label="Insert PivotTable"
                testId="action-insert-pivot"
                disabled={disabled || !props.selection}
                onClick={props.onOpenInsertPivot}
              />
            ),
          },
          {
            id: "illustrations",
            label: "Illustrationen",
            render: ({ props }) => (
              <ActionBtn
                icon={<ImageIcon size={14} />}
                label="Insert image"
                testId="action-insert-image"
                disabled={disabled}
                onClick={props.onInsertImage}
              />
            ),
          },
          {
            id: "charts",
            label: "Diagramme",
            render: ({ props }) => (
              <ActionBtn
                icon={<BarChart3 size={14} />}
                label="Insert chart"
                testId="action-insert-chart"
                disabled={disabled || !props.selection}
                onClick={props.onOpenInsertChart}
              />
            ),
          },
          {
            id: "comments",
            label: "Kommentare",
            render: ({ props }) => (
              <ActionBtn
                icon={<MessageSquarePlus size={14} />}
                label="Add comment"
                testId="action-add-comment"
                disabled={disabled || !props.selection}
                onClick={props.onAddComment}
              />
            ),
          },
          {
            // Phase 9c §4d — Insert depth.
            //
            // `add-sparkline` / `remove-sparkline` need an
            // `<x14:sparklineGroups>` parser + sheet-cell SparklineRenderer
            // (Canvas overlay over the existing grid renderer);
            // `set-page-header-footer` writes `<headerFooter>` into
            // `sheet1.xml`; "Recommended Charts" is a UI heuristic over
            // the existing `xlsx:add-chart` plus a chart-shape inference
            // step. Each is small but net-new and gets its own follow-up
            // plan. We surface the planned ribbon group as
            // disabled triggers (same pattern as DOCX Design / PPTX
            // shape outline) so the gap is honest, not invisible.
            id: "depth-coming-soon",
            label: "Mehr",
            render: () => (
              <>
                <ActionBtn
                  icon={<TrendingUp size={14} />}
                  label="Sparkline — tracked for the next milestone (needs xlsx:add-sparkline)."
                  testId="xlsx-add-sparkline-coming-soon"
                  disabled
                  onClick={() => {}}
                />
                <ActionBtn
                  icon={<Heading size={14} />}
                  label="Header / footer — tracked for the next milestone (needs <headerFooter> typed model)."
                  testId="xlsx-set-page-header-footer-coming-soon"
                  disabled
                  onClick={() => {}}
                />
                <ActionBtn
                  icon={<Lightbulb size={14} />}
                  label="Recommended charts — tracked for the next milestone (needs heuristics over selection)."
                  testId="xlsx-recommended-charts-coming-soon"
                  disabled
                  onClick={() => {}}
                />
              </>
            ),
          },
        ],
      },
      {
        id: "data",
        label: "Daten",
        groups: [
          {
            id: "tools",
            label: "Datentools",
            render: ({ props }) => (
              <ActionBtn
                icon={<TableProperties size={14} />}
                label="Text to Columns"
                testId="data-text-to-columns"
                disabled={disabled || !props.canTextToColumns}
                onClick={props.onTextToColumns}
              />
            ),
          },
          {
            id: "filter",
            label: "Sortieren und Filtern",
            render: ({ props }) => (
              <>
                <ActionBtn
                  icon={<ArrowUpAZ size={14} />}
                  label="Sort A → Z"
                  testId="data-sort-asc"
                  disabled={disabled || !props.canSort}
                  onClick={() => props.onSort("asc")}
                />
                <ActionBtn
                  icon={<ArrowDownZA size={14} />}
                  label="Sort Z → A"
                  testId="data-sort-desc"
                  disabled={disabled || !props.canSort}
                  onClick={() => props.onSort("desc")}
                />
                <ToggleBtn
                  icon={<Filter size={14} />}
                  label={props.filterActive ? "Remove AutoFilter" : "Apply AutoFilter"}
                  testId="data-filter-toggle"
                  active={props.filterActive}
                  disabled={disabled}
                  onClick={props.onToggleFilter}
                />
              </>
            ),
          },
          {
            // Phase 9c §4e — Data depth.
            //
            // Remove-duplicates is fully wired (backend handler +
            // dialog). Group-rows / ungroup-rows / group-columns /
            // ungroup-columns / add-subtotal / goal-seek / flash-fill
            // / advanced-filter still need new backend handlers
            // (typed model fields for outline levels, a bisection
            // runtime for goal-seek, a string-interpolation heuristic
            // for flash-fill). Surfacing the planned ribbon group
            // with disabled triggers keeps the gap visible and lets
            // a follow-up plan flip `disabled` and wire `onClick`
            // once the backends land.
            id: "data-depth-coming-soon",
            label: "Datenwerkzeuge",
            render: ({ props }) => (
              <>
                <ActionBtn
                  icon={<CopyMinus size={14} />}
                  label="Duplikate entfernen"
                  testId="xlsx-remove-duplicates"
                  disabled={disabled || !props.canRemoveDuplicates}
                  onClick={props.onOpenRemoveDuplicates}
                />
                <ActionBtn
                  icon={<Group size={14} />}
                  label="Group rows / columns — tracked for the next milestone (needs row/col outlineLevel typed model)."
                  testId="xlsx-group-coming-soon"
                  disabled
                  onClick={() => {}}
                />
                <ActionBtn
                  icon={<Ungroup size={14} />}
                  label="Ungroup rows / columns — tracked for the next milestone (needs row/col outlineLevel typed model)."
                  testId="xlsx-ungroup-coming-soon"
                  disabled
                  onClick={() => {}}
                />
                <ActionBtn
                  icon={<SigmaIcon size={14} />}
                  label="Subtotal — tracked for the next milestone (needs SUBTOTAL formula synthesis at group breaks)."
                  testId="xlsx-subtotal-coming-soon"
                  disabled
                  onClick={() => {}}
                />
                <ActionBtn
                  icon={<Target size={14} />}
                  label="Goal seek — tracked for the next milestone (needs bisection runtime over recalc)."
                  testId="xlsx-goal-seek-coming-soon"
                  disabled
                  onClick={() => {}}
                />
                <ActionBtn
                  icon={<Wand2 size={14} />}
                  label="Flash fill — tracked for the next milestone (needs string-interpolation heuristic)."
                  testId="xlsx-flash-fill-coming-soon"
                  disabled
                  onClick={() => {}}
                />
                <ActionBtn
                  icon={<FilterIcon size={14} />}
                  label="Advanced filter — tracked for the next milestone (needs criteria-range evaluator)."
                  testId="xlsx-advanced-filter-coming-soon"
                  disabled
                  onClick={() => {}}
                />
              </>
            ),
          },
        ],
      },
      {
        id: "page-layout",
        label: "Seitenlayout",
        groups: [
          {
            id: "page-setup",
            label: "Seite einrichten",
            render: ({ props }) => (
              <>
                <MarginsMenu
                  disabled={disabled}
                  onApplyPreset={props.onApplyMarginsPreset}
                  onOpenCustom={() => props.onOpenPageSetup("margins")}
                />
                <OrientationMenu
                  disabled={disabled}
                  onApply={props.onApplyOrientation}
                  onOpenMore={() => props.onOpenPageSetup("page")}
                />
                <PaperSizeMenu
                  disabled={disabled}
                  onApply={props.onApplyPaperSize}
                  onOpenMore={() => props.onOpenPageSetup("page")}
                />
                <ActionBtn
                  icon={<FileText size={14} />}
                  label="Page Setup…"
                  testId="action-page-setup"
                  disabled={disabled}
                  onClick={() => props.onOpenPageSetup()}
                />
              </>
            ),
          },
          {
            id: "print-area",
            label: "Drucken",
            render: ({ props }) => (
              <>
                <PrintAreaMenu
                  disabled={disabled}
                  onSet={() => props.onPrintArea("set")}
                  onClear={() => props.onPrintArea("clear")}
                  onAdd={() => props.onPrintArea("add")}
                />
                <ActionBtn
                  icon={<Repeat size={14} />}
                  label="Print Titles…"
                  testId="action-print-titles"
                  disabled={disabled}
                  onClick={() => props.onOpenPageSetup("sheet")}
                />
              </>
            ),
          },
          {
            id: "sheet-options",
            label: "Blattoptionen",
            render: ({ props }) => (
              <>
                <ToggleBtn
                  icon={<Printer size={14} />}
                  label="Print gridlines"
                  testId="action-print-gridlines"
                  active={props.printGridLines}
                  disabled={disabled}
                  onClick={() => props.onTogglePrintFlag("gridLines", !props.printGridLines)}
                />
                <ToggleBtn
                  icon={<PrinterCheck size={14} />}
                  label="Print headings"
                  testId="action-print-headings"
                  active={props.printHeadings}
                  disabled={disabled}
                  onClick={() => props.onTogglePrintFlag("headings", !props.printHeadings)}
                />
              </>
            ),
          },
        ],
      },
      {
        id: "formulas",
        label: "Formeln",
        groups: [
          {
            id: "named-cells",
            label: "Definierte Namen",
            render: ({ props }) => (
              <ActionBtn
                icon={<Tag size={14} />}
                label="Name Manager"
                testId="action-name-manager"
                disabled={disabled}
                onClick={props.onOpenNameManager}
              />
            ),
          },
          {
            id: "formula-auditing",
            label: "Formelüberwachung",
            render: ({ props }) => (
              <ToggleBtn
                icon={props.showFormulas ? <EyeOff size={14} /> : <Eye size={14} />}
                label={props.showFormulas ? "Hide formulas" : "Show formulas"}
                testId="action-show-formulas"
                active={props.showFormulas}
                disabled={disabled}
                onClick={props.onToggleShowFormulas}
              />
            ),
          },
          {
            id: "calculation",
            label: "Berechnung",
            render: ({ props }) => (
              <CalcModeMenu
                disabled={disabled}
                mode={props.calcMode}
                calcOnSave={props.calcOnSave}
                onSetMode={props.onSetCalcMode}
                onSetCalcOnSave={props.onSetCalcOnSave}
              />
            ),
          },
        ],
      },
      {
        id: "review",
        label: "Überprüfen",
        groups: [
          {
            id: "comments-review",
            label: "Kommentare",
            render: ({ props }) => (
              <ActionBtn
                icon={<MessageSquarePlus size={14} />}
                label="Add comment"
                testId="action-add-comment-review"
                disabled={disabled || !props.selection}
                onClick={props.onAddComment}
              />
            ),
          },
          {
            id: "protect",
            label: "Schützen",
            render: ({ props }) => (
              <>
                {/* Protect Sheet/Workbook open a confirmation dialog rather
                 * than acting as instant toggles, so we render them as
                 * `ActionBtn` (no `aria-pressed`, no toggled background).
                 * The Lock icon flips to "Unprotect…" wording when the
                 * resource is already locked so users can tell at a glance
                 * what the click will do — the dialog handles confirm. */}
                <ActionBtn
                  icon={
                    props.sheetProtected ? (
                      <span className="relative inline-flex">
                        <Lock size={14} />
                        <span className="absolute -bottom-0.5 -right-0.5 h-1.5 w-1.5 rounded-full bg-accent" />
                      </span>
                    ) : (
                      <Lock size={14} />
                    )
                  }
                  label={props.sheetProtected ? "Unprotect Sheet…" : "Protect Sheet…"}
                  testId="action-protect-sheet"
                  disabled={disabled}
                  onClick={props.onOpenProtectSheet}
                />
                <ActionBtn
                  icon={
                    props.workbookProtected ? (
                      <span className="relative inline-flex">
                        <Lock size={14} />
                        <span className="absolute -bottom-0.5 -right-0.5 h-1.5 w-1.5 rounded-full bg-accent" />
                      </span>
                    ) : (
                      <Lock size={14} />
                    )
                  }
                  label={props.workbookProtected ? "Unprotect Workbook…" : "Protect Workbook…"}
                  testId="action-protect-workbook"
                  disabled={disabled}
                  onClick={props.onOpenProtectWorkbook}
                />
              </>
            ),
          },
        ],
      },
      {
        id: "view",
        label: "Ansicht",
        groups: [
          {
            id: "view-modes",
            label: "Arbeitsmappenansichten",
            render: ({ props }) => (
              <ViewModeMenu disabled={disabled} mode={props.viewMode} onSet={props.onSetSheetView} />
            ),
          },
          {
            id: "view-show",
            label: "Anzeigen",
            render: ({ props }) => (
              <>
                <ToggleBtn
                  icon={<Ruler size={14} />}
                  label={props.showRulerView ? "Hide ruler" : "Show ruler"}
                  testId="action-view-ruler"
                  active={props.showRulerView}
                  disabled={disabled}
                  onClick={() => props.onToggleViewFlag("showRuler", !props.showRulerView)}
                />
                <ToggleBtn
                  icon={<Grid3x3 size={14} />}
                  label={props.showGridLinesView ? "Hide gridlines" : "Show gridlines"}
                  testId="action-view-gridlines"
                  active={props.showGridLinesView}
                  disabled={disabled}
                  onClick={() => props.onToggleViewFlag("showGridLines", !props.showGridLinesView)}
                />
                <ToggleBtn
                  icon={<TableProperties size={14} />}
                  label={props.showRowColHeadersView ? "Hide headings" : "Show headings"}
                  testId="action-view-headings"
                  active={props.showRowColHeadersView}
                  disabled={disabled}
                  onClick={() => props.onToggleViewFlag("showRowColHeaders", !props.showRowColHeadersView)}
                />
              </>
            ),
          },
          {
            id: "view-zoom",
            label: "Zoom",
            render: ({ props }) => (
              <button
                type="button"
                data-testid="action-zoom"
                title={`Zoom (${props.zoomScale}%)`}
                aria-label="Zoom"
                disabled={disabled}
                onMouseDown={(e) => e.preventDefault()}
                onClick={props.onOpenZoom}
                className="inline-flex h-7 items-center gap-1 rounded px-2 text-xs text-foreground hover:bg-hover disabled:opacity-50"
              >
                <ZoomIn size={14} />
                <span className="tabular-nums">{props.zoomScale}%</span>
              </button>
            ),
          },
          {
            id: "window",
            label: "Fenster",
            render: ({ props }) => (
              <FreezeMenu
                disabled={disabled}
                freeze={props.freeze}
                anchor={props.freezeAnchor}
                onFreeze={props.onFreeze}
              />
            ),
          },
        ],
      },
      {
        id: "chart-tools",
        label: "Diagrammtools",
        contextual: { accent: "chart" },
        visible: ({ props }) => props.selectedChartId !== null,
        autoActivateWhen: ({ props }) => props.selectedChartId !== null,
        groups: [
          {
            id: "chart-actions",
            label: "Diagramm",
            render: ({ props }) => (
              <ActionBtn
                icon={<Pencil size={14} />}
                label="Edit chart"
                testId="action-edit-chart"
                disabled={disabled || !props.selectedChartId}
                onClick={props.onEditSelectedChart}
              />
            ),
          },
        ],
      },
    ],
  };
}

function Divider(): ReactNode {
  return <div className="mx-1 h-5 w-px bg-divider" aria-hidden />;
}

interface ToggleBtnProps {
  readonly icon: ReactNode;
  readonly label: string;
  readonly testId: string;
  readonly active: boolean;
  readonly disabled: boolean;
  readonly onClick: () => void;
}

function ToggleBtn(props: ToggleBtnProps): ReactNode {
  const { icon, label, testId, active, disabled, onClick } = props;
  return (
    <button
      type="button"
      data-testid={testId}
      title={label}
      aria-label={label}
      aria-pressed={active}
      disabled={disabled}
      onMouseDown={(e) => {
        // Don't steal focus from the surface — toolbar clicks should
        // leave the active cell selection (and therefore the next
        // type-to-edit target) intact.
        e.preventDefault();
      }}
      onClick={onClick}
      className={cn(
        "inline-flex h-7 w-7 items-center justify-center rounded text-foreground hover:bg-hover disabled:opacity-50",
        active && "bg-[var(--ai-violet-light)] text-[var(--ai-violet)]"
      )}
    >
      {icon}
    </button>
  );
}

interface ActionBtnProps {
  readonly icon: ReactNode;
  readonly label: string;
  readonly testId: string;
  readonly disabled: boolean;
  readonly onClick: () => void;
}

function ActionBtn(props: ActionBtnProps): ReactNode {
  const { icon, label, testId, disabled, onClick } = props;
  return (
    <button
      type="button"
      data-testid={testId}
      title={label}
      aria-label={label}
      disabled={disabled}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
      className={cn(
        "inline-flex h-7 w-7 items-center justify-center rounded text-foreground hover:bg-hover disabled:opacity-40"
      )}
    >
      {icon}
    </button>
  );
}

interface FreezeMenuProps {
  readonly disabled: boolean;
  readonly freeze: { rows: number; cols: number } | undefined;
  readonly anchor: { row: number; col: number } | null;
  readonly onFreeze: (rows: number, cols: number) => void;
}

/**
 * C3 — Freeze panes splitter button. Mirrors Excel's "View > Freeze
 * Panes" dropdown:
 *   - Freeze top row             → rows=1, cols=0
 *   - Freeze first column        → rows=0, cols=1
 *   - Freeze panes (at selection) → rows=anchor.row, cols=anchor.col
 *   - Unfreeze panes             → rows=0, cols=0 (only when frozen)
 *
 * The button's pressed state mirrors whether the active sheet has any
 * freeze configured — important for users who jump between sheets and
 * want a glanceable indicator of "is this one frozen?".
 */
function FreezeMenu(props: FreezeMenuProps): ReactNode {
  const { disabled, freeze, anchor, onFreeze } = props;
  const [open, setOpen] = useState(false);
  const triggerRef = useRef<HTMLButtonElement | null>(null);

  const isFrozen = !!freeze && (freeze.rows > 0 || freeze.cols > 0);
  // "Freeze panes at selection" only makes sense when the anchor is
  // away from row 0 / column 0 — otherwise it's identical to one of
  // the simpler presets and Excel hides it.
  const canFreezeAtSelection = !!anchor && (anchor.row > 0 || anchor.col > 0);

  return (
    <>
      <button
        ref={triggerRef}
        type="button"
        data-testid="action-freeze-toggle"
        title="Freeze panes"
        aria-label="Freeze panes"
        aria-expanded={open}
        aria-haspopup="menu"
        disabled={disabled}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => setOpen((v) => !v)}
        className={cn(
          "inline-flex h-7 items-center gap-0.5 rounded px-1 text-foreground hover:bg-hover disabled:opacity-50",
          isFrozen && "bg-[var(--ai-violet-light)] text-[var(--ai-violet)]"
        )}
      >
        <Snowflake size={14} />
        <ChevronDown size={10} />
      </button>
      <ToolbarMenu
        open={open}
        onClose={() => setOpen(false)}
        triggerRef={triggerRef}
        role="menu"
        testId="freeze-menu"
        className="min-w-[220px] rounded-md border border-divider bg-surface p-1 shadow-lg"
      >
        <FreezeMenuItem
          label="Freeze top row"
          shortcut=""
          checked={isFrozen && freeze!.rows === 1 && freeze!.cols === 0}
          onClick={() => {
            setOpen(false);
            onFreeze(1, 0);
          }}
        />
        <FreezeMenuItem
          label="Freeze first column"
          shortcut=""
          checked={isFrozen && freeze!.rows === 0 && freeze!.cols === 1}
          onClick={() => {
            setOpen(false);
            onFreeze(0, 1);
          }}
        />
        <FreezeMenuItem
          label="Freeze panes at selection"
          shortcut={canFreezeAtSelection ? `${anchor!.row}r×${anchor!.col}c` : ""}
          disabled={!canFreezeAtSelection}
          checked={!!anchor && isFrozen && freeze!.rows === anchor.row && freeze!.cols === anchor.col}
          onClick={() => {
            if (!anchor) return;
            setOpen(false);
            onFreeze(anchor.row, anchor.col);
          }}
        />
        <div className="my-1 h-px bg-divider" />
        <FreezeMenuItem
          label="Unfreeze panes"
          shortcut=""
          disabled={!isFrozen}
          onClick={() => {
            setOpen(false);
            onFreeze(0, 0);
          }}
        />
      </ToolbarMenu>
    </>
  );
}

interface FreezeMenuItemProps {
  readonly label: string;
  readonly shortcut: string;
  readonly checked?: boolean;
  readonly disabled?: boolean;
  readonly onClick: () => void;
}

function FreezeMenuItem(props: FreezeMenuItemProps): ReactNode {
  const { label, shortcut, checked, disabled, onClick } = props;
  return (
    <button
      type="button"
      role="menuitem"
      disabled={disabled}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
      className={cn(
        "flex w-full items-center justify-between gap-3 rounded px-2 py-1.5 text-left text-xs text-foreground hover:bg-hover disabled:opacity-40",
        checked && "bg-[var(--ai-violet-light)]"
      )}
    >
      <span>{label}</span>
      {shortcut ? <span className="text-tertiary tabular-nums">{shortcut}</span> : null}
    </button>
  );
}

interface BordersMenuProps {
  readonly disabled: boolean;
  /** Last-used preset (used as the splitter's primary action). */
  readonly last: BorderPreset;
  readonly onApply: (preset: BorderPreset) => void;
  readonly onOpenMore: () => void;
}

/**
 * C6 — Borders splitter button.
 *
 * The icon side re-applies the last preset (Excel parity); the
 * chevron side opens a menu of presets that mirror Excel's "Home >
 * Borders" dropdown:
 *
 *   - All borders
 *   - Outside borders
 *   - Thick outside borders
 *   - Top / Bottom / Left / Right border
 *   - Top and bottom border
 *   - Top and thick bottom border
 *   - No border
 *   - More borders…  → opens Format Cells dialog on the Border tab
 *
 * The icon swaps to a faint outline when the last preset is "none"
 * so the user can tell at a glance that re-clicking will *clear*
 * borders rather than draw them.
 */
function BordersMenu(props: BordersMenuProps): ReactNode {
  const { disabled, last, onApply, onOpenMore } = props;
  const [open, setOpen] = useState(false);
  const triggerRef = useRef<HTMLButtonElement | null>(null);

  const Icon = last === "none" ? SquareDashed : last === "all" ? Grid3x3 : Square;

  return (
    <span className="inline-flex items-center">
      <button
        type="button"
        data-testid="format-borders-apply"
        title={`Apply ${BORDER_PRESET_LABEL[last]}`}
        aria-label={`Apply ${BORDER_PRESET_LABEL[last]}`}
        disabled={disabled}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => onApply(last)}
        className="inline-flex h-7 w-7 items-center justify-center rounded text-foreground hover:bg-hover disabled:opacity-50"
      >
        <Icon size={14} />
      </button>
      <button
        ref={triggerRef}
        type="button"
        data-testid="format-borders-toggle"
        title="More borders"
        aria-label="More borders"
        aria-expanded={open}
        aria-haspopup="menu"
        disabled={disabled}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => setOpen((v) => !v)}
        className="inline-flex h-7 w-4 items-center justify-center rounded text-foreground hover:bg-hover disabled:opacity-50"
      >
        <ChevronDown size={10} />
      </button>
      <ToolbarMenu
        open={open}
        onClose={() => setOpen(false)}
        triggerRef={triggerRef}
        role="menu"
        testId="borders-menu"
        className="min-w-[220px] rounded-md border border-divider bg-surface p-1 shadow-lg"
      >
        <BordersMenuItem
          preset="all"
          onPick={(p) => {
            setOpen(false);
            onApply(p);
          }}
        />
        <BordersMenuItem
          preset="outside"
          onPick={(p) => {
            setOpen(false);
            onApply(p);
          }}
        />
        <BordersMenuItem
          preset="thick-outside"
          onPick={(p) => {
            setOpen(false);
            onApply(p);
          }}
        />
        <div className="my-1 h-px bg-divider" />
        <BordersMenuItem
          preset="top"
          onPick={(p) => {
            setOpen(false);
            onApply(p);
          }}
        />
        <BordersMenuItem
          preset="bottom"
          onPick={(p) => {
            setOpen(false);
            onApply(p);
          }}
        />
        <BordersMenuItem
          preset="left"
          onPick={(p) => {
            setOpen(false);
            onApply(p);
          }}
        />
        <BordersMenuItem
          preset="right"
          onPick={(p) => {
            setOpen(false);
            onApply(p);
          }}
        />
        <div className="my-1 h-px bg-divider" />
        <BordersMenuItem
          preset="top-bottom"
          onPick={(p) => {
            setOpen(false);
            onApply(p);
          }}
        />
        <BordersMenuItem
          preset="top-thick-bottom"
          onPick={(p) => {
            setOpen(false);
            onApply(p);
          }}
        />
        <div className="my-1 h-px bg-divider" />
        <BordersMenuItem
          preset="none"
          onPick={(p) => {
            setOpen(false);
            onApply(p);
          }}
        />
        <div className="my-1 h-px bg-divider" />
        <button
          type="button"
          role="menuitem"
          data-testid="borders-menu-more"
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => {
            setOpen(false);
            onOpenMore();
          }}
          className="flex w-full items-center justify-between gap-3 rounded px-2 py-1.5 text-left text-xs text-foreground hover:bg-hover"
        >
          <span>More borders…</span>
        </button>
      </ToolbarMenu>
    </span>
  );
}

const BORDER_PRESET_LABEL: Readonly<Record<BorderPreset, string>> = {
  all: "All borders",
  outside: "Outside borders",
  "thick-outside": "Thick outside borders",
  top: "Top border",
  bottom: "Bottom border",
  left: "Left border",
  right: "Right border",
  "top-bottom": "Top and bottom border",
  "top-thick-bottom": "Top and thick bottom border",
  none: "No border",
};

interface BordersMenuItemProps {
  readonly preset: BorderPreset;
  readonly onPick: (preset: BorderPreset) => void;
}

function BordersMenuItem(props: BordersMenuItemProps): ReactNode {
  const { preset, onPick } = props;
  return (
    <button
      type="button"
      role="menuitem"
      data-testid={`borders-menu-${preset}`}
      onMouseDown={(e) => e.preventDefault()}
      onClick={() => onPick(preset)}
      className="flex w-full items-center justify-between gap-3 rounded px-2 py-1.5 text-left text-xs text-foreground hover:bg-hover"
    >
      <span>{BORDER_PRESET_LABEL[preset]}</span>
    </button>
  );
}

interface SimpleMenuItemProps {
  readonly label: string;
  readonly checked?: boolean;
  readonly disabled?: boolean;
  readonly testId?: string;
  readonly onClick: () => void;
}

function SimpleMenuItem(props: SimpleMenuItemProps): ReactNode {
  const { label, checked, disabled, testId, onClick } = props;
  return (
    <button
      type="button"
      role="menuitem"
      disabled={disabled}
      data-testid={testId}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
      className={cn(
        "flex w-full items-center gap-2 rounded px-2 py-1.5 text-left text-xs text-foreground hover:bg-hover disabled:opacity-40",
        checked && "bg-[var(--ai-violet-light)]"
      )}
    >
      <span className="inline-block w-3 text-center">{checked ? "✓" : ""}</span>
      <span className="flex-1">{label}</span>
    </button>
  );
}

interface MarginsMenuProps {
  readonly disabled: boolean;
  readonly onApplyPreset: (preset: "normal" | "wide" | "narrow") => void;
  readonly onOpenCustom: () => void;
}

function MarginsMenu(props: MarginsMenuProps): ReactNode {
  const { disabled, onApplyPreset, onOpenCustom } = props;
  const [open, setOpen] = useState(false);
  const triggerRef = useRef<HTMLButtonElement | null>(null);
  return (
    <span className="inline-flex items-center">
      <button
        ref={triggerRef}
        type="button"
        data-testid="action-margins"
        title="Margins"
        aria-label="Margins"
        aria-expanded={open}
        aria-haspopup="menu"
        disabled={disabled}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => setOpen((v) => !v)}
        className="inline-flex h-7 items-center gap-0.5 rounded px-1 text-foreground hover:bg-hover disabled:opacity-50"
      >
        <Frame size={14} />
        <ChevronDown size={10} />
      </button>
      <ToolbarMenu
        open={open}
        onClose={() => setOpen(false)}
        triggerRef={triggerRef}
        role="menu"
        testId="margins-menu"
        className="min-w-[220px] rounded-md border border-divider bg-surface p-1 shadow-lg"
      >
        <SimpleMenuItem
          label="Normal"
          testId="margins-menu-normal"
          onClick={() => {
            setOpen(false);
            onApplyPreset("normal");
          }}
        />
        <SimpleMenuItem
          label="Wide"
          testId="margins-menu-wide"
          onClick={() => {
            setOpen(false);
            onApplyPreset("wide");
          }}
        />
        <SimpleMenuItem
          label="Narrow"
          testId="margins-menu-narrow"
          onClick={() => {
            setOpen(false);
            onApplyPreset("narrow");
          }}
        />
        <div className="my-1 h-px bg-divider" />
        <SimpleMenuItem
          label="Custom margins…"
          testId="margins-menu-custom"
          onClick={() => {
            setOpen(false);
            onOpenCustom();
          }}
        />
      </ToolbarMenu>
    </span>
  );
}

interface OrientationMenuProps {
  readonly disabled: boolean;
  readonly onApply: (orientation: "portrait" | "landscape") => void;
  readonly onOpenMore: () => void;
}

function OrientationMenu(props: OrientationMenuProps): ReactNode {
  const { disabled, onApply, onOpenMore } = props;
  const [open, setOpen] = useState(false);
  const triggerRef = useRef<HTMLButtonElement | null>(null);
  return (
    <span className="inline-flex items-center">
      <button
        ref={triggerRef}
        type="button"
        data-testid="action-orientation"
        title="Orientation"
        aria-label="Orientation"
        aria-expanded={open}
        aria-haspopup="menu"
        disabled={disabled}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => setOpen((v) => !v)}
        className="inline-flex h-7 items-center gap-0.5 rounded px-1 text-foreground hover:bg-hover disabled:opacity-50"
      >
        <FileText size={14} />
        <ChevronDown size={10} />
      </button>
      <ToolbarMenu
        open={open}
        onClose={() => setOpen(false)}
        triggerRef={triggerRef}
        role="menu"
        testId="orientation-menu"
        className="min-w-[180px] rounded-md border border-divider bg-surface p-1 shadow-lg"
      >
        <SimpleMenuItem
          label="Portrait"
          testId="orientation-menu-portrait"
          onClick={() => {
            setOpen(false);
            onApply("portrait");
          }}
        />
        <SimpleMenuItem
          label="Landscape"
          testId="orientation-menu-landscape"
          onClick={() => {
            setOpen(false);
            onApply("landscape");
          }}
        />
        <div className="my-1 h-px bg-divider" />
        <SimpleMenuItem
          label="More…"
          testId="orientation-menu-more"
          onClick={() => {
            setOpen(false);
            onOpenMore();
          }}
        />
      </ToolbarMenu>
    </span>
  );
}

interface PaperSizeMenuProps {
  readonly disabled: boolean;
  readonly onApply: (paperSize: number) => void;
  readonly onOpenMore: () => void;
}

const PAPER_QUICK = [
  { id: 1, label: 'Letter (8.5" × 11")' },
  { id: 5, label: 'Legal (8.5" × 14")' },
  { id: 9, label: "A4 (21 × 29.7 cm)" },
  { id: 8, label: "A3 (29.7 × 42 cm)" },
];

function PaperSizeMenu(props: PaperSizeMenuProps): ReactNode {
  const { disabled, onApply, onOpenMore } = props;
  const [open, setOpen] = useState(false);
  const triggerRef = useRef<HTMLButtonElement | null>(null);
  return (
    <span className="inline-flex items-center">
      <button
        ref={triggerRef}
        type="button"
        data-testid="action-paper-size"
        title="Paper size"
        aria-label="Paper size"
        aria-expanded={open}
        aria-haspopup="menu"
        disabled={disabled}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => setOpen((v) => !v)}
        className="inline-flex h-7 items-center gap-0.5 rounded px-1 text-foreground hover:bg-hover disabled:opacity-50"
      >
        <Square size={14} />
        <ChevronDown size={10} />
      </button>
      <ToolbarMenu
        open={open}
        onClose={() => setOpen(false)}
        triggerRef={triggerRef}
        role="menu"
        testId="paper-size-menu"
        className="min-w-[260px] rounded-md border border-divider bg-surface p-1 shadow-lg"
      >
        {PAPER_QUICK.map((p) => (
          <SimpleMenuItem
            key={p.id}
            label={p.label}
            testId={`paper-size-menu-${p.id}`}
            onClick={() => {
              setOpen(false);
              onApply(p.id);
            }}
          />
        ))}
        <div className="my-1 h-px bg-divider" />
        <SimpleMenuItem
          label="More paper sizes…"
          testId="paper-size-menu-more"
          onClick={() => {
            setOpen(false);
            onOpenMore();
          }}
        />
      </ToolbarMenu>
    </span>
  );
}

interface PrintAreaMenuProps {
  readonly disabled: boolean;
  readonly onSet: () => void;
  readonly onClear: () => void;
  readonly onAdd: () => void;
}

function PrintAreaMenu(props: PrintAreaMenuProps): ReactNode {
  const { disabled, onSet, onClear, onAdd } = props;
  const [open, setOpen] = useState(false);
  const triggerRef = useRef<HTMLButtonElement | null>(null);
  return (
    <span className="inline-flex items-center">
      <button
        ref={triggerRef}
        type="button"
        data-testid="action-print-area"
        title="Print Area"
        aria-label="Print Area"
        aria-expanded={open}
        aria-haspopup="menu"
        disabled={disabled}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => setOpen((v) => !v)}
        className="inline-flex h-7 items-center gap-0.5 rounded px-1 text-foreground hover:bg-hover disabled:opacity-50"
      >
        <PrinterCheck size={14} />
        <ChevronDown size={10} />
      </button>
      <ToolbarMenu
        open={open}
        onClose={() => setOpen(false)}
        triggerRef={triggerRef}
        role="menu"
        testId="print-area-menu"
        className="min-w-[220px] rounded-md border border-divider bg-surface p-1 shadow-lg"
      >
        <SimpleMenuItem
          label="Set print area"
          testId="print-area-menu-set"
          onClick={() => {
            setOpen(false);
            onSet();
          }}
        />
        <SimpleMenuItem
          label="Clear print area"
          testId="print-area-menu-clear"
          onClick={() => {
            setOpen(false);
            onClear();
          }}
        />
        <SimpleMenuItem
          label="Add to print area"
          testId="print-area-menu-add"
          onClick={() => {
            setOpen(false);
            onAdd();
          }}
        />
      </ToolbarMenu>
    </span>
  );
}

interface CalcModeMenuProps {
  readonly disabled: boolean;
  readonly mode: "auto" | "autoNoTable" | "manual";
  readonly calcOnSave: boolean;
  readonly onSetMode: (mode: "auto" | "autoNoTable" | "manual") => void;
  readonly onSetCalcOnSave: (value: boolean) => void;
}

function CalcModeMenu(props: CalcModeMenuProps): ReactNode {
  const { disabled, mode, calcOnSave, onSetMode, onSetCalcOnSave } = props;
  const [open, setOpen] = useState(false);
  const triggerRef = useRef<HTMLButtonElement | null>(null);
  return (
    <span className="inline-flex items-center">
      <button
        ref={triggerRef}
        type="button"
        data-testid="action-calc-mode"
        title="Calculation Options"
        aria-label="Calculation Options"
        aria-expanded={open}
        aria-haspopup="menu"
        disabled={disabled}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => setOpen((v) => !v)}
        className="inline-flex h-7 items-center gap-0.5 rounded px-1 text-foreground hover:bg-hover disabled:opacity-50"
      >
        <Calculator size={14} />
        <ChevronDown size={10} />
      </button>
      <ToolbarMenu
        open={open}
        onClose={() => setOpen(false)}
        triggerRef={triggerRef}
        role="menu"
        testId="calc-mode-menu"
        className="min-w-[260px] rounded-md border border-divider bg-surface p-1 shadow-lg"
      >
        <SimpleMenuItem
          label="Automatic"
          testId="calc-mode-menu-auto"
          checked={mode === "auto"}
          onClick={() => {
            setOpen(false);
            onSetMode("auto");
          }}
        />
        <SimpleMenuItem
          label="Automatic except for data tables"
          testId="calc-mode-menu-auto-no-table"
          checked={mode === "autoNoTable"}
          onClick={() => {
            setOpen(false);
            onSetMode("autoNoTable");
          }}
        />
        <SimpleMenuItem
          label="Manual"
          testId="calc-mode-menu-manual"
          checked={mode === "manual"}
          onClick={() => {
            setOpen(false);
            onSetMode("manual");
          }}
        />
        <div className="my-1 h-px bg-divider" />
        <SimpleMenuItem
          label="Recalculate before saving"
          testId="calc-mode-menu-on-save"
          checked={calcOnSave}
          onClick={() => {
            setOpen(false);
            onSetCalcOnSave(!calcOnSave);
          }}
        />
      </ToolbarMenu>
    </span>
  );
}

interface ViewModeMenuProps {
  readonly disabled: boolean;
  readonly mode: "normal" | "pageBreakPreview" | "pageLayout";
  readonly onSet: (mode: "normal" | "pageBreakPreview" | "pageLayout") => void;
}

const VIEW_MODE_LABEL: Readonly<Record<"normal" | "pageBreakPreview" | "pageLayout", string>> = {
  normal: "Normal",
  pageBreakPreview: "Page Break Preview",
  pageLayout: "Page Layout",
};

function ViewModeMenu(props: ViewModeMenuProps): ReactNode {
  const { disabled, mode, onSet } = props;
  const [open, setOpen] = useState(false);
  const triggerRef = useRef<HTMLButtonElement | null>(null);
  return (
    <span className="inline-flex items-center">
      <button
        ref={triggerRef}
        type="button"
        data-testid="action-view-mode"
        title={`View: ${VIEW_MODE_LABEL[mode]}`}
        aria-label="View mode"
        aria-expanded={open}
        aria-haspopup="menu"
        disabled={disabled}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => setOpen((v) => !v)}
        className="inline-flex h-7 items-center gap-1 rounded px-2 text-xs text-foreground hover:bg-hover disabled:opacity-50"
      >
        <Eye size={14} />
        <span>{VIEW_MODE_LABEL[mode]}</span>
        <ChevronDown size={10} />
      </button>
      <ToolbarMenu
        open={open}
        onClose={() => setOpen(false)}
        triggerRef={triggerRef}
        role="menu"
        testId="view-mode-menu"
        className="min-w-[220px] rounded-md border border-divider bg-surface p-1 shadow-lg"
      >
        <SimpleMenuItem
          label="Normal"
          testId="view-mode-menu-normal"
          checked={mode === "normal"}
          onClick={() => {
            setOpen(false);
            onSet("normal");
          }}
        />
        <SimpleMenuItem
          label="Page Break Preview"
          testId="view-mode-menu-page-break-preview"
          checked={mode === "pageBreakPreview"}
          onClick={() => {
            setOpen(false);
            onSet("pageBreakPreview");
          }}
        />
        <SimpleMenuItem
          label="Page Layout"
          testId="view-mode-menu-page-layout"
          checked={mode === "pageLayout"}
          onClick={() => {
            setOpen(false);
            onSet("pageLayout");
          }}
        />
      </ToolbarMenu>
    </span>
  );
}

/**
 * Excel-parity AutoSum splitter. Single click defaults to SUM (the
 * green Σ button on the ribbon); the chevron drops down the rest of
 * the aggregate menu (Average, Count Numbers, Max, Min). Each entry
 * dispatches `xlsx:set-cell-formula` against the active anchor; the
 * parent decides which range to feed into the function via the
 * standard "walk up while non-empty" Excel heuristic.
 */
function AutoSumSplit(props: {
  readonly disabled: boolean;
  readonly onPick: (kind: "sum" | "average" | "count" | "max" | "min") => void;
}): ReactNode {
  const [open, setOpen] = useState(false);
  const triggerRef = useRef<HTMLButtonElement | null>(null);
  return (
    <span className="inline-flex">
      <button
        type="button"
        data-testid="action-autosum"
        title="AutoSum (Σ)"
        aria-label="AutoSum"
        disabled={props.disabled}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => props.onPick("sum")}
        className="inline-flex h-7 items-center gap-1 rounded-l border border-divider px-1.5 text-foreground hover:bg-hover disabled:opacity-40"
      >
        <Sigma size={14} />
        <span className="text-[11px]">Sum</span>
      </button>
      <button
        ref={triggerRef}
        type="button"
        data-testid="action-autosum-chevron"
        title="More AutoSum options"
        aria-label="More AutoSum options"
        aria-haspopup="menu"
        aria-expanded={open}
        disabled={props.disabled}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => setOpen((v) => !v)}
        className="inline-flex h-7 items-center justify-center rounded-r border border-l-0 border-divider px-1 text-foreground hover:bg-hover disabled:opacity-40"
      >
        <ChevronDown size={10} />
      </button>
      <ToolbarMenu
        open={open}
        onClose={() => setOpen(false)}
        triggerRef={triggerRef}
        role="menu"
        testId="autosum-menu"
        className="min-w-[180px] rounded-md border border-divider bg-surface p-1 shadow-lg"
      >
        <SimpleMenuItem
          label="Sum"
          testId="autosum-menu-sum"
          onClick={() => {
            setOpen(false);
            props.onPick("sum");
          }}
        />
        <SimpleMenuItem
          label="Average"
          testId="autosum-menu-average"
          onClick={() => {
            setOpen(false);
            props.onPick("average");
          }}
        />
        <SimpleMenuItem
          label="Count Numbers"
          testId="autosum-menu-count"
          onClick={() => {
            setOpen(false);
            props.onPick("count");
          }}
        />
        <SimpleMenuItem
          label="Max"
          testId="autosum-menu-max"
          onClick={() => {
            setOpen(false);
            props.onPick("max");
          }}
        />
        <SimpleMenuItem
          label="Min"
          testId="autosum-menu-min"
          onClick={() => {
            setOpen(false);
            props.onPick("min");
          }}
        />
      </ToolbarMenu>
    </span>
  );
}
