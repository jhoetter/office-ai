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
} from "lucide-react";
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
            render: ({ onNumberFormat }) => (
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
              <ToggleBtn
                icon={<Filter size={14} />}
                label={props.filterActive ? "Remove AutoFilter" : "Apply AutoFilter"}
                testId="data-filter-toggle"
                active={props.filterActive}
                disabled={disabled}
                onClick={props.onToggleFilter}
              />
            ),
          },
        ],
      },
      {
        id: "view",
        label: "Ansicht",
        groups: [
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
