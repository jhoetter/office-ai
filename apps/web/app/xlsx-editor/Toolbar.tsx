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
  Keyboard,
  MessageSquarePlus,
  Filter,
  Image as ImageIcon,
} from "lucide-react";
import { TextFormatBar, cn } from "@officeai/ui";
import type { ActiveTextFormat, TextFormatProvider } from "@officeai/text-formatting";
import type { CellFormatPatch, EffectiveStyle, StyleTable } from "@officeai/xlsx";
import { flattenCellXf } from "@officeai/xlsx";
import { NUMBER_FORMAT_PRESETS, type NumberFormatPresetId } from "./styles";
import type { Selection } from "./selection";

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
  /** Open the keyboard-shortcuts help dialog. */
  readonly onOpenShortcuts: () => void;
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
}

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
  const {
    disabled,
    anchorStyleId,
    styles,
    selection,
    onApply,
    textFormatProvider,
    textFormatActive,
    canMerge,
    canUnmerge,
    onMerge,
    onUnmerge,
    canUndo,
    canRedo,
    onUndo,
    onRedo,
    onTextToColumns,
    canTextToColumns,
    onOpenShortcuts,
    onAddComment,
    onToggleFilter,
    filterActive,
    onInsertImage,
  } = props;

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

  return (
    <div
      data-testid="xlsx-toolbar"
      className="flex flex-wrap items-center gap-1 rounded-md border border-divider bg-surface px-2 py-1.5"
    >
      <ActionBtn
        icon={<Undo2 size={14} />}
        label="Undo (⌘Z)"
        testId="action-undo"
        disabled={!canUndo}
        onClick={onUndo}
      />
      <ActionBtn
        icon={<Redo2 size={14} />}
        label="Redo (⇧⌘Z)"
        testId="action-redo"
        disabled={!canRedo}
        onClick={onRedo}
      />

      <Divider />

      <TextFormatBar
        provider={textFormatProvider}
        active={textFormatActive}
        disabled={disabled}
        testIdPrefix="format"
      />

      <Divider />

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

      <Divider />

      <ActionBtn
        icon={<Merge size={14} />}
        label="Merge cells"
        testId="format-merge"
        disabled={disabled || !canMerge}
        onClick={onMerge}
      />
      <ActionBtn
        icon={<Split size={14} />}
        label="Unmerge cells"
        testId="format-unmerge"
        disabled={disabled || !canUnmerge}
        onClick={onUnmerge}
      />

      <Divider />

      <ActionBtn
        icon={<TableProperties size={14} />}
        label="Text to Columns"
        testId="data-text-to-columns"
        disabled={disabled || !canTextToColumns}
        onClick={onTextToColumns}
      />

      <ToggleBtn
        icon={<Filter size={14} />}
        label={filterActive ? "Remove AutoFilter" : "Apply AutoFilter"}
        testId="data-filter-toggle"
        active={filterActive}
        disabled={disabled}
        onClick={onToggleFilter}
      />

      <Divider />

      <ActionBtn
        icon={<MessageSquarePlus size={14} />}
        label="Add comment"
        testId="action-add-comment"
        disabled={disabled || !selection}
        onClick={onAddComment}
      />

      <ActionBtn
        icon={<ImageIcon size={14} />}
        label="Insert image"
        testId="action-insert-image"
        disabled={disabled}
        onClick={onInsertImage}
      />

      <Divider />

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

      <div className="ml-auto flex items-center">
        <ActionBtn
          icon={<Keyboard size={14} />}
          label="Keyboard shortcuts (⌘/)"
          testId="action-shortcuts"
          disabled={false}
          onClick={onOpenShortcuts}
        />
      </div>
    </div>
  );
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
