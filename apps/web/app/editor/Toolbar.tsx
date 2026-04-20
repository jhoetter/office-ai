"use client";

import { useEffect, useRef, useState, type ReactNode } from "react";
import {
  Image as ImageIcon,
  MessageSquarePlus,
  AlignLeft,
  AlignCenter,
  AlignRight,
  AlignJustify,
  Indent,
  Outdent,
  List,
  ListOrdered,
  ChevronDown,
  Pencil,
  PenLine,
  Eye,
  Pilcrow,
  Check,
  X,
  ScrollText,
  SeparatorHorizontal,
} from "lucide-react";
import { TextFormatBar, cn } from "@officeai/ui";
import { InsertTableMenu } from "./InsertTableMenu";
import type { ActiveTextFormat, TextFormatProvider } from "@officeai/text-formatting";
import { ToolbarMenu, ToolbarRow } from "../lib/shell";
import { useTranslator } from "@/lib/i18n";

export interface ToolbarStyleOption {
  value: string;
  label: string;
}

export type AlignmentValue = "left" | "center" | "right" | "justify";

/**
 * Effective spacing values surfaced to the toolbar for display. All
 * fields are optional because the cascade may not resolve them (e.g. a
 * paragraph in a synthetic doc with no styles part may have no `line`).
 */
export interface ResolvedSpacingDisplay {
  readonly line?: number;
  readonly lineRule?: "auto" | "exact" | "atLeast";
  readonly before?: number;
  readonly after?: number;
}

export interface ToolbarProps {
  agentReady: boolean;
  docInfo: { paragraphs: number; revision: number; commentThreads: number } | null;
  activeStyle: string;
  /**
   * Shared text-formatting provider. Owns selection-aware reads
   * (B/I/U/S, font family/size/color/highlight) and dispatches the
   * corresponding `docx:format-range` patches via the agent. Wired
   * up in `DocxEditor.tsx` per render.
   */
  textFormatProvider: TextFormatProvider;
  textFormatActive: ActiveTextFormat;
  /** Alignment of the paragraph containing the caret, or `null`. */
  activeAlignment: AlignmentValue | null;
  /**
   * Effective `<w:spacing>` for the paragraph containing the caret, after
   * the style cascade has been resolved. Drives the line-spacing dropdown
   * and the before/after numeric readouts. `null` when no paragraph is
   * focused.
   */
  activeSpacing: ResolvedSpacingDisplay | null;
  /**
   * Effective left indent in twips for the paragraph containing the
   * caret. `null` when no paragraph is focused. Pure display — the
   * existing ±360 buttons mutate via `onAdjustIndent`.
   */
  activeIndentLeft: number | null;
  /** Style picker contents derived from the loaded document. */
  styleOptions: ReadonlyArray<ToolbarStyleOption>;
  onInsertImage: () => void;
  onInsertTable: (rows: number, cols: number) => void;
  onSetParagraphStyle: (style: string) => void;
  onSetAlignment: (alignment: AlignmentValue) => void;
  onAdjustIndent: (deltaTwips: number) => void;
  /**
   * Apply a `docx:set-paragraph-spacing` to the paragraph the caret is in.
   * Each field follows the command's `null = clear` / `undefined = leave`
   * semantics. The toolbar always sends `lineRule` alongside `line` to
   * keep the OOXML pair consistent.
   */
  onSetParagraphSpacing: (patch: {
    line?: number | null;
    lineRule?: "auto" | "exact" | "atLeast" | null;
    before?: number | null;
    after?: number | null;
  }) => void;
  onToggleList: (kind: "bullet" | "ordered") => void;
  onAddComment: () => void;
  /**
   * B11 — Section break menu. The four OOXML section types map to
   * Word's Insert › Breaks submenu: Next page, Continuous, Even
   * page, Odd page. The toolbar dispatches this with `paragraphIndex`
   * resolved at the editor layer.
   */
  onInsertSectionBreak: (type: "nextPage" | "continuous" | "evenPage" | "oddPage") => void;
  /**
   * Surface "not yet supported" toasts for buttons whose backing command
   * does not yet exist.
   */
  onUnsupported: (label: string) => void;
  /**
   * Editor interaction mode (Word "Track Changes" surface):
   *   - `"edit"` — direct edits;
   *   - `"suggest"` — every insert/delete becomes a tracked
   *     `<w:ins>` / `<w:del>` revision (the Suggesting / redlining
   *     mode);
   *   - `"view"` — read-only.
   * The picker writes through to PM's `setEditMode` without
   * remounting so the cursor and selection survive mode changes.
   */
  editMode: EditModeValue;
  onSetEditMode: (mode: EditModeValue) => void;
  /** Word's pilcrow toggle — show/hide nonprinting characters. */
  formattingMarksOn: boolean;
  onToggleFormattingMarks: () => void;
  /**
   * B8 — Review tab. Number of unresolved tracked-change wrappers in
   * the document; when 0 the menu disables its Accept/Reject items.
   */
  trackedChangesCount: number;
  onAcceptAllChanges: () => void;
  onRejectAllChanges: () => void;
}

export type EditModeValue = "edit" | "suggest" | "view";

/**
 * Editor toolbar. Layout is intentionally close to Word's: file/style on
 * the left, inline marks + color in the middle, alignment / indent /
 * list on the right, and the export action pinned to the far right
 * with the doc metadata strip.
 *
 * Selection-binding contract (P2.2): every dropdown / pressed-state
 * button derives its value from `activeXxx` props that are recomputed
 * on every selection change in `DocxEditor.tsx`. `MIXED` selections
 * render as "—" in the FontSize/FontFamily pickers; `undefined` means
 * "no mark on this run" and renders blank.
 */
export function Toolbar(props: ToolbarProps): ReactNode {
  const { t } = useTranslator();
  return (
    <ToolbarRow
      ariaLabel={t("docx.toolbar.ariaLabel")}
      testId="docx-toolbar"
      leadingClassName="editor-toolbar"
      trailing={
        <div className="flex items-center gap-3 text-xs text-secondary">
          {props.docInfo && (
            <span className="hidden whitespace-nowrap md:inline">
              {t(
                props.docInfo.commentThreads === 1
                  ? "docx.toolbar.paragraphCountOne"
                  : "docx.toolbar.paragraphCount",
                {
                  n: props.docInfo.paragraphs,
                  revision: props.docInfo.revision,
                  comments: props.docInfo.commentThreads,
                }
              )}
            </span>
          )}
          <ReviewMenu
            count={props.trackedChangesCount}
            onAcceptAll={props.onAcceptAllChanges}
            onRejectAll={props.onRejectAllChanges}
          />
          <EditModePicker value={props.editMode} onChange={props.onSetEditMode} />
        </div>
      }
    >
      {/* Paragraph style — derived from snapshot.root.body. */}
      <ParagraphStylePicker
        value={props.activeStyle}
        options={props.styleOptions}
        onChange={props.onSetParagraphStyle}
        disabled={!props.agentReady}
      />

      <TextFormatBar
        provider={props.textFormatProvider}
        active={props.textFormatActive}
        disabled={!props.agentReady}
      />

      <Divider />

      {/* Alignment */}
      <ToolbarBtn
        label={t("docx.toolbar.alignLeft")}
        active={props.activeAlignment === "left"}
        onClick={() => props.onSetAlignment("left")}
      >
        <AlignLeft size={14} />
      </ToolbarBtn>
      <ToolbarBtn
        label={t("docx.toolbar.alignCenter")}
        active={props.activeAlignment === "center"}
        onClick={() => props.onSetAlignment("center")}
      >
        <AlignCenter size={14} />
      </ToolbarBtn>
      <ToolbarBtn
        label={t("docx.toolbar.alignRight")}
        active={props.activeAlignment === "right"}
        onClick={() => props.onSetAlignment("right")}
      >
        <AlignRight size={14} />
      </ToolbarBtn>
      <ToolbarBtn
        label={t("docx.toolbar.alignJustify")}
        active={props.activeAlignment === "justify"}
        onClick={() => props.onSetAlignment("justify")}
      >
        <AlignJustify size={14} />
      </ToolbarBtn>

      <Divider />

      {/* Indentation — ±360 twips per click (¼ inch, matches Word). */}
      <ToolbarBtn label={t("docx.toolbar.decreaseIndent")} onClick={() => props.onAdjustIndent(-360)}>
        <Outdent size={14} />
      </ToolbarBtn>
      <ToolbarBtn label={t("docx.toolbar.increaseIndent")} onClick={() => props.onAdjustIndent(360)}>
        <Indent size={14} />
      </ToolbarBtn>
      {/*
        Left-indent readout. Reserves a fixed-width slot regardless of
        whether the active paragraph has any indent, so toggling
        between indented and non-indented selections never reflows
        the toolbar.
      */}
      <span
        className="inline-block min-w-[3.25rem] px-1 text-[11px] tabular-nums text-secondary"
        title={t("docx.toolbar.leftIndent")}
        aria-hidden={props.activeIndentLeft === null || props.activeIndentLeft <= 0}
      >
        {props.activeIndentLeft !== null && props.activeIndentLeft > 0
          ? `${twipsToInches(props.activeIndentLeft)}"`
          : ""}
      </span>

      <Divider />

      <SpacingMenu
        spacing={props.activeSpacing}
        onApply={props.onSetParagraphSpacing}
        disabled={!props.agentReady || props.activeSpacing === null}
      />

      <Divider />

      {/* Lists */}
      <ToolbarBtn label={t("docx.toolbar.bulletList")} onClick={() => props.onToggleList("bullet")}>
        <List size={14} />
      </ToolbarBtn>
      <ToolbarBtn label={t("docx.toolbar.numberedList")} onClick={() => props.onToggleList("ordered")}>
        <ListOrdered size={14} />
      </ToolbarBtn>

      {/* Show formatting marks (Word's pilcrow toggle) */}
      <ToolbarBtn
        label={t("docx.toolbar.showFormattingMarks")}
        active={props.formattingMarksOn}
        onClick={props.onToggleFormattingMarks}
      >
        <Pilcrow size={14} />
      </ToolbarBtn>

      <Divider />

      {/* Image insert */}
      <ToolbarBtn label={t("docx.toolbar.insertImage")} onClick={props.onInsertImage}>
        <ImageIcon size={14} />
      </ToolbarBtn>

      {/* Insert table — Word-style grid picker. */}
      <InsertTableMenu disabled={!props.agentReady} onInsert={props.onInsertTable} />

      {/* Section break — Word-style submenu (B11). */}
      <SectionBreakMenu disabled={!props.agentReady} onInsert={props.onInsertSectionBreak} />

      {/* Comment */}
      <ToolbarBtn label={t("docx.toolbar.addComment")} onClick={props.onAddComment}>
        <MessageSquarePlus size={14} />
      </ToolbarBtn>
    </ToolbarRow>
  );
}

/**
 * Edit-mode picker (Editing / Suggesting / Viewing — the Word and
 * Google Docs surface). Visualises the current mode with a colored
 * pill so the user always knows whether their next keystroke goes
 * to plain text, becomes a tracked suggestion, or is rejected
 * outright by the read-only surface.
 *
 * Click the pill to open a simple menu; the previously-selected
 * option is highlighted. Mode changes are dispatched immediately
 * (no confirm step) — the user can always revert by re-opening the
 * picker.
 */
function EditModePicker(props: { value: EditModeValue; onChange: (v: EditModeValue) => void }): ReactNode {
  const { t } = useTranslator();
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    if (!open) return;
    const onDoc = (e: MouseEvent) => {
      if (!ref.current) return;
      if (!ref.current.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, [open]);

  const meta = EDIT_MODE_META[props.value];
  const modeLabel = t(`docx.toolbar.${props.value === "edit" ? "editing" : props.value === "suggest" ? "suggesting" : "viewing"}`);
  return (
    <div ref={ref} className="relative">
      <button
        type="button"
        title={t("docx.toolbar.editModeLabel", { mode: modeLabel })}
        aria-label={t("docx.toolbar.editModeLabel", { mode: modeLabel })}
        aria-haspopup="menu"
        aria-expanded={open}
        onClick={() => setOpen((v) => !v)}
        className={cn(
          "inline-flex items-center gap-1.5 rounded-md border px-2 py-1 text-xs font-medium transition-colors",
          meta.pillClass
        )}
        data-testid="edit-mode-picker"
        data-edit-mode={props.value}
      >
        <meta.Icon size={12} />
        {modeLabel}
        <ChevronDown size={10} />
      </button>
      {open && (
        <div
          role="menu"
          className="absolute right-0 z-30 mt-1 w-56 rounded-md border border-divider bg-surface p-1 text-xs shadow-md"
        >
          {(Object.keys(EDIT_MODE_META) as EditModeValue[]).map((key) => {
            const m = EDIT_MODE_META[key];
            const active = key === props.value;
            const labelKey = key === "edit" ? "editing" : key === "suggest" ? "suggesting" : "viewing";
            const descKey =
              key === "edit"
                ? "editingDescription"
                : key === "suggest"
                  ? "suggestingDescription"
                  : "viewingDescription";
            return (
              <button
                key={key}
                type="button"
                role="menuitemradio"
                aria-checked={active}
                onClick={() => {
                  props.onChange(key);
                  setOpen(false);
                }}
                data-testid={`edit-mode-option-${key}`}
                className={cn(
                  "flex w-full items-start gap-2 rounded px-2 py-1.5 text-left hover:bg-hover",
                  active && "bg-hover"
                )}
              >
                <m.Icon size={12} className={cn("mt-0.5 shrink-0", m.iconColorClass)} />
                <span className="min-w-0 flex-1">
                  <span className="block font-medium text-foreground">{t(`docx.toolbar.${labelKey}`)}</span>
                  <span className="block text-[11px] text-secondary">{t(`docx.toolbar.${descKey}`)}</span>
                </span>
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}

/**
 * B11 — Section break menu. Mirrors Word's Insert › Breaks submenu
 * with the four legal OOXML section types. The shortcut for the
 * common "Next page" entry (Mod+Shift+Enter) is surfaced inline so
 * keyboard-first users can discover it.
 */
function SectionBreakMenu(props: {
  disabled: boolean;
  onInsert: (type: "nextPage" | "continuous" | "evenPage" | "oddPage") => void;
}): ReactNode {
  const { t } = useTranslator();
  const [open, setOpen] = useState(false);
  const triggerRef = useRef<HTMLButtonElement | null>(null);

  const choose = (type: "nextPage" | "continuous" | "evenPage" | "oddPage") => {
    props.onInsert(type);
    setOpen(false);
  };

  return (
    <span className="inline-flex">
      <button
        ref={triggerRef}
        type="button"
        title={t("docx.toolbar.insertSectionBreak")}
        aria-label={t("docx.toolbar.insertSectionBreak")}
        aria-haspopup="menu"
        aria-expanded={open}
        disabled={props.disabled}
        onClick={() => setOpen((v) => !v)}
        className={cn(
          "inline-flex h-7 items-center gap-1 rounded-md border border-transparent px-1.5 text-xs text-foreground transition-colors hover:bg-hover",
          props.disabled && "cursor-not-allowed opacity-50 hover:bg-transparent"
        )}
        data-testid="section-break-menu-button"
      >
        <SeparatorHorizontal size={14} />
        <ChevronDown size={10} />
      </button>
      <ToolbarMenu
        open={open}
        onClose={() => setOpen(false)}
        triggerRef={triggerRef}
        role="menu"
        className="w-64 rounded-md border border-divider bg-surface p-1 text-xs shadow-md"
      >
        <SectionBreakMenuItem
          label={t("docx.section.nextPage")}
          description={t("docx.section.nextPageDescription")}
          shortcut="Mod+Shift+Enter"
          onClick={() => choose("nextPage")}
          testId="section-break-next-page"
        />
        <SectionBreakMenuItem
          label={t("docx.section.continuous")}
          description={t("docx.section.continuousDescription")}
          onClick={() => choose("continuous")}
          testId="section-break-continuous"
        />
        <SectionBreakMenuItem
          label={t("docx.section.evenPage")}
          description={t("docx.section.evenPageDescription")}
          onClick={() => choose("evenPage")}
          testId="section-break-even"
        />
        <SectionBreakMenuItem
          label={t("docx.section.oddPage")}
          description={t("docx.section.oddPageDescription")}
          onClick={() => choose("oddPage")}
          testId="section-break-odd"
        />
      </ToolbarMenu>
    </span>
  );
}

function SectionBreakMenuItem(props: {
  label: string;
  description: string;
  shortcut?: string;
  onClick: () => void;
  testId: string;
}): ReactNode {
  return (
    <button
      type="button"
      role="menuitem"
      onClick={props.onClick}
      data-testid={props.testId}
      className="flex w-full items-start gap-2 rounded px-2 py-1.5 text-left hover:bg-hover"
    >
      <span className="min-w-0 flex-1">
        <span className="block font-medium text-foreground">{props.label}</span>
        <span className="block text-[11px] text-secondary">{props.description}</span>
      </span>
      {props.shortcut ? (
        <kbd className="ml-2 shrink-0 rounded bg-hover px-1 py-0.5 font-mono text-[10px] text-secondary">
          {props.shortcut}
        </kbd>
      ) : null}
    </button>
  );
}

/**
 * B8 — Review menu.
 *
 * Word's "Review" tab condensed into a single menu next to the edit
 * mode picker so it stays visible at every viewport width. Carries
 * a small badge with the live count of unresolved revisions; the
 * Accept-all / Reject-all entries are disabled when the count is
 * zero so the menu becomes a stable affordance instead of jumping
 * in and out of the layout.
 */
function ReviewMenu(props: { count: number; onAcceptAll: () => void; onRejectAll: () => void }): ReactNode {
  const { t } = useTranslator();
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    if (!open) return;
    const onDoc = (e: MouseEvent) => {
      if (!ref.current) return;
      if (!ref.current.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, [open]);

  const empty = props.count === 0;
  return (
    <div ref={ref} className="relative">
      <button
        type="button"
        title={t("docx.toolbar.reviewTracked")}
        aria-label={t("docx.toolbar.reviewTracked")}
        aria-haspopup="menu"
        aria-expanded={open}
        onClick={() => setOpen((v) => !v)}
        className={cn(
          "inline-flex items-center gap-1.5 rounded-md border border-divider bg-surface px-2 py-1 text-xs font-medium text-foreground transition-colors hover:bg-hover"
        )}
        data-testid="review-menu-button"
      >
        <ScrollText size={12} />
        {t("docx.toolbar.review")}
        {!empty && (
          <span
            className="rounded-full bg-accent px-1.5 py-0.5 text-[10px] font-semibold leading-none text-white"
            aria-label={t(
              props.count === 1 ? "docx.toolbar.unresolvedChange" : "docx.toolbar.unresolvedChanges",
              { n: props.count }
            )}
          >
            {props.count}
          </span>
        )}
        <ChevronDown size={10} />
      </button>
      {open && (
        <div
          role="menu"
          className="absolute right-0 z-30 mt-1 w-64 rounded-md border border-divider bg-surface p-1 text-xs shadow-md"
        >
          <ReviewMenuItem
            disabled={empty}
            icon={<Check size={12} className="text-[var(--success)]" />}
            label={t("docx.toolbar.acceptAll")}
            description={t("docx.toolbar.acceptAllDescription")}
            onClick={() => {
              if (empty) return;
              props.onAcceptAll();
              setOpen(false);
            }}
            testId="review-accept-all"
          />
          <ReviewMenuItem
            disabled={empty}
            icon={<X size={12} className="text-[var(--error)]" />}
            label={t("docx.toolbar.rejectAll")}
            description={t("docx.toolbar.rejectAllDescription")}
            onClick={() => {
              if (empty) return;
              props.onRejectAll();
              setOpen(false);
            }}
            testId="review-reject-all"
          />
        </div>
      )}
    </div>
  );
}

function ReviewMenuItem(props: {
  disabled: boolean;
  icon: ReactNode;
  label: string;
  description: string;
  onClick: () => void;
  testId: string;
}): ReactNode {
  return (
    <button
      type="button"
      role="menuitem"
      disabled={props.disabled}
      onClick={props.onClick}
      data-testid={props.testId}
      className={cn(
        "flex w-full items-start gap-2 rounded px-2 py-1.5 text-left",
        props.disabled ? "cursor-not-allowed opacity-50" : "hover:bg-hover"
      )}
    >
      <span className="mt-0.5 shrink-0">{props.icon}</span>
      <span className="min-w-0 flex-1">
        <span className="block font-medium text-foreground">{props.label}</span>
        <span className="block text-[11px] text-secondary">{props.description}</span>
      </span>
    </button>
  );
}

const EDIT_MODE_META: Record<
  EditModeValue,
  {
    Icon: typeof Pencil;
    pillClass: string;
    iconColorClass: string;
  }
> = {
  edit: {
    Icon: Pencil,
    pillClass: "border-divider bg-surface text-foreground hover:bg-hover",
    iconColorClass: "text-foreground",
  },
  suggest: {
    Icon: PenLine,
    pillClass:
      "border-[var(--ai-violet)] bg-[var(--ai-violet-light)] text-[var(--ai-violet)] hover:brightness-95",
    iconColorClass: "text-[var(--ai-violet)]",
  },
  view: {
    Icon: Eye,
    pillClass: "border-divider bg-hover text-secondary hover:bg-divider",
    iconColorClass: "text-secondary",
  },
};

function Divider(): ReactNode {
  return <div className="mx-1 h-4 w-px bg-divider" aria-hidden />;
}

function ToolbarBtn(props: {
  label: string;
  onClick: () => void;
  active?: boolean;
  children: ReactNode;
}): ReactNode {
  return (
    <button
      type="button"
      title={props.label}
      aria-label={props.label}
      aria-pressed={props.active ?? undefined}
      onClick={props.onClick}
      className={cn(
        "rounded-md p-1.5 text-secondary hover:bg-hover hover:text-foreground",
        props.active && "bg-accent-light text-foreground"
      )}
    >
      {props.children}
    </button>
  );
}

function ParagraphStylePicker(props: {
  value: string;
  options: ReadonlyArray<ToolbarStyleOption>;
  onChange: (v: string) => void;
  disabled?: boolean;
}): ReactNode {
  const { t } = useTranslator();
  // Ensure the active value is in the option list so the <select> doesn't
  // silently drop the displayed value to "Normal" when the doc carries a
  // style id we haven't surfaced yet.
  const hasActive = props.options.some((o) => o.value === props.value);
  const items = hasActive
    ? props.options
    : [{ value: props.value, label: props.value || "—" }, ...props.options];
  return (
    <label className="inline-flex items-center gap-1 text-xs text-secondary">
      <span className="sr-only">{t("docx.toolbar.style")}</span>
      <select
        title={t("docx.toolbar.style")}
        aria-label={t("docx.toolbar.style")}
        value={props.value}
        disabled={props.disabled}
        onChange={(e) => props.onChange(e.target.value)}
        className="h-7 max-w-40 rounded-md border border-divider bg-surface px-2 text-xs text-foreground hover:bg-hover focus:outline-none"
      >
        {items.map((s) => (
          <option key={s.value} value={s.value}>
            {s.label}
          </option>
        ))}
      </select>
    </label>
  );
}

/**
 * Spacing dropdown — line spacing (single / 1.15 / 1.5 / double) and
 * per-paragraph before/after spacing in points. Reads the resolved
 * effective spacing for display so a Heading style's inherited "1.5
 * line" surfaces here even when the paragraph has no direct
 * `<w:spacing>`.
 *
 * Word stores `<w:spacing w:line>` in twentieths of a line for
 * `lineRule="auto"` (so 240 = 1.0, 360 = 1.5) and in twips for
 * `exact` / `atLeast`. We mutate the `auto` family from the picker
 * because that covers the >95% case; the existing exact/atLeast
 * value is shown but the picker resets `lineRule` to `auto` on
 * change. Before/after stay as twips for OOXML parity (1pt = 20
 * twips); inputs accept points and convert.
 */
function SpacingMenu(props: {
  spacing: ResolvedSpacingDisplay | null;
  onApply: (patch: {
    line?: number | null;
    lineRule?: "auto" | "exact" | "atLeast" | null;
    before?: number | null;
    after?: number | null;
  }) => void;
  disabled?: boolean;
}): ReactNode {
  const { t } = useTranslator();
  const [open, setOpen] = useState(false);
  const triggerRef = useRef<HTMLButtonElement | null>(null);

  const lineDisplay = formatLineSpacing(props.spacing, t);
  const beforePts = toPoints(props.spacing?.before);
  const afterPts = toPoints(props.spacing?.after);

  const presets: ReadonlyArray<{ key: string; line: number }> = [
    { key: "lineSingle", line: 240 },
    { key: "line115", line: 276 },
    { key: "line150", line: 360 },
    { key: "lineDouble", line: 480 },
  ];

  return (
    <span className="inline-flex">
      <button
        ref={triggerRef}
        type="button"
        title={t("docx.toolbar.lineParagraphSpacing")}
        aria-label={t("docx.toolbar.lineParagraphSpacing")}
        aria-haspopup="menu"
        aria-expanded={open}
        disabled={props.disabled}
        onClick={() => setOpen((v) => !v)}
        className="inline-flex items-center gap-1 rounded-md px-2 py-1.5 text-xs text-secondary hover:bg-hover hover:text-foreground disabled:opacity-50"
      >
        <span className="tabular-nums">{lineDisplay}</span>
        <ChevronDown size={10} />
      </button>
      <ToolbarMenu
        open={open}
        onClose={() => setOpen(false)}
        triggerRef={triggerRef}
        align="right"
        role="menu"
        className="w-56 rounded-md border border-divider bg-surface p-2 text-xs shadow-md"
      >
        <div className="mb-2 font-medium text-foreground">{t("docx.toolbar.lineSpacing")}</div>
        {presets.map((preset) => (
          <button
            key={preset.key}
            type="button"
            role="menuitem"
            onClick={() => {
              props.onApply({ line: preset.line, lineRule: "auto" });
              setOpen(false);
            }}
            className={cn(
              "flex w-full items-center justify-between rounded px-2 py-1 text-left hover:bg-hover",
              isActiveLine(props.spacing, preset.line) && "bg-accent-light text-foreground"
            )}
          >
            <span>{t(`docx.toolbar.${preset.key}`)}</span>
            <span className="text-secondary tabular-nums">{(preset.line / 240).toFixed(2)}×</span>
          </button>
        ))}
        <div className="mt-3 border-t border-divider pt-2">
          <div className="mb-1 font-medium text-foreground">{t("docx.toolbar.paragraphSpacingPt")}</div>
          <label className="mb-1 flex items-center justify-between gap-2">
            <span className="text-secondary">{t("docx.toolbar.before")}</span>
            <input
              type="number"
              min={0}
              step={1}
              defaultValue={beforePts ?? ""}
              onBlur={(e) => {
                const v = e.target.value === "" ? null : Number(e.target.value);
                if (v === null) props.onApply({ before: null });
                else if (Number.isFinite(v) && v >= 0) props.onApply({ before: Math.round(v * 20) });
              }}
              className="h-6 w-16 rounded border border-divider bg-surface px-1 text-right tabular-nums focus:outline-none"
            />
          </label>
          <label className="flex items-center justify-between gap-2">
            <span className="text-secondary">{t("docx.toolbar.after")}</span>
            <input
              type="number"
              min={0}
              step={1}
              defaultValue={afterPts ?? ""}
              onBlur={(e) => {
                const v = e.target.value === "" ? null : Number(e.target.value);
                if (v === null) props.onApply({ after: null });
                else if (Number.isFinite(v) && v >= 0) props.onApply({ after: Math.round(v * 20) });
              }}
              className="h-6 w-16 rounded border border-divider bg-surface px-1 text-right tabular-nums focus:outline-none"
            />
          </label>
        </div>
      </ToolbarMenu>
    </span>
  );
}

function formatLineSpacing(
  s: ResolvedSpacingDisplay | null,
  t: (key: string) => string
): string {
  if (!s || s.line === undefined) return t("docx.toolbar.spacingButton");
  if (s.lineRule && s.lineRule !== "auto") return `${(s.line / 20).toFixed(1)}pt`;
  return `${(s.line / 240).toFixed(2)}×`;
}

function isActiveLine(s: ResolvedSpacingDisplay | null, line: number): boolean {
  if (!s) return false;
  if (s.lineRule && s.lineRule !== "auto") return false;
  return s.line === line;
}

function toPoints(twips: number | undefined): number | undefined {
  if (twips === undefined) return undefined;
  return Math.round(twips / 20);
}

function twipsToInches(twips: number): string {
  return (twips / 1440).toFixed(2);
}
