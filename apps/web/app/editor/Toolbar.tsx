"use client";

import { useEffect, useRef, useState, type ReactNode } from "react";
import {
  Bold,
  Italic,
  Underline,
  Strikethrough,
  Download,
  FileUp,
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
  Palette,
  Highlighter,
} from "lucide-react";
import { Button, cn } from "@officeai/ui";
import {
  COLOR_PALETTE,
  FONT_FAMILIES,
  FONT_SIZES,
  HIGHLIGHT_PALETTE,
  MIXED,
  type MaybeMixed,
} from "@/lib/format-helpers";
import type { TextFormat } from "@officeai/docx";

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
  activeMarks: ReadonlySet<string>;
  /** Half-points value of the run-level font-size mark active at the
   *  selection. `MIXED` means the selection straddles different sizes;
   *  `undefined` means no run carries a `font_size` mark. */
  activeFontSize: MaybeMixed<number>;
  activeFontFamily: MaybeMixed<string>;
  /** Hex string (no `#`) of the active `color` mark. */
  activeColor: MaybeMixed<string>;
  /** OOXML highlight name (`yellow`, `green`, ...) of the active highlight mark. */
  activeHighlight: MaybeMixed<string>;
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
  onOpenFile: () => void;
  onInsertImage: () => void;
  onExport: () => void;
  onSetParagraphStyle: (style: string) => void;
  onApplyFormat: (format: TextFormat) => void;
  onToggleMark: (mark: "bold" | "italic" | "underline" | "strike") => void;
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
   * Surface "not yet supported" toasts for buttons whose backing command
   * does not yet exist.
   */
  onUnsupported: (label: string) => void;
}

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
  return (
    <div
      role="toolbar"
      aria-label="Document toolbar"
      className="editor-toolbar flex flex-wrap items-center gap-1 border-b border-divider pb-3"
    >
      {/* File group */}
      <button
        type="button"
        onClick={props.onOpenFile}
        className="inline-flex items-center gap-1.5 rounded-md border border-divider bg-surface px-2.5 py-1 text-xs text-foreground hover:bg-hover"
      >
        <FileUp size={14} />
        Open .docx
      </button>

      <Divider />

      {/* Paragraph style — derived from snapshot.root.body. */}
      <ParagraphStylePicker
        value={props.activeStyle}
        options={props.styleOptions}
        onChange={props.onSetParagraphStyle}
        disabled={!props.agentReady}
      />

      {/* Font family + size, both controlled. */}
      <FontFamilyPicker
        value={props.activeFontFamily}
        onChange={(family) => props.onApplyFormat({ fontFamily: family })}
        disabled={!props.agentReady}
      />
      <FontSizePicker
        value={props.activeFontSize}
        onChange={(halfPoints) => props.onApplyFormat({ fontSize: halfPoints })}
        disabled={!props.agentReady}
      />

      <Divider />

      {/* Inline marks */}
      <ToolbarBtn
        label="Bold"
        active={props.activeMarks.has("bold")}
        onClick={() => props.onToggleMark("bold")}
      >
        <Bold size={14} />
      </ToolbarBtn>
      <ToolbarBtn
        label="Italic"
        active={props.activeMarks.has("italic")}
        onClick={() => props.onToggleMark("italic")}
      >
        <Italic size={14} />
      </ToolbarBtn>
      <ToolbarBtn
        label="Underline"
        active={props.activeMarks.has("underline")}
        onClick={() => props.onToggleMark("underline")}
      >
        <Underline size={14} />
      </ToolbarBtn>
      <ToolbarBtn
        label="Strike"
        active={props.activeMarks.has("strikethrough")}
        onClick={() => props.onToggleMark("strike")}
      >
        <Strikethrough size={14} />
      </ToolbarBtn>

      <Divider />

      {/* Color + highlight */}
      <ColorPicker
        label="Font color"
        icon={<Palette size={14} />}
        value={typeof props.activeColor === "string" ? `#${props.activeColor}` : undefined}
        items={COLOR_PALETTE.map((c) => ({ value: c.hex, label: c.name, swatch: `#${c.hex}` }))}
        onPick={(hex) => props.onApplyFormat({ color: hex })}
        disabled={!props.agentReady}
      />
      <ColorPicker
        label="Highlight"
        icon={<Highlighter size={14} />}
        value={
          typeof props.activeHighlight === "string"
            ? HIGHLIGHT_PALETTE.find((h) => h.name === props.activeHighlight)?.swatch
            : undefined
        }
        items={HIGHLIGHT_PALETTE.map((h) => ({ value: h.name, label: h.label, swatch: h.swatch }))}
        onPick={(name) => props.onApplyFormat({ highlight: name })}
        disabled={!props.agentReady}
      />

      <Divider />

      {/* Alignment */}
      <ToolbarBtn
        label="Align left"
        active={props.activeAlignment === "left"}
        onClick={() => props.onSetAlignment("left")}
      >
        <AlignLeft size={14} />
      </ToolbarBtn>
      <ToolbarBtn
        label="Align center"
        active={props.activeAlignment === "center"}
        onClick={() => props.onSetAlignment("center")}
      >
        <AlignCenter size={14} />
      </ToolbarBtn>
      <ToolbarBtn
        label="Align right"
        active={props.activeAlignment === "right"}
        onClick={() => props.onSetAlignment("right")}
      >
        <AlignRight size={14} />
      </ToolbarBtn>
      <ToolbarBtn
        label="Align justify"
        active={props.activeAlignment === "justify"}
        onClick={() => props.onSetAlignment("justify")}
      >
        <AlignJustify size={14} />
      </ToolbarBtn>

      <Divider />

      {/* Indentation — ±360 twips per click (¼ inch, matches Word). */}
      <ToolbarBtn label="Decrease indent" onClick={() => props.onAdjustIndent(-360)}>
        <Outdent size={14} />
      </ToolbarBtn>
      <ToolbarBtn label="Increase indent" onClick={() => props.onAdjustIndent(360)}>
        <Indent size={14} />
      </ToolbarBtn>
      {props.activeIndentLeft !== null && props.activeIndentLeft > 0 && (
        <span className="px-1 text-[11px] tabular-nums text-secondary" title="Left indent">
          {twipsToInches(props.activeIndentLeft)}"
        </span>
      )}

      <Divider />

      <SpacingMenu
        spacing={props.activeSpacing}
        onApply={props.onSetParagraphSpacing}
        disabled={!props.agentReady || props.activeSpacing === null}
      />

      <Divider />

      {/* Lists */}
      <ToolbarBtn label="Bullet list" onClick={() => props.onToggleList("bullet")}>
        <List size={14} />
      </ToolbarBtn>
      <ToolbarBtn label="Numbered list" onClick={() => props.onToggleList("ordered")}>
        <ListOrdered size={14} />
      </ToolbarBtn>

      <Divider />

      {/* Image insert */}
      <ToolbarBtn label="Insert image" onClick={props.onInsertImage}>
        <ImageIcon size={14} />
      </ToolbarBtn>

      {/* Comment */}
      <ToolbarBtn label="Add comment" onClick={props.onAddComment}>
        <MessageSquarePlus size={14} />
      </ToolbarBtn>

      <div className="ml-auto flex items-center gap-3 text-xs text-secondary">
        {props.docInfo && (
          <span className="hidden whitespace-nowrap md:inline">
            {props.docInfo.paragraphs} paragraphs · rev {props.docInfo.revision} ·{" "}
            {props.docInfo.commentThreads} comment{props.docInfo.commentThreads === 1 ? "" : "s"}
          </span>
        )}
        <Button variant="accent" size="sm" onClick={props.onExport}>
          <Download size={14} />
          Export
        </Button>
      </div>
    </div>
  );
}

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
  // Ensure the active value is in the option list so the <select> doesn't
  // silently drop the displayed value to "Normal" when the doc carries a
  // style id we haven't surfaced yet.
  const hasActive = props.options.some((o) => o.value === props.value);
  const items = hasActive
    ? props.options
    : [{ value: props.value, label: props.value || "—" }, ...props.options];
  return (
    <label className="inline-flex items-center gap-1 text-xs text-secondary">
      <span className="sr-only">Paragraph style</span>
      <select
        title="Paragraph style"
        aria-label="Paragraph style"
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

function FontSizePicker(props: {
  value: MaybeMixed<number>;
  onChange: (halfPoints: number) => void;
  disabled?: boolean;
}): ReactNode {
  const display = formatFontSize(props.value);
  // Selected option must be a string in the option list. Use "" for the
  // mixed/empty cases so the placeholder option remains selected.
  const selectValue =
    typeof props.value === "number" && FONT_SIZES.includes(props.value) ? String(props.value) : "";
  return (
    <label className="inline-flex items-center gap-1 text-xs text-secondary">
      <span className="sr-only">Font size</span>
      <select
        title="Font size"
        aria-label="Font size"
        value={selectValue}
        disabled={props.disabled}
        onChange={(e) => {
          const n = Number(e.target.value);
          if (!Number.isFinite(n) || n <= 0) return;
          props.onChange(n);
        }}
        className="h-7 w-16 rounded-md border border-divider bg-surface px-2 text-xs text-foreground hover:bg-hover focus:outline-none"
      >
        <option value="" disabled>
          {display}
        </option>
        {FONT_SIZES.map((halfPts) => (
          <option key={halfPts} value={halfPts}>
            {halfPts / 2}
          </option>
        ))}
      </select>
    </label>
  );
}

function formatFontSize(value: MaybeMixed<number>): string {
  if (value === MIXED) return "—";
  if (typeof value !== "number") return "Size";
  return String(value / 2);
}

function FontFamilyPicker(props: {
  value: MaybeMixed<string>;
  onChange: (family: string) => void;
  disabled?: boolean;
}): ReactNode {
  const families =
    props.value && typeof props.value === "string" && !FONT_FAMILIES.includes(props.value)
      ? [props.value, ...FONT_FAMILIES]
      : FONT_FAMILIES;
  const selectValue = typeof props.value === "string" ? props.value : "";
  const placeholder = props.value === MIXED ? "—" : "Font";
  return (
    <label className="inline-flex items-center gap-1 text-xs text-secondary">
      <span className="sr-only">Font family</span>
      <select
        title="Font family"
        aria-label="Font family"
        value={selectValue}
        disabled={props.disabled}
        onChange={(e) => {
          if (!e.target.value) return;
          props.onChange(e.target.value);
        }}
        className="h-7 w-32 rounded-md border border-divider bg-surface px-2 text-xs text-foreground hover:bg-hover focus:outline-none"
      >
        <option value="" disabled>
          {placeholder}
        </option>
        {families.map((f) => (
          <option key={f} value={f}>
            {f}
          </option>
        ))}
      </select>
    </label>
  );
}

interface ColorItem {
  value: string;
  label: string;
  swatch: string;
}

function ColorPicker(props: {
  label: string;
  icon: ReactNode;
  /** Active swatch (hex including `#`, or a CSS-named color). Undefined → no stripe. */
  value: string | undefined;
  items: ReadonlyArray<ColorItem>;
  onPick: (value: string) => void;
  disabled?: boolean;
}): ReactNode {
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

  return (
    <div ref={ref} className="relative inline-flex">
      <button
        type="button"
        title={props.label}
        aria-label={props.label}
        aria-haspopup="menu"
        aria-expanded={open}
        disabled={props.disabled}
        onClick={() => setOpen((v) => !v)}
        className="flex flex-col items-center gap-0 rounded-md p-1.5 text-secondary hover:bg-hover hover:text-foreground"
      >
        <span className="flex items-center gap-0.5">
          {props.icon}
          <ChevronDown size={10} />
        </span>
        <span
          aria-hidden
          className="mt-0.5 block h-0.5 w-4 rounded-sm"
          style={{ background: props.value ?? "transparent" }}
        />
      </button>
      {open && (
        <div
          role="menu"
          className="absolute left-0 top-full z-30 mt-1 grid w-44 grid-cols-4 gap-1 rounded-md border border-divider bg-surface p-2 shadow-md"
        >
          {props.items.map((item) => (
            <button
              key={item.value}
              type="button"
              role="menuitem"
              title={item.label}
              aria-label={`${props.label}: ${item.label}`}
              onClick={() => {
                props.onPick(item.value);
                setOpen(false);
              }}
              className="flex h-7 w-7 items-center justify-center rounded border border-divider hover:scale-110"
              style={{ background: item.swatch }}
            />
          ))}
        </div>
      )}
    </div>
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

  const lineDisplay = formatLineSpacing(props.spacing);
  const beforePts = toPoints(props.spacing?.before);
  const afterPts = toPoints(props.spacing?.after);

  return (
    <div ref={ref} className="relative inline-flex">
      <button
        type="button"
        title="Line and paragraph spacing"
        aria-label="Line and paragraph spacing"
        aria-haspopup="menu"
        aria-expanded={open}
        disabled={props.disabled}
        onClick={() => setOpen((v) => !v)}
        className="inline-flex items-center gap-1 rounded-md px-2 py-1.5 text-xs text-secondary hover:bg-hover hover:text-foreground disabled:opacity-50"
      >
        <span className="tabular-nums">{lineDisplay}</span>
        <ChevronDown size={10} />
      </button>
      {open && (
        <div
          role="menu"
          className="absolute right-0 top-full z-30 mt-1 w-56 rounded-md border border-divider bg-surface p-2 text-xs shadow-md"
        >
          <div className="mb-2 font-medium text-foreground">Line spacing</div>
          {LINE_SPACING_PRESETS.map((preset) => (
            <button
              key={preset.label}
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
              <span>{preset.label}</span>
              <span className="text-secondary tabular-nums">{(preset.line / 240).toFixed(2)}×</span>
            </button>
          ))}
          <div className="mt-3 border-t border-divider pt-2">
            <div className="mb-1 font-medium text-foreground">Paragraph spacing (pt)</div>
            <label className="mb-1 flex items-center justify-between gap-2">
              <span className="text-secondary">Before</span>
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
              <span className="text-secondary">After</span>
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
        </div>
      )}
    </div>
  );
}

const LINE_SPACING_PRESETS: ReadonlyArray<{ label: string; line: number }> = [
  { label: "Single", line: 240 },
  { label: "1.15", line: 276 },
  { label: "1.5", line: 360 },
  { label: "Double", line: 480 },
];

function formatLineSpacing(s: ResolvedSpacingDisplay | null): string {
  if (!s || s.line === undefined) return "Spacing";
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
