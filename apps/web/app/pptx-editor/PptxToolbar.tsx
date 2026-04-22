"use client";

import * as React from "react";
import {
  AlignCenter,
  AlignCenterHorizontal,
  AlignCenterVertical,
  AlignEndHorizontal,
  AlignEndVertical,
  AlignHorizontalDistributeCenter,
  AlignJustify,
  AlignLeft,
  AlignRight,
  AlignStartHorizontal,
  AlignStartVertical,
  AlignVerticalDistributeCenter,
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  AlignVerticalJustifyStart,
  ChevronDown,
  ChevronsDown,
  ChevronsUp,
  Circle,
  Copy,
  CornerDownRight,
  Diamond,
  Group,
  Image as ImageIcon,
  Layers,
  MessageSquarePlus,
  Minus,
  MoveDown,
  MoveRight,
  MoveUp,
  Play,
  Plus,
  Shapes,
  Spline,
  Square,
  StickyNote,
  Table2,
  Trash2,
  Triangle,
  Type,
  Ungroup,
  Video,
  Wand2,
} from "lucide-react";
import { TextFormatBar } from "@officeai/ui";
import type { ActiveTextFormat, TextFormatProvider } from "@officeai/text-formatting";
import type { AlignMode, LayoutKindPayload, ReorderShapeMode, ShapePreset, TextAnchor } from "@officeai/pptx";

export type TextAlignment = "left" | "center" | "right" | "justify";
import { LayoutTemplate } from "lucide-react";
import { ToolbarMenu, ToolbarRow } from "../lib/shell";

export interface PptxToolbarProps {
  readonly disabled: boolean;
  readonly slideCount: number;
  readonly hasSelection: boolean;
  /** Total number of selected shapes; drives align/distribute enablement. */
  readonly selectionCount: number;
  readonly currentFill: string | null;
  readonly textFormatProvider: TextFormatProvider;
  readonly textFormatActive: ActiveTextFormat;
  readonly onAddSlide: () => void;
  readonly onAddSlideWithLayout: (kind: LayoutKindPayload) => void;
  readonly onSetSlideLayout: (kind: LayoutKindPayload) => void;
  readonly onDeleteSlide: () => void;
  readonly onDuplicateSlide: () => void;
  readonly onAddTextBox: () => void;
  readonly onAddShape: (preset: ShapePreset) => void;
  readonly onAddConnector: (connectorType: "straight" | "elbow" | "curved") => void;
  /**
   * Currently armed connector tool, if any. Drives the menu trigger's
   * "active" styling so the user gets a clear toggle indicator while
   * the tool is engaged (matches Figma / Slides where a pressed
   * shape-tool button signals "click on the canvas to draw").
   */
  readonly connectorToolType: "straight" | "elbow" | "curved" | null;
  readonly onInsertImage: (file: File) => void;
  /**
   * Embed a video / audio file on the active slide. Mirrors
   * `onInsertImage` but accepts `video/*` or `audio/*` and dispatches
   * `pptx:insert-media` (which writes the binary part, registers
   * slide-rels, and stamps a typed `MediaShape`).
   */
  readonly onInsertMedia: (file: File) => void;
  /**
   * Slide transition assigned to the active slide. Used by the
   * `TransitionMenu` to highlight the current selection. `null` means
   * the transition is unknown / unsupported.
   */
  readonly currentTransitionKind?: import("@officeai/pptx").TransitionKind | null;
  readonly currentTransitionSpeed?: import("@officeai/pptx").TransitionSpeed | null;
  /**
   * Apply a transition to the active slide. The toolbar menu lets the
   * user pick a kind directly; pass `null` for `speed` to keep the
   * current speed (Word/PowerPoint default is "medium").
   */
  readonly onSetSlideTransition: (
    kind: import("@officeai/pptx").TransitionKind,
    speed: import("@officeai/pptx").TransitionSpeed | null
  ) => void;
  /**
   * Open the XlsxRangePicker dialog so the user can pick a .xlsx
   * file + sheet + range and decide whether to insert it as a
   * native table, a live-OLE embedded workbook, or a chart. Mirrors
   * the three "Insert from xlsx" entries in the action catalogue.
   */
  readonly onInsertFromXlsx: () => void;
  /**
   * D9 — replace the picture's bitmap behind the currently-selected
   * `Picture` shape. Only enabled when `selectedIsPicture` is true.
   * The shell file picker reuses the same MIME allowlist as
   * `onInsertImage` so users can't introduce types the model wouldn't
   * round-trip cleanly.
   */
  readonly onReplacePicture: (file: File) => void;
  readonly selectedIsPicture: boolean;
  readonly onDeleteShape: () => void;
  readonly onAlign: (mode: AlignMode, relativeTo: "selection" | "slide") => void;
  readonly onDistribute: (axis: "horizontal" | "vertical") => void;
  readonly onReorder: (mode: ReorderShapeMode) => void;
  readonly onGroup: () => void;
  readonly onUngroup: () => void;
  readonly onDuplicateShape: () => void;
  readonly canGroup: boolean;
  readonly canUngroup: boolean;
  readonly onChangeFill: (hex: string | null) => void;
  /**
   * Per-paragraph horizontal text alignment inside the active text
   * shape (or the paragraph the text-edit caret is in). Pass `null`
   * to clear an existing override.
   */
  readonly onSetTextAlignment: (alignment: TextAlignment | null) => void;
  /** Shape-wide vertical text anchor for the active text shape. */
  readonly onSetTextAnchor: (anchor: TextAnchor | null) => void;
  /**
   * Currently effective alignment / anchor on the active text shape,
   * used to render pressed-state highlighting. `null` means "no text
   * shape is active" (controls render disabled) or "no explicit
   * value" (no button is highlighted).
   */
  readonly activeTextAlignment: TextAlignment | null;
  readonly activeTextAnchor: TextAnchor | null;
  /** True when the toolbar should enable the text-alignment cluster. */
  readonly hasTextShapeFocus: boolean;
  /**
   * Open the comments composer. The editor decides where the new pin
   * lands: shape-anchored when a shape is selected, free-pin centred
   * on the slide otherwise.
   */
  readonly onAddComment: () => void;
  /**
   * D10 — kick off Present mode. The button is rightmost in the
   * toolbar so users can scan the toolbar left-to-right edit-then-
   * present, mirroring the PowerPoint ribbon's Slide Show tab
   * placement. Zoom now lives in the shared status-bar `ZoomControl`
   * so it sits in the same place across DOCX / XLSX / PPTX.
   */
  readonly onPresent: () => void;
  readonly canPresent: boolean;
  /**
   * Toggle the speaker-notes panel beneath the slide. The panel is
   * hidden by default — PowerPoint's behaviour — and the user opts in
   * via this trailing toolbar button. `notesOpen` drives the
   * `aria-pressed` / active styling so the toggle reads as a clear
   * on/off control.
   */
  readonly onToggleNotes: () => void;
  readonly notesOpen: boolean;
}

interface ShapeOption {
  readonly preset: ShapePreset;
  readonly label: string;
  readonly icon: React.ReactNode;
}

const SHAPE_OPTIONS: ReadonlyArray<ShapeOption> = [
  { preset: "rect", label: "Rectangle", icon: <Square size={14} /> },
  { preset: "roundRect", label: "Rounded Rectangle", icon: <Square size={14} className="rounded-sm" /> },
  { preset: "ellipse", label: "Ellipse", icon: <Circle size={14} /> },
  { preset: "triangle", label: "Triangle", icon: <Triangle size={14} /> },
  { preset: "diamond", label: "Diamond", icon: <Diamond size={14} /> },
  { preset: "line", label: "Line", icon: <Minus size={14} /> },
  { preset: "rightArrow", label: "Arrow", icon: <MoveRight size={14} /> },
];

export function PptxToolbar(props: PptxToolbarProps) {
  const { disabled, hasSelection, selectionCount, currentFill, textFormatProvider, textFormatActive } = props;
  const [alignRelativeTo, setAlignRelativeTo] = React.useState<"selection" | "slide">("selection");
  const minAlignSelection = alignRelativeTo === "slide" ? 1 : 2;
  const canAlign = !disabled && selectionCount >= minAlignSelection;
  const canDistribute = !disabled && selectionCount >= 3;
  const imageInputRef = React.useRef<HTMLInputElement | null>(null);
  const mediaInputRef = React.useRef<HTMLInputElement | null>(null);
  return (
    <ToolbarRow
      ariaLabel="Slide toolbar"
      testId="pptx-toolbar"
      trailing={
        <div className="inline-flex items-center gap-1">
          <button
            type="button"
            onClick={props.onToggleNotes}
            disabled={disabled}
            aria-pressed={props.notesOpen}
            title={props.notesOpen ? "Hide speaker notes" : "Show speaker notes"}
            data-testid="pptx-notes-toggle"
            className={`inline-flex items-center gap-1 rounded border border-divider px-2 py-1 text-xs text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40 ${
              props.notesOpen ? "bg-hover ring-1 ring-[var(--accent)]/40" : ""
            }`}
          >
            <StickyNote size={14} />
            <span className="hidden sm:inline">Notes</span>
          </button>
          <button
            type="button"
            onClick={props.onPresent}
            disabled={!props.canPresent}
            title="Start presentation (F5)"
            data-testid="pptx-present"
            className="inline-flex items-center gap-1 rounded border border-divider px-2 py-1 text-xs text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
          >
            <Play size={14} />
            <span className="hidden sm:inline">Present</span>
          </button>
        </div>
      }
    >
      <ToolbarButton
        onClick={props.onAddSlide}
        icon={<Plus size={14} />}
        label="Add slide"
        disabled={disabled}
      />
      <ToolbarButton
        onClick={props.onDuplicateSlide}
        icon={<Copy size={14} />}
        label="Duplicate"
        disabled={disabled || props.slideCount < 1}
      />
      <ToolbarButton
        onClick={props.onDeleteSlide}
        icon={<Trash2 size={14} />}
        label="Delete slide"
        disabled={disabled || props.slideCount <= 1}
      />
      <LayoutMenu
        disabled={disabled}
        onAddWithLayout={props.onAddSlideWithLayout}
        onSetLayout={props.onSetSlideLayout}
        canSetLayout={props.slideCount > 0}
      />
      <TransitionMenu
        disabled={disabled || props.slideCount < 1}
        currentKind={props.currentTransitionKind ?? "none"}
        currentSpeed={props.currentTransitionSpeed ?? null}
        onSetTransition={props.onSetSlideTransition}
      />
      <Sep />
      <ToolbarButton
        onClick={props.onAddTextBox}
        icon={<Type size={14} />}
        label="Text box"
        disabled={disabled}
      />
      <ShapeMenu disabled={disabled} onPick={props.onAddShape} />
      <ConnectorMenu disabled={disabled} onPick={props.onAddConnector} activeType={props.connectorToolType} />
      <ToolbarButton
        onClick={() => imageInputRef.current?.click()}
        icon={<ImageIcon size={14} />}
        label="Image"
        disabled={disabled}
      />
      <input
        ref={imageInputRef}
        type="file"
        accept="image/png,image/jpeg,image/gif,image/webp,image/bmp"
        className="hidden"
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) props.onInsertImage(f);
          e.target.value = "";
        }}
      />
      <PictureReplaceButton
        onReplace={props.onReplacePicture}
        disabled={disabled || !props.selectedIsPicture}
      />
      <ToolbarButton
        onClick={() => mediaInputRef.current?.click()}
        icon={<Video size={14} />}
        label="Media"
        title="Insert video or audio"
        disabled={disabled}
        testId="pptx-insert-media"
      />
      <input
        ref={mediaInputRef}
        type="file"
        accept="video/*,audio/*"
        className="hidden"
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) props.onInsertMedia(f);
          e.target.value = "";
        }}
      />
      <ToolbarButton
        onClick={props.onInsertFromXlsx}
        icon={<Table2 size={14} />}
        label="From xlsx"
        disabled={disabled}
      />
      <ToolbarButton
        onClick={props.onDeleteShape}
        icon={<Trash2 size={14} />}
        label="Delete"
        disabled={disabled || !hasSelection}
      />
      <Sep />
      <IconButton
        onClick={() => props.onAlign("left", alignRelativeTo)}
        icon={<AlignStartVertical size={14} />}
        label={`Align left (${alignRelativeTo === "slide" ? "to slide" : "to selection"})`}
        disabled={!canAlign}
      />
      <IconButton
        onClick={() => props.onAlign("center-h", alignRelativeTo)}
        icon={<AlignCenterVertical size={14} />}
        label={`Align center (${alignRelativeTo === "slide" ? "to slide" : "to selection"})`}
        disabled={!canAlign}
      />
      <IconButton
        onClick={() => props.onAlign("right", alignRelativeTo)}
        icon={<AlignEndVertical size={14} />}
        label={`Align right (${alignRelativeTo === "slide" ? "to slide" : "to selection"})`}
        disabled={!canAlign}
      />
      <IconButton
        onClick={() => props.onAlign("top", alignRelativeTo)}
        icon={<AlignStartHorizontal size={14} />}
        label={`Align top (${alignRelativeTo === "slide" ? "to slide" : "to selection"})`}
        disabled={!canAlign}
      />
      <IconButton
        onClick={() => props.onAlign("middle-v", alignRelativeTo)}
        icon={<AlignCenterHorizontal size={14} />}
        label={`Align middle (${alignRelativeTo === "slide" ? "to slide" : "to selection"})`}
        disabled={!canAlign}
      />
      <IconButton
        onClick={() => props.onAlign("bottom", alignRelativeTo)}
        icon={<AlignEndHorizontal size={14} />}
        label={`Align bottom (${alignRelativeTo === "slide" ? "to slide" : "to selection"})`}
        disabled={!canAlign}
      />
      <AlignTargetToggle value={alignRelativeTo} onChange={setAlignRelativeTo} disabled={disabled} />
      <IconButton
        onClick={() => props.onDistribute("horizontal")}
        icon={<AlignHorizontalDistributeCenter size={14} />}
        label="Distribute horizontally"
        disabled={!canDistribute}
      />
      <IconButton
        onClick={() => props.onDistribute("vertical")}
        icon={<AlignVerticalDistributeCenter size={14} />}
        label="Distribute vertically"
        disabled={!canDistribute}
      />
      <ArrangeMenu
        disabled={disabled || !hasSelection}
        onReorder={props.onReorder}
        onGroup={props.onGroup}
        onUngroup={props.onUngroup}
        onDuplicateShape={props.onDuplicateShape}
        canGroup={props.canGroup}
        canUngroup={props.canUngroup}
        canDuplicate={hasSelection}
      />
      <Sep />
      {/*
        The wrapper opts the entire shared text-format bar into the
        contenteditable's "keep editing" guard (see TextEditOverlay's
        onBlur in SlideCanvas). Without it, mousedown on a toggle would
        blur the editable, commit a plain-text overwrite, and *also*
        wipe the selection ref the provider depends on.
      */}
      <div data-pptx-keep-edit className="inline-flex items-center">
        <TextFormatBar
          provider={textFormatProvider}
          active={textFormatActive}
          disabled={disabled}
          testIdPrefix="pptx-format"
        />
      </div>
      <Sep />
      {/*
        Text-alignment cluster. PowerPoint's "Align Text" controls:
        the four horizontal options drive `pptx:set-paragraph-alignment`
        and the three vertical ones drive `pptx:set-text-anchor`. Both
        target the currently active text shape (selection caret's shape
        if a text-edit overlay is open, otherwise the canvas-selected
        shape). The `data-pptx-keep-edit` wrapper is the same trick the
        TextFormatBar uses — it stops mousedown from blurring an open
        contenteditable.
      */}
      <div data-pptx-keep-edit className="inline-flex items-center gap-0.5">
        <TextAlignToggle
          label="Align text left"
          icon={<AlignLeft size={14} />}
          active={props.activeTextAlignment === "left"}
          disabled={disabled || !props.hasTextShapeFocus}
          onClick={() => props.onSetTextAlignment("left")}
          testId="pptx-text-align-left"
        />
        <TextAlignToggle
          label="Align text center"
          icon={<AlignCenter size={14} />}
          active={props.activeTextAlignment === "center"}
          disabled={disabled || !props.hasTextShapeFocus}
          onClick={() => props.onSetTextAlignment("center")}
          testId="pptx-text-align-center"
        />
        <TextAlignToggle
          label="Align text right"
          icon={<AlignRight size={14} />}
          active={props.activeTextAlignment === "right"}
          disabled={disabled || !props.hasTextShapeFocus}
          onClick={() => props.onSetTextAlignment("right")}
          testId="pptx-text-align-right"
        />
        <TextAlignToggle
          label="Justify text"
          icon={<AlignJustify size={14} />}
          active={props.activeTextAlignment === "justify"}
          disabled={disabled || !props.hasTextShapeFocus}
          onClick={() => props.onSetTextAlignment("justify")}
          testId="pptx-text-align-justify"
        />
        <span className="mx-0.5 h-5 w-px bg-divider" aria-hidden />
        <TextAlignToggle
          label="Align text top"
          icon={<AlignVerticalJustifyStart size={14} />}
          active={props.activeTextAnchor === "top"}
          disabled={disabled || !props.hasTextShapeFocus}
          onClick={() => props.onSetTextAnchor("top")}
          testId="pptx-text-anchor-top"
        />
        <TextAlignToggle
          label="Align text middle"
          icon={<AlignVerticalJustifyCenter size={14} />}
          active={props.activeTextAnchor === "middle"}
          disabled={disabled || !props.hasTextShapeFocus}
          onClick={() => props.onSetTextAnchor("middle")}
          testId="pptx-text-anchor-middle"
        />
        <TextAlignToggle
          label="Align text bottom"
          icon={<AlignVerticalJustifyEnd size={14} />}
          active={props.activeTextAnchor === "bottom"}
          disabled={disabled || !props.hasTextShapeFocus}
          onClick={() => props.onSetTextAnchor("bottom")}
          testId="pptx-text-anchor-bottom"
        />
      </div>
      <ColorWell
        label="Fill"
        value={currentFill ?? "#ffffff"}
        showClear={hasSelection && currentFill !== null}
        disabled={disabled || !hasSelection}
        onChange={(hex) => props.onChangeFill(hex)}
        onClear={() => props.onChangeFill(null)}
      />
      <Sep />
      <ToolbarButton
        onClick={props.onAddComment}
        icon={<MessageSquarePlus size={14} />}
        label="Comment"
        title={hasSelection ? "Comment on selected shape" : "Add comment to slide"}
        disabled={disabled}
        testId="pptx-add-comment"
      />
    </ToolbarRow>
  );
}

function Sep() {
  return <span className="mx-2 h-5 w-px bg-divider" />;
}

interface ToolbarButtonProps {
  readonly onClick: () => void;
  readonly icon: React.ReactNode;
  readonly label: string;
  /** Optional tooltip override; defaults to `label`. Use this when the
   * tooltip wording depends on selection state but the visible label
   * must stay fixed-width to keep the toolbar layout stable. */
  readonly title?: string;
  readonly disabled?: boolean;
  readonly testId?: string;
}

function ToolbarButton({ onClick, icon, label, title, disabled, testId }: ToolbarButtonProps) {
  return (
    <button
      type="button"
      onClick={onClick}
      disabled={disabled}
      title={title ?? label}
      data-testid={testId}
      className="inline-flex shrink-0 items-center gap-1.5 rounded px-2 py-1 text-xs text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
    >
      {icon}
      <span className="hidden sm:inline">{label}</span>
    </button>
  );
}

/**
 * Icon-only toolbar button — keeps wide button rows (like the eight
 * align/distribute controls) compact. The label is still surfaced via
 * `title` and `aria-label` so screen readers and tooltips work.
 */
function IconButton({ onClick, icon, label, disabled }: ToolbarButtonProps) {
  return (
    <button
      type="button"
      onClick={onClick}
      disabled={disabled}
      title={label}
      aria-label={label}
      data-testid={`pptx-${label.toLowerCase().replace(/\s+/g, "-")}`}
      className="inline-flex h-7 w-7 items-center justify-center rounded text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
    >
      {icon}
    </button>
  );
}

interface TextAlignToggleProps {
  readonly label: string;
  readonly icon: React.ReactNode;
  readonly active: boolean;
  readonly disabled: boolean;
  readonly onClick: () => void;
  readonly testId: string;
}

/**
 * Pressed-state aware icon button for the text-alignment cluster.
 * Mirrors `IconButton`'s sizing so the row reads as one cohesive
 * group, and renders an active ring when its alignment / anchor
 * matches the currently focused shape.
 */
function TextAlignToggle({ label, icon, active, disabled, onClick, testId }: TextAlignToggleProps) {
  return (
    <button
      type="button"
      onClick={onClick}
      disabled={disabled}
      title={label}
      aria-label={label}
      aria-pressed={active}
      data-testid={testId}
      className={`inline-flex h-7 w-7 items-center justify-center rounded text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40 ${
        active ? "bg-hover ring-1 ring-[var(--accent)]/40" : ""
      }`}
    >
      {icon}
    </button>
  );
}

interface ShapeMenuProps {
  readonly disabled: boolean;
  readonly onPick: (preset: ShapePreset) => void;
}

interface LayoutOption {
  readonly kind: LayoutKindPayload;
  readonly label: string;
}

const LAYOUT_OPTIONS: ReadonlyArray<LayoutOption> = [
  { kind: "title", label: "Title Slide" },
  { kind: "titleAndContent", label: "Title and Content" },
  { kind: "sectionHeader", label: "Section Header" },
  { kind: "twoContent", label: "Two Content" },
  { kind: "comparison", label: "Comparison" },
  { kind: "titleOnly", label: "Title Only" },
  { kind: "blank", label: "Blank" },
  { kind: "contentWithCaption", label: "Content with Caption" },
  { kind: "pictureWithCaption", label: "Picture with Caption" },
  { kind: "titleSlide", label: "Title with Body" },
  { kind: "bigNumber", label: "Big Number" },
];

interface LayoutSketchProps {
  readonly kind: LayoutKindPayload;
  readonly className?: string;
}

/**
 * Tiny 16:9 SVG wireframe used in the Layout dropdown. Each block
 * abstracts a placeholder slot: bars hint at title/text shapes, the
 * dashed rectangles hint at picture/content shapes, the centred numeral
 * hints at the "Big Number" template. The styling intentionally uses
 * theme tokens (`currentColor`) so the sketch tracks light/dark mode
 * without needing per-theme palettes — a plain monochromatic wireframe
 * reads as "rough preview" instead of pretending to be a real render.
 */
function LayoutSketch({ kind, className }: LayoutSketchProps): React.ReactElement {
  const stroke = "currentColor";
  const fill = "currentColor";
  return (
    <svg
      viewBox="0 0 80 45"
      role="presentation"
      aria-hidden="true"
      className={`shrink-0 text-secondary ${className ?? ""}`}
    >
      <rect
        x="0.5"
        y="0.5"
        width="79"
        height="44"
        rx="2"
        fill="var(--surface, #fff)"
        stroke={stroke}
        strokeOpacity="0.45"
        strokeWidth="0.75"
      />
      <g fill={fill} fillOpacity="0.55" stroke={stroke} strokeOpacity="0.35" strokeWidth="0.4">
        {renderLayoutSketchBody(kind)}
      </g>
    </svg>
  );
}

function renderLayoutSketchBody(kind: LayoutKindPayload): React.ReactElement {
  switch (kind) {
    case "title":
      return (
        <>
          <rect x="14" y="17" width="52" height="6" rx="0.5" />
          <rect x="22" y="26" width="36" height="3" rx="0.5" fillOpacity="0.3" />
        </>
      );
    case "titleAndContent":
      return (
        <>
          <rect x="6" y="6" width="52" height="5" rx="0.5" />
          <rect x="6" y="15" width="68" height="24" rx="1" fillOpacity="0.18" strokeDasharray="1.2 1" />
        </>
      );
    case "sectionHeader":
      return (
        <>
          <rect x="6" y="14" width="20" height="2.5" rx="0.5" fillOpacity="0.35" />
          <rect x="6" y="19" width="50" height="7" rx="0.5" />
          <rect x="6" y="29" width="36" height="2.5" rx="0.5" fillOpacity="0.3" />
        </>
      );
    case "twoContent":
      return (
        <>
          <rect x="6" y="6" width="52" height="5" rx="0.5" />
          <rect x="6" y="15" width="32" height="24" rx="1" fillOpacity="0.18" strokeDasharray="1.2 1" />
          <rect x="42" y="15" width="32" height="24" rx="1" fillOpacity="0.18" strokeDasharray="1.2 1" />
        </>
      );
    case "comparison":
      return (
        <>
          <rect x="6" y="6" width="52" height="5" rx="0.5" />
          <rect x="6" y="14" width="32" height="3" rx="0.5" fillOpacity="0.35" />
          <rect x="42" y="14" width="32" height="3" rx="0.5" fillOpacity="0.35" />
          <rect x="6" y="20" width="32" height="19" rx="1" fillOpacity="0.18" strokeDasharray="1.2 1" />
          <rect x="42" y="20" width="32" height="19" rx="1" fillOpacity="0.18" strokeDasharray="1.2 1" />
        </>
      );
    case "titleOnly":
      return <rect x="6" y="6" width="52" height="5" rx="0.5" />;
    case "blank":
      return <></>;
    case "contentWithCaption":
      return (
        <>
          <rect x="6" y="6" width="22" height="4" rx="0.5" />
          <rect x="6" y="13" width="22" height="2" rx="0.5" fillOpacity="0.3" />
          <rect x="6" y="16" width="22" height="2" rx="0.5" fillOpacity="0.3" />
          <rect x="6" y="19" width="18" height="2" rx="0.5" fillOpacity="0.3" />
          <rect x="32" y="6" width="42" height="33" rx="1" fillOpacity="0.18" strokeDasharray="1.2 1" />
        </>
      );
    case "pictureWithCaption":
      return (
        <>
          <rect x="6" y="6" width="68" height="24" rx="1" fillOpacity="0.18" strokeDasharray="1.2 1" />
          <rect x="6" y="33" width="34" height="3" rx="0.5" />
          <rect x="6" y="38" width="50" height="2" rx="0.5" fillOpacity="0.3" />
        </>
      );
    case "titleSlide":
      return (
        <>
          <rect x="6" y="6" width="52" height="5" rx="0.5" />
          <rect x="6" y="15" width="68" height="3" rx="0.5" fillOpacity="0.3" />
          <rect x="6" y="20" width="68" height="3" rx="0.5" fillOpacity="0.3" />
          <rect x="6" y="25" width="62" height="3" rx="0.5" fillOpacity="0.3" />
          <rect x="6" y="30" width="58" height="3" rx="0.5" fillOpacity="0.3" />
        </>
      );
    case "bigNumber":
      return (
        <>
          <text
            x="40"
            y="32"
            textAnchor="middle"
            fontSize="28"
            fontWeight="700"
            fontFamily="ui-sans-serif, system-ui, sans-serif"
            fill="currentColor"
            fillOpacity="0.55"
            stroke="none"
          >
            42
          </text>
          <rect x="22" y="36" width="36" height="2.5" rx="0.5" fillOpacity="0.3" />
        </>
      );
  }
}

interface LayoutMenuProps {
  readonly disabled: boolean;
  readonly canSetLayout: boolean;
  readonly onAddWithLayout: (kind: LayoutKindPayload) => void;
  readonly onSetLayout: (kind: LayoutKindPayload) => void;
}

function LayoutMenu({ disabled, canSetLayout, onAddWithLayout, onSetLayout }: LayoutMenuProps) {
  const [open, setOpen] = React.useState(false);
  const [mode, setMode] = React.useState<"add" | "set">("add");
  const triggerRef = React.useRef<HTMLButtonElement>(null);
  return (
    <>
      <button
        ref={triggerRef}
        type="button"
        disabled={disabled}
        onClick={() => setOpen((v) => !v)}
        title="Layout"
        data-testid="pptx-layout-menu-trigger"
        className="inline-flex items-center gap-1 rounded px-2 py-1 text-xs text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
      >
        <LayoutTemplate size={14} />
        <span className="hidden sm:inline">Layout</span>
        <ChevronDown size={12} />
      </button>
      <ToolbarMenu
        open={open}
        onClose={() => setOpen(false)}
        triggerRef={triggerRef}
        role="menu"
        testId="pptx-layout-menu"
        className="w-[22rem] rounded-md border border-divider bg-surface p-2 shadow-lg"
      >
        <div className="mb-2 inline-flex w-full overflow-hidden rounded border border-divider text-[10px]">
          <button
            type="button"
            onClick={() => setMode("add")}
            className={`flex-1 px-2 py-1 ${mode === "add" ? "bg-hover" : ""}`}
          >
            New slide
          </button>
          <button
            type="button"
            onClick={() => setMode("set")}
            disabled={!canSetLayout}
            className={`flex-1 px-2 py-1 disabled:opacity-40 ${mode === "set" ? "bg-hover" : ""}`}
          >
            Apply to current
          </button>
        </div>
        <div className="grid grid-cols-2 gap-1.5">
          {LAYOUT_OPTIONS.map((opt) => (
            <button
              key={opt.kind}
              type="button"
              role="menuitem"
              data-testid={`pptx-layout-${opt.kind}`}
              onClick={() => {
                setOpen(false);
                if (mode === "add") onAddWithLayout(opt.kind);
                else onSetLayout(opt.kind);
              }}
              title={opt.label}
              className="group flex flex-col items-stretch gap-1 rounded border border-divider/60 bg-transparent p-1.5 text-left text-foreground hover:border-divider hover:bg-hover focus:outline-none focus-visible:ring-1 focus-visible:ring-[var(--accent)]/60"
            >
              <LayoutSketch kind={opt.kind} className="aspect-[16/9] w-full" />
              <span className="truncate text-[11px] leading-tight text-secondary group-hover:text-foreground">
                {opt.label}
              </span>
            </button>
          ))}
        </div>
      </ToolbarMenu>
    </>
  );
}

interface TransitionMenuProps {
  readonly disabled: boolean;
  readonly currentKind: import("@officeai/pptx").TransitionKind;
  readonly currentSpeed: import("@officeai/pptx").TransitionSpeed | null;
  readonly onSetTransition: (
    kind: import("@officeai/pptx").TransitionKind,
    speed: import("@officeai/pptx").TransitionSpeed | null
  ) => void;
}

const TRANSITION_OPTIONS: ReadonlyArray<{
  readonly kind: import("@officeai/pptx").TransitionKind;
  readonly label: string;
  readonly hint: string;
}> = [
  { kind: "none", label: "None", hint: "Hard cut, no animation" },
  { kind: "fade", label: "Fade", hint: "Cross-fade between slides" },
  { kind: "push", label: "Push", hint: "New slide pushes the old one off" },
  { kind: "wipe", label: "Wipe", hint: "Wipe across the slide" },
  { kind: "split", label: "Split", hint: "Split open from the centre" },
  { kind: "cut", label: "Cut", hint: "Instant cut" },
];

const TRANSITION_SPEED_OPTIONS: ReadonlyArray<{
  readonly value: import("@officeai/pptx").TransitionSpeed;
  readonly label: string;
}> = [
  { value: "slow", label: "Slow" },
  { value: "med", label: "Medium" },
  { value: "fast", label: "Fast" },
];

function TransitionMenu({ disabled, currentKind, currentSpeed, onSetTransition }: TransitionMenuProps) {
  const [open, setOpen] = React.useState(false);
  const triggerRef = React.useRef<HTMLButtonElement>(null);
  const activeLabel =
    currentKind === "unsupported"
      ? "Custom"
      : (TRANSITION_OPTIONS.find((o) => o.kind === currentKind)?.label ?? "None");
  return (
    <>
      <button
        ref={triggerRef}
        type="button"
        disabled={disabled}
        onClick={() => setOpen((v) => !v)}
        title={`Slide transition (currently ${activeLabel})`}
        data-testid="pptx-transition-menu-trigger"
        className="inline-flex items-center gap-1 rounded px-2 py-1 text-xs text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
      >
        <Wand2 size={14} />
        <span className="hidden sm:inline">Transition</span>
        {currentKind !== "none" && currentKind !== "unsupported" ? (
          <span className="hidden text-[10px] text-secondary md:inline">· {activeLabel}</span>
        ) : null}
        <ChevronDown size={12} />
      </button>
      <ToolbarMenu
        open={open}
        onClose={() => setOpen(false)}
        triggerRef={triggerRef}
        role="menu"
        testId="pptx-transition-menu"
        className="w-[16rem] rounded-md border border-divider bg-surface p-2 shadow-lg"
      >
        <div className="mb-2 px-1 text-[10px] uppercase tracking-wide text-secondary">Effect</div>
        <div className="mb-2 grid grid-cols-2 gap-1.5">
          {TRANSITION_OPTIONS.map((opt) => {
            const isActive = opt.kind === currentKind;
            return (
              <button
                key={opt.kind}
                type="button"
                role="menuitem"
                data-testid={`pptx-transition-${opt.kind}`}
                onClick={() => {
                  setOpen(false);
                  onSetTransition(opt.kind, currentSpeed ?? "med");
                }}
                title={opt.hint}
                className={`group flex flex-col items-start gap-0.5 rounded border px-2 py-1.5 text-left ${
                  isActive
                    ? "border-[var(--accent)] bg-[color-mix(in_srgb,var(--accent)_8%,transparent)]"
                    : "border-divider/60 bg-transparent hover:border-divider hover:bg-hover"
                }`}
              >
                <span className="text-[11px] font-medium text-foreground">{opt.label}</span>
                <span className="truncate text-[10px] leading-tight text-secondary">{opt.hint}</span>
              </button>
            );
          })}
        </div>
        <div className="mb-1 px-1 text-[10px] uppercase tracking-wide text-secondary">Speed</div>
        <div className="inline-flex w-full overflow-hidden rounded border border-divider text-[10px]">
          {TRANSITION_SPEED_OPTIONS.map((s) => {
            const isActive = (currentSpeed ?? "med") === s.value;
            return (
              <button
                key={s.value}
                type="button"
                disabled={currentKind === "none" || currentKind === "unsupported"}
                onClick={() => onSetTransition(currentKind, s.value)}
                className={`flex-1 px-2 py-1 disabled:opacity-40 ${isActive ? "bg-hover" : ""}`}
                data-testid={`pptx-transition-speed-${s.value}`}
              >
                {s.label}
              </button>
            );
          })}
        </div>
      </ToolbarMenu>
    </>
  );
}

interface ConnectorOption {
  readonly type: "straight" | "elbow" | "curved";
  readonly label: string;
  readonly icon: React.ReactNode;
}

const CONNECTOR_OPTIONS: ReadonlyArray<ConnectorOption> = [
  { type: "straight", label: "Straight", icon: <Minus size={14} /> },
  { type: "elbow", label: "Elbow", icon: <CornerDownRight size={14} /> },
  { type: "curved", label: "Curved", icon: <Spline size={14} /> },
];

interface ConnectorMenuProps {
  readonly disabled: boolean;
  readonly onPick: (type: "straight" | "elbow" | "curved") => void;
  readonly activeType: "straight" | "elbow" | "curved" | null;
}

function ConnectorMenu({ disabled, onPick, activeType }: ConnectorMenuProps) {
  const [open, setOpen] = React.useState(false);
  const triggerRef = React.useRef<HTMLButtonElement>(null);
  // Trigger reads "armed" when the connector tool is currently
  // engaged. We use a tinted background instead of a heavy outline so
  // the toolbar still feels tidy while clearly signalling the active
  // tool — re-clicking the trigger then exits the tool, matching
  // Slides/Figma's toggle pattern.
  const armedClass = activeType ? "bg-hover ring-1 ring-purple-500/60" : "";
  return (
    <>
      <button
        ref={triggerRef}
        type="button"
        disabled={disabled}
        onClick={() => setOpen((v) => !v)}
        title={activeType ? `Drawing ${activeType} connector — Esc to cancel` : "Connector"}
        data-testid="pptx-connector-menu-trigger"
        data-active-type={activeType ?? ""}
        className={`inline-flex items-center gap-1 rounded px-2 py-1 text-xs text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40 ${armedClass}`}
      >
        <Spline size={14} />
        <span className="hidden sm:inline">Connector</span>
        <ChevronDown size={12} />
      </button>
      <ToolbarMenu
        open={open}
        onClose={() => setOpen(false)}
        triggerRef={triggerRef}
        role="menu"
        testId="pptx-connector-menu"
        className="grid w-40 grid-cols-1 gap-0.5 rounded-md border border-divider bg-surface p-1 shadow-lg"
      >
        {CONNECTOR_OPTIONS.map((opt) => {
          const isActive = activeType === opt.type;
          return (
            <button
              key={opt.type}
              type="button"
              role="menuitem"
              data-testid={`pptx-connector-${opt.type}`}
              aria-pressed={isActive}
              onClick={() => {
                setOpen(false);
                onPick(opt.type);
              }}
              className={`flex items-center gap-2 rounded px-2 py-1 text-left text-xs text-foreground hover:bg-hover ${isActive ? "bg-hover" : ""}`}
            >
              {opt.icon}
              <span>{opt.label}</span>
              {isActive ? <span className="ml-auto text-[10px] text-purple-500">●</span> : null}
            </button>
          );
        })}
      </ToolbarMenu>
    </>
  );
}

/**
 * Tiny click-outside dropdown. We deliberately don't pull in a popover
 * library here — the rest of the editor stays dependency-free and the
 * menu is only seven items, so a hand-rolled implementation is cheaper
 * than negotiating a positioning library's portal/z-index quirks.
 */
function ShapeMenu({ disabled, onPick }: ShapeMenuProps) {
  const [open, setOpen] = React.useState(false);
  const triggerRef = React.useRef<HTMLButtonElement>(null);
  return (
    <>
      <button
        ref={triggerRef}
        type="button"
        disabled={disabled}
        onClick={() => setOpen((v) => !v)}
        title="Shape"
        data-testid="pptx-shape-menu-trigger"
        className="inline-flex items-center gap-1 rounded px-2 py-1 text-xs text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
      >
        <Shapes size={14} />
        <span className="hidden sm:inline">Shape</span>
        <ChevronDown size={12} />
      </button>
      <ToolbarMenu
        open={open}
        onClose={() => setOpen(false)}
        triggerRef={triggerRef}
        role="menu"
        testId="pptx-shape-menu"
        className="grid w-40 grid-cols-1 gap-0.5 rounded-md border border-divider bg-surface p-1 shadow-lg"
      >
        {SHAPE_OPTIONS.map((opt) => (
          <button
            key={opt.preset}
            type="button"
            role="menuitem"
            data-testid={`pptx-shape-${opt.preset}`}
            onClick={() => {
              setOpen(false);
              onPick(opt.preset);
            }}
            className="flex items-center gap-2 rounded px-2 py-1 text-left text-xs text-foreground hover:bg-hover"
          >
            {opt.icon}
            <span>{opt.label}</span>
          </button>
        ))}
      </ToolbarMenu>
    </>
  );
}

interface ColorWellProps {
  readonly label: string;
  readonly value: string;
  readonly disabled?: boolean;
  readonly showClear?: boolean;
  readonly onChange: (hex: string) => void;
  readonly onClear?: () => void;
}

function ColorWell({ label, value, disabled, showClear, onChange, onClear }: ColorWellProps) {
  // The clear ("×") affordance is always rendered and uses
  // `visibility: hidden` when irrelevant. That way the toolbar's
  // width doesn't change between "no fill" and "has fill" states
  // (which would push every right-of-here button by ~14 px).
  return (
    <span className="ml-1 inline-flex shrink-0 items-center gap-1" title={label}>
      <label className="sr-only">{label}</label>
      <input
        type="color"
        disabled={disabled}
        value={normaliseColor(value)}
        onChange={(e) => onChange(e.target.value)}
        data-testid={`pptx-color-${label.toLowerCase().replace(/\s+/g, "-")}`}
        className="h-6 w-7 cursor-pointer rounded border border-divider bg-surface p-0 disabled:cursor-not-allowed disabled:opacity-40"
      />
      <button
        type="button"
        onClick={onClear}
        aria-hidden={!showClear}
        tabIndex={showClear ? 0 : -1}
        title={`Clear ${label.toLowerCase()}`}
        className="rounded px-1 text-[10px] text-secondary hover:bg-hover"
        style={{ visibility: showClear ? "visible" : "hidden" }}
      >
        ×
      </button>
    </span>
  );
}

interface ArrangeMenuProps {
  readonly disabled: boolean;
  readonly onReorder: (mode: ReorderShapeMode) => void;
  readonly onGroup: () => void;
  readonly onUngroup: () => void;
  readonly onDuplicateShape: () => void;
  readonly canGroup: boolean;
  readonly canUngroup: boolean;
  readonly canDuplicate: boolean;
}

interface ArrangeOption {
  readonly mode: ReorderShapeMode;
  readonly label: string;
  readonly icon: React.ReactNode;
}

const ARRANGE_OPTIONS: ReadonlyArray<ArrangeOption> = [
  { mode: "to-front", label: "Bring to front", icon: <ChevronsUp size={14} /> },
  { mode: "forward", label: "Bring forward", icon: <MoveUp size={14} /> },
  { mode: "backward", label: "Send backward", icon: <MoveDown size={14} /> },
  { mode: "to-back", label: "Send to back", icon: <ChevronsDown size={14} /> },
];

function ArrangeMenu({
  disabled,
  onReorder,
  onGroup,
  onUngroup,
  onDuplicateShape,
  canGroup,
  canUngroup,
  canDuplicate,
}: ArrangeMenuProps) {
  const [open, setOpen] = React.useState(false);
  const triggerRef = React.useRef<HTMLButtonElement>(null);
  return (
    <>
      <button
        ref={triggerRef}
        type="button"
        disabled={disabled}
        onClick={() => setOpen((v) => !v)}
        title="Arrange"
        data-testid="pptx-arrange-menu-trigger"
        className="inline-flex items-center gap-1 rounded px-2 py-1 text-xs text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
      >
        <Layers size={14} />
        <span className="hidden sm:inline">Arrange</span>
        <ChevronDown size={12} />
      </button>
      <ToolbarMenu
        open={open}
        onClose={() => setOpen(false)}
        triggerRef={triggerRef}
        role="menu"
        testId="pptx-arrange-menu"
        className="grid w-48 grid-cols-1 gap-0.5 rounded-md border border-divider bg-surface p-1 shadow-lg"
      >
        {ARRANGE_OPTIONS.map((opt) => (
          <button
            key={opt.mode}
            type="button"
            role="menuitem"
            data-testid={`pptx-arrange-${opt.mode}`}
            onClick={() => {
              setOpen(false);
              onReorder(opt.mode);
            }}
            className="flex items-center gap-2 rounded px-2 py-1 text-left text-xs text-foreground hover:bg-hover"
          >
            {opt.icon}
            <span>{opt.label}</span>
          </button>
        ))}
        <div className="my-0.5 h-px bg-divider" role="separator" />
        <button
          type="button"
          role="menuitem"
          data-testid="pptx-arrange-group"
          disabled={!canGroup}
          onClick={() => {
            setOpen(false);
            onGroup();
          }}
          className="flex items-center gap-2 rounded px-2 py-1 text-left text-xs text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
        >
          <Group size={14} />
          <span>Group</span>
          <span className="ml-auto text-[10px] text-muted">⇧⌘G</span>
        </button>
        <button
          type="button"
          role="menuitem"
          data-testid="pptx-arrange-ungroup"
          disabled={!canUngroup}
          onClick={() => {
            setOpen(false);
            onUngroup();
          }}
          className="flex items-center gap-2 rounded px-2 py-1 text-left text-xs text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
        >
          <Ungroup size={14} />
          <span>Ungroup</span>
          <span className="ml-auto text-[10px] text-muted">⌥⇧⌘G</span>
        </button>
        <button
          type="button"
          role="menuitem"
          data-testid="pptx-arrange-duplicate"
          disabled={!canDuplicate}
          onClick={() => {
            setOpen(false);
            onDuplicateShape();
          }}
          className="flex items-center gap-2 rounded px-2 py-1 text-left text-xs text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
        >
          <Copy size={14} />
          <span>Duplicate</span>
          <span className="ml-auto text-[10px] text-muted">⌘D</span>
        </button>
      </ToolbarMenu>
    </>
  );
}

interface AlignTargetToggleProps {
  readonly value: "selection" | "slide";
  readonly onChange: (value: "selection" | "slide") => void;
  readonly disabled: boolean;
}

/**
 * Tiny toggle pill (Selected ↔ Slide) controlling the reference frame
 * for the six align icons. Mirrors PowerPoint's "Align to Slide" /
 * "Align Selected Objects" choices in the Arrange menu.
 */
function AlignTargetToggle({ value, onChange, disabled }: AlignTargetToggleProps) {
  const next = value === "selection" ? "slide" : "selection";
  const label = value === "slide" ? "To slide" : "To selection";
  const title =
    value === "slide"
      ? "Aligning to slide bounds. Click to switch to align selected objects."
      : "Aligning to selection. Click to switch to align to slide.";
  return (
    <button
      type="button"
      disabled={disabled}
      onClick={() => onChange(next)}
      title={title}
      aria-pressed={value === "slide"}
      data-testid="pptx-align-target-toggle"
      data-align-relative-to={value}
      className="ml-0.5 inline-flex items-center gap-1 rounded border border-divider px-1.5 py-0.5 text-[10px] uppercase tracking-wide text-muted hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
    >
      {label}
    </button>
  );
}

interface PictureReplaceButtonProps {
  readonly onReplace: (file: File) => void;
  readonly disabled: boolean;
}

/**
 * D9 — contextual "Replace image…" button shown next to the generic
 * Image insert button when a `Picture` shape is selected. Pops the
 * native file picker and forwards the chosen file to the editor,
 * which dispatches `pptx:replace-picture-media`.
 *
 * Always rendered (with `disabled={true}` when no picture is
 * selected) so the toolbar's width never changes when selection
 * flips between picture and non-picture shapes — that previously
 * shifted every button to its right by ~70 px on each click.
 */
function PictureReplaceButton({ onReplace, disabled }: PictureReplaceButtonProps): React.ReactElement {
  const inputRef = React.useRef<HTMLInputElement | null>(null);
  return (
    <>
      <button
        type="button"
        onClick={() => inputRef.current?.click()}
        disabled={disabled}
        title="Replace image"
        data-testid="pptx-replace-image"
        className="inline-flex shrink-0 items-center gap-1 rounded px-2 py-1 text-xs text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
      >
        <ImageIcon size={14} />
        <span className="hidden sm:inline">Replace</span>
      </button>
      <input
        ref={inputRef}
        type="file"
        accept="image/png,image/jpeg,image/gif,image/webp,image/bmp"
        className="hidden"
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) onReplace(f);
          e.target.value = "";
        }}
      />
    </>
  );
}

function normaliseColor(v: string): string {
  // <input type="color"> only accepts 7-char `#rrggbb`, so coerce
  // anything else (uppercase / no hash / shorthand) into that format.
  const cleaned = v.trim().replace(/^#/, "");
  if (/^[0-9a-fA-F]{6}$/.test(cleaned)) return `#${cleaned.toLowerCase()}`;
  if (/^[0-9a-fA-F]{3}$/.test(cleaned)) {
    const [a, b, c] = cleaned.toLowerCase();
    return `#${a}${a}${b}${b}${c}${c}`;
  }
  return "#ffffff";
}
