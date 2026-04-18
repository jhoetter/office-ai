"use client";

import * as React from "react";
import {
  AlignCenterHorizontal,
  AlignCenterVertical,
  AlignEndHorizontal,
  AlignEndVertical,
  AlignHorizontalDistributeCenter,
  AlignStartHorizontal,
  AlignStartVertical,
  AlignVerticalDistributeCenter,
  ChevronDown,
  Circle,
  Copy,
  CornerDownRight,
  Diamond,
  Download,
  FileUp,
  Image as ImageIcon,
  Keyboard,
  Minus,
  MoveRight,
  Plus,
  Shapes,
  Spline,
  Square,
  Trash2,
  Triangle,
  Type,
  ZoomIn,
  ZoomOut,
} from "lucide-react";
import { TextFormatBar } from "@officeai/ui";
import type { ActiveTextFormat, TextFormatProvider } from "@officeai/text-formatting";
import type { AlignMode, LayoutKindPayload, ShapePreset } from "@officeai/pptx";
import { LayoutTemplate } from "lucide-react";

export interface PptxToolbarProps {
  readonly disabled: boolean;
  readonly slideCount: number;
  readonly hasSelection: boolean;
  /** Total number of selected shapes; drives align/distribute enablement. */
  readonly selectionCount: number;
  readonly currentFill: string | null;
  readonly textFormatProvider: TextFormatProvider;
  readonly textFormatActive: ActiveTextFormat;
  readonly onOpenFile: () => void;
  readonly onExport: () => void;
  readonly onAddSlide: () => void;
  readonly onAddSlideWithLayout: (kind: LayoutKindPayload) => void;
  readonly onSetSlideLayout: (kind: LayoutKindPayload) => void;
  readonly onDeleteSlide: () => void;
  readonly onDuplicateSlide: () => void;
  readonly onAddTextBox: () => void;
  readonly onAddShape: (preset: ShapePreset) => void;
  readonly onAddConnector: (connectorType: "straight" | "elbow" | "curved") => void;
  readonly onInsertImage: (file: File) => void;
  readonly onDeleteShape: () => void;
  readonly onAlign: (mode: AlignMode) => void;
  readonly onDistribute: (axis: "horizontal" | "vertical") => void;
  readonly onChangeFill: (hex: string | null) => void;
  readonly zoom: number;
  readonly minZoom: number;
  readonly maxZoom: number;
  readonly onZoomChange: (zoom: number) => void;
  readonly onZoomReset: () => void;
  /** Open the keyboard-shortcuts help dialog. */
  readonly onOpenShortcuts: () => void;
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
  const {
    disabled,
    hasSelection,
    selectionCount,
    currentFill,
    textFormatProvider,
    textFormatActive,
    zoom,
    minZoom,
    maxZoom,
    onZoomChange,
    onZoomReset,
  } = props;
  const canAlign = !disabled && selectionCount >= 2;
  const canDistribute = !disabled && selectionCount >= 3;
  const zoomPct = Math.round(zoom * 100);
  const imageInputRef = React.useRef<HTMLInputElement | null>(null);
  return (
    <div className="flex flex-wrap items-center gap-1 border-b border-divider pb-2">
      <ToolbarButton onClick={props.onOpenFile} icon={<FileUp size={14} />} label="Open" />
      <ToolbarButton
        onClick={props.onExport}
        icon={<Download size={14} />}
        label="Export"
        disabled={disabled}
      />
      <Sep />
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
      <Sep />
      <ToolbarButton
        onClick={props.onAddTextBox}
        icon={<Type size={14} />}
        label="Text box"
        disabled={disabled}
      />
      <ShapeMenu disabled={disabled} onPick={props.onAddShape} />
      <ConnectorMenu disabled={disabled} onPick={props.onAddConnector} />
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
      <ToolbarButton
        onClick={props.onDeleteShape}
        icon={<Trash2 size={14} />}
        label="Delete"
        disabled={disabled || !hasSelection}
      />
      <Sep />
      <IconButton
        onClick={() => props.onAlign("left")}
        icon={<AlignStartVertical size={14} />}
        label="Align left"
        disabled={!canAlign}
      />
      <IconButton
        onClick={() => props.onAlign("center-h")}
        icon={<AlignCenterVertical size={14} />}
        label="Align center"
        disabled={!canAlign}
      />
      <IconButton
        onClick={() => props.onAlign("right")}
        icon={<AlignEndVertical size={14} />}
        label="Align right"
        disabled={!canAlign}
      />
      <IconButton
        onClick={() => props.onAlign("top")}
        icon={<AlignStartHorizontal size={14} />}
        label="Align top"
        disabled={!canAlign}
      />
      <IconButton
        onClick={() => props.onAlign("middle-v")}
        icon={<AlignCenterHorizontal size={14} />}
        label="Align middle"
        disabled={!canAlign}
      />
      <IconButton
        onClick={() => props.onAlign("bottom")}
        icon={<AlignEndHorizontal size={14} />}
        label="Align bottom"
        disabled={!canAlign}
      />
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
      <ColorWell
        label="Fill"
        value={currentFill ?? "#ffffff"}
        showClear={hasSelection && currentFill !== null}
        disabled={disabled || !hasSelection}
        onChange={(hex) => props.onChangeFill(hex)}
        onClear={() => props.onChangeFill(null)}
      />
      <Sep />
      <div className="ml-auto flex items-center gap-2 text-xs text-secondary">
        <ToolbarButton
          onClick={() => onZoomChange(Math.max(minZoom, zoom - 0.1))}
          icon={<ZoomOut size={14} />}
          label="Zoom out"
          disabled={disabled || zoom <= minZoom + 1e-6}
        />
        <input
          type="range"
          aria-label="Zoom"
          data-testid="pptx-zoom-slider"
          min={minZoom}
          max={maxZoom}
          step={0.05}
          value={zoom}
          disabled={disabled}
          onChange={(e) => onZoomChange(Number(e.target.value))}
          className="h-1 w-28 accent-[var(--accent)]"
        />
        <ToolbarButton
          onClick={() => onZoomChange(Math.min(maxZoom, zoom + 0.1))}
          icon={<ZoomIn size={14} />}
          label="Zoom in"
          disabled={disabled || zoom >= maxZoom - 1e-6}
        />
        <button
          type="button"
          onClick={onZoomReset}
          disabled={disabled}
          title="Reset zoom"
          data-testid="pptx-zoom-reset"
          className="min-w-[44px] rounded px-1.5 py-0.5 text-xs tabular-nums text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
        >
          {zoomPct}%
        </button>
        <ToolbarButton
          onClick={props.onOpenShortcuts}
          icon={<Keyboard size={14} />}
          label="Keyboard shortcuts (⌘/)"
          testId="pptx-shortcuts"
        />
      </div>
    </div>
  );
}

function Sep() {
  return <span className="mx-2 h-5 w-px bg-divider" />;
}

interface ToolbarButtonProps {
  readonly onClick: () => void;
  readonly icon: React.ReactNode;
  readonly label: string;
  readonly disabled?: boolean;
  readonly testId?: string;
}

function ToolbarButton({ onClick, icon, label, disabled, testId }: ToolbarButtonProps) {
  return (
    <button
      type="button"
      onClick={onClick}
      disabled={disabled}
      title={label}
      data-testid={testId}
      className="inline-flex items-center gap-1.5 rounded px-2 py-1 text-xs text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
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

interface LayoutMenuProps {
  readonly disabled: boolean;
  readonly canSetLayout: boolean;
  readonly onAddWithLayout: (kind: LayoutKindPayload) => void;
  readonly onSetLayout: (kind: LayoutKindPayload) => void;
}

function LayoutMenu({ disabled, canSetLayout, onAddWithLayout, onSetLayout }: LayoutMenuProps) {
  const [open, setOpen] = React.useState(false);
  const [mode, setMode] = React.useState<"add" | "set">("add");
  const ref = React.useRef<HTMLDivElement>(null);
  React.useEffect(() => {
    if (!open) return;
    function onDoc(e: MouseEvent) {
      if (!ref.current) return;
      if (!ref.current.contains(e.target as Node)) setOpen(false);
    }
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, [open]);
  return (
    <div ref={ref} className="relative">
      <button
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
      {open ? (
        <div
          role="menu"
          data-testid="pptx-layout-menu"
          className="absolute left-0 top-full z-30 mt-1 w-64 rounded-md border border-divider bg-surface p-1 shadow-lg"
        >
          <div className="mb-1 inline-flex w-full overflow-hidden rounded border border-divider text-[10px]">
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
          <div className="grid grid-cols-1 gap-0.5">
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
                className="flex items-center gap-2 rounded px-2 py-1 text-left text-xs text-foreground hover:bg-hover"
              >
                <LayoutTemplate size={14} />
                <span>{opt.label}</span>
              </button>
            ))}
          </div>
        </div>
      ) : null}
    </div>
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
}

function ConnectorMenu({ disabled, onPick }: ConnectorMenuProps) {
  const [open, setOpen] = React.useState(false);
  const ref = React.useRef<HTMLDivElement>(null);
  React.useEffect(() => {
    if (!open) return;
    function onDoc(e: MouseEvent) {
      if (!ref.current) return;
      if (!ref.current.contains(e.target as Node)) setOpen(false);
    }
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, [open]);
  return (
    <div ref={ref} className="relative">
      <button
        type="button"
        disabled={disabled}
        onClick={() => setOpen((v) => !v)}
        title="Connector"
        data-testid="pptx-connector-menu-trigger"
        className="inline-flex items-center gap-1 rounded px-2 py-1 text-xs text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
      >
        <Spline size={14} />
        <span className="hidden sm:inline">Connector</span>
        <ChevronDown size={12} />
      </button>
      {open ? (
        <div
          role="menu"
          data-testid="pptx-connector-menu"
          className="absolute left-0 top-full z-30 mt-1 grid w-40 grid-cols-1 gap-0.5 rounded-md border border-divider bg-surface p-1 shadow-lg"
        >
          {CONNECTOR_OPTIONS.map((opt) => (
            <button
              key={opt.type}
              type="button"
              role="menuitem"
              data-testid={`pptx-connector-${opt.type}`}
              onClick={() => {
                setOpen(false);
                onPick(opt.type);
              }}
              className="flex items-center gap-2 rounded px-2 py-1 text-left text-xs text-foreground hover:bg-hover"
            >
              {opt.icon}
              <span>{opt.label}</span>
            </button>
          ))}
        </div>
      ) : null}
    </div>
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
  const ref = React.useRef<HTMLDivElement>(null);
  React.useEffect(() => {
    if (!open) return;
    function onDoc(e: MouseEvent) {
      if (!ref.current) return;
      if (!ref.current.contains(e.target as Node)) setOpen(false);
    }
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, [open]);
  return (
    <div ref={ref} className="relative">
      <button
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
      {open ? (
        <div
          role="menu"
          data-testid="pptx-shape-menu"
          className="absolute left-0 top-full z-30 mt-1 grid w-40 grid-cols-1 gap-0.5 rounded-md border border-divider bg-surface p-1 shadow-lg"
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
        </div>
      ) : null}
    </div>
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
  return (
    <span className="ml-1 inline-flex items-center gap-1" title={label}>
      <label className="sr-only">{label}</label>
      <input
        type="color"
        disabled={disabled}
        value={normaliseColor(value)}
        onChange={(e) => onChange(e.target.value)}
        data-testid={`pptx-color-${label.toLowerCase().replace(/\s+/g, "-")}`}
        className="h-6 w-7 cursor-pointer rounded border border-divider bg-surface p-0 disabled:cursor-not-allowed disabled:opacity-40"
      />
      {showClear ? (
        <button
          type="button"
          onClick={onClear}
          title={`Clear ${label.toLowerCase()}`}
          className="rounded px-1 text-[10px] text-secondary hover:bg-hover"
        >
          ×
        </button>
      ) : null}
    </span>
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
