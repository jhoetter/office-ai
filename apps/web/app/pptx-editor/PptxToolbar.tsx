"use client";

import * as React from "react";
import {
  Bold,
  ChevronDown,
  Circle,
  Copy,
  Diamond,
  Download,
  FileUp,
  Image as ImageIcon,
  Italic,
  Minus,
  MoveRight,
  Plus,
  Shapes,
  Square,
  Trash2,
  Triangle,
  Type,
  Underline,
  ZoomIn,
  ZoomOut,
} from "lucide-react";
import type { ShapePreset } from "@officeai/pptx";

export interface PptxToolbarProps {
  readonly disabled: boolean;
  readonly slideCount: number;
  readonly hasSelection: boolean;
  readonly currentFill: string | null;
  readonly currentFontPt: number | null;
  readonly onOpenFile: () => void;
  readonly onExport: () => void;
  readonly onAddSlide: () => void;
  readonly onDeleteSlide: () => void;
  readonly onDuplicateSlide: () => void;
  readonly onAddTextBox: () => void;
  readonly onAddShape: (preset: ShapePreset) => void;
  readonly onInsertImage: (file: File) => void;
  readonly onDeleteShape: () => void;
  readonly onToggleBold: () => void;
  readonly onToggleItalic: () => void;
  readonly onToggleUnderline: () => void;
  readonly onChangeFill: (hex: string | null) => void;
  readonly onChangeTextColor: (hex: string) => void;
  readonly onChangeFontSize: (pt: number) => void;
  readonly zoom: number;
  readonly minZoom: number;
  readonly maxZoom: number;
  readonly onZoomChange: (zoom: number) => void;
  readonly onZoomReset: () => void;
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

const FONT_SIZES: ReadonlyArray<number> = [8, 10, 12, 14, 16, 18, 20, 24, 28, 32, 36, 44, 54, 66, 80, 96];

export function PptxToolbar(props: PptxToolbarProps) {
  const {
    disabled,
    hasSelection,
    currentFill,
    currentFontPt,
    zoom,
    minZoom,
    maxZoom,
    onZoomChange,
    onZoomReset,
  } = props;
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
      <Sep />
      <ToolbarButton
        onClick={props.onAddTextBox}
        icon={<Type size={14} />}
        label="Text box"
        disabled={disabled}
      />
      <ShapeMenu disabled={disabled} onPick={props.onAddShape} />
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
      <ToolbarButton
        onClick={props.onToggleBold}
        icon={<Bold size={14} />}
        label="Bold"
        disabled={disabled}
      />
      <ToolbarButton
        onClick={props.onToggleItalic}
        icon={<Italic size={14} />}
        label="Italic"
        disabled={disabled}
      />
      <ToolbarButton
        onClick={props.onToggleUnderline}
        icon={<Underline size={14} />}
        label="Underline"
        disabled={disabled}
      />
      <FontSizePicker disabled={disabled} valuePt={currentFontPt} onChange={props.onChangeFontSize} />
      <ColorWell
        label="Text color"
        // Text color always represents the next-applied color; no current-state preview.
        value="#111111"
        disabled={disabled}
        onChange={(hex) => props.onChangeTextColor(hex)}
      />
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
}

function ToolbarButton({ onClick, icon, label, disabled }: ToolbarButtonProps) {
  return (
    <button
      type="button"
      onClick={onClick}
      disabled={disabled}
      title={label}
      className="inline-flex items-center gap-1.5 rounded px-2 py-1 text-xs text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
    >
      {icon}
      <span className="hidden sm:inline">{label}</span>
    </button>
  );
}

interface ShapeMenuProps {
  readonly disabled: boolean;
  readonly onPick: (preset: ShapePreset) => void;
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

interface FontSizePickerProps {
  readonly disabled: boolean;
  readonly valuePt: number | null;
  readonly onChange: (pt: number) => void;
}

function FontSizePicker({ disabled, valuePt, onChange }: FontSizePickerProps) {
  const initial = valuePt != null ? String(Math.round(valuePt)) : "";
  const [draft, setDraft] = React.useState<string>(initial);
  // Reset the local draft whenever the external selection switches to a
  // different font size. Done inline (per React docs: "Adjusting state
  // when a prop changes") rather than in an effect so we don't trip
  // `react-hooks/set-state-in-effect`. We track the *prop* (not the
  // derived string) so callers that change the rounded value reset us
  // even if `valuePt` is a different unrounded number.
  const [prevPt, setPrevPt] = React.useState(valuePt);
  if (valuePt !== prevPt) {
    setPrevPt(valuePt);
    setDraft(initial);
  }
  const commit = React.useCallback(() => {
    const n = Number(draft);
    if (Number.isFinite(n) && n >= 1 && n <= 400) onChange(n);
  }, [draft, onChange]);
  return (
    <div className="ml-1 flex items-center">
      <input
        type="number"
        min={1}
        max={400}
        step={1}
        list="pptx-font-sizes"
        disabled={disabled}
        value={draft}
        onChange={(e) => setDraft(e.target.value)}
        onBlur={commit}
        onKeyDown={(e) => {
          if (e.key === "Enter") {
            e.preventDefault();
            commit();
            (e.currentTarget as HTMLInputElement).blur();
          }
        }}
        title="Font size (pt)"
        data-testid="pptx-font-size"
        className="h-7 w-14 rounded border border-divider bg-surface px-1 text-xs text-foreground disabled:cursor-not-allowed disabled:opacity-40"
        placeholder="pt"
      />
      <datalist id="pptx-font-sizes">
        {FONT_SIZES.map((s) => (
          <option key={s} value={s} />
        ))}
      </datalist>
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
