"use client";

import * as React from "react";
import { ArrowRight, Circle, Minus, MoreHorizontal, Spline, Triangle, X } from "lucide-react";
import type {
  ConnectorDashStyle,
  ConnectorEndShape,
  ConnectorShape,
  ConnectorType,
} from "@officeai/pptx";

/**
 * Floating "format" mini-bar shown when a single connector is
 * selected. Keeps the controls users actually reach for (line type,
 * weight, dash, head/tail style, color) one click away — without
 * opening the full toolbar Format panel.
 *
 * Wiring contract:
 *   - `connector` is the live shape from the agent snapshot. We read
 *     defaults off it on every render so undo / redo / external edits
 *     re-flow into the bar's pickers automatically.
 *   - Every change produces a single `pptx:set-connector-style`
 *     command via `onPatch`. The editor owns the agent reference; the
 *     bar is intentionally agent-agnostic so the same component can
 *     be reused in a future preview pane / compare-view.
 *   - Position is computed by the parent (typically anchored to the
 *     selected shape's bbox via `slideToClient` math) and passed in
 *     as `style.transform` on the wrapper. The bar itself only owns
 *     its internal layout.
 */
export interface ConnectorContextBarProps {
  readonly connector: ConnectorShape;
  /**
   * Apply a partial style update. The implementation should dispatch
   * `pptx:set-connector-style` for the connector currently in focus.
   * Returning a promise is fine — the bar treats the call as fire-
   * and-forget; errors propagate via the editor's existing toast
   * channel.
   */
  readonly onPatch: (patch: ConnectorStylePatch) => void;
  /** Optional positioning style applied to the wrapper. */
  readonly style?: React.CSSProperties;
}

export interface ConnectorStylePatch {
  readonly connectorType?: ConnectorType;
  readonly strokeColor?: string;
  readonly strokeWidthEmu?: number;
  readonly strokeDash?: ConnectorDashStyle;
  readonly headEnd?: ConnectorEndShape;
  readonly tailEnd?: ConnectorEndShape;
}

const TYPE_OPTIONS: ReadonlyArray<{ type: ConnectorType; label: string; icon: React.ReactNode }> = [
  { type: "straight", label: "Straight", icon: <Minus size={14} /> },
  { type: "elbow", label: "Elbow", icon: <Spline size={14} /> },
  { type: "curved", label: "Curved", icon: <Spline size={14} className="rotate-45" /> },
];

const DASH_OPTIONS: ReadonlyArray<{ dash: ConnectorDashStyle; label: string }> = [
  { dash: "solid", label: "Solid" },
  { dash: "dashed", label: "Dashed" },
  { dash: "dotted", label: "Dotted" },
];

const WIDTH_OPTIONS: ReadonlyArray<{ widthEmu: number; label: string; previewPx: number }> = [
  { widthEmu: 6_350, label: "0.5 pt", previewPx: 1 },
  { widthEmu: 9_525, label: "0.75 pt", previewPx: 1 },
  { widthEmu: 12_700, label: "1 pt", previewPx: 2 },
  { widthEmu: 19_050, label: "1.5 pt", previewPx: 2 },
  { widthEmu: 25_400, label: "2 pt", previewPx: 3 },
  { widthEmu: 38_100, label: "3 pt", previewPx: 4 },
  { widthEmu: 57_150, label: "4.5 pt", previewPx: 5 },
];

const END_OPTIONS: ReadonlyArray<{
  end: ConnectorEndShape;
  label: string;
  icon: React.ReactNode;
}> = [
  { end: "none", label: "None", icon: <X size={12} /> },
  { end: "arrow", label: "Arrow", icon: <ArrowRight size={12} /> },
  { end: "triangle", label: "Triangle", icon: <Triangle size={12} /> },
  { end: "oval", label: "Oval", icon: <Circle size={12} /> },
];

const COLOR_PALETTE: ReadonlyArray<string> = [
  "374151",
  "111827",
  "0ea5e9",
  "7c3aed",
  "ef4444",
  "f97316",
  "10b981",
  "facc15",
];

export function ConnectorContextBar({
  connector,
  onPatch,
  style,
}: ConnectorContextBarProps): React.ReactElement {
  const stroke = connector.stroke;
  const currentColor = stroke?.color ?? "374151";
  const currentWidth = stroke?.widthEmu ?? 9_525;
  const currentDash: ConnectorDashStyle = stroke?.dash ?? "solid";
  const currentHead = connector.headEnd ?? "none";
  const currentTail = connector.tailEnd ?? "none";

  return (
    <div
      data-testid="pptx-connector-context-bar"
      className="flex items-center gap-1 rounded-md border border-divider bg-surface px-1.5 py-1 shadow-md"
      style={style}
      onMouseDown={(e) => {
        // Prevent the canvas from clearing selection / starting a
        // marquee when the user clicks one of our pickers.
        e.stopPropagation();
      }}
    >
      <SegmentedTypePicker
        value={connector.connectorType}
        onChange={(t) => onPatch({ connectorType: t })}
      />
      <Sep />
      <WidthPicker valueEmu={currentWidth} onChange={(w) => onPatch({ strokeWidthEmu: w })} />
      <DashPicker value={currentDash} onChange={(d) => onPatch({ strokeDash: d })} />
      <Sep />
      <EndPicker
        kind="tail"
        value={currentTail}
        onChange={(e) => onPatch({ tailEnd: e })}
      />
      <EndPicker
        kind="head"
        value={currentHead}
        onChange={(e) => onPatch({ headEnd: e })}
      />
      <Sep />
      <ColorPicker value={currentColor} onChange={(c) => onPatch({ strokeColor: c })} />
    </div>
  );
}

function Sep(): React.ReactElement {
  return <div className="mx-0.5 h-5 w-px bg-divider" />;
}

interface SegmentedTypePickerProps {
  readonly value: ConnectorType;
  readonly onChange: (next: ConnectorType) => void;
}

function SegmentedTypePicker({ value, onChange }: SegmentedTypePickerProps): React.ReactElement {
  return (
    <div className="flex items-center gap-0.5 rounded bg-background/40 p-0.5">
      {TYPE_OPTIONS.map((opt) => {
        const active = value === opt.type;
        return (
          <button
            key={opt.type}
            type="button"
            title={opt.label}
            data-testid={`pptx-connector-bar-type-${opt.type}`}
            aria-pressed={active}
            onClick={() => onChange(opt.type)}
            className={`flex h-6 w-7 items-center justify-center rounded text-foreground hover:bg-hover ${active ? "bg-hover ring-1 ring-purple-500/60" : ""}`}
          >
            {opt.icon}
          </button>
        );
      })}
    </div>
  );
}

interface WidthPickerProps {
  readonly valueEmu: number;
  readonly onChange: (next: number) => void;
}

function WidthPicker({ valueEmu, onChange }: WidthPickerProps): React.ReactElement {
  const [open, setOpen] = React.useState(false);
  const ref = React.useRef<HTMLDivElement>(null);
  React.useEffect(() => {
    if (!open) return;
    function onDoc(e: MouseEvent) {
      if (!ref.current?.contains(e.target as Node)) setOpen(false);
    }
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, [open]);
  const closest = WIDTH_OPTIONS.reduce((acc, o) =>
    Math.abs(o.widthEmu - valueEmu) < Math.abs(acc.widthEmu - valueEmu) ? o : acc
  );
  return (
    <div ref={ref} className="relative">
      <button
        type="button"
        data-testid="pptx-connector-bar-weight"
        onClick={() => setOpen((v) => !v)}
        className="flex h-6 items-center gap-1 rounded px-1.5 text-xs text-foreground hover:bg-hover"
        title="Stroke weight"
      >
        <span
          aria-hidden
          className="block w-5 rounded-sm bg-foreground"
          style={{ height: closest.previewPx }}
        />
        <span className="text-[10px] tabular-nums text-secondary">{closest.label}</span>
      </button>
      {open ? (
        <div
          role="menu"
          data-testid="pptx-connector-bar-weight-menu"
          className="absolute left-0 top-full z-30 mt-1 w-32 rounded-md border border-divider bg-surface p-1 shadow-lg"
        >
          {WIDTH_OPTIONS.map((opt) => (
            <button
              key={opt.widthEmu}
              type="button"
              data-testid={`pptx-connector-bar-weight-${opt.widthEmu}`}
              onClick={() => {
                onChange(opt.widthEmu);
                setOpen(false);
              }}
              className={`flex w-full items-center gap-2 rounded px-2 py-1 text-left text-xs text-foreground hover:bg-hover ${opt.widthEmu === closest.widthEmu ? "bg-hover" : ""}`}
            >
              <span
                aria-hidden
                className="block w-10 rounded-sm bg-foreground"
                style={{ height: opt.previewPx }}
              />
              <span className="text-[10px] tabular-nums text-secondary">{opt.label}</span>
            </button>
          ))}
        </div>
      ) : null}
    </div>
  );
}

interface DashPickerProps {
  readonly value: ConnectorDashStyle;
  readonly onChange: (next: ConnectorDashStyle) => void;
}

function DashPicker({ value, onChange }: DashPickerProps): React.ReactElement {
  return (
    <div className="flex items-center gap-0.5 rounded bg-background/40 p-0.5">
      {DASH_OPTIONS.map((opt) => {
        const active = value === opt.dash;
        // Render a tiny line preview matching the dash style so the
        // picker is recognisable without a tooltip on every hover.
        const dashStyle =
          opt.dash === "dashed" ? "4 3" : opt.dash === "dotted" ? "1 2" : undefined;
        return (
          <button
            key={opt.dash}
            type="button"
            title={opt.label}
            data-testid={`pptx-connector-bar-dash-${opt.dash}`}
            aria-pressed={active}
            onClick={() => onChange(opt.dash)}
            className={`flex h-6 w-7 items-center justify-center rounded hover:bg-hover ${active ? "bg-hover ring-1 ring-purple-500/60" : ""}`}
          >
            <svg width={18} height={6} aria-hidden>
              <line
                x1={1}
                y1={3}
                x2={17}
                y2={3}
                stroke="currentColor"
                strokeWidth={1.5}
                strokeDasharray={dashStyle}
              />
            </svg>
          </button>
        );
      })}
    </div>
  );
}

interface EndPickerProps {
  readonly kind: "head" | "tail";
  readonly value: ConnectorEndShape;
  readonly onChange: (next: ConnectorEndShape) => void;
}

function EndPicker({ kind, value, onChange }: EndPickerProps): React.ReactElement {
  const [open, setOpen] = React.useState(false);
  const ref = React.useRef<HTMLDivElement>(null);
  React.useEffect(() => {
    if (!open) return;
    function onDoc(e: MouseEvent) {
      if (!ref.current?.contains(e.target as Node)) setOpen(false);
    }
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, [open]);
  const current = END_OPTIONS.find((o) => o.end === value) ?? END_OPTIONS[0];
  return (
    <div ref={ref} className="relative">
      <button
        type="button"
        data-testid={`pptx-connector-bar-${kind}`}
        title={`${kind === "head" ? "Head (end)" : "Tail (start)"} marker`}
        onClick={() => setOpen((v) => !v)}
        className="flex h-6 w-8 items-center justify-center rounded text-foreground hover:bg-hover"
      >
        {kind === "tail" ? <span className="rotate-180">{current.icon}</span> : current.icon}
      </button>
      {open ? (
        <div
          role="menu"
          className="absolute left-0 top-full z-30 mt-1 w-28 rounded-md border border-divider bg-surface p-1 shadow-lg"
        >
          {END_OPTIONS.map((opt) => (
            <button
              key={opt.end}
              type="button"
              data-testid={`pptx-connector-bar-${kind}-${opt.end}`}
              onClick={() => {
                onChange(opt.end);
                setOpen(false);
              }}
              className={`flex w-full items-center gap-2 rounded px-2 py-1 text-left text-xs text-foreground hover:bg-hover ${opt.end === current.end ? "bg-hover" : ""}`}
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

interface ColorPickerProps {
  readonly value: string;
  readonly onChange: (next: string) => void;
}

function ColorPicker({ value, onChange }: ColorPickerProps): React.ReactElement {
  const [open, setOpen] = React.useState(false);
  const ref = React.useRef<HTMLDivElement>(null);
  React.useEffect(() => {
    if (!open) return;
    function onDoc(e: MouseEvent) {
      if (!ref.current?.contains(e.target as Node)) setOpen(false);
    }
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, [open]);
  return (
    <div ref={ref} className="relative">
      <button
        type="button"
        data-testid="pptx-connector-bar-color"
        title="Stroke color"
        onClick={() => setOpen((v) => !v)}
        className="flex h-6 w-8 items-center justify-center rounded hover:bg-hover"
      >
        <span
          aria-hidden
          className="block h-4 w-4 rounded-full border border-divider"
          style={{ backgroundColor: `#${value}` }}
        />
        <MoreHorizontal size={10} className="ml-0.5 text-secondary" />
      </button>
      {open ? (
        <div
          role="menu"
          data-testid="pptx-connector-bar-color-menu"
          className="absolute left-0 top-full z-30 mt-1 w-40 rounded-md border border-divider bg-surface p-2 shadow-lg"
        >
          <div className="grid grid-cols-4 gap-1">
            {COLOR_PALETTE.map((hex) => (
              <button
                key={hex}
                type="button"
                data-testid={`pptx-connector-bar-color-${hex}`}
                onClick={() => {
                  onChange(hex);
                  setOpen(false);
                }}
                title={`#${hex}`}
                className={`block h-6 w-6 rounded border ${hex === value ? "border-purple-500 ring-1 ring-purple-500" : "border-divider"}`}
                style={{ backgroundColor: `#${hex}` }}
              />
            ))}
          </div>
          <label className="mt-2 flex items-center gap-1 text-[10px] text-secondary">
            Custom
            <input
              type="color"
              value={`#${value}`}
              onChange={(e) => onChange(e.target.value.replace(/^#/, "").toLowerCase())}
              data-testid="pptx-connector-bar-color-custom"
              className="h-6 w-full"
            />
          </label>
        </div>
      ) : null}
    </div>
  );
}
