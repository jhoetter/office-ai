"use client";

/**
 * Office-style fill picker for slide backgrounds and shape fills.
 *
 * Renders a popover-driven UI that mirrors PowerPoint's "Format
 * Background" / "Format Shape" pane. Sections:
 *   • None    — clear the fill (`{ type: "none" }`).
 *   • Solid   — theme + standard color rows, custom hex input, alpha
 *               slider (the rightmost theme row is reserved for the
 *               eight Office accent colours; the standard row matches
 *               the colour grid the rest of the app already uses for
 *               text).
 *   • Gradient — kind toggle (linear / radial), angle dial for linear,
 *                multi-stop editor with add/remove/per-stop colour and
 *                position controls.
 *
 * Pattern + Picture sections are stubbed out and rendered behind a
 * "Coming soon" hint — the data model already supports them, so the
 * follow-up session can drop the UI in without further command/render
 * work.
 *
 * The component is controlled: pass `value` (a `FillSpec | null`, where
 * `null` means "no explicit fill — inherits") and react to `onChange`.
 * Validation is performed by the caller (the underlying command
 * normalises every spec via `normaliseFillSpec`).
 */

import * as React from "react";
import { ChevronDown, X } from "lucide-react";
import { Popover } from "@officeai/ui";
import type { FillSpec, GradientFillSpec, GradientStop } from "@officeai/pptx";

// ─── Palettes ─────────────────────────────────────────────────────────────

interface Swatch {
  readonly name: string;
  /** Lowercase RRGGBB without `#`. */
  readonly hex: string;
}

/**
 * Office "Theme Colors" row — eight accents the toolbar exposes as the
 * top row. We keep them as literal hex (vs `schemeClr` references) so
 * the picker always paints predictably regardless of which theme the
 * deck uses; the underlying command stores the literal hex too.
 */
const THEME_SWATCHES: ReadonlyArray<Swatch> = [
  { name: "White", hex: "ffffff" },
  { name: "Black", hex: "000000" },
  { name: "Tan", hex: "eeece1" },
  { name: "Dark Blue 2", hex: "1f497d" },
  { name: "Accent 1", hex: "4f81bd" },
  { name: "Accent 2", hex: "c0504d" },
  { name: "Accent 3", hex: "9bbb59" },
  { name: "Accent 4", hex: "8064a2" },
  { name: "Accent 5", hex: "4bacc6" },
  { name: "Accent 6", hex: "f79646" },
];

/** Office "Standard Colors" row — the ten classic preset hues. */
const STANDARD_SWATCHES: ReadonlyArray<Swatch> = [
  { name: "Dark Red", hex: "c00000" },
  { name: "Red", hex: "ff0000" },
  { name: "Orange", hex: "ffc000" },
  { name: "Yellow", hex: "ffff00" },
  { name: "Light Green", hex: "92d050" },
  { name: "Green", hex: "00b050" },
  { name: "Light Blue", hex: "00b0f0" },
  { name: "Blue", hex: "0070c0" },
  { name: "Dark Blue", hex: "002060" },
  { name: "Purple", hex: "7030a0" },
];

/** A friendly two-stop linear gradient used as the "default" gradient. */
const DEFAULT_GRADIENT: GradientFillSpec = {
  type: "gradient",
  kind: "linear",
  angleDeg: 90,
  stops: [
    { pos: 0, color: "ffffff" },
    { pos: 1, color: "1f497d" },
  ],
};

// ─── Public API ───────────────────────────────────────────────────────────

export interface FillPickerProps {
  readonly label: string;
  readonly value: FillSpec | null;
  /**
   * Called when the user picks a new fill. `null` means "clear" —
   * the underlying shape inherits its style (or, for slide
   * backgrounds, the layout/master).
   */
  readonly onChange: (next: FillSpec | null) => void;
  readonly disabled?: boolean;
  /** Hide the trailing "×" clear button. Default: shown. */
  readonly hideClear?: boolean;
  /** test-id prefix for E2E hooks. */
  readonly testId?: string;
  /**
   * Trigger style. `"chip"` shows a small colour swatch + chevron
   * (toolbar default). `"button"` shows a labelled button (used by
   * the Design tab where we want a wider hit target).
   */
  readonly variant?: "chip" | "button";
}

type Section = "solid" | "gradient" | "pattern" | "picture" | "none";

export function FillPicker({
  label,
  value,
  onChange,
  disabled,
  hideClear,
  testId,
  variant = "chip",
}: FillPickerProps): React.ReactNode {
  const stripeStyle = useStripeStyle(value);
  const initialSection: Section = (value?.type ?? "solid") as Section;

  return (
    <span className="inline-flex shrink-0 items-center gap-1" title={label}>
      <label className="sr-only">{label}</label>
      <Popover
        panelClassName="w-[260px] !p-3"
        trigger={
          <button
            type="button"
            disabled={disabled}
            data-testid={testId ?? "pptx-fill-trigger"}
            onMouseDown={(e) => e.preventDefault()}
            className={
              variant === "button"
                ? "inline-flex items-center gap-1.5 rounded border border-divider px-2 py-1 text-xs text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
                : "inline-flex items-center gap-0.5 rounded border border-divider bg-surface px-1.5 py-0.5 hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
            }
          >
            <span
              aria-hidden
              className="block h-4 w-5 rounded-sm border border-divider"
              style={stripeStyle}
            />
            {variant === "button" && <span className="text-xs">{label}</span>}
            <ChevronDown size={10} className="text-secondary" />
          </button>
        }
      >
        <FillPanel initialSection={initialSection} value={value} onChange={onChange} testId={testId} />
      </Popover>
      {!hideClear && (
        <button
          type="button"
          onClick={() => onChange(null)}
          aria-hidden={value === null}
          tabIndex={value === null ? -1 : 0}
          title={`Clear ${label.toLowerCase()}`}
          className="rounded px-1 text-[10px] text-secondary hover:bg-hover"
          style={{ visibility: value === null ? "hidden" : "visible" }}
          data-testid={testId ? `${testId}-clear` : undefined}
        >
          <X size={12} />
        </button>
      )}
    </span>
  );
}

// ─── Panel ────────────────────────────────────────────────────────────────

interface FillPanelProps {
  readonly initialSection: Section;
  readonly value: FillSpec | null;
  readonly onChange: (next: FillSpec | null) => void;
  readonly testId?: string;
}

function FillPanel({ initialSection, value, onChange, testId }: FillPanelProps): React.ReactNode {
  const [section, setSection] = React.useState<Section>(initialSection);

  // When the user opens the popover after the value changed externally,
  // jump to the section matching the new value.
  React.useEffect(() => {
    if (!value) return;
    setSection(value.type as Section);
  }, [value]);

  return (
    <div data-popover-close="false" className="flex flex-col gap-2">
      <SectionTabs current={section} onSelect={setSection} testId={testId} />
      <div className="border-t border-divider" />
      {section === "none" && <NoneSection onChange={onChange} />}
      {section === "solid" && (
        <SolidSection
          value={value?.type === "solid" ? value : undefined}
          onChange={(spec) => onChange(spec)}
          testId={testId}
        />
      )}
      {section === "gradient" && (
        <GradientSection
          value={value?.type === "gradient" ? value : DEFAULT_GRADIENT}
          onChange={(spec) => onChange(spec)}
          testId={testId}
        />
      )}
      {section === "pattern" && <ComingSoonSection label="Pattern fill" />}
      {section === "picture" && <ComingSoonSection label="Picture fill" />}
    </div>
  );
}

interface SectionTabsProps {
  readonly current: Section;
  readonly onSelect: (s: Section) => void;
  readonly testId?: string;
}

function SectionTabs({ current, onSelect, testId }: SectionTabsProps): React.ReactNode {
  const tabs: ReadonlyArray<{ id: Section; label: string }> = [
    { id: "solid", label: "Solid" },
    { id: "gradient", label: "Gradient" },
    { id: "pattern", label: "Pattern" },
    { id: "picture", label: "Picture" },
    { id: "none", label: "None" },
  ];
  return (
    <div role="tablist" aria-label="Fill type" className="flex flex-wrap gap-1">
      {tabs.map((t) => (
        <button
          key={t.id}
          type="button"
          role="tab"
          aria-selected={current === t.id}
          onClick={() => onSelect(t.id)}
          data-testid={testId ? `${testId}-tab-${t.id}` : undefined}
          className={`rounded px-2 py-0.5 text-xs ${
            current === t.id
              ? "bg-[var(--accent)]/15 text-foreground ring-1 ring-[var(--accent)]/40"
              : "text-secondary hover:bg-hover hover:text-foreground"
          }`}
        >
          {t.label}
        </button>
      ))}
    </div>
  );
}

// ─── Sections ─────────────────────────────────────────────────────────────

function NoneSection({ onChange }: { onChange: (v: FillSpec | null) => void }): React.ReactNode {
  return (
    <div className="flex flex-col gap-2 py-1 text-xs text-secondary">
      <p>No fill — the shape will render as transparent.</p>
      <div className="flex gap-1">
        <button
          type="button"
          onClick={() => onChange({ type: "none" })}
          className="rounded border border-divider px-2 py-1 text-foreground hover:bg-hover"
        >
          Apply No Fill
        </button>
        <button
          type="button"
          onClick={() => onChange(null)}
          className="rounded border border-divider px-2 py-1 text-foreground hover:bg-hover"
          title="Clear the explicit fill (inherit from layout / style)"
        >
          Reset (inherit)
        </button>
      </div>
    </div>
  );
}

interface SolidSectionProps {
  readonly value: Extract<FillSpec, { type: "solid" }> | undefined;
  readonly onChange: (spec: FillSpec) => void;
  readonly testId?: string;
}

function SolidSection({ value, onChange, testId }: SolidSectionProps): React.ReactNode {
  const color = value?.color?.toLowerCase() ?? "ffffff";
  const alpha = value?.alpha ?? 1;
  const setColor = (hex: string) => onChange({ type: "solid", color: hex, alpha: alpha < 1 ? alpha : undefined });
  const setAlpha = (a: number) =>
    onChange({ type: "solid", color, alpha: a >= 1 ? undefined : a });
  return (
    <div className="flex flex-col gap-2">
      <SwatchRow label="Theme Colors" swatches={THEME_SWATCHES} active={color} onPick={setColor} />
      <SwatchRow label="Standard Colors" swatches={STANDARD_SWATCHES} active={color} onPick={setColor} />
      <CustomColorRow color={color} onChange={setColor} testId={testId ? `${testId}-solid` : undefined} />
      <AlphaRow alpha={alpha} onChange={setAlpha} />
    </div>
  );
}

interface GradientSectionProps {
  readonly value: GradientFillSpec;
  readonly onChange: (spec: GradientFillSpec) => void;
  readonly testId?: string;
}

function GradientSection({ value, onChange, testId }: GradientSectionProps): React.ReactNode {
  const updateStop = (i: number, patch: Partial<GradientStop>) => {
    const stops = value.stops.map((s, j) => (j === i ? { ...s, ...patch } : s));
    onChange({ ...value, stops });
  };
  const addStop = () => {
    if (value.stops.length >= 10) return;
    const last = value.stops[value.stops.length - 1];
    const second = value.stops[value.stops.length - 2] ?? last;
    const pos = Math.min(1, (last.pos + (last.pos - second.pos || 0.25)) / 1);
    onChange({ ...value, stops: [...value.stops, { pos: Math.min(1, pos), color: last.color }] });
  };
  const removeStop = (i: number) => {
    if (value.stops.length <= 2) return;
    onChange({ ...value, stops: value.stops.filter((_, j) => j !== i) });
  };

  return (
    <div className="flex flex-col gap-2">
      <div className="flex items-center gap-2">
        <span className="text-[10px] uppercase tracking-wide text-secondary">Type</span>
        <div className="inline-flex rounded border border-divider">
          {(["linear", "radial"] as const).map((kind) => (
            <button
              key={kind}
              type="button"
              onClick={() => onChange({ ...value, kind })}
              data-testid={testId ? `${testId}-grad-kind-${kind}` : undefined}
              className={`px-2 py-0.5 text-xs ${
                value.kind === kind
                  ? "bg-[var(--accent)]/15 text-foreground"
                  : "text-secondary hover:bg-hover"
              }`}
            >
              {kind === "linear" ? "Linear" : "Radial"}
            </button>
          ))}
        </div>
      </div>

      {value.kind === "linear" && (
        <div className="flex items-center gap-2">
          <span className="text-[10px] uppercase tracking-wide text-secondary">Angle</span>
          <input
            type="range"
            min={0}
            max={359}
            value={value.angleDeg}
            onChange={(e) => onChange({ ...value, angleDeg: Number(e.target.value) })}
            data-testid={testId ? `${testId}-grad-angle` : undefined}
            className="flex-1"
          />
          <span className="w-10 text-right text-xs tabular-nums text-secondary">{value.angleDeg}°</span>
        </div>
      )}

      <GradientPreview spec={value} />

      <div className="flex flex-col gap-1.5">
        <div className="flex items-center justify-between">
          <span className="text-[10px] uppercase tracking-wide text-secondary">Stops</span>
          <button
            type="button"
            onClick={addStop}
            disabled={value.stops.length >= 10}
            className="rounded border border-divider px-1.5 py-0.5 text-[11px] text-foreground hover:bg-hover disabled:opacity-40"
          >
            + Add
          </button>
        </div>
        {value.stops.map((stop, i) => (
          <div key={i} className="flex items-center gap-1.5">
            <input
              type="color"
              value={`#${stop.color}`}
              onChange={(e) => updateStop(i, { color: e.target.value.replace(/^#/, "") })}
              data-testid={testId ? `${testId}-grad-stop-${i}-color` : undefined}
              className="h-6 w-7 cursor-pointer rounded border border-divider bg-surface p-0"
            />
            <input
              type="range"
              min={0}
              max={100}
              value={Math.round(stop.pos * 100)}
              onChange={(e) => updateStop(i, { pos: Number(e.target.value) / 100 })}
              data-testid={testId ? `${testId}-grad-stop-${i}-pos` : undefined}
              className="flex-1"
            />
            <span className="w-9 text-right text-[11px] tabular-nums text-secondary">
              {Math.round(stop.pos * 100)}%
            </span>
            <button
              type="button"
              onClick={() => removeStop(i)}
              disabled={value.stops.length <= 2}
              title="Remove stop"
              className="rounded px-1 text-[11px] text-secondary hover:bg-hover disabled:opacity-30"
            >
              <X size={11} />
            </button>
          </div>
        ))}
      </div>
    </div>
  );
}

function ComingSoonSection({ label }: { label: string }): React.ReactNode {
  return (
    <div className="rounded border border-dashed border-divider px-3 py-4 text-center text-[11px] text-secondary">
      {label} is coming in a future update.
    </div>
  );
}

// ─── Building blocks ──────────────────────────────────────────────────────

interface SwatchRowProps {
  readonly label: string;
  readonly swatches: ReadonlyArray<Swatch>;
  readonly active: string;
  readonly onPick: (hex: string) => void;
}

function SwatchRow({ label, swatches, active, onPick }: SwatchRowProps): React.ReactNode {
  return (
    <div className="flex flex-col gap-1">
      <span className="text-[10px] uppercase tracking-wide text-secondary">{label}</span>
      <div className="flex flex-wrap gap-1">
        {swatches.map((s) => {
          const isActive = active.toLowerCase() === s.hex.toLowerCase();
          return (
            <button
              key={s.hex}
              type="button"
              onClick={() => onPick(s.hex)}
              title={s.name}
              aria-label={`${label}: ${s.name}`}
              className={`h-5 w-5 rounded-sm border ${
                isActive ? "border-[var(--accent)] ring-1 ring-[var(--accent)]" : "border-divider"
              } hover:scale-110`}
              style={{ background: `#${s.hex}` }}
            />
          );
        })}
      </div>
    </div>
  );
}

interface CustomColorRowProps {
  readonly color: string;
  readonly onChange: (hex: string) => void;
  readonly testId?: string;
}

function CustomColorRow({ color, onChange, testId }: CustomColorRowProps): React.ReactNode {
  return (
    <div className="flex items-center gap-2">
      <span className="text-[10px] uppercase tracking-wide text-secondary">Custom</span>
      <input
        type="color"
        value={`#${color}`}
        onChange={(e) => onChange(e.target.value.replace(/^#/, "").toLowerCase())}
        data-testid={testId ? `${testId}-custom-color` : undefined}
        className="h-6 w-8 cursor-pointer rounded border border-divider bg-surface p-0"
      />
      <span className="text-xs uppercase tabular-nums text-secondary">#{color}</span>
    </div>
  );
}

interface AlphaRowProps {
  readonly alpha: number;
  readonly onChange: (alpha: number) => void;
}

function AlphaRow({ alpha, onChange }: AlphaRowProps): React.ReactNode {
  const pct = Math.round(alpha * 100);
  return (
    <div className="flex items-center gap-2">
      <span className="text-[10px] uppercase tracking-wide text-secondary">Opacity</span>
      <input
        type="range"
        min={0}
        max={100}
        value={pct}
        onChange={(e) => onChange(Number(e.target.value) / 100)}
        className="flex-1"
      />
      <span className="w-9 text-right text-[11px] tabular-nums text-secondary">{pct}%</span>
    </div>
  );
}

interface GradientPreviewProps {
  readonly spec: GradientFillSpec;
}

function GradientPreview({ spec }: GradientPreviewProps): React.ReactNode {
  const css = gradientToCss(spec);
  return (
    <div
      className="h-8 w-full rounded border border-divider"
      style={{ background: css }}
      aria-label="Gradient preview"
    />
  );
}

// ─── CSS preview ──────────────────────────────────────────────────────────

function gradientToCss(spec: GradientFillSpec): string {
  const stops = spec.stops
    .slice()
    .sort((a, b) => a.pos - b.pos)
    .map((s) => `rgba(${hexToRgb(s.color)}, ${s.alpha ?? 1}) ${(s.pos * 100).toFixed(1)}%`)
    .join(", ");
  if (spec.kind === "linear") {
    return `linear-gradient(${spec.angleDeg}deg, ${stops})`;
  }
  return `radial-gradient(circle at center, ${stops})`;
}

function hexToRgb(hex: string): string {
  const v = hex.padStart(6, "0");
  const r = parseInt(v.slice(0, 2), 16);
  const g = parseInt(v.slice(2, 4), 16);
  const b = parseInt(v.slice(4, 6), 16);
  return `${r}, ${g}, ${b}`;
}

function useStripeStyle(value: FillSpec | null): React.CSSProperties {
  return React.useMemo(() => {
    // We never combine the `background` shorthand with `backgroundImage`
    // here — React warns about shorthand/longhand mixing because the
    // shorthand resets the longhand on rerender. For alpha previews we
    // therefore set `backgroundColor` (longhand) alongside the checker
    // `backgroundImage`. For full-opacity solids and gradients a single
    // shorthand is fine.
    if (!value) return { backgroundColor: "transparent", backgroundImage: checkerImage() };
    switch (value.type) {
      case "none":
        return { backgroundColor: "transparent", backgroundImage: checkerImage() };
      case "solid": {
        const opacity = value.alpha ?? 1;
        if (opacity >= 1) return { background: `#${value.color}` };
        return {
          backgroundColor: `rgba(${hexToRgb(value.color)}, ${opacity})`,
          backgroundImage: checkerImage(),
          backgroundBlendMode: "normal",
        };
      }
      case "gradient":
        return { background: gradientToCss(value) };
      case "pattern":
        return { background: `#${value.fgColor}` };
      case "picture":
        return { background: "#dddddd" };
      default:
        return { background: "transparent" };
    }
  }, [value]);
}

/**
 * 6-pixel checker pattern used to indicate transparency / alpha. Inline
 * SVG so the component carries no external asset dependency.
 */
function checkerImage(): string {
  const svg =
    '<svg xmlns="http://www.w3.org/2000/svg" width="8" height="8"><rect width="8" height="8" fill="#fff"/><rect width="4" height="4" fill="#ddd"/><rect x="4" y="4" width="4" height="4" fill="#ddd"/></svg>';
  return `url("data:image/svg+xml;utf8,${encodeURIComponent(svg)}")`;
}
