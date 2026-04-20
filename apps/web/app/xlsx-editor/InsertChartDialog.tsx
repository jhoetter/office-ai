"use client";

import { useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import { useFocusTrap } from "@officeai/ui";
import type { ChartKind, ChartPalette, SheetChart } from "@officeai/xlsx";
import {
  normalizeRangeForStorage,
  parseChartRangeShape,
  pickToggleDefaults,
  validateChartShape,
} from "./chartShape";

/**
 * Result of submitting the dialog. Carries the *user-intended* shape;
 * callers translate to either `xlsx:add-chart` (insert mode) or
 * `xlsx:update-chart` (edit mode). `title` is `undefined` when the
 * field is empty so insert can omit it; for edit, `XlsxEditor`
 * remaps that to `null` so the update handler clears the title.
 *
 * Style fields follow the same convention: `undefined` for cleared
 * axis titles (so insert omits, edit clears via `null`); the booleans
 * and palette are always reported because the dialog always exposes
 * an explicit toggle / picker.
 */
export interface ChartDialogSubmit {
  readonly kind: ChartKind;
  readonly dataRange: string;
  readonly title?: string;
  readonly hasHeaderRow: boolean;
  readonly hasCategoryColumn: boolean;
  readonly palette: ChartPalette;
  readonly showLegend: boolean;
  readonly showDataLabels: boolean;
  readonly showGridlines: boolean;
  readonly xAxisTitle?: string;
  readonly yAxisTitle?: string;
}

interface ChartDialogProps {
  readonly open: boolean;
  /** `"insert"` opens with selection-derived defaults; `"edit"` pre-fills from `initial`. */
  readonly mode: "insert" | "edit";
  /** Default range when `mode === "insert"`. Ignored in edit mode. */
  readonly defaultRange?: string;
  readonly defaultKind?: ChartKind;
  /** Existing chart, required when `mode === "edit"`. */
  readonly initial?: SheetChart;
  readonly onCancel: () => void;
  readonly onSubmit: (args: ChartDialogSubmit) => void;
}

interface InsertChartDialogProps {
  readonly open: boolean;
  readonly defaultRange: string;
  readonly defaultKind?: ChartKind;
  readonly onCancel: () => void;
  readonly onSubmit: (args: ChartDialogSubmit) => void;
}

const CHART_KINDS: ReadonlyArray<{ readonly kind: ChartKind; readonly label: string }> = [
  { kind: "column", label: "Column" },
  { kind: "bar", label: "Bar" },
  { kind: "line", label: "Line" },
  { kind: "pie", label: "Pie" },
];

/**
 * Palette swatches shown in the Appearance section. Each swatch
 * paints the first three colors of the cycle so the picker is
 * comparable at a glance. Swatch hex values intentionally mirror
 * the renderer's `PALETTES` map; if you adjust one, adjust the
 * other.
 */
const PALETTE_OPTIONS: ReadonlyArray<{
  readonly id: ChartPalette;
  readonly label: string;
  readonly swatch: ReadonlyArray<string>;
}> = [
  { id: "default", label: "Default", swatch: ["#5b8def", "#f6c34a", "#7bc274"] },
  { id: "vibrant", label: "Vibrant", swatch: ["#2563eb", "#dc2626", "#16a34a"] },
  { id: "pastel", label: "Pastel", swatch: ["#a8c5f0", "#fbd38d", "#9ed5a6"] },
  { id: "warm", label: "Warm", swatch: ["#f97316", "#ef4444", "#fbbf24"] },
  { id: "cool", label: "Cool", swatch: ["#0ea5e9", "#10b981", "#6366f1"] },
  { id: "mono", label: "Mono", swatch: ["#1f2937", "#4b5563", "#9ca3af"] },
];

/**
 * Insert / Edit chart dialog (C15 + Bug B fix).
 *
 *   - Pick a chart type (column / bar / line / pie).
 *   - Confirm the data range (defaults to the active selection on
 *     insert, to the chart's stored range on edit).
 *   - Toggle "First row is header" / "First column is categories".
 *   - Optional title.
 *
 * Submitting calls `onSubmit` with the canonical shape; the parent
 * decides whether to dispatch `xlsx:add-chart` or `xlsx:update-chart`.
 * In insert mode the chart appears anchored just to the right of the
 * source range (handler default); in edit mode the existing anchor
 * is preserved.
 */
export function ChartDialog(props: ChartDialogProps): ReactNode {
  const { open, mode, defaultRange, defaultKind, initial, onCancel, onSubmit } = props;
  const seedRange =
    mode === "edit" ? (initial?.dataRange ?? defaultRange ?? "A1:B5") : (defaultRange ?? "A1:B5");
  const seedKind: ChartKind = mode === "edit" ? (initial?.kind ?? "column") : (defaultKind ?? "column");
  const seedTitle = mode === "edit" ? (initial?.title ?? "") : "";

  const [kind, setKind] = useState<ChartKind>(seedKind);
  const [range, setRange] = useState(seedRange);
  const [title, setTitle] = useState(seedTitle);
  const [hasHeaderRow, setHasHeaderRow] = useState(initial?.hasHeaderRow ?? true);
  const [hasCategoryColumn, setHasCategoryColumn] = useState(initial?.hasCategoryColumn ?? true);
  const [palette, setPalette] = useState<ChartPalette>(initial?.palette ?? "default");
  const [showLegend, setShowLegend] = useState(initial?.showLegend !== false);
  const [showDataLabels, setShowDataLabels] = useState(initial?.showDataLabels === true);
  const [showGridlines, setShowGridlines] = useState(initial?.showGridlines !== false);
  const [xAxisTitle, setXAxisTitle] = useState(initial?.xAxisTitle ?? "");
  const [yAxisTitle, setYAxisTitle] = useState(initial?.yAxisTitle ?? "");
  const panelRef = useRef<HTMLDivElement>(null);
  useFocusTrap(panelRef, { enabled: open, onEscape: onCancel });

  useEffect(() => {
    if (!open) return;
    if (mode === "edit" && initial) {
      setRange(initial.dataRange);
      setKind(initial.kind);
      setTitle(initial.title ?? "");
      setHasHeaderRow(initial.hasHeaderRow);
      setHasCategoryColumn(initial.hasCategoryColumn);
      setPalette(initial.palette ?? "default");
      setShowLegend(initial.showLegend !== false);
      setShowDataLabels(initial.showDataLabels === true);
      setShowGridlines(initial.showGridlines !== false);
      setXAxisTitle(initial.xAxisTitle ?? "");
      setYAxisTitle(initial.yAxisTitle ?? "");
      return;
    }
    const fallback = defaultRange ?? "A1:B5";
    setRange(fallback);
    setKind(defaultKind ?? "column");
    setTitle("");
    const defaults = pickToggleDefaults(parseChartRangeShape(fallback));
    setHasHeaderRow(defaults.hasHeaderRow);
    setHasCategoryColumn(defaults.hasCategoryColumn);
    setPalette("default");
    setShowLegend(true);
    setShowDataLabels(false);
    setShowGridlines(true);
    setXAxisTitle("");
    setYAxisTitle("");
  }, [open, mode, defaultRange, defaultKind, initial]);

  const shape = useMemo(() => parseChartRangeShape(range), [range]);
  const validation = useMemo(
    () => validateChartShape(range, shape, hasHeaderRow, hasCategoryColumn),
    [range, shape, hasHeaderRow, hasCategoryColumn]
  );

  if (!open) return null;

  const canSubmit = validation.kind === "ok";
  const isEdit = mode === "edit";
  const heading = isEdit ? "Edit chart" : "Insert chart";
  const submitLabel = isEdit ? "Save" : "Insert";
  const ariaLabel = isEdit ? "Edit chart" : "Insert chart";
  const testId = isEdit ? "edit-chart-dialog" : "insert-chart-dialog";

  const submit = () => {
    if (!canSubmit) return;
    onSubmit({
      kind,
      dataRange: normalizeRangeForStorage(range),
      title: title.trim() ? title.trim() : undefined,
      hasHeaderRow,
      hasCategoryColumn,
      palette,
      showLegend,
      showDataLabels,
      showGridlines,
      xAxisTitle: xAxisTitle.trim() ? xAxisTitle.trim() : undefined,
      yAxisTitle: yAxisTitle.trim() ? yAxisTitle.trim() : undefined,
    });
  };

  const hint = (() => {
    switch (validation.kind) {
      case "ok":
      case "empty":
        return null;
      case "invalid":
        return `"${range.trim()}" is not a valid range (e.g. A1:B7).`;
      case "single-cell":
        return "Pick a range of at least two cells.";
      case "no-values":
        if (validation.axis === "column") {
          return 'Selection has no value cells with these toggles — turn off "First column contains category labels".';
        }
        return 'Selection has no value cells with these toggles — turn off "First row contains series labels".';
    }
  })();

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label={ariaLabel}
      data-testid={testId}
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(0,0,0,0.35)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        zIndex: 1000,
      }}
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onCancel();
      }}
    >
      <div
        ref={panelRef}
        tabIndex={-1}
        style={{
          width: 460,
          maxHeight: "90vh",
          overflowY: "auto",
          background: "var(--background)",
          color: "var(--foreground)",
          border: "1px solid var(--border)",
          borderRadius: 8,
          boxShadow: "0 12px 32px rgba(0,0,0,0.18)",
          padding: 20,
          display: "flex",
          flexDirection: "column",
          gap: 14,
          outline: "none",
        }}
      >
        <div style={{ fontSize: 14, fontWeight: 600 }}>{heading}</div>

        <div>
          <div style={{ fontSize: 12, fontWeight: 500, marginBottom: 6 }}>Chart type</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 6 }}>
            {CHART_KINDS.map((entry) => (
              <button
                key={entry.kind}
                type="button"
                onClick={() => setKind(entry.kind)}
                style={{
                  padding: "8px 0",
                  border: entry.kind === kind ? "1.5px solid var(--ai-violet)" : "1px solid var(--border)",
                  borderRadius: 6,
                  background: entry.kind === kind ? "var(--ai-violet-light)" : "var(--background)",
                  color: "var(--foreground)",
                  fontSize: 12,
                  cursor: "pointer",
                }}
              >
                {entry.label}
              </button>
            ))}
          </div>
        </div>

        <label style={{ display: "flex", flexDirection: "column", gap: 4 }}>
          <span style={{ fontSize: 12, fontWeight: 500 }}>Data range</span>
          <input
            type="text"
            value={range}
            onChange={(e) => setRange(e.target.value)}
            placeholder="A1:B7"
            style={{
              padding: "6px 8px",
              border: "1px solid var(--border)",
              borderRadius: 4,
              fontSize: 12,
              fontFamily: "monospace",
              background: "var(--background)",
              color: "var(--foreground)",
            }}
            autoFocus
          />
        </label>

        <label style={{ display: "flex", flexDirection: "column", gap: 4 }}>
          <span style={{ fontSize: 12, fontWeight: 500 }}>Title (optional)</span>
          <input
            type="text"
            value={title}
            onChange={(e) => setTitle(e.target.value)}
            placeholder="e.g. Q4 sales"
            style={{
              padding: "6px 8px",
              border: "1px solid var(--border)",
              borderRadius: 4,
              fontSize: 12,
              background: "var(--background)",
              color: "var(--foreground)",
            }}
          />
        </label>

        <div style={{ display: "flex", flexDirection: "column", gap: 6, fontSize: 12 }}>
          <label style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <input
              type="checkbox"
              checked={hasHeaderRow}
              onChange={(e) => setHasHeaderRow(e.target.checked)}
            />
            First row contains series labels
          </label>
          <label style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <input
              type="checkbox"
              checked={hasCategoryColumn}
              onChange={(e) => setHasCategoryColumn(e.target.checked)}
            />
            First column contains category labels
          </label>
        </div>

        <div
          data-testid="chart-dialog-appearance"
          style={{
            display: "flex",
            flexDirection: "column",
            gap: 10,
            paddingTop: 10,
            borderTop: "1px solid var(--border)",
          }}
        >
          <div style={{ fontSize: 12, fontWeight: 600, color: "var(--muted-foreground)" }}>Appearance</div>

          <div>
            <div style={{ fontSize: 12, fontWeight: 500, marginBottom: 6 }}>Color palette</div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 6 }}>
              {PALETTE_OPTIONS.map((opt) => {
                const active = palette === opt.id;
                return (
                  <button
                    key={opt.id}
                    type="button"
                    data-testid={`chart-palette-${opt.id}`}
                    aria-pressed={active}
                    onClick={() => setPalette(opt.id)}
                    style={{
                      padding: "6px 8px",
                      border: active ? "1.5px solid var(--ai-violet)" : "1px solid var(--border)",
                      borderRadius: 6,
                      background: active ? "var(--ai-violet-light)" : "var(--background)",
                      color: "var(--foreground)",
                      fontSize: 11,
                      cursor: "pointer",
                      display: "flex",
                      flexDirection: "column",
                      alignItems: "center",
                      gap: 4,
                    }}
                  >
                    <span style={{ display: "flex", gap: 2 }}>
                      {opt.swatch.map((color, i) => (
                        <span
                          key={`${opt.id}-${i}`}
                          style={{
                            width: 14,
                            height: 14,
                            borderRadius: 3,
                            background: color,
                            border: "1px solid rgba(0,0,0,0.06)",
                          }}
                        />
                      ))}
                    </span>
                    <span>{opt.label}</span>
                  </button>
                );
              })}
            </div>
          </div>

          <div style={{ display: "flex", flexDirection: "column", gap: 6, fontSize: 12 }}>
            <label style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <input type="checkbox" checked={showLegend} onChange={(e) => setShowLegend(e.target.checked)} />
              Show legend
            </label>
            <label style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <input
                type="checkbox"
                checked={showDataLabels}
                onChange={(e) => setShowDataLabels(e.target.checked)}
              />
              Show data labels
            </label>
            <label style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <input
                type="checkbox"
                checked={showGridlines}
                onChange={(e) => setShowGridlines(e.target.checked)}
              />
              Show gridlines
            </label>
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
            <label style={{ display: "flex", flexDirection: "column", gap: 4 }}>
              <span style={{ fontSize: 12, fontWeight: 500 }}>X axis title</span>
              <input
                type="text"
                value={xAxisTitle}
                onChange={(e) => setXAxisTitle(e.target.value)}
                placeholder="e.g. Quarter"
                style={{
                  padding: "6px 8px",
                  border: "1px solid var(--border)",
                  borderRadius: 4,
                  fontSize: 12,
                  background: "var(--background)",
                  color: "var(--foreground)",
                }}
              />
            </label>
            <label style={{ display: "flex", flexDirection: "column", gap: 4 }}>
              <span style={{ fontSize: 12, fontWeight: 500 }}>Y axis title</span>
              <input
                type="text"
                value={yAxisTitle}
                onChange={(e) => setYAxisTitle(e.target.value)}
                placeholder="e.g. Revenue"
                style={{
                  padding: "6px 8px",
                  border: "1px solid var(--border)",
                  borderRadius: 4,
                  fontSize: 12,
                  background: "var(--background)",
                  color: "var(--foreground)",
                }}
              />
            </label>
          </div>
        </div>

        {hint ? (
          <div
            data-testid="chart-dialog-hint"
            role="status"
            style={{
              fontSize: 11,
              color: "var(--destructive, #c0392b)",
              background: "var(--surface)",
              border: "1px solid var(--border)",
              borderRadius: 4,
              padding: "6px 8px",
              lineHeight: 1.4,
            }}
          >
            {hint}
          </div>
        ) : null}

        <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
          <button
            type="button"
            onClick={onCancel}
            style={{
              padding: "6px 12px",
              border: "1px solid var(--border)",
              borderRadius: 4,
              background: "var(--background)",
              color: "var(--foreground)",
              fontSize: 12,
              cursor: "pointer",
            }}
          >
            Cancel
          </button>
          <button
            type="button"
            onClick={submit}
            disabled={!canSubmit}
            style={{
              padding: "6px 12px",
              border: "1px solid var(--ai-violet)",
              borderRadius: 4,
              background: canSubmit ? "var(--ai-violet)" : "var(--muted, #b8b8c4)",
              color: "white",
              fontSize: 12,
              cursor: canSubmit ? "pointer" : "not-allowed",
              opacity: canSubmit ? 1 : 0.7,
            }}
          >
            {submitLabel}
          </button>
        </div>
      </div>
    </div>
  );
}

/**
 * Back-compat wrapper for the original `InsertChartDialog` callers
 * (XlsxEditor's command-palette integration). Just delegates to
 * `ChartDialog` with `mode="insert"`.
 */
export function InsertChartDialog(props: InsertChartDialogProps): ReactNode {
  return (
    <ChartDialog
      open={props.open}
      mode="insert"
      defaultRange={props.defaultRange}
      defaultKind={props.defaultKind}
      onCancel={props.onCancel}
      onSubmit={props.onSubmit}
    />
  );
}
