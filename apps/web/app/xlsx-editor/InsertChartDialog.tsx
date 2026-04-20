"use client";

import { useEffect, useRef, useState, type ReactNode } from "react";
import { useFocusTrap } from "@officeai/ui";
import type { ChartKind } from "@officeai/xlsx";
import { useTranslator } from "@/lib/i18n";

interface InsertChartDialogProps {
  readonly open: boolean;
  readonly defaultRange: string;
  readonly defaultKind?: ChartKind;
  readonly onCancel: () => void;
  readonly onSubmit: (args: {
    readonly kind: ChartKind;
    readonly dataRange: string;
    readonly title?: string;
    readonly hasHeaderRow: boolean;
    readonly hasCategoryColumn: boolean;
  }) => void;
}

const CHART_KINDS: ReadonlyArray<{ readonly kind: ChartKind; readonly labelKey: string }> = [
  { kind: "column", labelKey: "xlsx.insertChart.column" },
  { kind: "bar", labelKey: "xlsx.insertChart.bar" },
  { kind: "line", labelKey: "xlsx.insertChart.line" },
  { kind: "pie", labelKey: "xlsx.insertChart.pie" },
];

/**
 * Insert Chart dialog (C15) — Excel parity:
 *   - Pick a chart type (column / bar / line / pie).
 *   - Confirm the data range (defaults to the active selection).
 *   - Toggle "First row is header" / "First column is categories".
 *   - Optional title.
 *
 * Submitting dispatches `xlsx:add-chart`; the chart appears
 * anchored just to the right of the source range.
 */
export function InsertChartDialog(props: InsertChartDialogProps): ReactNode {
  const { open, defaultRange, defaultKind, onCancel, onSubmit } = props;
  const { t } = useTranslator();
  const [kind, setKind] = useState<ChartKind>(defaultKind ?? "column");
  const [range, setRange] = useState(defaultRange);
  const [title, setTitle] = useState("");
  const [hasHeaderRow, setHasHeaderRow] = useState(true);
  const [hasCategoryColumn, setHasCategoryColumn] = useState(true);
  const panelRef = useRef<HTMLDivElement>(null);
  useFocusTrap(panelRef, { enabled: open, onEscape: onCancel });

  useEffect(() => {
    if (open) {
      setRange(defaultRange);
      setKind(defaultKind ?? "column");
      setTitle("");
      setHasHeaderRow(true);
      setHasCategoryColumn(true);
    }
  }, [open, defaultRange, defaultKind]);

  if (!open) return null;

  const submit = () => {
    if (!range.trim()) return;
    onSubmit({
      kind,
      dataRange: range.trim().toUpperCase(),
      title: title.trim() ? title.trim() : undefined,
      hasHeaderRow,
      hasCategoryColumn,
    });
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label={t("xlsx.insertChart.title")}
      data-testid="insert-chart-dialog"
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
          width: 420,
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
        <div style={{ fontSize: 14, fontWeight: 600 }}>{t("xlsx.insertChart.title")}</div>

        <div>
          <div style={{ fontSize: 12, fontWeight: 500, marginBottom: 6 }}>{t("xlsx.insertChart.chartType")}</div>
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
                {t(entry.labelKey)}
              </button>
            ))}
          </div>
        </div>

        <label style={{ display: "flex", flexDirection: "column", gap: 4 }}>
          <span style={{ fontSize: 12, fontWeight: 500 }}>{t("xlsx.insertChart.dataRange")}</span>
          <input
            type="text"
            value={range}
            onChange={(e) => setRange(e.target.value)}
            placeholder={t("xlsx.insertChart.rangePlaceholder")}
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
          <span style={{ fontSize: 12, fontWeight: 500 }}>{t("xlsx.insertChart.titleOptional")}</span>
          <input
            type="text"
            value={title}
            onChange={(e) => setTitle(e.target.value)}
            placeholder={t("xlsx.insertChart.titlePlaceholder")}
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
            {t("xlsx.insertChart.hasHeaderRow")}
          </label>
          <label style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <input
              type="checkbox"
              checked={hasCategoryColumn}
              onChange={(e) => setHasCategoryColumn(e.target.checked)}
            />
            {t("xlsx.insertChart.hasCategoryColumn")}
          </label>
        </div>

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
            {t("common.cancel")}
          </button>
          <button
            type="button"
            onClick={submit}
            style={{
              padding: "6px 12px",
              border: "1px solid var(--ai-violet)",
              borderRadius: 4,
              background: "var(--ai-violet)",
              color: "white",
              fontSize: 12,
              cursor: "pointer",
            }}
          >
            {t("xlsx.insertChart.insert")}
          </button>
        </div>
      </div>
    </div>
  );
}
