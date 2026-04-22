"use client";

import { useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import { useFocusTrap } from "@officeai/ui";

/**
 * Pivot creation dialog (Phase-1 minimum-viable build).
 *
 * Real OOXML pivot tables ship with a separate cache part + view
 * definition that round-trip the user's row/column/value/filter axes.
 * We model the read side already (`PivotTablePart` + `PivotCachePart`)
 * but creating one from scratch requires emitting valid pivotCache /
 * pivotTable / pivotCacheRecords XML — multi-day work that's its own
 * spec milestone.
 *
 * To unblock "how do I make a pivot now?" today, the dialog drives a
 * **values-only summary**: pick the source range, the column to group
 * by, the column to aggregate, the aggregation, and the destination
 * top-left. The runner computes the result client-side and writes the
 * summary as plain values via `xlsx:set-range-values`. This is not a
 * round-trippable pivot (re-saving doesn't preserve pivot metadata)
 * but it covers the "summarize my data" need users actually have on
 * day one. The full OOXML write path is tracked in
 * `spec/xlsx/pivot-tables.md` Phase 2.
 */

export type PivotAggregation = "sum" | "count" | "average" | "min" | "max";

export interface PivotDialogSubmit {
  readonly sourceRange: string;
  readonly hasHeaderRow: boolean;
  readonly groupColumnIndex: number;
  readonly valueColumnIndex: number;
  readonly aggregation: PivotAggregation;
  readonly destinationTopLeftA1: string;
}

interface InsertPivotDialogProps {
  readonly open: boolean;
  readonly defaultSourceRange: string;
  readonly defaultDestination: string;
  /**
   * Header labels for each column in the source range. Length = number
   * of columns in `defaultSourceRange`. The dialog uses these to label
   * the group/value pickers; falls back to "Column 1", "Column 2", …
   * when the user toggles "First row is header" off.
   */
  readonly headerLabels: ReadonlyArray<string>;
  readonly onCancel: () => void;
  readonly onSubmit: (args: PivotDialogSubmit) => void;
}

const AGGREGATIONS: ReadonlyArray<{ readonly id: PivotAggregation; readonly label: string }> = [
  { id: "sum", label: "Sum" },
  { id: "count", label: "Count" },
  { id: "average", label: "Average" },
  { id: "min", label: "Min" },
  { id: "max", label: "Max" },
];

export function InsertPivotTableDialog(props: InsertPivotDialogProps): ReactNode {
  const { open, defaultSourceRange, defaultDestination, headerLabels, onCancel, onSubmit } = props;

  const [sourceRange, setSourceRange] = useState(defaultSourceRange);
  const [hasHeaderRow, setHasHeaderRow] = useState(true);
  const [groupColumnIndex, setGroupColumnIndex] = useState(0);
  const [valueColumnIndex, setValueColumnIndex] = useState(headerLabels.length > 1 ? 1 : 0);
  const [aggregation, setAggregation] = useState<PivotAggregation>("sum");
  const [destination, setDestination] = useState(defaultDestination);

  const panelRef = useRef<HTMLDivElement>(null);
  useFocusTrap(panelRef, { enabled: open, onEscape: onCancel });

  useEffect(() => {
    if (!open) return;
    setSourceRange(defaultSourceRange);
    setHasHeaderRow(true);
    setGroupColumnIndex(0);
    setValueColumnIndex(headerLabels.length > 1 ? 1 : 0);
    setAggregation("sum");
    setDestination(defaultDestination);
  }, [open, defaultSourceRange, defaultDestination, headerLabels.length]);

  const labels = useMemo(() => {
    if (hasHeaderRow) {
      return headerLabels.map((h, i) => (h.trim().length > 0 ? h : `Column ${i + 1}`));
    }
    return headerLabels.map((_, i) => `Column ${i + 1}`);
  }, [headerLabels, hasHeaderRow]);

  if (!open) return null;

  const canSubmit =
    sourceRange.trim().length > 0 &&
    destination.trim().length > 0 &&
    groupColumnIndex >= 0 &&
    valueColumnIndex >= 0 &&
    labels.length > 0;

  const submit = () => {
    if (!canSubmit) return;
    onSubmit({
      sourceRange: sourceRange.trim(),
      hasHeaderRow,
      groupColumnIndex,
      valueColumnIndex,
      aggregation,
      destinationTopLeftA1: destination.trim(),
    });
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label="Insert pivot table"
      data-testid="insert-pivot-dialog"
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/40"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onCancel();
      }}
    >
      <div
        ref={panelRef}
        className="w-[480px] max-w-full rounded-md border border-divider bg-surface p-4 text-sm shadow-xl"
      >
        <h2 className="mb-3 text-base font-semibold text-foreground">Insert pivot table</h2>

        <div className="mb-3">
          <label className="mb-1 block text-xs font-medium text-secondary" htmlFor="pivot-source">
            Source data range
          </label>
          <input
            id="pivot-source"
            data-testid="pivot-source-range"
            value={sourceRange}
            onChange={(e) => setSourceRange(e.target.value)}
            className="w-full rounded border border-divider bg-background px-2 py-1.5 font-mono text-xs text-foreground focus:border-[var(--accent)] focus:outline-none"
            placeholder="A1:C10"
          />
        </div>

        <label className="mb-3 inline-flex items-center gap-2 text-xs text-foreground">
          <input
            type="checkbox"
            checked={hasHeaderRow}
            onChange={(e) => setHasHeaderRow(e.target.checked)}
            className="h-3 w-3 accent-[var(--accent)]"
          />
          First row is a header
        </label>

        <div className="mb-3 grid grid-cols-2 gap-3">
          <div>
            <label className="mb-1 block text-xs font-medium text-secondary" htmlFor="pivot-group">
              Group rows by
            </label>
            <select
              id="pivot-group"
              data-testid="pivot-group-column"
              value={groupColumnIndex}
              onChange={(e) => setGroupColumnIndex(Number.parseInt(e.target.value, 10))}
              className="w-full rounded border border-divider bg-background px-2 py-1.5 text-xs text-foreground focus:border-[var(--accent)] focus:outline-none"
            >
              {labels.map((l, i) => (
                <option key={i} value={i}>
                  {l}
                </option>
              ))}
            </select>
          </div>

          <div>
            <label className="mb-1 block text-xs font-medium text-secondary" htmlFor="pivot-value">
              Aggregate column
            </label>
            <select
              id="pivot-value"
              data-testid="pivot-value-column"
              value={valueColumnIndex}
              onChange={(e) => setValueColumnIndex(Number.parseInt(e.target.value, 10))}
              className="w-full rounded border border-divider bg-background px-2 py-1.5 text-xs text-foreground focus:border-[var(--accent)] focus:outline-none"
            >
              {labels.map((l, i) => (
                <option key={i} value={i}>
                  {l}
                </option>
              ))}
            </select>
          </div>
        </div>

        <div className="mb-3 grid grid-cols-2 gap-3">
          <div>
            <label className="mb-1 block text-xs font-medium text-secondary" htmlFor="pivot-agg">
              Aggregation
            </label>
            <select
              id="pivot-agg"
              data-testid="pivot-aggregation"
              value={aggregation}
              onChange={(e) => setAggregation(e.target.value as PivotAggregation)}
              className="w-full rounded border border-divider bg-background px-2 py-1.5 text-xs text-foreground focus:border-[var(--accent)] focus:outline-none"
            >
              {AGGREGATIONS.map((a) => (
                <option key={a.id} value={a.id}>
                  {a.label}
                </option>
              ))}
            </select>
          </div>

          <div>
            <label className="mb-1 block text-xs font-medium text-secondary" htmlFor="pivot-dest">
              Destination top-left
            </label>
            <input
              id="pivot-dest"
              data-testid="pivot-destination"
              value={destination}
              onChange={(e) => setDestination(e.target.value)}
              className="w-full rounded border border-divider bg-background px-2 py-1.5 font-mono text-xs text-foreground focus:border-[var(--accent)] focus:outline-none"
              placeholder="E1"
            />
          </div>
        </div>

        <p className="mb-4 text-[11px] text-tertiary">
          Phase 1 writes the result as plain values. Full OOXML pivot round-trip (slicers, calculated fields,
          drill-through) lands in a follow-up.
        </p>

        <div className="flex justify-end gap-2">
          <button
            type="button"
            onClick={onCancel}
            className="rounded border border-divider px-3 py-1.5 text-xs hover:bg-hover"
          >
            Cancel
          </button>
          <button
            type="button"
            onClick={submit}
            disabled={!canSubmit}
            data-testid="pivot-submit"
            className="rounded bg-[var(--accent)] px-3 py-1.5 text-xs font-medium text-white hover:opacity-90 disabled:opacity-40"
          >
            Insert
          </button>
        </div>
      </div>
    </div>
  );
}
