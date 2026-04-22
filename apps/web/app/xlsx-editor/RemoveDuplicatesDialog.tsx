"use client";

/**
 * Excel-style "Remove Duplicates" dialog.
 *
 * Mirrors Data → Remove Duplicates: the user picks which columns
 * inside the active range identify a duplicate, and may toggle
 * "My data has headers" — when on, the first row is treated as a
 * header (matching Excel) and the column labels come from that row;
 * when off, the columns are labeled "Column A / B / …" and the
 * full range is deduped including row 1.
 *
 * Submission dispatches `xlsx:remove-duplicates`. The handler keeps
 * the first occurrence of each unique key tuple, repacks survivors
 * to the top of the body, and clears trailing rows — matching
 * Excel's selection-only behaviour.
 */

import { useEffect, useMemo, useRef, useState } from "react";
import type { ReactNode } from "react";
import { X } from "lucide-react";
import { useFocusTrap } from "@officeai/ui";

export interface RemoveDuplicatesValues {
  readonly hasHeaders: boolean;
  /** 0-based offsets from the range's first column. */
  readonly keyCols: ReadonlyArray<number>;
}

export interface RemoveDuplicatesDialogProps {
  readonly open: boolean;
  readonly sheetName: string | null;
  /** A1 range covered by the active selection (e.g. "A1:E25"). */
  readonly rangeRef: string | null;
  /** First-row labels (length = column span); used when hasHeaders=true. */
  readonly headerLabels: ReadonlyArray<string>;
  /** Spreadsheet column letters for the same span (always provided). */
  readonly columnLetters: ReadonlyArray<string>;
  readonly onClose: () => void;
  readonly onSubmit: (values: RemoveDuplicatesValues) => void;
}

export function RemoveDuplicatesDialog(props: RemoveDuplicatesDialogProps): ReactNode {
  const { open, sheetName, rangeRef, headerLabels, columnLetters, onClose, onSubmit } = props;
  const [hasHeaders, setHasHeaders] = useState<boolean>(true);
  const [picked, setPicked] = useState<ReadonlySet<number>>(() => new Set());
  const panelRef = useRef<HTMLDivElement>(null);
  useFocusTrap(panelRef, { enabled: open, onEscape: onClose });

  // Reset state every time the dialog re-opens so the prior session
  // doesn't leak its column selection into a fresh range.
  useEffect(() => {
    if (!open) return;
    setHasHeaders(true);
    setPicked(new Set(columnLetters.map((_, i) => i)));
  }, [open, columnLetters]);

  const labels = useMemo(() => {
    return columnLetters.map((letter, i) => {
      if (hasHeaders) {
        const raw = headerLabels[i];
        const trimmed = (raw ?? "").toString().trim();
        return trimmed.length > 0 ? trimmed : `(Column ${letter})`;
      }
      return `Column ${letter}`;
    });
  }, [columnLetters, headerLabels, hasHeaders]);

  if (!open) return null;

  const colCount = columnLetters.length;
  const allChecked = picked.size === colCount && colCount > 0;
  const noneChecked = picked.size === 0;
  const canSubmit = !!rangeRef && colCount > 0 && !noneChecked;

  const toggle = (idx: number, on: boolean): void => {
    setPicked((prev) => {
      const next = new Set(prev);
      if (on) next.add(idx);
      else next.delete(idx);
      return next;
    });
  };

  const handleSelectAll = (): void => {
    setPicked(new Set(columnLetters.map((_, i) => i)));
  };
  const handleUnselectAll = (): void => {
    setPicked(new Set());
  };

  const handleSubmit = (): void => {
    if (!canSubmit) return;
    onSubmit({
      hasHeaders,
      keyCols: Array.from(picked).sort((a, b) => a - b),
    });
    onClose();
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="xlsx-remove-duplicates-title"
      className="fixed inset-0 z-[60] flex items-center justify-center bg-black/40 px-4 py-6"
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div
        ref={panelRef}
        tabIndex={-1}
        className="flex max-h-[85vh] w-full max-w-md flex-col rounded-lg border border-divider bg-surface shadow-2xl outline-none"
        onClick={(e) => e.stopPropagation()}
      >
        <header className="flex items-center justify-between gap-3 border-b border-divider px-5 py-3">
          <div>
            <h2 id="xlsx-remove-duplicates-title" className="text-base font-semibold">
              Duplikate entfernen{sheetName ? ` — ${sheetName}` : ""}
            </h2>
            <p className="text-xs text-secondary">
              {rangeRef
                ? `Wende Vergleich auf den Bereich ${rangeRef} an. Wähle Spalten für die Eindeutigkeitsprüfung.`
                : "Markiere zuerst einen Bereich auf dem Blatt."}
            </p>
          </div>
          <button
            type="button"
            aria-label="Close"
            onClick={onClose}
            className="rounded p-1 text-secondary transition-colors hover:bg-hover hover:text-default"
          >
            <X size={16} />
          </button>
        </header>

        <div className="flex flex-1 flex-col gap-4 overflow-y-auto px-5 py-4 text-sm">
          <label className="flex items-center gap-2">
            <input
              type="checkbox"
              data-testid="xlsx-remove-duplicates-headers"
              checked={hasHeaders}
              onChange={(e) => setHasHeaders(e.currentTarget.checked)}
            />
            <span>Daten haben Überschriften</span>
          </label>

          <fieldset className="flex flex-col gap-2">
            <div className="flex items-center justify-between">
              <legend className="text-xs font-medium uppercase tracking-wide text-secondary">
                Spalten
              </legend>
              <div className="flex items-center gap-2 text-xs">
                <button
                  type="button"
                  onClick={handleSelectAll}
                  disabled={allChecked}
                  className="rounded border border-divider bg-background px-2 py-0.5 hover:bg-hover disabled:opacity-40"
                >
                  Alle auswählen
                </button>
                <button
                  type="button"
                  onClick={handleUnselectAll}
                  disabled={noneChecked}
                  className="rounded border border-divider bg-background px-2 py-0.5 hover:bg-hover disabled:opacity-40"
                >
                  Auswahl aufheben
                </button>
              </div>
            </div>
            <div className="flex max-h-64 flex-col gap-1 overflow-y-auto rounded border border-divider bg-background px-2 py-1.5">
              {labels.map((label, idx) => (
                <label key={idx} className="flex items-center gap-2">
                  <input
                    type="checkbox"
                    data-testid={`xlsx-remove-duplicates-col-${idx}`}
                    checked={picked.has(idx)}
                    onChange={(e) => toggle(idx, e.currentTarget.checked)}
                  />
                  <span className="truncate" title={label}>
                    {label}
                  </span>
                </label>
              ))}
              {colCount === 0 ? (
                <span className="text-xs text-secondary">Kein Bereich ausgewählt.</span>
              ) : null}
            </div>
          </fieldset>
        </div>

        <footer className="flex items-center justify-end gap-2 border-t border-divider px-5 py-3">
          <button
            type="button"
            onClick={onClose}
            className="rounded border border-divider bg-background px-3 py-1.5 text-sm hover:bg-hover"
          >
            Abbrechen
          </button>
          <button
            type="button"
            onClick={handleSubmit}
            disabled={!canSubmit}
            data-testid="xlsx-remove-duplicates-apply"
            className="rounded bg-accent px-3 py-1.5 text-sm font-medium text-white hover:opacity-90 disabled:opacity-40"
          >
            OK
          </button>
        </footer>
      </div>
    </div>
  );
}
