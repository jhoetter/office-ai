"use client";

/**
 * Excel-style "Protect Sheet" dialog.
 *
 * Mirrors Review → Protect Sheet. The user picks which actions remain
 * permitted while the sheet is protected, and optionally supplies a
 * password (we hash it in the bus handler — the dialog ships the
 * literal password and the editor turns it into a precomputed hash
 * before dispatching, matching how Excel treats password-protect).
 *
 * For this milestone we ship the password verbatim as `passwordHash`
 * because the bus handler does not yet derive ECMA-376 SHA-512 hashes
 * for us. The CLI / MCP path uses the same payload shape.
 *
 * Submission dispatches `xlsx:set-sheet-protection`. To unprotect,
 * the dialog also exposes a "Remove protection" button that
 * dispatches `enabled: false`.
 */

import { useEffect, useRef, useState } from "react";
import type { ReactNode } from "react";
import { X } from "@officeai/ui/sonaloop-icons";
import { useFocusTrap } from "@officeai/ui";

export interface ProtectSheetValues {
  readonly passwordHash?: string;
  readonly selectLockedCells: boolean;
  readonly selectUnlockedCells: boolean;
  readonly formatCells: boolean;
  readonly formatColumns: boolean;
  readonly formatRows: boolean;
  readonly insertColumns: boolean;
  readonly insertRows: boolean;
  readonly insertHyperlinks: boolean;
  readonly deleteColumns: boolean;
  readonly deleteRows: boolean;
  readonly sort: boolean;
  readonly autoFilter: boolean;
  readonly pivotTables: boolean;
}

export interface ProtectSheetDialogProps {
  readonly open: boolean;
  readonly sheetName: string | null;
  readonly currentlyProtected: boolean;
  readonly onClose: () => void;
  readonly onProtect: (values: ProtectSheetValues) => void;
  readonly onUnprotect: () => void;
}

const DEFAULT_VALUES: ProtectSheetValues = {
  selectLockedCells: true,
  selectUnlockedCells: true,
  formatCells: false,
  formatColumns: false,
  formatRows: false,
  insertColumns: false,
  insertRows: false,
  insertHyperlinks: false,
  deleteColumns: false,
  deleteRows: false,
  sort: false,
  autoFilter: false,
  pivotTables: false,
};

const PERMISSION_ROWS: ReadonlyArray<{ id: keyof ProtectSheetValues; label: string }> = [
  { id: "selectLockedCells", label: "Select locked cells" },
  { id: "selectUnlockedCells", label: "Select unlocked cells" },
  { id: "formatCells", label: "Format cells" },
  { id: "formatColumns", label: "Format columns" },
  { id: "formatRows", label: "Format rows" },
  { id: "insertColumns", label: "Insert columns" },
  { id: "insertRows", label: "Insert rows" },
  { id: "insertHyperlinks", label: "Insert hyperlinks" },
  { id: "deleteColumns", label: "Delete columns" },
  { id: "deleteRows", label: "Delete rows" },
  { id: "sort", label: "Sort" },
  { id: "autoFilter", label: "Use AutoFilter" },
  { id: "pivotTables", label: "Use PivotTable & PivotChart" },
];

export function ProtectSheetDialog(props: ProtectSheetDialogProps): ReactNode {
  const { open, sheetName, currentlyProtected, onClose, onProtect, onUnprotect } = props;
  const [password, setPassword] = useState<string>("");
  const [values, setValues] = useState<ProtectSheetValues>(DEFAULT_VALUES);
  const panelRef = useRef<HTMLDivElement>(null);
  useFocusTrap(panelRef, { enabled: open, onEscape: onClose });

  useEffect(() => {
    if (!open) return;
    setPassword("");
    setValues(DEFAULT_VALUES);
  }, [open]);

  if (!open) return null;

  const handleApply = (): void => {
    onProtect({
      ...values,
      ...(password ? { passwordHash: password } : {}),
    });
    onClose();
  };

  const handleUnprotect = (): void => {
    onUnprotect();
    onClose();
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="xlsx-protect-sheet-title"
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
            <h2 id="xlsx-protect-sheet-title" className="text-base font-semibold">
              Protect sheet — {sheetName ?? ""}
            </h2>
            <p className="text-xs text-secondary">
              {currentlyProtected
                ? "This sheet is currently protected."
                : "Choose which actions are still allowed while this sheet is protected."}
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
          <label className="flex flex-col gap-1">
            <span className="text-xs font-medium uppercase tracking-wide text-secondary">
              Password (optional)
            </span>
            <input
              type="password"
              value={password}
              onChange={(e) => setPassword(e.currentTarget.value)}
              className="rounded border border-divider bg-background px-2 py-1"
              data-testid="protect-sheet-password"
            />
          </label>
          <fieldset className="flex flex-col gap-1.5">
            <legend className="text-xs font-medium uppercase tracking-wide text-secondary">
              Allow all users of this worksheet to:
            </legend>
            <div className="grid grid-cols-1 gap-1">
              {PERMISSION_ROWS.map((row) => (
                <label key={row.id} className="flex items-center gap-2">
                  <input
                    type="checkbox"
                    data-testid={`protect-sheet-${row.id}`}
                    checked={Boolean(values[row.id])}
                    onChange={(e) =>
                      setValues({ ...values, [row.id]: e.currentTarget.checked } as ProtectSheetValues)
                    }
                  />
                  <span>{row.label}</span>
                </label>
              ))}
            </div>
          </fieldset>
        </div>

        <footer className="flex items-center justify-between gap-2 border-t border-divider px-5 py-3">
          <button
            type="button"
            onClick={handleUnprotect}
            disabled={!currentlyProtected}
            data-testid="protect-sheet-unprotect"
            className="rounded border border-divider bg-background px-3 py-1.5 text-sm text-secondary hover:bg-hover disabled:opacity-40"
          >
            Remove protection
          </button>
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={onClose}
              className="rounded border border-divider bg-background px-3 py-1.5 text-sm hover:bg-hover"
            >
              Cancel
            </button>
            <button
              type="button"
              onClick={handleApply}
              data-testid="protect-sheet-apply"
              className="rounded bg-accent px-3 py-1.5 text-sm font-medium text-white hover:opacity-90"
            >
              OK
            </button>
          </div>
        </footer>
      </div>
    </div>
  );
}
