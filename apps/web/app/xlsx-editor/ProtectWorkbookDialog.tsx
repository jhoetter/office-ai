"use client";

/**
 * Excel-style "Protect Workbook" dialog.
 *
 * Mirrors Review → Protect Workbook → Structure. The user picks
 * which workbook-level surfaces to lock and optionally supplies a
 * password. We only expose the two flags that meaningfully reach
 * users in current Excel (`structure`, `windows`); the rest live on
 * the bus payload for CLI/MCP completeness.
 */

import { useEffect, useRef, useState } from "react";
import type { ReactNode } from "react";
import { X } from "lucide-react";
import { useFocusTrap } from "@officeai/ui";

export interface ProtectWorkbookValues {
  readonly workbookPasswordHash?: string;
  readonly lockStructure: boolean;
  readonly lockWindows: boolean;
}

export interface ProtectWorkbookDialogProps {
  readonly open: boolean;
  readonly currentlyProtected: boolean;
  readonly onClose: () => void;
  readonly onProtect: (values: ProtectWorkbookValues) => void;
  readonly onUnprotect: () => void;
}

export function ProtectWorkbookDialog(props: ProtectWorkbookDialogProps): ReactNode {
  const { open, currentlyProtected, onClose, onProtect, onUnprotect } = props;
  const [password, setPassword] = useState<string>("");
  const [lockStructure, setLockStructure] = useState<boolean>(true);
  const [lockWindows, setLockWindows] = useState<boolean>(false);
  const panelRef = useRef<HTMLDivElement>(null);
  useFocusTrap(panelRef, { enabled: open, onEscape: onClose });

  useEffect(() => {
    if (!open) return;
    setPassword("");
    setLockStructure(true);
    setLockWindows(false);
  }, [open]);

  if (!open) return null;

  const handleApply = (): void => {
    onProtect({
      lockStructure,
      lockWindows,
      ...(password ? { workbookPasswordHash: password } : {}),
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
      aria-labelledby="xlsx-protect-workbook-title"
      className="fixed inset-0 z-[60] flex items-center justify-center bg-black/40 px-4 py-6"
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div
        ref={panelRef}
        tabIndex={-1}
        className="flex w-full max-w-md flex-col rounded-lg border border-divider bg-surface shadow-2xl outline-none"
        onClick={(e) => e.stopPropagation()}
      >
        <header className="flex items-center justify-between gap-3 border-b border-divider px-5 py-3">
          <div>
            <h2 id="xlsx-protect-workbook-title" className="text-base font-semibold">
              Protect workbook structure
            </h2>
            <p className="text-xs text-secondary">
              {currentlyProtected
                ? "This workbook is currently protected."
                : "Choose what to lock at the workbook level."}
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

        <div className="flex flex-col gap-3 px-5 py-4 text-sm">
          <label className="flex flex-col gap-1">
            <span className="text-xs font-medium uppercase tracking-wide text-secondary">
              Password (optional)
            </span>
            <input
              type="password"
              value={password}
              onChange={(e) => setPassword(e.currentTarget.value)}
              className="rounded border border-divider bg-background px-2 py-1"
              data-testid="protect-workbook-password"
            />
          </label>
          <fieldset className="flex flex-col gap-1.5">
            <legend className="text-xs font-medium uppercase tracking-wide text-secondary">
              Protect workbook for
            </legend>
            <label className="flex items-center gap-2">
              <input
                type="checkbox"
                checked={lockStructure}
                onChange={(e) => setLockStructure(e.currentTarget.checked)}
                data-testid="protect-workbook-structure"
              />
              <span>Structure (sheets, hidden sheets, ordering)</span>
            </label>
            <label className="flex items-center gap-2">
              <input
                type="checkbox"
                checked={lockWindows}
                onChange={(e) => setLockWindows(e.currentTarget.checked)}
                data-testid="protect-workbook-windows"
              />
              <span>Windows</span>
            </label>
          </fieldset>
        </div>

        <footer className="flex items-center justify-between gap-2 border-t border-divider px-5 py-3">
          <button
            type="button"
            onClick={handleUnprotect}
            disabled={!currentlyProtected}
            data-testid="protect-workbook-unprotect"
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
              data-testid="protect-workbook-apply"
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
