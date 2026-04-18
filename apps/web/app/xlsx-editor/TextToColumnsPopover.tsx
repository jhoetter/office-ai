"use client";

import { useEffect, useRef, useState, type ReactNode } from "react";
import { cn } from "@officeai/ui";

/**
 * Modal-style popover that mimics Excel's "Convert Text to Columns
 * Wizard" but compressed into a single dialog. Captures:
 *
 *   - delimiter preset (Tab / Comma / Semicolon / Space) or custom
 *   - "Treat consecutive delimiters as one"
 *
 * The destination defaults to the source range's top-left in the
 * caller; we don't surface a destination picker yet to keep this a
 * one-shot dialog.
 */
export interface TextToColumnsPopoverProps {
  readonly open: boolean;
  readonly defaultDelimiter?: string;
  readonly onCancel: () => void;
  readonly onConfirm: (opts: { delimiter: string; treatConsecutiveAsOne: boolean }) => void;
}

type Preset = "tab" | "comma" | "semicolon" | "space" | "custom";

const PRESET_DELIMS: Record<Exclude<Preset, "custom">, string> = {
  tab: "\t",
  comma: ",",
  semicolon: ";",
  space: " ",
};

function detectPreset(delim: string): Preset {
  for (const [key, value] of Object.entries(PRESET_DELIMS)) {
    if (value === delim) return key as Preset;
  }
  return "custom";
}

export function TextToColumnsPopover(props: TextToColumnsPopoverProps): ReactNode {
  const { open, defaultDelimiter = ",", onCancel, onConfirm } = props;
  const [preset, setPreset] = useState<Preset>(() => detectPreset(defaultDelimiter));
  const [custom, setCustom] = useState<string>(defaultDelimiter);
  const [collapse, setCollapse] = useState<boolean>(false);
  const customRef = useRef<HTMLInputElement>(null);
  const dialogRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;
    setPreset(detectPreset(defaultDelimiter));
    setCustom(defaultDelimiter);
    setCollapse(false);
  }, [open, defaultDelimiter]);

  useEffect(() => {
    if (!open) return;
    function onKey(e: KeyboardEvent) {
      if (e.key === "Escape") {
        e.preventDefault();
        onCancel();
      }
    }
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [open, onCancel]);

  // Auto-focus the first interactive control so keyboard users land
  // somewhere useful without a tabindex hunt.
  useEffect(() => {
    if (!open) return;
    const root = dialogRef.current;
    if (!root) return;
    const first = root.querySelector<HTMLButtonElement | HTMLInputElement>("button, input");
    first?.focus();
  }, [open]);

  if (!open) return null;

  const delim = preset === "custom" ? custom : PRESET_DELIMS[preset];
  const canSubmit = delim.length > 0;

  return (
    <div
      role="dialog"
      aria-modal
      aria-labelledby="ttoc-title"
      data-testid="text-to-columns-popover"
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/30"
      onClick={(e) => {
        if (e.target === e.currentTarget) onCancel();
      }}
    >
      <div
        ref={dialogRef}
        className="w-[320px] rounded-lg border border-divider bg-surface p-4 shadow-lg"
      >
        <h2 id="ttoc-title" className="mb-2 text-sm font-medium text-foreground">
          Text to Columns
        </h2>
        <p className="mb-3 text-xs text-secondary">
          Split each cell in the selection on the chosen delimiter.
        </p>

        <fieldset className="mb-3 space-y-1">
          <legend className="text-xs text-secondary mb-1">Delimiter</legend>
          {(["tab", "comma", "semicolon", "space", "custom"] as ReadonlyArray<Preset>).map((p) => (
            <label
              key={p}
              className={cn(
                "flex items-center gap-2 rounded px-2 py-1 text-xs text-foreground hover:bg-hover cursor-pointer"
              )}
            >
              <input
                type="radio"
                name="ttoc-preset"
                value={p}
                checked={preset === p}
                onChange={() => {
                  setPreset(p);
                  if (p === "custom") {
                    setTimeout(() => customRef.current?.focus(), 0);
                  }
                }}
                data-testid={`ttoc-preset-${p}`}
              />
              <span className="capitalize">{p}</span>
              {p === "custom" ? (
                <input
                  ref={customRef}
                  type="text"
                  value={custom}
                  onChange={(e) => setCustom(e.target.value)}
                  onFocus={() => setPreset("custom")}
                  placeholder="e.g. ::"
                  data-testid="ttoc-custom-input"
                  className="ml-auto h-6 w-24 rounded border border-divider bg-background px-1 font-mono text-xs"
                />
              ) : (
                <span className="ml-auto font-mono text-secondary">
                  {p === "tab" ? "\\t" : p === "space" ? "␣" : PRESET_DELIMS[p as Exclude<Preset, "custom">]}
                </span>
              )}
            </label>
          ))}
        </fieldset>

        <label className="mb-3 flex items-center gap-2 text-xs text-foreground">
          <input
            type="checkbox"
            checked={collapse}
            onChange={(e) => setCollapse(e.target.checked)}
            data-testid="ttoc-collapse"
          />
          Treat consecutive delimiters as one
        </label>

        <div className="flex justify-end gap-2">
          <button
            type="button"
            onClick={onCancel}
            data-testid="ttoc-cancel"
            className="h-8 rounded border border-divider bg-background px-3 text-xs text-foreground hover:bg-hover"
          >
            Cancel
          </button>
          <button
            type="button"
            disabled={!canSubmit}
            onClick={() => onConfirm({ delimiter: delim, treatConsecutiveAsOne: collapse })}
            data-testid="ttoc-confirm"
            className="h-8 rounded bg-[var(--ai-violet)] px-3 text-xs font-medium text-white hover:opacity-90 disabled:opacity-50"
          >
            Split
          </button>
        </div>
      </div>
    </div>
  );
}
