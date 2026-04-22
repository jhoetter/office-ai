"use client";

/**
 * PowerPoint-style "Slide Size" dialog.
 *
 * Mirrors Design → Slide Size → Custom Slide Size. Lets the user
 * pick from the standard widescreen / standard / paper presets or
 * enter a custom width × height (in cm or inches), and apply.
 *
 * Submission dispatches `pptx:set-slide-size`.
 */

import { useEffect, useMemo, useRef, useState } from "react";
import type { ReactNode } from "react";
import { X } from "lucide-react";
import { useFocusTrap } from "@officeai/ui";

const EMU_PER_INCH = 914400;
const EMU_PER_CM = 360000;

export type SlideSizePreset = "widescreen" | "standard" | "a4" | "letter" | "custom";

const PRESET_LABEL: Readonly<Record<SlideSizePreset, string>> = {
  widescreen: "On-screen show (16:9)",
  standard: "On-screen show (4:3)",
  a4: "A4 paper (210 × 297 mm)",
  letter: "Letter paper (8.5 × 11 in)",
  custom: "Custom",
};

const PRESET_DIMS: Readonly<Record<Exclude<SlideSizePreset, "custom">, { cxEmu: number; cyEmu: number }>> = {
  widescreen: { cxEmu: 12192000, cyEmu: 6858000 },
  standard: { cxEmu: 9144000, cyEmu: 6858000 },
  a4: { cxEmu: 9906000, cyEmu: 7560000 },
  letter: { cxEmu: 9144000, cyEmu: 6858000 },
};

export interface SlideSizeDialogProps {
  readonly open: boolean;
  readonly currentCxEmu: number;
  readonly currentCyEmu: number;
  readonly onClose: () => void;
  readonly onSubmit: (payload: {
    preset: SlideSizePreset;
    cxEmu: number;
    cyEmu: number;
  }) => void;
}

export function SlideSizeDialog(props: SlideSizeDialogProps): ReactNode {
  const { open, currentCxEmu, currentCyEmu, onClose, onSubmit } = props;
  const useMetric = useMemo(() => {
    try {
      return new Intl.NumberFormat().resolvedOptions().locale?.startsWith("en-US") !== true;
    } catch {
      return false;
    }
  }, []);
  const unit: "in" | "cm" = useMetric ? "cm" : "in";

  const [preset, setPreset] = useState<SlideSizePreset>(matchPreset(currentCxEmu, currentCyEmu));
  const [cxEmu, setCxEmu] = useState<number>(currentCxEmu);
  const [cyEmu, setCyEmu] = useState<number>(currentCyEmu);

  useEffect(() => {
    if (!open) return;
    setPreset(matchPreset(currentCxEmu, currentCyEmu));
    setCxEmu(currentCxEmu);
    setCyEmu(currentCyEmu);
  }, [open, currentCxEmu, currentCyEmu]);

  const panelRef = useRef<HTMLDivElement>(null);
  useFocusTrap(panelRef, { enabled: open, onEscape: onClose });

  if (!open) return null;

  const onSelectPreset = (id: SlideSizePreset): void => {
    setPreset(id);
    if (id !== "custom") {
      setCxEmu(PRESET_DIMS[id].cxEmu);
      setCyEmu(PRESET_DIMS[id].cyEmu);
    }
  };

  const handleApply = (): void => {
    onSubmit({ preset, cxEmu, cyEmu });
    onClose();
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="pptx-slide-size-title"
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
          <h2 id="pptx-slide-size-title" className="text-base font-semibold">
            Slide size
          </h2>
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
              Slides sized for
            </span>
            <select
              value={preset}
              onChange={(e) => onSelectPreset(e.currentTarget.value as SlideSizePreset)}
              className="rounded border border-divider bg-background px-2 py-1.5"
              data-testid="slide-size-preset"
            >
              {(Object.keys(PRESET_LABEL) as SlideSizePreset[]).map((id) => (
                <option key={id} value={id}>
                  {PRESET_LABEL[id]}
                </option>
              ))}
            </select>
          </label>

          <div className="grid grid-cols-2 gap-3">
            <UnitField
              label={`Width (${unit})`}
              emu={cxEmu}
              unit={unit}
              testId="slide-size-width"
              onChange={(v) => {
                setCxEmu(v);
                setPreset("custom");
              }}
            />
            <UnitField
              label={`Height (${unit})`}
              emu={cyEmu}
              unit={unit}
              testId="slide-size-height"
              onChange={(v) => {
                setCyEmu(v);
                setPreset("custom");
              }}
            />
          </div>
        </div>

        <footer className="flex items-center justify-end gap-2 border-t border-divider px-5 py-3">
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
            data-testid="slide-size-apply"
            className="rounded bg-accent px-3 py-1.5 text-sm font-medium text-white hover:opacity-90"
          >
            OK
          </button>
        </footer>
      </div>
    </div>
  );
}

function UnitField(props: {
  label: string;
  emu: number;
  unit: "in" | "cm";
  testId: string;
  onChange: (emu: number) => void;
}): ReactNode {
  const { label, emu, unit, testId, onChange } = props;
  const factor = unit === "in" ? EMU_PER_INCH : EMU_PER_CM;
  const display = (emu / factor).toFixed(2);
  return (
    <label className="flex flex-col gap-1">
      <span className="text-xs text-secondary">{label}</span>
      <input
        type="number"
        inputMode="decimal"
        step={unit === "in" ? 0.1 : 0.5}
        min={0}
        value={display}
        onChange={(e) => {
          const v = Number(e.currentTarget.value);
          if (Number.isFinite(v) && v >= 0) onChange(Math.round(v * factor));
        }}
        className="rounded border border-divider bg-background px-2 py-1 tabular-nums"
        data-testid={testId}
      />
    </label>
  );
}

function matchPreset(cx: number, cy: number): SlideSizePreset {
  for (const id of ["widescreen", "standard", "a4", "letter"] as const) {
    const p = PRESET_DIMS[id];
    if (Math.abs(p.cxEmu - cx) < 1000 && Math.abs(p.cyEmu - cy) < 1000) return id;
  }
  return "custom";
}
