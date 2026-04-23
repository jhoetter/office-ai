"use client";

/**
 * PowerPoint-style "Set Up Show" dialog.
 *
 * Mirrors Slideshow → Set Up Slide Show. Lets the user pick the
 * playback mode (presented by speaker / browsed by individual /
 * kiosk) and toggle the most common boolean flags (loop, narration,
 * animation, timings).
 *
 * Submission dispatches `pptx:set-show-options`. A "Reset all" link
 * dispatches the same command with `clear: true` to remove the
 * `<p:showPr>` element entirely (revert to PowerPoint defaults).
 */

import { useEffect, useRef, useState } from "react";
import type { ReactNode } from "react";
import { X } from "lucide-react";
import { useFocusTrap } from "@officeai/ui";

export type ShowType = "presenter" | "browse" | "kiosk";

export interface SetUpShowValues {
  readonly showType: ShowType;
  readonly loop: boolean;
  readonly showNarration: boolean;
  readonly showAnimation: boolean;
  readonly useTimings: boolean;
}

export interface SetUpShowDialogProps {
  readonly open: boolean;
  readonly current: SetUpShowValues;
  readonly onClose: () => void;
  readonly onSubmit: (values: SetUpShowValues) => void;
  readonly onClear: () => void;
}

export function SetUpShowDialog(props: SetUpShowDialogProps): ReactNode {
  const { open, current, onClose, onSubmit, onClear } = props;
  const [values, setValues] = useState<SetUpShowValues>(current);
  const panelRef = useRef<HTMLDivElement>(null);
  useFocusTrap(panelRef, { enabled: open, onEscape: onClose });

  useEffect(() => {
    if (open) setValues(current);
  }, [open, current]);

  if (!open) return null;

  const handleApply = (): void => {
    onSubmit(values);
    onClose();
  };

  const handleReset = (): void => {
    onClear();
    onClose();
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="pptx-show-title"
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
          <h2 id="pptx-show-title" className="text-base font-semibold">
            Set up slide show
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

        <div className="flex flex-col gap-4 px-5 py-4 text-sm">
          <fieldset className="flex flex-col gap-1">
            <legend className="text-xs font-medium uppercase tracking-wide text-secondary">Show type</legend>
            <Radio
              label="Presented by a speaker (full screen)"
              checked={values.showType === "presenter"}
              onChange={() => setValues({ ...values, showType: "presenter" })}
              testId="show-type-presenter"
            />
            <Radio
              label="Browsed by an individual (window)"
              checked={values.showType === "browse"}
              onChange={() => setValues({ ...values, showType: "browse" })}
              testId="show-type-browse"
            />
            <Radio
              label="Browsed at a kiosk (full screen, loop)"
              checked={values.showType === "kiosk"}
              onChange={() => setValues({ ...values, showType: "kiosk" })}
              testId="show-type-kiosk"
            />
          </fieldset>

          <fieldset className="flex flex-col gap-1.5">
            <legend className="text-xs font-medium uppercase tracking-wide text-secondary">
              Show options
            </legend>
            <Check
              label="Loop continuously until 'Esc'"
              checked={values.loop}
              onChange={(v) => setValues({ ...values, loop: v })}
              testId="show-opt-loop"
            />
            <Check
              label="Show without narration"
              checked={!values.showNarration}
              onChange={(v) => setValues({ ...values, showNarration: !v })}
              testId="show-opt-no-narration"
            />
            <Check
              label="Show without animation"
              checked={!values.showAnimation}
              onChange={(v) => setValues({ ...values, showAnimation: !v })}
              testId="show-opt-no-animation"
            />
            <Check
              label="Use timings, if present"
              checked={values.useTimings}
              onChange={(v) => setValues({ ...values, useTimings: v })}
              testId="show-opt-use-timings"
            />
          </fieldset>
        </div>

        <footer className="flex items-center justify-between gap-2 border-t border-divider px-5 py-3">
          <button
            type="button"
            onClick={handleReset}
            data-testid="show-reset"
            className="rounded text-xs text-secondary underline hover:text-default"
          >
            Reset to defaults
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
              data-testid="show-apply"
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

function Radio(props: { label: string; checked: boolean; onChange: () => void; testId: string }): ReactNode {
  const { label, checked, onChange, testId } = props;
  return (
    <label className="flex items-center gap-2">
      <input type="radio" checked={checked} onChange={onChange} data-testid={testId} />
      <span>{label}</span>
    </label>
  );
}

function Check(props: {
  label: string;
  checked: boolean;
  onChange: (v: boolean) => void;
  testId: string;
}): ReactNode {
  const { label, checked, onChange, testId } = props;
  return (
    <label className="flex items-center gap-2">
      <input
        type="checkbox"
        checked={checked}
        onChange={(e) => onChange(e.currentTarget.checked)}
        data-testid={testId}
      />
      <span>{label}</span>
    </label>
  );
}
