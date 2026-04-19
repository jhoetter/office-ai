"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { X } from "lucide-react";
import { useFocusTrap } from "@officeai/ui";
import type { DocxSnapshot } from "@officeai/docx";

/**
 * B3 — Page Setup dialog.
 *
 * Renders a modal that maps the typed `<w:sectPr>` projection onto a
 * Word-style "Page Setup" surface: orientation, paper size (presets +
 * custom width/height) and margins (presets + custom values). Submit
 * dispatches `docx:set-page-setup` for the section that owns the
 * caret's paragraph index — so a multi-section document can have its
 * sections configured independently.
 *
 * Units default to inches in en-US/en-GB and centimetres elsewhere,
 * matching the same heuristic used by `PageRuler`. Internally
 * everything is twips so the command stays unit-agnostic.
 */

const TWIPS_PER_INCH = 1440;
const TWIPS_PER_CM = 567;

export interface PageSetupValues {
  pgSz: { w: number; h: number; orient: "portrait" | "landscape" };
  pgMar: { top: number; right: number; bottom: number; left: number; header: number; footer: number };
}

export interface PageSetupDialogProps {
  readonly open: boolean;
  readonly snapshot: DocxSnapshot | null;
  readonly paragraphIndex: number;
  readonly onClose: () => void;
  readonly onSubmit: (next: PageSetupValues) => void;
}

interface PaperPreset {
  readonly id: string;
  readonly label: string;
  readonly w: number; // twips, portrait
  readonly h: number; // twips, portrait
}

const PAPER_PRESETS: ReadonlyArray<PaperPreset> = [
  { id: "letter", label: 'Letter (8.5" × 11")', w: 12240, h: 15840 },
  { id: "legal", label: 'Legal (8.5" × 14")', w: 12240, h: 20160 },
  { id: "tabloid", label: 'Tabloid (11" × 17")', w: 15840, h: 24480 },
  { id: "a3", label: "A3 (29.7 × 42 cm)", w: 16838, h: 23811 },
  { id: "a4", label: "A4 (21 × 29.7 cm)", w: 11906, h: 16838 },
  { id: "a5", label: "A5 (14.8 × 21 cm)", w: 8390, h: 11906 },
  { id: "b5", label: "B5 (17.6 × 25 cm)", w: 9978, h: 14173 },
];

interface MarginPreset {
  readonly id: string;
  readonly label: string;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
  readonly left: number;
}

const MARGIN_PRESETS: ReadonlyArray<MarginPreset> = [
  { id: "normal", label: 'Normal (1" all sides)', top: 1440, right: 1440, bottom: 1440, left: 1440 },
  { id: "narrow", label: 'Narrow (0.5" all sides)', top: 720, right: 720, bottom: 720, left: 720 },
  {
    id: "moderate",
    label: 'Moderate (1" / 0.75")',
    top: 1440,
    right: 1080,
    bottom: 1440,
    left: 1080,
  },
  { id: "wide", label: 'Wide (1" / 2")', top: 1440, right: 2880, bottom: 1440, left: 2880 },
];

export function PageSetupDialog(props: PageSetupDialogProps) {
  const { open, snapshot, paragraphIndex, onClose, onSubmit } = props;
  const initial = useMemo(() => resolveSectionValues(snapshot, paragraphIndex), [snapshot, paragraphIndex]);
  const useMetric = useMemo(() => isMetricLocale(), []);
  const unit: "in" | "cm" = useMetric ? "cm" : "in";

  const [orient, setOrient] = useState<"portrait" | "landscape">(initial.pgSz.orient);
  const [paperW, setPaperW] = useState(initial.pgSz.w);
  const [paperH, setPaperH] = useState(initial.pgSz.h);
  const [top, setTop] = useState(initial.pgMar.top);
  const [right, setRight] = useState(initial.pgMar.right);
  const [bottom, setBottom] = useState(initial.pgMar.bottom);
  const [left, setLeft] = useState(initial.pgMar.left);
  const [header, setHeader] = useState(initial.pgMar.header);
  const [footer, setFooter] = useState(initial.pgMar.footer);

  useEffect(() => {
    if (!open) return;
    setOrient(initial.pgSz.orient);
    setPaperW(initial.pgSz.w);
    setPaperH(initial.pgSz.h);
    setTop(initial.pgMar.top);
    setRight(initial.pgMar.right);
    setBottom(initial.pgMar.bottom);
    setLeft(initial.pgMar.left);
    setHeader(initial.pgMar.header);
    setFooter(initial.pgMar.footer);
  }, [open, initial]);

  const panelRef = useRef<HTMLDivElement>(null);
  useFocusTrap(panelRef, { enabled: open, onEscape: onClose });

  if (!open) return null;

  const matchedPaper = matchPaper(paperW, paperH, orient);

  const handleApply = () => {
    const oriented = applyOrientation(paperW, paperH, orient);
    onSubmit({
      pgSz: { w: oriented.w, h: oriented.h, orient },
      pgMar: { top, right, bottom, left, header, footer },
    });
    onClose();
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="page-setup-title"
      data-page-setup-dialog
      className="fixed inset-0 z-[60] flex items-center justify-center bg-black/40 px-4 py-6"
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div
        ref={panelRef}
        tabIndex={-1}
        className="flex max-h-[85vh] w-full max-w-xl flex-col rounded-lg border border-divider bg-surface shadow-2xl outline-none"
        onClick={(e) => e.stopPropagation()}
      >
        <header className="flex items-center justify-between gap-3 border-b border-divider px-5 py-3">
          <div>
            <h2 id="page-setup-title" className="text-base font-semibold">
              Page setup
            </h2>
            <p className="text-xs text-secondary">
              Edit paper size, orientation and margins for the active section.
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
          <fieldset className="flex flex-col gap-2">
            <legend className="text-xs font-medium uppercase tracking-wide text-secondary">
              Orientation
            </legend>
            <div className="flex gap-2">
              <OrientButton
                active={orient === "portrait"}
                label="Portrait"
                onClick={() => setOrient("portrait")}
              />
              <OrientButton
                active={orient === "landscape"}
                label="Landscape"
                onClick={() => setOrient("landscape")}
              />
            </div>
          </fieldset>

          <fieldset className="flex flex-col gap-2">
            <legend className="text-xs font-medium uppercase tracking-wide text-secondary">Paper size</legend>
            <select
              value={matchedPaper?.id ?? "custom"}
              onChange={(e) => {
                const sel = PAPER_PRESETS.find((p) => p.id === e.target.value);
                if (!sel) return;
                setPaperW(sel.w);
                setPaperH(sel.h);
              }}
              className="rounded border border-divider bg-background px-2 py-1.5"
              data-testid="page-setup-paper"
            >
              {PAPER_PRESETS.map((p) => (
                <option key={p.id} value={p.id}>
                  {p.label}
                </option>
              ))}
              <option value="custom">Custom…</option>
            </select>
            <div className="grid grid-cols-2 gap-2">
              <UnitField
                label={`Width (${unit})`}
                twips={orient === "landscape" ? Math.max(paperW, paperH) : Math.min(paperW, paperH)}
                onChange={(t) => {
                  if (orient === "landscape") {
                    setPaperW(Math.max(paperW, paperH));
                    setPaperH(Math.min(paperW, paperH));
                  }
                  setPaperW(t);
                }}
                unit={unit}
              />
              <UnitField
                label={`Height (${unit})`}
                twips={orient === "landscape" ? Math.min(paperW, paperH) : Math.max(paperW, paperH)}
                onChange={(t) => setPaperH(t)}
                unit={unit}
              />
            </div>
          </fieldset>

          <fieldset className="flex flex-col gap-2">
            <legend className="text-xs font-medium uppercase tracking-wide text-secondary">Margins</legend>
            <div className="flex flex-wrap gap-1.5">
              {MARGIN_PRESETS.map((m) => (
                <button
                  key={m.id}
                  type="button"
                  onClick={() => {
                    setTop(m.top);
                    setRight(m.right);
                    setBottom(m.bottom);
                    setLeft(m.left);
                  }}
                  className="rounded border border-divider bg-background px-2 py-1 text-xs hover:bg-hover"
                >
                  {m.label}
                </button>
              ))}
            </div>
            <div className="grid grid-cols-2 gap-2">
              <UnitField label={`Top (${unit})`} twips={top} onChange={setTop} unit={unit} />
              <UnitField label={`Bottom (${unit})`} twips={bottom} onChange={setBottom} unit={unit} />
              <UnitField label={`Left (${unit})`} twips={left} onChange={setLeft} unit={unit} />
              <UnitField label={`Right (${unit})`} twips={right} onChange={setRight} unit={unit} />
              <UnitField label={`Header (${unit})`} twips={header} onChange={setHeader} unit={unit} />
              <UnitField label={`Footer (${unit})`} twips={footer} onChange={setFooter} unit={unit} />
            </div>
          </fieldset>
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
            className="rounded bg-accent px-3 py-1.5 text-sm font-medium text-white hover:opacity-90"
            data-testid="page-setup-apply"
          >
            Apply
          </button>
        </footer>
      </div>
    </div>
  );
}

function OrientButton(props: { active: boolean; label: string; onClick: () => void }) {
  const { active, label, onClick } = props;
  return (
    <button
      type="button"
      onClick={onClick}
      aria-pressed={active}
      className={`flex-1 rounded border px-3 py-1.5 text-sm transition-colors ${
        active
          ? "border-accent bg-accent-light text-default"
          : "border-divider bg-background text-secondary hover:bg-hover"
      }`}
    >
      {label}
    </button>
  );
}

function UnitField(props: {
  label: string;
  twips: number;
  unit: "in" | "cm";
  onChange: (twips: number) => void;
}) {
  const { label, twips, unit, onChange } = props;
  const factor = unit === "in" ? TWIPS_PER_INCH : TWIPS_PER_CM;
  const display = (twips / factor).toFixed(2);
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
      />
    </label>
  );
}

interface ResolvedValues {
  pgSz: { w: number; h: number; orient: "portrait" | "landscape" };
  pgMar: { top: number; right: number; bottom: number; left: number; header: number; footer: number };
}

function resolveSectionValues(snapshot: DocxSnapshot | null, paragraphIndex: number): ResolvedValues {
  const props = locateSectionProps(snapshot, paragraphIndex);
  const pgSz = props?.pgSz ?? { w: 12240, h: 15840 };
  const pgMar = props?.pgMar ?? {
    top: 1440,
    right: 1440,
    bottom: 1440,
    left: 1440,
    header: 720,
    footer: 720,
  };
  const orient: "portrait" | "landscape" = pgSz.orient ?? (pgSz.w > pgSz.h ? "landscape" : "portrait");
  return {
    pgSz: { w: pgSz.w, h: pgSz.h, orient },
    pgMar: {
      top: pgMar.top,
      right: pgMar.right,
      bottom: pgMar.bottom,
      left: pgMar.left,
      header: pgMar.header,
      footer: pgMar.footer,
    },
  };
}

function locateSectionProps(snapshot: DocxSnapshot | null, paragraphIndex: number) {
  if (!snapshot) return null;
  const body = snapshot.root.body;
  for (let i = paragraphIndex; i < body.length; i++) {
    const block = body[i];
    if (block.kind === "section-break") return block.properties;
  }
  for (let i = body.length - 1; i >= 0; i--) {
    const block = body[i];
    if (block.kind === "section-break") return block.properties;
  }
  return null;
}

function applyOrientation(w: number, h: number, orient: "portrait" | "landscape"): { w: number; h: number } {
  if (orient === "landscape") {
    return { w: Math.max(w, h), h: Math.min(w, h) };
  }
  return { w: Math.min(w, h), h: Math.max(w, h) };
}

function matchPaper(w: number, h: number, orient: "portrait" | "landscape"): PaperPreset | null {
  const oriented = applyOrientation(w, h, orient);
  for (const p of PAPER_PRESETS) {
    const portraitMatch = oriented.w === Math.min(p.w, p.h) && oriented.h === Math.max(p.w, p.h);
    if (portraitMatch) return p;
  }
  return null;
}

function isMetricLocale(): boolean {
  if (typeof navigator === "undefined") return false;
  const lang = (navigator.language || "en-US").toLowerCase();
  if (lang.startsWith("en-us")) return false;
  if (lang.startsWith("en-gb")) return false;
  if (lang.startsWith("en-lr")) return false;
  if (lang === "my-mm") return false;
  return true;
}
