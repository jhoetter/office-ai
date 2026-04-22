"use client";

/**
 * Excel-style "Page Setup" dialog.
 *
 * Surface for the four backend commands that together reproduce the
 * Page Layout → Page Setup launcher in Excel:
 *   - `xlsx:set-page-setup`   (Page tab: orientation, paper, scaling)
 *   - `xlsx:set-page-margins` (Margins tab: presets + per-edge inputs)
 *   - `xlsx:set-print-options`(Sheet tab: gridlines, headings, centering)
 *   - `xlsx:set-print-area`   (Sheet tab: print area range)
 *   - `xlsx:set-print-titles` (Sheet tab: rows/cols repeated on every page)
 *
 * The dialog reads the current sheet's opaque `<pageSetup …/>`,
 * `<pageMargins …/>` and `<printOptions …/>` via tiny regex parsers
 * so the form can show truth instead of defaults; submission only
 * dispatches the commands whose tab values actually changed.
 */

import { useEffect, useMemo, useRef, useState } from "react";
import type { ReactNode } from "react";
import { X } from "lucide-react";
import { useFocusTrap } from "@officeai/ui";
import type { Sheet, XlsxSnapshot } from "@officeai/xlsx";

export type PageSetupTab = "page" | "margins" | "sheet";

export interface PageSetupSubmit {
  readonly setup?: PageSetupValues;
  readonly margins?: MarginsValues;
  readonly printOptions?: PrintOptionsValues;
  readonly printArea?: { range: string | null };
  readonly printTitles?: { rows: string | null; cols: string | null };
}

export interface PageSetupValues {
  readonly orientation: "default" | "portrait" | "landscape";
  readonly paperSize: number;
  readonly scale: number | null;
  readonly fitToWidth: number | null;
  readonly fitToHeight: number | null;
  readonly blackAndWhite: boolean;
  readonly draft: boolean;
}

export interface MarginsValues {
  readonly preset: "normal" | "wide" | "narrow" | "custom";
  readonly leftIn: number;
  readonly rightIn: number;
  readonly topIn: number;
  readonly bottomIn: number;
  readonly headerIn: number;
  readonly footerIn: number;
}

export interface PrintOptionsValues {
  readonly horizontalCentered: boolean;
  readonly verticalCentered: boolean;
  readonly headings: boolean;
  readonly gridLines: boolean;
}

export interface PageSetupDialogProps {
  readonly open: boolean;
  readonly initialTab?: PageSetupTab;
  readonly snapshot: XlsxSnapshot | null;
  readonly sheetName: string | null;
  readonly onClose: () => void;
  readonly onSubmit: (submit: PageSetupSubmit) => void;
}

interface PaperPreset {
  readonly id: number;
  readonly label: string;
}

const PAPER_PRESETS: ReadonlyArray<PaperPreset> = [
  { id: 1, label: 'Letter (8.5" × 11")' },
  { id: 5, label: 'Legal (8.5" × 14")' },
  { id: 3, label: 'Tabloid (11" × 17")' },
  { id: 8, label: "A3 (29.7 × 42 cm)" },
  { id: 9, label: "A4 (21 × 29.7 cm)" },
  { id: 11, label: "A5 (14.8 × 21 cm)" },
  { id: 13, label: "B5 (17.6 × 25 cm)" },
  { id: 70, label: "Envelope DL (11 × 22 cm)" },
];

const MARGIN_PRESETS: Record<"normal" | "wide" | "narrow", Omit<MarginsValues, "preset">> = {
  normal: { leftIn: 0.7, rightIn: 0.7, topIn: 0.75, bottomIn: 0.75, headerIn: 0.3, footerIn: 0.3 },
  wide: { leftIn: 1, rightIn: 1, topIn: 1, bottomIn: 1, headerIn: 0.5, footerIn: 0.5 },
  narrow: { leftIn: 0.25, rightIn: 0.25, topIn: 0.75, bottomIn: 0.75, headerIn: 0.3, footerIn: 0.3 },
};

export function PageSetupDialog(props: PageSetupDialogProps): ReactNode {
  const { open, initialTab, snapshot, sheetName, onClose, onSubmit } = props;
  const sheet = useMemo(
    () => (snapshot && sheetName ? snapshot.root.sheets.find((s) => s.name === sheetName) ?? null : null),
    [snapshot, sheetName]
  );

  const initial = useMemo(() => readSheetState(snapshot, sheet), [snapshot, sheet]);

  const [tab, setTab] = useState<PageSetupTab>(initialTab ?? "page");
  const [setup, setSetup] = useState<PageSetupValues>(initial.setup);
  const [margins, setMargins] = useState<MarginsValues>(initial.margins);
  const [printOpts, setPrintOpts] = useState<PrintOptionsValues>(initial.printOptions);
  const [printArea, setPrintArea] = useState<string>(initial.printArea ?? "");
  const [titleRows, setTitleRows] = useState<string>(initial.titleRows ?? "");
  const [titleCols, setTitleCols] = useState<string>(initial.titleCols ?? "");

  useEffect(() => {
    if (!open) return;
    setTab(initialTab ?? "page");
    setSetup(initial.setup);
    setMargins(initial.margins);
    setPrintOpts(initial.printOptions);
    setPrintArea(initial.printArea ?? "");
    setTitleRows(initial.titleRows ?? "");
    setTitleCols(initial.titleCols ?? "");
  }, [open, initialTab, initial]);

  const panelRef = useRef<HTMLDivElement>(null);
  useFocusTrap(panelRef, { enabled: open, onEscape: onClose });

  if (!open || !sheet) return null;

  const handleApply = (): void => {
    const submit: {
      -readonly [K in keyof PageSetupSubmit]: PageSetupSubmit[K];
    } = {};
    if (!shallowEqualSetup(setup, initial.setup)) {
      submit.setup = setup;
    }
    if (!shallowEqualMargins(margins, initial.margins)) {
      submit.margins = margins;
    }
    if (!shallowEqualPrintOptions(printOpts, initial.printOptions)) {
      submit.printOptions = printOpts;
    }
    const trimmedArea = printArea.trim();
    if (trimmedArea !== (initial.printArea ?? "")) {
      submit.printArea = { range: trimmedArea ? trimmedArea : null };
    }
    const trimmedRows = titleRows.trim();
    const trimmedCols = titleCols.trim();
    if (
      trimmedRows !== (initial.titleRows ?? "") ||
      trimmedCols !== (initial.titleCols ?? "")
    ) {
      submit.printTitles = {
        rows: trimmedRows ? trimmedRows : null,
        cols: trimmedCols ? trimmedCols : null,
      };
    }
    onSubmit(submit);
    onClose();
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="xlsx-page-setup-title"
      data-page-setup-dialog
      className="fixed inset-0 z-[60] flex items-center justify-center bg-black/40 px-4 py-6"
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div
        ref={panelRef}
        tabIndex={-1}
        className="flex max-h-[85vh] w-full max-w-2xl flex-col rounded-lg border border-divider bg-surface shadow-2xl outline-none"
        onClick={(e) => e.stopPropagation()}
      >
        <header className="flex items-center justify-between gap-3 border-b border-divider px-5 py-3">
          <div>
            <h2 id="xlsx-page-setup-title" className="text-base font-semibold">
              Page setup — {sheet.name}
            </h2>
            <p className="text-xs text-secondary">
              Configure orientation, margins, print options and titles for this worksheet.
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

        <div className="flex gap-4 border-b border-divider px-5">
          <TabBtn label="Page" active={tab === "page"} onClick={() => setTab("page")} testId="page-setup-tab-page" />
          <TabBtn label="Margins" active={tab === "margins"} onClick={() => setTab("margins")} testId="page-setup-tab-margins" />
          <TabBtn label="Sheet" active={tab === "sheet"} onClick={() => setTab("sheet")} testId="page-setup-tab-sheet" />
        </div>

        <div className="flex flex-1 flex-col gap-4 overflow-y-auto px-5 py-4 text-sm">
          {tab === "page" ? <PageTab values={setup} onChange={setSetup} /> : null}
          {tab === "margins" ? <MarginsTab values={margins} onChange={setMargins} /> : null}
          {tab === "sheet" ? (
            <SheetTab
              printOpts={printOpts}
              onPrintOptsChange={setPrintOpts}
              printArea={printArea}
              onPrintAreaChange={setPrintArea}
              titleRows={titleRows}
              titleCols={titleCols}
              onTitleRowsChange={setTitleRows}
              onTitleColsChange={setTitleCols}
            />
          ) : null}
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

function TabBtn(props: { label: string; active: boolean; onClick: () => void; testId: string }): ReactNode {
  const { label, active, onClick, testId } = props;
  return (
    <button
      type="button"
      data-testid={testId}
      role="tab"
      aria-selected={active}
      onClick={onClick}
      className={`-mb-px border-b-2 px-1 py-2 text-xs font-medium transition-colors ${
        active
          ? "border-accent text-default"
          : "border-transparent text-secondary hover:text-default"
      }`}
    >
      {label}
    </button>
  );
}

function PageTab(props: { values: PageSetupValues; onChange: (v: PageSetupValues) => void }): ReactNode {
  const { values, onChange } = props;
  const fitMode = (values.fitToWidth ?? 0) > 0 || (values.fitToHeight ?? 0) > 0;
  return (
    <>
      <fieldset className="flex flex-col gap-2">
        <legend className="text-xs font-medium uppercase tracking-wide text-secondary">Orientation</legend>
        <div className="flex gap-2">
          <OrientButton
            active={values.orientation === "portrait"}
            label="Portrait"
            onClick={() => onChange({ ...values, orientation: "portrait" })}
          />
          <OrientButton
            active={values.orientation === "landscape"}
            label="Landscape"
            onClick={() => onChange({ ...values, orientation: "landscape" })}
          />
        </div>
      </fieldset>

      <fieldset className="flex flex-col gap-2">
        <legend className="text-xs font-medium uppercase tracking-wide text-secondary">Paper size</legend>
        <select
          value={values.paperSize}
          onChange={(e) => onChange({ ...values, paperSize: Number(e.target.value) })}
          className="rounded border border-divider bg-background px-2 py-1.5"
          data-testid="page-setup-paper"
        >
          {PAPER_PRESETS.map((p) => (
            <option key={p.id} value={p.id}>
              {p.label}
            </option>
          ))}
        </select>
      </fieldset>

      <fieldset className="flex flex-col gap-2">
        <legend className="text-xs font-medium uppercase tracking-wide text-secondary">Scaling</legend>
        <div className="flex flex-col gap-2">
          <label className="flex items-center gap-2">
            <input
              type="radio"
              checked={!fitMode}
              onChange={() =>
                onChange({ ...values, fitToWidth: null, fitToHeight: null, scale: values.scale ?? 100 })
              }
            />
            <span>Adjust to</span>
            <input
              type="number"
              min={10}
              max={400}
              value={values.scale ?? 100}
              onChange={(e) => onChange({ ...values, scale: Number(e.currentTarget.value) })}
              disabled={fitMode}
              className="w-20 rounded border border-divider bg-background px-2 py-1 tabular-nums disabled:opacity-50"
              data-testid="page-setup-scale"
            />
            <span>% normal size</span>
          </label>
          <label className="flex items-center gap-2">
            <input
              type="radio"
              checked={fitMode}
              onChange={() => onChange({ ...values, fitToWidth: 1, fitToHeight: 1, scale: null })}
            />
            <span>Fit to</span>
            <input
              type="number"
              min={1}
              max={100}
              value={values.fitToWidth ?? 1}
              onChange={(e) => onChange({ ...values, fitToWidth: Number(e.currentTarget.value) })}
              disabled={!fitMode}
              className="w-16 rounded border border-divider bg-background px-2 py-1 tabular-nums disabled:opacity-50"
              data-testid="page-setup-fit-width"
            />
            <span>page(s) wide by</span>
            <input
              type="number"
              min={1}
              max={100}
              value={values.fitToHeight ?? 1}
              onChange={(e) => onChange({ ...values, fitToHeight: Number(e.currentTarget.value) })}
              disabled={!fitMode}
              className="w-16 rounded border border-divider bg-background px-2 py-1 tabular-nums disabled:opacity-50"
              data-testid="page-setup-fit-height"
            />
            <span>tall</span>
          </label>
        </div>
      </fieldset>

      <fieldset className="flex flex-col gap-2">
        <legend className="text-xs font-medium uppercase tracking-wide text-secondary">Print quality</legend>
        <label className="flex items-center gap-2">
          <input
            type="checkbox"
            checked={values.blackAndWhite}
            onChange={(e) => onChange({ ...values, blackAndWhite: e.currentTarget.checked })}
            data-testid="page-setup-bw"
          />
          <span>Black and white</span>
        </label>
        <label className="flex items-center gap-2">
          <input
            type="checkbox"
            checked={values.draft}
            onChange={(e) => onChange({ ...values, draft: e.currentTarget.checked })}
            data-testid="page-setup-draft"
          />
          <span>Draft quality</span>
        </label>
      </fieldset>
    </>
  );
}

function MarginsTab(props: { values: MarginsValues; onChange: (v: MarginsValues) => void }): ReactNode {
  const { values, onChange } = props;
  return (
    <>
      <fieldset className="flex flex-col gap-2">
        <legend className="text-xs font-medium uppercase tracking-wide text-secondary">Preset</legend>
        <div className="flex flex-wrap gap-1.5">
          {(["normal", "wide", "narrow"] as const).map((id) => (
            <button
              key={id}
              type="button"
              data-testid={`page-setup-margin-${id}`}
              aria-pressed={values.preset === id}
              onClick={() => onChange({ preset: id, ...MARGIN_PRESETS[id] })}
              className={`rounded border px-2 py-1 text-xs ${
                values.preset === id
                  ? "border-accent bg-accent-light text-default"
                  : "border-divider bg-background hover:bg-hover"
              }`}
            >
              {id[0]?.toUpperCase()}{id.slice(1)}
            </button>
          ))}
        </div>
      </fieldset>
      <fieldset className="grid grid-cols-2 gap-2">
        <legend className="col-span-2 text-xs font-medium uppercase tracking-wide text-secondary">
          Custom (inches)
        </legend>
        <NumField label="Top" value={values.topIn} onChange={(v) => onChange({ ...values, preset: "custom", topIn: v })} />
        <NumField label="Bottom" value={values.bottomIn} onChange={(v) => onChange({ ...values, preset: "custom", bottomIn: v })} />
        <NumField label="Left" value={values.leftIn} onChange={(v) => onChange({ ...values, preset: "custom", leftIn: v })} />
        <NumField label="Right" value={values.rightIn} onChange={(v) => onChange({ ...values, preset: "custom", rightIn: v })} />
        <NumField label="Header" value={values.headerIn} onChange={(v) => onChange({ ...values, preset: "custom", headerIn: v })} />
        <NumField label="Footer" value={values.footerIn} onChange={(v) => onChange({ ...values, preset: "custom", footerIn: v })} />
      </fieldset>
    </>
  );
}

function SheetTab(props: {
  printOpts: PrintOptionsValues;
  onPrintOptsChange: (v: PrintOptionsValues) => void;
  printArea: string;
  onPrintAreaChange: (v: string) => void;
  titleRows: string;
  titleCols: string;
  onTitleRowsChange: (v: string) => void;
  onTitleColsChange: (v: string) => void;
}): ReactNode {
  const { printOpts, onPrintOptsChange, printArea, onPrintAreaChange, titleRows, titleCols, onTitleRowsChange, onTitleColsChange } = props;
  return (
    <>
      <fieldset className="flex flex-col gap-2">
        <legend className="text-xs font-medium uppercase tracking-wide text-secondary">Print area</legend>
        <input
          type="text"
          placeholder="A1:D20"
          value={printArea}
          onChange={(e) => onPrintAreaChange(e.currentTarget.value)}
          className="rounded border border-divider bg-background px-2 py-1 font-mono text-xs"
          data-testid="page-setup-print-area"
        />
        <p className="text-[11px] text-tertiary">Leave blank to clear the print area.</p>
      </fieldset>

      <fieldset className="grid grid-cols-2 gap-2">
        <legend className="col-span-2 text-xs font-medium uppercase tracking-wide text-secondary">
          Print titles (repeat on each page)
        </legend>
        <label className="flex flex-col gap-1">
          <span className="text-xs text-secondary">Rows to repeat at top</span>
          <input
            type="text"
            placeholder="1:1"
            value={titleRows}
            onChange={(e) => onTitleRowsChange(e.currentTarget.value)}
            className="rounded border border-divider bg-background px-2 py-1 font-mono text-xs"
            data-testid="page-setup-title-rows"
          />
        </label>
        <label className="flex flex-col gap-1">
          <span className="text-xs text-secondary">Columns to repeat at left</span>
          <input
            type="text"
            placeholder="A:A"
            value={titleCols}
            onChange={(e) => onTitleColsChange(e.currentTarget.value)}
            className="rounded border border-divider bg-background px-2 py-1 font-mono text-xs"
            data-testid="page-setup-title-cols"
          />
        </label>
      </fieldset>

      <fieldset className="flex flex-col gap-2">
        <legend className="text-xs font-medium uppercase tracking-wide text-secondary">Print</legend>
        <CheckRow
          label="Gridlines"
          checked={printOpts.gridLines}
          onChange={(v) => onPrintOptsChange({ ...printOpts, gridLines: v })}
          testId="page-setup-print-gridlines"
        />
        <CheckRow
          label="Row and column headings"
          checked={printOpts.headings}
          onChange={(v) => onPrintOptsChange({ ...printOpts, headings: v })}
          testId="page-setup-print-headings"
        />
      </fieldset>

      <fieldset className="flex flex-col gap-2">
        <legend className="text-xs font-medium uppercase tracking-wide text-secondary">Center on page</legend>
        <CheckRow
          label="Horizontally"
          checked={printOpts.horizontalCentered}
          onChange={(v) => onPrintOptsChange({ ...printOpts, horizontalCentered: v })}
          testId="page-setup-center-h"
        />
        <CheckRow
          label="Vertically"
          checked={printOpts.verticalCentered}
          onChange={(v) => onPrintOptsChange({ ...printOpts, verticalCentered: v })}
          testId="page-setup-center-v"
        />
      </fieldset>
    </>
  );
}

function OrientButton(props: { active: boolean; label: string; onClick: () => void }): ReactNode {
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

function NumField(props: { label: string; value: number; onChange: (v: number) => void }): ReactNode {
  const { label, value, onChange } = props;
  return (
    <label className="flex flex-col gap-1">
      <span className="text-xs text-secondary">{label}</span>
      <input
        type="number"
        inputMode="decimal"
        step={0.05}
        min={0}
        value={value.toFixed(2)}
        onChange={(e) => {
          const v = Number(e.currentTarget.value);
          if (Number.isFinite(v) && v >= 0) onChange(Math.round(v * 100) / 100);
        }}
        className="rounded border border-divider bg-background px-2 py-1 tabular-nums"
      />
    </label>
  );
}

function CheckRow(props: { label: string; checked: boolean; onChange: (v: boolean) => void; testId: string }): ReactNode {
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

interface InitialState {
  readonly setup: PageSetupValues;
  readonly margins: MarginsValues;
  readonly printOptions: PrintOptionsValues;
  readonly printArea: string | null;
  readonly titleRows: string | null;
  readonly titleCols: string | null;
}

function readSheetState(snapshot: XlsxSnapshot | null, sheet: Sheet | null): InitialState {
  return {
    setup: parsePageSetup(sheet?.pageSetupXml),
    margins: parsePageMargins(sheet?.pageMarginsXml),
    printOptions: parsePrintOptions(sheet?.printOptionsXml),
    ...parsePrintNames(snapshot, sheet),
  };
}

function parsePageSetup(xml: string | undefined): PageSetupValues {
  const attrs = parseAttrs(xml);
  const orientation = (attrs.get("orientation") as PageSetupValues["orientation"]) ?? "default";
  const paperSize = Number(attrs.get("paperSize") ?? 9);
  const scaleAttr = attrs.get("scale");
  const fwAttr = attrs.get("fitToWidth");
  const fhAttr = attrs.get("fitToHeight");
  return {
    orientation,
    paperSize: Number.isFinite(paperSize) && paperSize > 0 ? paperSize : 9,
    scale: scaleAttr !== undefined ? Number(scaleAttr) : 100,
    fitToWidth: fwAttr !== undefined ? Number(fwAttr) : null,
    fitToHeight: fhAttr !== undefined ? Number(fhAttr) : null,
    blackAndWhite: attrs.get("blackAndWhite") === "1" || attrs.get("blackAndWhite") === "true",
    draft: attrs.get("draft") === "1" || attrs.get("draft") === "true",
  };
}

function parsePageMargins(xml: string | undefined): MarginsValues {
  const attrs = parseAttrs(xml);
  const margins = {
    leftIn: Number(attrs.get("left") ?? 0.7),
    rightIn: Number(attrs.get("right") ?? 0.7),
    topIn: Number(attrs.get("top") ?? 0.75),
    bottomIn: Number(attrs.get("bottom") ?? 0.75),
    headerIn: Number(attrs.get("header") ?? 0.3),
    footerIn: Number(attrs.get("footer") ?? 0.3),
  };
  const preset = matchMarginPreset(margins);
  return { preset, ...margins };
}

function matchMarginPreset(m: Omit<MarginsValues, "preset">): MarginsValues["preset"] {
  for (const id of ["normal", "wide", "narrow"] as const) {
    const p = MARGIN_PRESETS[id];
    if (
      Math.abs(p.leftIn - m.leftIn) < 0.005 &&
      Math.abs(p.rightIn - m.rightIn) < 0.005 &&
      Math.abs(p.topIn - m.topIn) < 0.005 &&
      Math.abs(p.bottomIn - m.bottomIn) < 0.005 &&
      Math.abs(p.headerIn - m.headerIn) < 0.005 &&
      Math.abs(p.footerIn - m.footerIn) < 0.005
    ) {
      return id;
    }
  }
  return "custom";
}

function parsePrintOptions(xml: string | undefined): PrintOptionsValues {
  const attrs = parseAttrs(xml);
  const truthy = (v: string | undefined): boolean => v === "1" || v === "true";
  return {
    horizontalCentered: truthy(attrs.get("horizontalCentered")),
    verticalCentered: truthy(attrs.get("verticalCentered")),
    headings: truthy(attrs.get("headings")),
    gridLines: truthy(attrs.get("gridLines")),
  };
}

function parsePrintNames(
  snapshot: XlsxSnapshot | null,
  sheet: Sheet | null
): { printArea: string | null; titleRows: string | null; titleCols: string | null } {
  if (!snapshot || !sheet) return { printArea: null, titleRows: null, titleCols: null };
  const print = snapshot.root.definedNames.find(
    (n) => n.name === "_xlnm.Print_Area" && n.scope === sheet.name
  );
  const titles = snapshot.root.definedNames.find(
    (n) => n.name === "_xlnm.Print_Titles" && n.scope === sheet.name
  );
  const printArea = print ? stripSheetPrefix(print.refersTo, sheet.name) : null;
  const { rows, cols } = titles ? splitTitles(titles.refersTo, sheet.name) : { rows: null, cols: null };
  return { printArea, titleRows: rows, titleCols: cols };
}

function stripSheetPrefix(refersTo: string, sheetName: string): string {
  const re = new RegExp(`(?:^|,)\\s*(?:'?${escapeForRegex(sheetName)}'?!)([^,]+)`, "g");
  const parts: string[] = [];
  let m: RegExpExecArray | null;
  while ((m = re.exec(refersTo)) !== null) {
    if (m[1]) parts.push(m[1].replace(/\$/g, ""));
  }
  return parts.length ? parts.join(",") : refersTo;
}

function splitTitles(refersTo: string, sheetName: string): { rows: string | null; cols: string | null } {
  const stripped = stripSheetPrefix(refersTo, sheetName);
  let rows: string | null = null;
  let cols: string | null = null;
  for (const part of stripped.split(",").map((s) => s.trim())) {
    if (!part) continue;
    if (/^\d+:\d+$/.test(part)) rows = part;
    else if (/^[A-Za-z]+:[A-Za-z]+$/.test(part)) cols = part;
  }
  return { rows, cols };
}

function escapeForRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function parseAttrs(xml: string | undefined): Map<string, string> {
  const out = new Map<string, string>();
  if (!xml) return out;
  const re = /([a-zA-Z:]+)\s*=\s*"([^"]*)"/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml)) !== null) {
    if (m[1] && m[2] !== undefined) out.set(m[1], m[2]);
  }
  return out;
}

function shallowEqualSetup(a: PageSetupValues, b: PageSetupValues): boolean {
  return (
    a.orientation === b.orientation &&
    a.paperSize === b.paperSize &&
    a.scale === b.scale &&
    a.fitToWidth === b.fitToWidth &&
    a.fitToHeight === b.fitToHeight &&
    a.blackAndWhite === b.blackAndWhite &&
    a.draft === b.draft
  );
}

function shallowEqualMargins(a: MarginsValues, b: MarginsValues): boolean {
  return (
    Math.abs(a.leftIn - b.leftIn) < 0.001 &&
    Math.abs(a.rightIn - b.rightIn) < 0.001 &&
    Math.abs(a.topIn - b.topIn) < 0.001 &&
    Math.abs(a.bottomIn - b.bottomIn) < 0.001 &&
    Math.abs(a.headerIn - b.headerIn) < 0.001 &&
    Math.abs(a.footerIn - b.footerIn) < 0.001
  );
}

function shallowEqualPrintOptions(a: PrintOptionsValues, b: PrintOptionsValues): boolean {
  return (
    a.horizontalCentered === b.horizontalCentered &&
    a.verticalCentered === b.verticalCentered &&
    a.headings === b.headings &&
    a.gridLines === b.gridLines
  );
}
