"use client";

import {
  useEffect,
  useMemo,
  useRef,
  useState,
  type ReactNode,
} from "react";
import {
  FileCode,
  FileImage,
  FileSpreadsheet,
  FileText,
  FileType2,
  Presentation,
  X,
} from "lucide-react";
import { cn } from "@officeai/ui";
import { InlineSpinner } from "./InlineSpinner";
import type {
  ExportFormat,
  ExportFormatGroup,
  ExportFormatIcon,
  ExportFormatOptionField,
  ExportOptionValue,
  ExportOptionValues,
} from "./types";

export interface ExportDialogProps {
  readonly open: boolean;
  readonly onClose: () => void;
  /** All formats the product supports — both `instant` and `dialog`
   * formats are listed; the dialog is the place users go when they
   * want to compare options. */
  readonly formats: ReadonlyArray<ExportFormat>;
  /** Initially selected format id. */
  readonly initialFormatId?: string;
  /** Document name without extension — used to preview the resulting
   * filename. */
  readonly baseFilename: string;
  /** Fired when the user clicks Export. The dialog stays open until
   * the promise resolves so the product can show progress; closes on
   * success, surfaces the error otherwise. */
  readonly onExport: (
    format: ExportFormat,
    options: ExportOptionValues
  ) => Promise<void> | void;
}

const GROUP_ORDER: ReadonlyArray<ExportFormatGroup> = [
  "deck",
  "native",
  "pdf-web",
  "data",
  "images",
  "current",
];

const GROUP_LABEL: Record<ExportFormatGroup, string> = {
  deck: "Whole deck",
  native: "Native format",
  "pdf-web": "PDF & web",
  data: "Data",
  images: "Images",
  current: "This slide",
};

/**
 * Rich Export dialog rendered above all editors when the user picks
 * a format with `kind: "dialog"` (or opens the dialog explicitly
 * from the dropdown footer).
 *
 * Layout:
 *   ┌────────────────────────────────────────────────┐
 *   │ Export                                       × │
 *   ├──────────────┬─────────────────────────────────┤
 *   │ Native       │  Format: PDF                    │
 *   │ • DOCX       │  Description …                  │
 *   │ PDF & web    │  ─────────────                  │
 *   │ • PDF        │  Page size [A4 ▾]               │
 *   │ • HTML       │  Orientation [Portrait ▾]       │
 *   │ Data         │                                 │
 *   │ • TXT        │                                 │
 *   │ • Markdown   │                                 │
 *   ├──────────────┴─────────────────────────────────┤
 *   │ filename.pdf            [ Cancel ] [ Export ]  │
 *   └────────────────────────────────────────────────┘
 */
export function ExportDialog({
  open,
  onClose,
  formats,
  initialFormatId,
  baseFilename,
  onExport,
}: ExportDialogProps): ReactNode {
  const grouped = useMemo(() => groupFormats(formats), [formats]);
  const [selectedId, setSelectedId] = useState<string | undefined>(
    initialFormatId ?? formats[0]?.id
  );
  const [optionState, setOptionState] = useState<
    Record<string, ExportOptionValues>
  >(() => seedOptionState(formats));
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const containerRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    if (!open) return;
    setSelectedId(initialFormatId ?? formats[0]?.id);
    setOptionState(seedOptionState(formats));
    setError(null);
    setBusy(false);
  }, [open, initialFormatId, formats]);

  useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape" && !busy) {
        e.preventDefault();
        onClose();
      }
    };
    document.addEventListener("keydown", onKey);
    return () => document.removeEventListener("keydown", onKey);
  }, [open, onClose, busy]);

  if (!open) return null;

  const selected = formats.find((f) => f.id === selectedId) ?? formats[0];
  if (!selected) return null;

  const filenamePreview = `${baseFilename || "Untitled"}.${selected.extension}`;
  const fields = selected.optionFields ?? [];
  const values = optionState[selected.id] ?? {};

  const handleConfirm = async () => {
    if (busy) return;
    setBusy(true);
    setError(null);
    try {
      await onExport(selected, values);
      setBusy(false);
      onClose();
    } catch (err) {
      setBusy(false);
      setError(err instanceof Error ? err.message : "Export failed.");
    }
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label="Export"
      className="fixed inset-0 z-40 flex items-start justify-center bg-black/30 px-4 pt-[10vh]"
      onMouseDown={(e) => {
        if (busy) return;
        if (e.target === e.currentTarget) onClose();
      }}
      data-testid="shell-export-dialog"
    >
      <div
        ref={containerRef}
        className="flex max-h-[80vh] w-full max-w-[720px] flex-col overflow-hidden rounded-lg border border-divider bg-background shadow-xl"
      >
        {/* Header */}
        <div className="flex h-11 items-center justify-between border-b border-divider px-3">
          <span className="text-sm font-medium text-foreground">Export</span>
          <button
            type="button"
            onClick={() => !busy && onClose()}
            disabled={busy}
            className="inline-flex h-7 w-7 items-center justify-center rounded-md text-secondary hover:bg-hover hover:text-foreground disabled:opacity-40"
            aria-label="Close"
            data-testid="shell-export-dialog-close"
          >
            <X size={14} />
          </button>
        </div>

        {/* Body: format list + options */}
        <div className="flex min-h-0 flex-1">
          <FormatList
            grouped={grouped}
            selectedId={selected.id}
            onSelect={setSelectedId}
          />
          <OptionsPane
            format={selected}
            fields={fields}
            values={values}
            onChange={(fieldId, next) =>
              setOptionState((prev) => ({
                ...prev,
                [selected.id]: { ...(prev[selected.id] ?? {}), [fieldId]: next },
              }))
            }
          />
        </div>

        {/* Footer */}
        <div className="flex items-center justify-between gap-2 border-t border-divider bg-surface px-3 py-2">
          <div className="min-w-0 flex-1 truncate text-xs text-secondary">
            <span className="text-tertiary">Saving as </span>
            <span className="font-medium text-foreground" data-testid="shell-export-filename-preview">
              {filenamePreview}
            </span>
            {error ? (
              <span className="ml-2 text-[color:var(--error)]" role="alert">
                {error}
              </span>
            ) : null}
          </div>
          <button
            type="button"
            onClick={() => !busy && onClose()}
            disabled={busy}
            className="inline-flex h-8 items-center rounded-md px-3 text-sm text-secondary hover:bg-hover hover:text-foreground disabled:opacity-40"
            data-testid="shell-export-cancel"
          >
            Cancel
          </button>
          <button
            type="button"
            onClick={() => void handleConfirm()}
            disabled={busy}
            className={cn(
              "inline-flex h-8 items-center gap-1.5 rounded-md bg-[color:var(--accent)] px-3 text-sm font-medium text-white shadow-sm",
              "hover:bg-[color:var(--accent-hover,var(--accent))]",
              "focus:outline-none focus-visible:ring-2 focus-visible:ring-[color:var(--accent)]/40",
              busy && "opacity-60"
            )}
            data-testid="shell-export-confirm"
          >
            {busy ? (
              <>
                <InlineSpinner size={12} />
                Exporting…
              </>
            ) : (
              `Export ${selected.extension.toUpperCase()}`
            )}
          </button>
        </div>
      </div>
    </div>
  );
}

function FormatList({
  grouped,
  selectedId,
  onSelect,
}: {
  readonly grouped: ReadonlyArray<{
    readonly group: ExportFormatGroup;
    readonly items: ReadonlyArray<ExportFormat>;
  }>;
  readonly selectedId: string;
  readonly onSelect: (id: string) => void;
}): ReactNode {
  return (
    <nav
      className="w-[220px] shrink-0 overflow-y-auto border-r border-divider bg-surface p-2"
      aria-label="Export format"
    >
      {grouped.map(({ group, items }) => (
        <div key={group} className="mb-2 last:mb-0">
          <div className="px-2 pb-1 text-[11px] font-semibold uppercase tracking-wide text-tertiary">
            {GROUP_LABEL[group]}
          </div>
          {items.map((fmt) => {
            const isActive = fmt.id === selectedId;
            return (
              <button
                key={fmt.id}
                type="button"
                onClick={() => onSelect(fmt.id)}
                className={cn(
                  "flex w-full items-center gap-2 rounded-md px-2 py-1.5 text-left text-sm",
                  isActive
                    ? "bg-[color:var(--accent-light,var(--hover))] text-foreground"
                    : "text-foreground hover:bg-hover"
                )}
                data-testid={`shell-export-format-${fmt.id}`}
                aria-current={isActive ? "true" : undefined}
              >
                <span className="text-secondary">
                  <FormatIcon icon={fmt.icon ?? guessIcon(fmt)} />
                </span>
                <span className="min-w-0 flex-1 truncate">{fmt.label}</span>
              </button>
            );
          })}
        </div>
      ))}
    </nav>
  );
}

function OptionsPane({
  format,
  fields,
  values,
  onChange,
}: {
  readonly format: ExportFormat;
  readonly fields: ReadonlyArray<ExportFormatOptionField>;
  readonly values: ExportOptionValues;
  readonly onChange: (fieldId: string, next: ExportOptionValue) => void;
}): ReactNode {
  return (
    <div className="flex min-h-0 flex-1 flex-col overflow-y-auto p-4">
      <div className="mb-3">
        <div className="text-base font-medium text-foreground">{format.label}</div>
        {format.description ? (
          <p className="mt-1 text-xs text-secondary">{format.description}</p>
        ) : null}
      </div>
      {fields.length === 0 ? (
        <p className="text-xs text-tertiary">No options for this format.</p>
      ) : (
        <div className="flex flex-col gap-3">
          {fields.map((field) => (
            <OptionField
              key={field.id}
              field={field}
              value={values[field.id]}
              onChange={(next) => onChange(field.id, next)}
            />
          ))}
        </div>
      )}
    </div>
  );
}

function OptionField({
  field,
  value,
  onChange,
}: {
  readonly field: ExportFormatOptionField;
  readonly value: ExportOptionValue | undefined;
  readonly onChange: (next: ExportOptionValue) => void;
}): ReactNode {
  const labelId = `export-opt-${field.id}`;
  switch (field.control.type) {
    case "select": {
      const current = typeof value === "string" ? value : field.control.defaultId;
      return (
        <label className="flex flex-col gap-1 text-sm text-foreground">
          <span id={labelId} className="text-xs font-medium text-secondary">
            {field.label}
          </span>
          <select
            aria-labelledby={labelId}
            value={current}
            onChange={(e) => onChange(e.target.value)}
            className="h-8 rounded-md border border-divider bg-background px-2 text-sm text-foreground focus:border-[color:var(--accent)] focus:outline-none"
            data-testid={`shell-export-option-${field.id}`}
          >
            {field.control.options.map((opt) => (
              <option key={opt.id} value={opt.id}>
                {opt.label}
              </option>
            ))}
          </select>
          {field.hint ? <span className="text-[11px] text-tertiary">{field.hint}</span> : null}
        </label>
      );
    }
    case "toggle": {
      const current =
        typeof value === "boolean" ? value : field.control.defaultValue;
      return (
        <label className="flex items-center gap-2 text-sm text-foreground">
          <input
            type="checkbox"
            checked={current}
            onChange={(e) => onChange(e.target.checked)}
            className="h-4 w-4 rounded border-divider text-[color:var(--accent)] focus:ring-[color:var(--accent)]/40"
            data-testid={`shell-export-option-${field.id}`}
          />
          <span>{field.label}</span>
          {field.hint ? <span className="text-[11px] text-tertiary">{field.hint}</span> : null}
        </label>
      );
    }
    case "text": {
      const current =
        typeof value === "string" ? value : field.control.defaultValue ?? "";
      return (
        <label className="flex flex-col gap-1 text-sm text-foreground">
          <span id={labelId} className="text-xs font-medium text-secondary">
            {field.label}
          </span>
          <input
            type="text"
            aria-labelledby={labelId}
            value={current}
            onChange={(e) => onChange(e.target.value)}
            placeholder={field.control.placeholder}
            className="h-8 rounded-md border border-divider bg-background px-2 text-sm text-foreground focus:border-[color:var(--accent)] focus:outline-none"
            data-testid={`shell-export-option-${field.id}`}
          />
          {field.hint ? <span className="text-[11px] text-tertiary">{field.hint}</span> : null}
        </label>
      );
    }
    case "multiSelect": {
      const current = Array.isArray(value) ? value : field.control.defaultIds;
      const set = new Set(current);
      return (
        <fieldset className="flex flex-col gap-1 text-sm text-foreground">
          <legend className="text-xs font-medium text-secondary">{field.label}</legend>
          <div
            className="flex max-h-[140px] flex-col gap-1 overflow-y-auto rounded-md border border-divider bg-background p-2"
            data-testid={`shell-export-option-${field.id}`}
          >
            {field.control.options.map((opt) => {
              const checked = set.has(opt.id);
              return (
                <label key={opt.id} className="flex items-center gap-2 text-sm text-foreground">
                  <input
                    type="checkbox"
                    checked={checked}
                    onChange={(e) => {
                      const nextSet = new Set(current);
                      if (e.target.checked) nextSet.add(opt.id);
                      else nextSet.delete(opt.id);
                      onChange(Array.from(nextSet));
                    }}
                    className="h-4 w-4 rounded border-divider text-[color:var(--accent)] focus:ring-[color:var(--accent)]/40"
                  />
                  <span>{opt.label}</span>
                </label>
              );
            })}
          </div>
          {field.hint ? <span className="text-[11px] text-tertiary">{field.hint}</span> : null}
        </fieldset>
      );
    }
    default: {
      const _exhaustive: never = field.control;
      void _exhaustive;
      return null;
    }
  }
}

function FormatIcon({ icon }: { readonly icon: ExportFormatIcon }): ReactNode {
  switch (icon) {
    case "doc":
      return <FileText size={14} />;
    case "sheet":
      return <FileSpreadsheet size={14} />;
    case "slides":
      return <Presentation size={14} />;
    case "pdf":
      return <FileType2 size={14} />;
    case "image":
      return <FileImage size={14} />;
    case "code":
      return <FileCode size={14} />;
    case "text":
      return <FileText size={14} />;
    default: {
      const _exhaustive: never = icon;
      void _exhaustive;
      return <FileText size={14} />;
    }
  }
}

function guessIcon(format: ExportFormat): ExportFormatIcon {
  switch (format.extension) {
    case "docx":
      return "doc";
    case "xlsx":
      return "sheet";
    case "pptx":
      return "slides";
    case "pdf":
      return "pdf";
    case "html":
    case "json":
      return "code";
    case "csv":
    case "tsv":
    case "txt":
    case "md":
      return "text";
    case "png":
    case "jpg":
    case "jpeg":
    case "svg":
    case "zip":
      return "image";
    default:
      return "doc";
  }
}

function groupFormats(
  formats: ReadonlyArray<ExportFormat>
): ReadonlyArray<{
  readonly group: ExportFormatGroup;
  readonly items: ReadonlyArray<ExportFormat>;
}> {
  const buckets = new Map<ExportFormatGroup, ExportFormat[]>();
  for (const fmt of formats) {
    const g = fmt.group ?? defaultGroup(fmt);
    const list = buckets.get(g) ?? [];
    list.push(fmt);
    buckets.set(g, list);
  }
  return GROUP_ORDER.filter((g) => (buckets.get(g)?.length ?? 0) > 0).map((g) => ({
    group: g,
    items: buckets.get(g) ?? [],
  }));
}

function defaultGroup(format: ExportFormat): ExportFormatGroup {
  switch (format.extension) {
    case "docx":
    case "xlsx":
    case "pptx":
      return "native";
    case "pdf":
    case "html":
      return "pdf-web";
    case "csv":
    case "tsv":
    case "json":
    case "txt":
    case "md":
      return "data";
    default:
      return "images";
  }
}

function seedOptionState(
  formats: ReadonlyArray<ExportFormat>
): Record<string, ExportOptionValues> {
  const out: Record<string, ExportOptionValues> = {};
  for (const fmt of formats) {
    if (!fmt.optionFields || fmt.optionFields.length === 0) continue;
    const seed: Record<string, ExportOptionValue> = {};
    for (const field of fmt.optionFields) {
      switch (field.control.type) {
        case "select":
          seed[field.id] = field.control.defaultId;
          break;
        case "toggle":
          seed[field.id] = field.control.defaultValue;
          break;
        case "text":
          seed[field.id] = field.control.defaultValue ?? "";
          break;
        case "multiSelect":
          seed[field.id] = field.control.defaultIds;
          break;
        default: {
          const _exhaustive: never = field.control;
          void _exhaustive;
        }
      }
    }
    out[fmt.id] = seed;
  }
  return out;
}
