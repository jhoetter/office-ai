"use client";

import * as React from "react";
import type { PdfMetadata } from "@officeai/pdf";
import { useTranslator } from "@/lib/i18n";

/**
 * Phase 9d "Document properties" dialog.
 *
 * Wraps `pdf:set-metadata` with the four user-facing fields that
 * Acrobat surfaces in its own properties sheet (title / author /
 * subject / keywords). The remaining metadata fields exposed by the
 * PDF spec — creator, producer, creation/modification dates,
 * pdfVersion, linearized, encryption — are read-only by convention
 * (Acrobat shows them in a "Description" tab but doesn't let users
 * edit them either) so we render them as plain text below the form.
 *
 * The dialog opens with the current snapshot's metadata pre-filled
 * and dispatches `pdf:set-metadata` only with the fields that
 * actually changed (so the resulting diff and the round-trip are as
 * minimal as possible).
 */
export interface PdfMetadataDialogProps {
  readonly open: boolean;
  readonly metadata: PdfMetadata;
  readonly onClose: () => void;
  readonly onSubmit: (patch: Partial<PdfMetadata>) => void;
}

export function PdfMetadataDialog({
  open,
  metadata,
  onClose,
  onSubmit,
}: PdfMetadataDialogProps): React.ReactElement | null {
  const { t } = useTranslator();
  const [title, setTitle] = React.useState(metadata.title ?? "");
  const [author, setAuthor] = React.useState(metadata.author ?? "");
  const [subject, setSubject] = React.useState(metadata.subject ?? "");
  const [keywords, setKeywords] = React.useState(metadata.keywords ?? "");
  const firstFieldRef = React.useRef<HTMLInputElement>(null);

  React.useEffect(() => {
    if (!open) return;
    setTitle(metadata.title ?? "");
    setAuthor(metadata.author ?? "");
    setSubject(metadata.subject ?? "");
    setKeywords(metadata.keywords ?? "");
    // Defer focus so the input exists in the DOM by the time we
    // ask for it (the dialog mounts on the same tick as `open`).
    queueMicrotask(() => firstFieldRef.current?.focus());
  }, [open, metadata]);

  React.useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === "Escape") onClose();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [open, onClose]);

  if (!open) return null;

  const handleSubmit = (e: React.FormEvent): void => {
    e.preventDefault();
    // `Partial<PdfMetadata>` inherits the parent's `readonly` field
    // modifiers, so we accumulate the diff in a mutable bag and cast
    // once at the boundary. The cast is sound — `onSubmit` only reads
    // the patch and merges it into a fresh metadata object inside
    // `pdf:set-metadata`'s handler.
    const mutable: Record<string, string> = {};
    if (title !== (metadata.title ?? "")) mutable.title = title;
    if (author !== (metadata.author ?? "")) mutable.author = author;
    if (subject !== (metadata.subject ?? "")) mutable.subject = subject;
    if (keywords !== (metadata.keywords ?? "")) mutable.keywords = keywords;
    onSubmit(mutable as Partial<PdfMetadata>);
    onClose();
  };

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/40"
      role="dialog"
      aria-modal="true"
      aria-label={t("pdf.documentPropertiesTitle")}
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
      data-testid="pdf-metadata-dialog"
    >
      <form
        onSubmit={handleSubmit}
        className="w-[420px] max-w-[90vw] rounded-lg border border-divider bg-background p-4 shadow-xl"
      >
        <h2 className="mb-3 text-sm font-semibold text-foreground">
          {t("pdf.documentPropertiesTitle")}
        </h2>
        <div className="flex flex-col gap-3">
          <Field label={t("pdf.metaTitle")} testId="pdf-meta-title">
            <input
              ref={firstFieldRef}
              type="text"
              value={title}
              onChange={(e) => setTitle(e.target.value)}
              className={inputCls}
            />
          </Field>
          <Field label={t("pdf.metaAuthor")} testId="pdf-meta-author">
            <input
              type="text"
              value={author}
              onChange={(e) => setAuthor(e.target.value)}
              className={inputCls}
            />
          </Field>
          <Field label={t("pdf.metaSubject")} testId="pdf-meta-subject">
            <input
              type="text"
              value={subject}
              onChange={(e) => setSubject(e.target.value)}
              className={inputCls}
            />
          </Field>
          <Field label={t("pdf.metaKeywords")} testId="pdf-meta-keywords">
            <input
              type="text"
              value={keywords}
              onChange={(e) => setKeywords(e.target.value)}
              className={inputCls}
            />
          </Field>
        </div>
        <ReadOnlyMeta metadata={metadata} />
        <div className="mt-4 flex justify-end gap-2">
          <button
            type="button"
            onClick={onClose}
            className="rounded border border-divider px-3 py-1 text-xs hover:bg-hover"
            data-testid="pdf-metadata-cancel"
          >
            {t("pdf.documentPropertiesCancel")}
          </button>
          <button
            type="submit"
            className="rounded bg-[var(--accent)] px-3 py-1 text-xs font-medium text-white hover:opacity-90"
            data-testid="pdf-metadata-save"
          >
            {t("pdf.documentPropertiesSave")}
          </button>
        </div>
      </form>
    </div>
  );
}

const inputCls =
  "w-full rounded border border-divider bg-background px-2 py-1 text-xs text-foreground focus:outline-none focus:ring-1 focus:ring-[var(--accent)]";

interface FieldProps {
  readonly label: string;
  readonly testId: string;
  readonly children: React.ReactNode;
}

function Field({ label, testId, children }: FieldProps): React.ReactElement {
  return (
    <label className="flex flex-col gap-1 text-[11px] text-secondary" data-testid={testId}>
      {label}
      {children}
    </label>
  );
}

function ReadOnlyMeta({ metadata }: { metadata: PdfMetadata }): React.ReactElement | null {
  const rows: Array<[string, string]> = [];
  if (metadata.creator) rows.push(["Creator", metadata.creator]);
  if (metadata.producer) rows.push(["Producer", metadata.producer]);
  if (metadata.creationDate) rows.push(["Created", metadata.creationDate]);
  if (metadata.modificationDate) rows.push(["Modified", metadata.modificationDate]);
  if (metadata.pdfVersion) rows.push(["PDF version", metadata.pdfVersion]);
  if (rows.length === 0) return null;
  return (
    <dl className="mt-3 grid grid-cols-[max-content,1fr] gap-x-3 gap-y-0.5 border-t border-divider pt-3 text-[10px] text-tertiary">
      {rows.map(([k, v]) => (
        <React.Fragment key={k}>
          <dt className="uppercase tracking-wide">{k}</dt>
          <dd className="truncate font-mono">{v}</dd>
        </React.Fragment>
      ))}
    </dl>
  );
}
