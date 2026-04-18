"use client";

import { useEffect, useRef, useState } from "react";
import { Hash, X } from "lucide-react";

/**
 * Inline editor popover for a page header / footer zone.
 *
 * Mounted by DocxEditor in response to a `pm-page-zone-edit`
 * CustomEvent fired by the page-decorations plugin (see
 * `apps/web/app/lib/page-decorations.ts`).
 *
 * The popover anchors to the zone DOM rectangle, drives a single
 * textarea, and submits via the existing typed commands:
 *  - `docx:set-header-text` / `docx:set-footer-text` for body text
 *  - `docx:insert-page-number` to add a `<w:fldSimple PAGE/>` field
 *
 * Both commands target paragraph 0 of the part. Multi-paragraph
 * authoring is a follow-up; this matches Word's "click into the
 * header area, type one line" experience.
 */
export interface PageZoneEditorProps {
  slot: "header" | "footer";
  partPath: string | null;
  pageNumber: number;
  initialText: string;
  anchorRect: { left: number; top: number; bottom: number; width: number };
  onSubmit: (text: string) => void | Promise<void>;
  onInsertPageNumber: () => void | Promise<void>;
  onCancel: () => void;
}

export function PageZoneEditor(props: PageZoneEditorProps): React.ReactNode {
  const { slot, partPath, pageNumber, initialText, anchorRect } = props;
  const [draft, setDraft] = useState(initialText);
  const taRef = useRef<HTMLTextAreaElement | null>(null);

  useEffect(() => {
    taRef.current?.focus();
    taRef.current?.select();
  }, []);

  const submit = () => {
    void props.onSubmit(draft);
  };

  const onKey = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === "Escape") {
      e.preventDefault();
      props.onCancel();
      return;
    }
    if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) {
      e.preventDefault();
      submit();
    }
  };

  // Position above the anchor rect for footer (so it doesn't slip
  // off the bottom of the viewport) and below it for header.
  const popoverTop =
    slot === "footer" ? Math.max(8, anchorRect.top - 200) : anchorRect.bottom + 6;

  const isAuthorable = partPath !== null;

  return (
    <div
      role="dialog"
      aria-label={`${slot === "header" ? "Header" : "Footer"} editor for page ${pageNumber}`}
      className="fixed z-50 w-[420px] max-w-[92vw] rounded-lg border border-divider bg-background p-3 shadow-2xl"
      style={{
        left: Math.max(8, Math.min(anchorRect.left, window.innerWidth - 440)),
        top: popoverTop,
      }}
      data-testid="page-zone-editor"
    >
      <div className="mb-1.5 flex items-center justify-between">
        <span className="text-[11px] font-medium uppercase tracking-wide text-secondary">
          {slot} · page {pageNumber}
          {partPath ? "" : " · no part"}
        </span>
        <button
          type="button"
          onClick={props.onCancel}
          aria-label="Close"
          className="rounded p-0.5 text-secondary hover:bg-hover"
        >
          <X size={12} />
        </button>
      </div>
      {!isAuthorable && (
        <p className="mb-2 rounded border border-warning/30 bg-warning/10 px-2 py-1 text-[11px] text-foreground">
          This document has no {slot} part for this section. Add one in
          Word, then re-open the file. Auto-creation is on the P4 list.
        </p>
      )}
      <textarea
        ref={taRef}
        rows={3}
        value={draft}
        onChange={(e) => setDraft(e.currentTarget.value)}
        onKeyDown={onKey}
        disabled={!isAuthorable}
        placeholder={
          isAuthorable
            ? slot === "header"
              ? "e.g. Project Apollo — Confidential"
              : "e.g. Page [PAGE]"
            : ""
        }
        className="w-full resize-y rounded border border-divider bg-background px-2 py-1.5 text-sm text-foreground focus:border-accent focus:outline-none disabled:opacity-50"
        data-testid="page-zone-editor-textarea"
      />
      <div className="mt-2 flex items-center justify-between gap-2">
        <button
          type="button"
          onClick={() => void props.onInsertPageNumber()}
          disabled={!isAuthorable}
          className="inline-flex items-center gap-1 rounded border border-divider px-2 py-1 text-[11px] text-foreground hover:bg-hover disabled:opacity-50"
          title="Insert PAGE field at end of paragraph"
        >
          <Hash size={11} /> Insert page number
        </button>
        <div className="flex items-center gap-1">
          <button
            type="button"
            onClick={props.onCancel}
            className="rounded px-2 py-1 text-[11px] text-secondary hover:bg-hover"
          >
            Cancel
          </button>
          <button
            type="button"
            onClick={submit}
            disabled={!isAuthorable}
            className="rounded bg-accent px-2.5 py-1 text-[11px] font-medium text-white hover:opacity-90 disabled:opacity-50"
          >
            Save · ⌘↵
          </button>
        </div>
      </div>
    </div>
  );
}
