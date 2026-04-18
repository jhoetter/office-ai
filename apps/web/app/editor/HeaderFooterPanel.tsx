"use client";

import { useMemo, useState, type ReactNode } from "react";
import { ChevronDown, FileMinus, Hash, PanelTop } from "lucide-react";
import { cn } from "@officeai/ui";
import type { DocxAgent, DocxSnapshot } from "@officeai/docx";

/**
 * P3.4 / W14 — header / footer authoring panel.
 *
 * Mounts above the editor surface. Lists every typed header / footer
 * part discovered in the snapshot and exposes Word's three most-asked
 * section authoring commands:
 *
 *  - Edit the first-paragraph text of any header/footer part
 *    (`docx:set-header-text` / `docx:set-footer-text`)
 *  - Insert a `<w:fldSimple w:instr=" PAGE "/>` field at the end of a
 *    header/footer paragraph (`docx:insert-page-number`)
 *  - Toggle "Different first page" on the trailing section
 *    (`docx:set-section-different-first`)
 *
 * The panel collapses by default to keep the editor chrome compact,
 * matching Word's "Header & Footer" ribbon being a one-click reveal.
 *
 * NOTE: this is the data-layer-first MVP for header/footer
 * authoring. A future iteration mounts each part as its own
 * ProseMirror instance and renders them visually at the top / bottom
 * of every page (the full focus model described in
 * `spec/docx/header-footer-authoring.md`). For now the panel is the
 * authoritative editing surface.
 */
export interface HeaderFooterPanelProps {
  agent: DocxAgent | null;
  snapshot: DocxSnapshot | null;
  onError: (msg: string) => void;
  onInfo: (msg: string) => void;
}

export function HeaderFooterPanel(props: HeaderFooterPanelProps): ReactNode {
  const [expanded, setExpanded] = useState(false);
  const { agent, snapshot } = props;

  const parts = useMemo(() => {
    if (!snapshot) return [];
    return snapshot.root.headersAndFooters.map((p) => ({
      partPath: p.partPath,
      kind: p.kind,
      target: p.target,
      firstParagraphText: extractFirstParagraphText(p),
      firstParagraphId:
        p.body[0] && p.body[0].kind === "paragraph" ? p.body[0].id : null,
    }));
  }, [snapshot]);

  const titlePgEnabled = useMemo(() => {
    if (!snapshot) return false;
    for (let i = snapshot.root.body.length - 1; i >= 0; i--) {
      const block = snapshot.root.body[i];
      if (block.kind === "section-break") return block.properties.titlePg ?? false;
    }
    return false;
  }, [snapshot]);

  const onEditPart = async (
    partPath: string,
    kind: "header" | "footer",
    text: string
  ) => {
    if (!agent) return;
    const cmdType = kind === "header" ? "docx:set-header-text" : "docx:set-footer-text";
    const m = await agent.applyCommand({
      type: cmdType,
      payload: { partId: partPath, paragraphIndex: 0, text },
      source: "human",
    });
    if (m.status !== "approved") {
      props.onError(m.rejection?.message ?? `${cmdType} rejected`);
    }
  };

  const onInsertPageNumber = async (paragraphId: string) => {
    if (!agent) return;
    const m = await agent.applyCommand({
      type: "docx:insert-page-number",
      payload: { paragraphId, offset: Number.MAX_SAFE_INTEGER },
      source: "human",
    });
    if (m.status === "approved") {
      props.onInfo("Inserted PAGE field");
    } else {
      props.onError(m.rejection?.message ?? "Insert page number rejected");
    }
  };

  const onToggleDifferentFirst = async (next: boolean) => {
    if (!agent) return;
    const m = await agent.applyCommand({
      type: "docx:set-section-different-first",
      payload: { paragraphIndex: 0, enabled: next },
      source: "human",
    });
    if (m.status === "approved") {
      props.onInfo(`Different first page: ${next ? "on" : "off"}`);
    } else {
      props.onError(m.rejection?.message ?? "Toggle rejected");
    }
  };

  if (!snapshot) return null;

  return (
    <div className="docx-header-footer-panel mt-2 rounded-md border border-divider bg-surface">
      <button
        type="button"
        onClick={() => setExpanded((e) => !e)}
        className="flex w-full items-center gap-2 px-3 py-2 text-left text-xs font-medium text-foreground hover:bg-hover"
        aria-expanded={expanded}
      >
        <PanelTop size={14} />
        <span>Headers, footers &amp; sections</span>
        <span className="ml-1 rounded bg-background px-1.5 py-0.5 text-[10px] text-secondary">
          {parts.length} part{parts.length === 1 ? "" : "s"}
        </span>
        <ChevronDown
          size={12}
          className={cn("ml-auto transition-transform", expanded && "rotate-180")}
        />
      </button>
      {expanded && (
        <div className="space-y-3 border-t border-divider px-3 py-3 text-xs">
          {parts.length === 0 ? (
            <p className="text-secondary">
              This document has no header or footer parts. Use Word to add one
              for now (P3.4 does not yet auto-create parts).
            </p>
          ) : (
            <ul className="space-y-3">
              {parts.map((p) => (
                <li
                  key={p.partPath}
                  className="rounded border border-divider bg-background p-2"
                >
                  <div className="mb-1.5 flex items-center justify-between">
                    <span className="text-[11px] font-medium uppercase tracking-wide text-secondary">
                      {p.kind} · {p.target}
                    </span>
                    <span className="text-[10px] text-secondary">{p.partPath}</span>
                  </div>
                  <textarea
                    rows={2}
                    defaultValue={p.firstParagraphText}
                    className="w-full resize-y rounded border border-divider bg-background px-2 py-1 text-xs text-foreground focus:border-accent focus:outline-none"
                    onBlur={(e) => {
                      const next = e.currentTarget.value;
                      if (next !== p.firstParagraphText) {
                        void onEditPart(p.partPath, p.kind, next);
                      }
                    }}
                  />
                  <div className="mt-1.5 flex items-center gap-1.5">
                    <button
                      type="button"
                      disabled={!p.firstParagraphId}
                      onClick={() => p.firstParagraphId && void onInsertPageNumber(p.firstParagraphId)}
                      className="inline-flex items-center gap-1 rounded border border-divider px-2 py-0.5 text-[11px] text-foreground hover:bg-hover disabled:opacity-50"
                    >
                      <Hash size={11} /> Insert PAGE field
                    </button>
                  </div>
                </li>
              ))}
            </ul>
          )}
          <div className="flex items-center justify-between border-t border-divider pt-2">
            <label className="flex items-center gap-2 text-xs text-foreground">
              <input
                type="checkbox"
                checked={titlePgEnabled}
                onChange={(e) => void onToggleDifferentFirst(e.currentTarget.checked)}
              />
              Different first page
              <span className="text-[10px] text-secondary">
                (toggles <code>w:titlePg</code> on the trailing section)
              </span>
            </label>
            <span className="inline-flex items-center gap-1 text-[10px] text-secondary">
              <FileMinus size={11} /> P3.4 minimum surface
            </span>
          </div>
        </div>
      )}
    </div>
  );
}

function extractFirstParagraphText(part: DocxSnapshot["root"]["headersAndFooters"][number]): string {
  const first = part.body[0];
  if (!first || first.kind !== "paragraph") return "";
  let out = "";
  for (const inline of first.children) {
    if (inline.kind === "run") {
      for (const c of inline.children) {
        if (c.kind === "text") out += c.text;
        else if (c.kind === "page-number-field") out += `[${c.field}]`;
      }
    }
  }
  return out;
}
