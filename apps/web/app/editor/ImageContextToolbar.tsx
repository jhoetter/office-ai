"use client";

import { useEffect, useState, type ReactNode } from "react";
import { Image as ImageIcon, Type, Trash2 } from "lucide-react";
import { cn } from "@officeai/ui";
import type { EditorView } from "prosemirror-view";

/**
 * B6 — contextual image toolbar.
 *
 * Inline images are rendered as PM atom nodes (`image`) and selecting
 * one puts a `NodeSelection` over the leaf. We mirror the selection
 * polling pattern used by `TableContextToolbar` and float a small bar
 * directly above the image with the only operations that survive a
 * round-trip through OOXML: editing alt text and removing the image.
 *
 * Resizing happens in-place via the corner handle drawn by
 * {@link ImageResizeOverlay}; this toolbar exposes the discrete
 * actions so they remain reachable via the keyboard.
 */

export interface SelectedImageInfo {
  readonly imageId: string;
  readonly widthPx: number;
  readonly heightPx: number;
  readonly altText: string;
}

export interface ImageContextToolbarProps {
  readonly view: EditorView | null;
  readonly host: HTMLElement | null;
  readonly onEditAlt: (info: SelectedImageInfo) => void;
  readonly onDelete: (imageId: string) => void;
}

interface SelectedImage {
  readonly info: SelectedImageInfo;
  readonly rect: { left: number; top: number; width: number };
}

export function ImageContextToolbar(props: ImageContextToolbarProps): ReactNode {
  const { view, host, onEditAlt, onDelete } = props;
  const [selected, setSelected] = useState<SelectedImage | null>(null);

  useEffect(() => {
    if (!view || !host) {
      setSelected(null);
      return;
    }
    const compute = () => {
      const sel = view.state.selection;
      const node = (sel as { node?: { type: { name: string }; attrs: Record<string, unknown> } }).node;
      if (!node || node.type.name !== "image") {
        setSelected(null);
        return;
      }
      const runId = typeof node.attrs.runId === "string" ? node.attrs.runId : null;
      if (!runId) {
        setSelected(null);
        return;
      }
      const dom = view.nodeDOM(sel.from) as HTMLElement | null;
      if (!dom) {
        setSelected(null);
        return;
      }
      const rect = dom.getBoundingClientRect();
      const width = typeof node.attrs.width === "number" ? node.attrs.width : Math.round(rect.width);
      const height = typeof node.attrs.height === "number" ? node.attrs.height : Math.round(rect.height);
      const alt = typeof node.attrs.alt === "string" ? node.attrs.alt : "";
      setSelected({
        info: { imageId: runId, widthPx: width, heightPx: height, altText: alt },
        rect: { left: rect.left, top: rect.top, width: rect.width },
      });
    };
    compute();
    const onSel = () => compute();
    document.addEventListener("selectionchange", onSel);
    const onScroll = () => compute();
    host.addEventListener("scroll", onScroll, { capture: true });
    window.addEventListener("resize", onSel);
    return () => {
      document.removeEventListener("selectionchange", onSel);
      host.removeEventListener("scroll", onScroll, { capture: true });
      window.removeEventListener("resize", onSel);
    };
  }, [view, host]);

  if (!selected) return null;

  const top = Math.max(8, selected.rect.top - 36);
  const left = selected.rect.left + selected.rect.width / 2;

  return (
    <div
      role="toolbar"
      aria-label="Image actions"
      className={cn(
        "fixed z-40 flex items-center gap-0.5 rounded-md border border-divider bg-surface p-1 shadow-md"
      )}
      style={{ top, left, transform: "translateX(-50%)" }}
      data-testid="image-context-toolbar"
    >
      <Btn label="Edit alt text" onClick={() => onEditAlt(selected.info)}>
        <Type size={14} />
        <span className="ml-1 text-xs">Alt text</span>
      </Btn>
      <span className="mx-1 h-4 w-px bg-divider" aria-hidden />
      <span className="px-1 text-xs text-secondary" aria-hidden>
        <ImageIcon size={12} className="mr-1 inline-block align-[-2px]" />
        {selected.info.widthPx} × {selected.info.heightPx} px
      </span>
      <span className="mx-1 h-4 w-px bg-divider" aria-hidden />
      <Btn label="Delete image" onClick={() => onDelete(selected.info.imageId)}>
        <Trash2 size={14} />
      </Btn>
    </div>
  );
}

function Btn(props: { label: string; onClick: () => void; children: ReactNode }) {
  return (
    <button
      type="button"
      title={props.label}
      aria-label={props.label}
      onClick={props.onClick}
      className="flex items-center rounded p-1 text-secondary hover:bg-hover hover:text-foreground"
    >
      {props.children}
    </button>
  );
}
