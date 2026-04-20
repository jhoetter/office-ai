"use client";

import { useCallback, useEffect, useRef, useState, type CSSProperties, type ReactNode } from "react";
import { ResizeHandles, cn, type ResizeHandleSide } from "@officeai/ui";
import type { EditorView } from "prosemirror-view";
import { useTranslator } from "@/lib/i18n";

/**
 * B6 — image resize overlay.
 *
 * Drawn over the selected inline image as a thin focus ring with eight
 * resize handles (Word/Pages style). Corner handles preserve aspect
 * ratio; side handles stretch one axis. While dragging we apply CSS
 * sizing to the underlying `<img>` so the user sees 60 fps feedback;
 * on `mouseup` we commit a single `docx:set-image-properties` command
 * with the rounded final dimensions.
 *
 * View-only: never mutates PM state, never touches the snapshot
 * directly. Sizing is committed via the `onResize` callback.
 */

type Handle = ResizeHandleSide;

const CORNER_HANDLES: ReadonlySet<Handle> = new Set(["nw", "ne", "se", "sw"]);
const MIN_PX = 16;
const MAX_PX = 4000;

export interface ImageResizeOverlayProps {
  readonly view: EditorView | null;
  readonly host: HTMLElement | null;
  readonly onResize: (imageId: string, widthPx: number, heightPx: number) => void;
}

interface SelectedImage {
  readonly imageId: string;
  readonly widthPx: number;
  readonly heightPx: number;
  readonly dom: HTMLImageElement;
  readonly rect: { left: number; top: number; width: number; height: number };
}

export function ImageResizeOverlay(props: ImageResizeOverlayProps): ReactNode {
  const { view, host, onResize } = props;
  const { t } = useTranslator();
  const [selected, setSelected] = useState<SelectedImage | null>(null);
  const [draft, setDraft] = useState<{ width: number; height: number } | null>(null);
  const draggingRef = useRef(false);

  useEffect(() => {
    if (!view || !host) {
      setSelected(null);
      return;
    }
    const compute = () => {
      if (draggingRef.current) return;
      const sel = view.state.selection;
      const node = (sel as { node?: { type: { name: string }; attrs: Record<string, unknown> } }).node;
      if (!node || node.type.name !== "image") {
        setSelected(null);
        setDraft(null);
        return;
      }
      const runId = typeof node.attrs.runId === "string" ? node.attrs.runId : null;
      if (!runId) {
        setSelected(null);
        return;
      }
      const dom = view.nodeDOM(sel.from) as HTMLElement | null;
      if (!dom || dom.tagName !== "IMG") {
        setSelected(null);
        return;
      }
      const img = dom as HTMLImageElement;
      const rect = img.getBoundingClientRect();
      const widthPx =
        typeof node.attrs.width === "number" && node.attrs.width > 0
          ? node.attrs.width
          : Math.round(rect.width);
      const heightPx =
        typeof node.attrs.height === "number" && node.attrs.height > 0
          ? node.attrs.height
          : Math.round(rect.height);
      setSelected({
        imageId: runId,
        widthPx,
        heightPx,
        dom: img,
        rect: { left: rect.left, top: rect.top, width: rect.width, height: rect.height },
      });
    };
    compute();
    const onSel = () => compute();
    document.addEventListener("selectionchange", onSel);
    host.addEventListener("scroll", onSel, { capture: true });
    window.addEventListener("resize", onSel);
    return () => {
      document.removeEventListener("selectionchange", onSel);
      host.removeEventListener("scroll", onSel, { capture: true });
      window.removeEventListener("resize", onSel);
    };
  }, [view, host]);

  const startDrag = useCallback(
    (handle: Handle, startX: number, startY: number) => {
      if (!selected) return;
      const startW = selected.widthPx;
      const startH = selected.heightPx;
      const aspect = startW > 0 && startH > 0 ? startW / startH : 1;
      const dom = selected.dom;
      let lastSize = { width: startW, height: startH };
      draggingRef.current = true;

      const onMove = (e: MouseEvent) => {
        const dx = e.clientX - startX;
        const dy = e.clientY - startY;
        lastSize = computeNextSize(handle, startW, startH, dx, dy, aspect);
        dom.style.width = `${lastSize.width}px`;
        dom.style.height = `${lastSize.height}px`;
        setDraft(lastSize);
      };
      const onUp = () => {
        window.removeEventListener("mousemove", onMove);
        window.removeEventListener("mouseup", onUp);
        dom.style.removeProperty("width");
        dom.style.removeProperty("height");
        draggingRef.current = false;
        setDraft(null);
        if (lastSize.width !== startW || lastSize.height !== startH) {
          onResize(selected.imageId, lastSize.width, lastSize.height);
        }
      };
      window.addEventListener("mousemove", onMove);
      window.addEventListener("mouseup", onUp);
    },
    [selected, onResize]
  );

  if (!selected) return null;

  const width = draft?.width ?? selected.rect.width;
  const height = draft?.height ?? selected.rect.height;

  return (
    <div
      className={cn("pointer-events-none fixed z-30 text-accent")}
      style={{ left: selected.rect.left, top: selected.rect.top, width, height }}
      data-testid="image-resize-overlay"
    >
      <div className="pointer-events-none absolute inset-0 ring-1 ring-accent" />
      {/* This consumer ignores `info.shiftKey` — DOCX images
          ALWAYS preserve aspect on a corner drag (matches Word /
          Pages behaviour). Edge handles stretch one axis as usual.
          See `computeNextSize` for the corner-aspect logic. */}
      <ResizeHandles
        handleSizePx={8}
        handleLabel={(side) => t("common.resizeHandle", { handle: side })}
        onHandleGrab={(info) => {
          startDrag(info.side, info.clientX, info.clientY);
        }}
      />
    </div>
  );
}

function computeNextSize(
  handle: Handle,
  startW: number,
  startH: number,
  dx: number,
  dy: number,
  aspect: number
): { width: number; height: number } {
  let width = startW;
  let height = startH;
  switch (handle) {
    case "e":
      width = startW + dx;
      break;
    case "w":
      width = startW - dx;
      break;
    case "s":
      height = startH + dy;
      break;
    case "n":
      height = startH - dy;
      break;
    case "se":
    case "ne":
    case "sw":
    case "nw": {
      const sx = handle === "ne" || handle === "se" ? 1 : -1;
      const sy = handle === "se" || handle === "sw" ? 1 : -1;
      const candidateW = Math.max(MIN_PX, startW + sx * dx);
      const candidateH = Math.max(MIN_PX, startH + sy * dy);
      const useWidthDriven = candidateW / candidateH >= aspect;
      if (useWidthDriven) {
        width = candidateW;
        height = Math.round(width / aspect);
      } else {
        height = candidateH;
        width = Math.round(height * aspect);
      }
      break;
    }
    default: {
      const _exhaustive: never = handle;
      void _exhaustive;
    }
  }
  return {
    width: clamp(Math.round(width), MIN_PX, MAX_PX),
    height: clamp(Math.round(height), MIN_PX, MAX_PX),
  };
}

function clamp(n: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, n));
}
