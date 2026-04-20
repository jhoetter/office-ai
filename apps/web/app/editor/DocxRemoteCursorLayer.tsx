"use client";

import { useEffect, useState, type ReactNode } from "react";
import type { EditorView } from "prosemirror-view";
import type { AwarenessState } from "@officeai/realtime";

/**
 * Remote-cursor overlay for the DOCX editor.
 *
 * Each remote peer publishes a `DocxCursor` (`head` + `anchor`
 * absolute ProseMirror positions) via the realtime awareness layer.
 * For every such peer we project those positions back to viewport
 * pixels with `view.coordsAtPos(...)` and render:
 *
 *   - A 2px-wide caret bar at `head` in the peer's identity color.
 *   - A thin highlight rect for the (head .. anchor) range when the
 *     selection is on a single visual line. Multi-line ranges only
 *     show the caret to avoid faking line widths we don't know.
 *   - A name tag pinned just above the caret so two peers in the
 *     same paragraph stay distinguishable.
 *
 * Positioned `fixed` so it composes with the page card's transform
 * scale (zoom). All elements are `pointer-events: none` — the
 * overlay never steals input.
 */

export interface DocxRemotePeer {
  readonly clientId: number;
  readonly state: AwarenessState;
}

export interface DocxRemoteCursorLayerProps {
  readonly view: EditorView | null;
  readonly host: HTMLElement | null;
  readonly peers: ReadonlyArray<DocxRemotePeer>;
}

interface CaretRect {
  readonly clientId: number;
  readonly name: string;
  readonly color: string;
  readonly caret: { readonly left: number; readonly top: number; readonly height: number };
  readonly band: {
    readonly left: number;
    readonly top: number;
    readonly width: number;
    readonly height: number;
  } | null;
}

export function DocxRemoteCursorLayer(props: DocxRemoteCursorLayerProps): ReactNode {
  const { view, host, peers } = props;
  const [rects, setRects] = useState<ReadonlyArray<CaretRect>>([]);

  useEffect(() => {
    if (!view || !host) {
      setRects([]);
      return;
    }
    let raf = 0;
    const compute = (): void => {
      cancelAnimationFrame(raf);
      raf = requestAnimationFrame(() => {
        const docSize = view.state.doc.content.size;
        const next: CaretRect[] = [];
        for (const peer of peers) {
          const c = peer.state.cursor;
          if (!c || c.product !== "docx") continue;
          const head = clamp(c.head, 0, docSize);
          const anchor = clamp(c.anchor, 0, docSize);
          let headCoords: CoordRect;
          let anchorCoords: CoordRect;
          try {
            headCoords = view.coordsAtPos(head);
            anchorCoords = head === anchor ? headCoords : view.coordsAtPos(anchor);
          } catch {
            continue;
          }
          const caret = {
            left: headCoords.left,
            top: headCoords.top,
            height: Math.max(12, headCoords.bottom - headCoords.top),
          };
          let band: CaretRect["band"] = null;
          if (head !== anchor) {
            const sameLine =
              Math.abs(headCoords.top - anchorCoords.top) < 2 &&
              Math.abs(headCoords.bottom - anchorCoords.bottom) < 2;
            if (sameLine) {
              const left = Math.min(headCoords.left, anchorCoords.left);
              const right = Math.max(headCoords.right, anchorCoords.right);
              band = {
                left,
                top: headCoords.top,
                width: Math.max(2, right - left),
                height: Math.max(12, headCoords.bottom - headCoords.top),
              };
            }
          }
          next.push({
            clientId: peer.clientId,
            name: peer.state.user.name,
            color: peer.state.user.color,
            caret,
            band,
          });
        }
        setRects(next);
      });
    };
    compute();
    const onScroll = (): void => compute();
    const onResize = (): void => compute();
    host.addEventListener("scroll", onScroll, { capture: true });
    window.addEventListener("scroll", onScroll, { capture: true });
    window.addEventListener("resize", onResize);
    document.addEventListener("selectionchange", onScroll);
    return () => {
      cancelAnimationFrame(raf);
      host.removeEventListener("scroll", onScroll, { capture: true });
      window.removeEventListener("scroll", onScroll, { capture: true });
      window.removeEventListener("resize", onResize);
      document.removeEventListener("selectionchange", onScroll);
    };
  }, [view, host, peers]);

  if (rects.length === 0) return null;

  return (
    <div
      className="pointer-events-none fixed inset-0 z-30"
      data-testid="docx-remote-cursor-layer"
      aria-hidden
    >
      {rects.map((r) => (
        <div key={r.clientId} data-testid="docx-remote-caret" data-peer-color={r.color}>
          {r.band && (
            <div
              className="pointer-events-none fixed"
              style={{
                left: `${r.band.left}px`,
                top: `${r.band.top}px`,
                width: `${r.band.width}px`,
                height: `${r.band.height}px`,
                backgroundColor: r.color,
                opacity: 0.18,
                borderRadius: "2px",
              }}
            />
          )}
          <div
            className="pointer-events-none fixed"
            style={{
              left: `${r.caret.left}px`,
              top: `${r.caret.top}px`,
              width: "2px",
              height: `${r.caret.height}px`,
              backgroundColor: r.color,
              boxShadow: `0 0 0 0.5px ${r.color}`,
            }}
          />
          <div
            className="pointer-events-none fixed whitespace-nowrap rounded-sm px-1 py-0.5 text-[10px] font-medium text-white shadow-sm"
            style={{
              left: `${r.caret.left}px`,
              top: `${r.caret.top - 14}px`,
              backgroundColor: r.color,
            }}
          >
            {r.name}
          </div>
        </div>
      ))}
    </div>
  );
}

function clamp(n: number, min: number, max: number): number {
  if (!Number.isFinite(n)) return min;
  if (n < min) return min;
  if (n > max) return max;
  return n;
}

interface CoordRect {
  readonly left: number;
  readonly right: number;
  readonly top: number;
  readonly bottom: number;
}
