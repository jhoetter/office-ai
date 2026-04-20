"use client";

import type { ReactNode } from "react";
import type { AwarenessState } from "@officeai/realtime";

/**
 * Shared "live cursor list" that any editor can drop in to surface
 * what each remote peer is doing right now. Each row reads e.g.
 * "Alice is on Sheet1!B4:D7" / "Bob has 2 shapes selected on slide 3"
 * / "Carol is at line 17". The colored dot uses the peer's identity
 * color so it lines up visually with the avatar in the top-bar
 * `PresenceStack`.
 *
 * This is intentionally lightweight (a fixed-size pill stack in the
 * lower-left of the editor) so it works in every editor without
 * having to plumb coordinate transforms (PM positions, EMU,
 * row/col-pixel offsets) into a shared overlay primitive. The
 * editor-specific overlays (caret line for DOCX, selection rect for
 * XLSX, shape outline for PPTX) live alongside this list and can be
 * adopted incrementally.
 */
export interface RemotePresencePeer {
  readonly clientId: number;
  readonly state: AwarenessState;
}

export interface RemotePresenceListProps {
  readonly peers: ReadonlyArray<RemotePresencePeer>;
}

export function RemotePresenceList({ peers }: RemotePresenceListProps): ReactNode {
  if (peers.length === 0) return null;
  return (
    <div
      className="pointer-events-none fixed bottom-3 left-3 z-30 flex max-w-[40vw] flex-col gap-1"
      data-testid="realtime-presence-list"
    >
      {peers.map((p) => {
        const summary = describePeer(p.state);
        if (!summary) return null;
        return (
          <div
            key={p.clientId}
            className="pointer-events-auto inline-flex items-center gap-2 rounded-md border border-divider bg-surface/95 px-2 py-1 text-xs text-secondary shadow-sm backdrop-blur"
          >
            <span
              aria-hidden
              className="inline-block size-2 rounded-full"
              style={{ backgroundColor: p.state.user.color }}
            />
            <span className="font-medium text-primary">{p.state.user.name}</span>
            <span>{summary}</span>
          </div>
        );
      })}
    </div>
  );
}

function describePeer(state: AwarenessState): string | null {
  const c = state.cursor;
  if (!c) return "is here";
  switch (c.product) {
    case "docx":
      return `is at position ${c.head}`;
    case "xlsx":
      return `is on ${c.sheetName}!${c.range}`;
    case "pptx":
      return c.shapeIds.length === 0
        ? `is on slide ${shortSlide(c.slideId)}`
        : `has ${c.shapeIds.length} shape${c.shapeIds.length === 1 ? "" : "s"} selected on slide ${shortSlide(c.slideId)}`;
    case "pdf":
      return c.normalizedRect ? `is reading page ${c.pageNumber}` : `is on page ${c.pageNumber}`;
    default: {
      const _never: never = c;
      void _never;
      return null;
    }
  }
}

function shortSlide(id: string): string {
  // PPTX presence publishes the slide's stable OOXML `partPath`,
  // e.g. `ppt/slides/slide3.xml`. Pull the trailing slide number out
  // for a human-friendly pill ("is on slide 3" instead of leaking
  // the part path or a UUID at the user).
  const part = /\/slide(\d+)\.xml$/.exec(id);
  if (part && part[1]) return part[1];
  // Legacy/test fallbacks.
  const legacy = /^slide-(\d+)$/.exec(id);
  if (legacy && legacy[1]) return legacy[1];
  return id;
}
