"use client";

import { useEffect, useRef } from "react";
import type {
  AwarenessState,
  DocxCursor,
  PdfSelection,
  PptxSelection,
  XlsxSelection,
} from "@officeai/realtime";
import type { RoomClient } from "./RoomClient";

/**
 * Throttled publisher for the per-editor cursor / selection slot of
 * the awareness payload.
 *
 * Why a hook (and not just inline `room.setAwareness(...)` calls):
 *
 *   - Awareness updates fire on every keystroke / mouse-move; without
 *     a throttle the websocket is hammered with diffs that the UI
 *     couldn't render fast enough to matter anyway.
 *   - We want a stable JSON-equality compare so identical successive
 *     selections (typical: arrow-key navigation that lands on the
 *     same cell / line) collapse to a single emit.
 *   - All four editors need the same logic; pulling it into a hook
 *     keeps `DocxEditor.tsx` / `XlsxEditor.tsx` / `PptxEditor.tsx` /
 *     `PdfEditor.tsx` each at one call-site instead of N copies.
 */
export type PresenceCursor = DocxCursor | XlsxSelection | PptxSelection | PdfSelection;

export interface UsePublishPresenceOptions {
  readonly room: RoomClient | null;
  /** When `null`, we publish a "no cursor" awareness state (peer present, no selection). */
  readonly cursor: PresenceCursor | null;
  /** Throttle window in ms. Defaults to 50 ms — fast enough to feel live, cheap enough to spam. */
  readonly throttleMs?: number;
}

export function usePublishPresence(opts: UsePublishPresenceOptions): void {
  const { room, cursor, throttleMs = 50 } = opts;
  const lastJsonRef = useRef<string | null>(null);
  const pendingTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  useEffect(() => {
    if (!room) return;
    const json = cursor ? JSON.stringify(cursor) : "";
    if (json === lastJsonRef.current) return;
    if (pendingTimerRef.current) clearTimeout(pendingTimerRef.current);
    pendingTimerRef.current = setTimeout(() => {
      lastJsonRef.current = json;
      const next: Partial<AwarenessState> = cursor ? { cursor } : { cursor: undefined };
      try {
        room.setAwareness(next);
      } catch (err) {
        console.warn("[realtime] setAwareness failed:", err);
      }
    }, throttleMs);

    return () => {
      if (pendingTimerRef.current) {
        clearTimeout(pendingTimerRef.current);
        pendingTimerRef.current = null;
      }
    };
  }, [room, cursor, throttleMs]);
}
