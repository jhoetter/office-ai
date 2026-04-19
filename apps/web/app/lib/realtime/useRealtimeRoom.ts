"use client";

import { useEffect, useRef, useState } from "react";
import type { AwarenessState } from "@officeai/realtime";
import { createRoomClient, type ProductKind, type RoomClient } from "./RoomClient";
import { resolveRealtimeUrl } from "./config";

export interface UseRealtimeRoomOptions {
  /**
   * Logical room id. Stable across reconnects (we suggest the file
   * URL or a synthetic `local-<sessionId>`). Pass `null` to opt out
   * of collaboration entirely (e.g. SSR or tests).
   */
  readonly roomId: string | null;
  readonly product: ProductKind;
  /**
   * When false the hook short-circuits to a no-op room (still
   * returns a sentinel so callers don't need to null-check). Useful
   * for explicitly-disabled environments.
   */
  readonly enabled?: boolean;
}

export interface RealtimeRoomState {
  readonly room: RoomClient | null;
  readonly status: "disabled" | "disconnected" | "connecting" | "connected";
  readonly remotePeers: ReadonlyArray<{ clientId: number; state: AwarenessState }>;
}

/**
 * React hook that manages a single `RoomClient` per editor mount.
 * Re-creates the client when `roomId` or `product` changes; tears
 * everything down on unmount. Subscribes to awareness so the
 * caller can render presence chips on every change.
 */
export function useRealtimeRoom(opts: UseRealtimeRoomOptions): RealtimeRoomState {
  const { roomId, product, enabled = true } = opts;
  const [room, setRoom] = useState<RoomClient | null>(null);
  const [status, setStatus] = useState<RealtimeRoomState["status"]>("disabled");
  const [peers, setPeers] = useState<ReadonlyArray<{ clientId: number; state: AwarenessState }>>([]);
  const ref = useRef<RoomClient | null>(null);

  useEffect(() => {
    if (!enabled || !roomId) {
      setRoom(null);
      setStatus("disabled");
      setPeers([]);
      return;
    }
    const url = resolveRealtimeUrl();
    const client = createRoomClient({ url, roomId, product });
    ref.current = client;
    setRoom(client);
    setStatus("connecting");

    const unsubStatus = client.onStatus((s) => setStatus(s));
    const unsubAwareness = client.onAwareness(() => {
      setPeers(client.getRemoteStates());
    });
    setPeers(client.getRemoteStates());

    return () => {
      unsubStatus();
      unsubAwareness();
      try {
        client.destroy();
      } catch {
        /* noop */
      }
      if (ref.current === client) ref.current = null;
      setRoom(null);
      setStatus("disabled");
      setPeers([]);
    };
  }, [enabled, roomId, product]);

  return { room, status, remotePeers: peers };
}
