"use client";

import { useEffect, useRef, useState } from "react";
import type { AwarenessState } from "@officeai/realtime";
import type { ProductKind, RoomClient } from "./RoomClient";
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
  /**
   * Optional host-supplied identity override. Forwarded straight to
   * `createRoomClient`; see `RoomClientOptions.identity` for the
   * field semantics. The hook re-creates the room when the `id`
   * changes so a host that swaps the user mid-session (rare but
   * possible — e.g. logout/login) re-publishes the new awareness.
   */
  readonly identity?: { readonly id: string; readonly name: string; readonly color?: string };
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
  const { roomId, product, enabled = true, identity } = opts;
  const [room, setRoom] = useState<RoomClient | null>(null);
  const [status, setStatus] = useState<RealtimeRoomState["status"]>("disabled");
  const [peers, setPeers] = useState<ReadonlyArray<{ clientId: number; state: AwarenessState }>>([]);
  const ref = useRef<RoomClient | null>(null);

  // Pull primitives out of `identity` so the effect's dep array is
  // stable: the host typically passes a fresh object each render
  // (`{ id, name }`), which would otherwise tear the room down on
  // every render.
  const identityId = identity?.id ?? null;
  const identityName = identity?.name ?? null;
  const identityColor = identity?.color ?? null;

  useEffect(() => {
    if (!enabled || !roomId) {
      setRoom(null);
      setStatus("disabled");
      setPeers([]);
      return;
    }
    setStatus("connecting");
    setRoom(null);
    setPeers([]);

    let cancelled = false;
    let client: RoomClient | null = null;
    let teardown: (() => void) | null = null;

    void import("./RoomClient")
      .then(({ createRoomClient }) => {
        if (cancelled) return;
        const url = resolveRealtimeUrl();
        client = createRoomClient({
          url,
          roomId,
          product,
          ...(identityId && identityName
            ? {
                identity: {
                  id: identityId,
                  name: identityName,
                  ...(identityColor ? { color: identityColor } : {}),
                },
              }
            : {}),
        });
        ref.current = client;
        setRoom(client);

        const unsubStatus = client.onStatus((s) => setStatus(s));
        const unsubAwareness = client.onAwareness(() => {
          setPeers(client?.getRemoteStates() ?? []);
        });
        setPeers(client.getRemoteStates());

        teardown = () => {
          unsubStatus();
          unsubAwareness();
          try {
            client?.destroy();
          } catch {
            /* noop */
          }
          if (ref.current === client) ref.current = null;
          client = null;
        };
      })
      .catch((err: unknown) => {
        if (cancelled) return;
        console.warn("[realtime] failed to load realtime room client:", err);
        setRoom(null);
        setStatus("disconnected");
        setPeers([]);
      });

    return () => {
      cancelled = true;
      teardown?.();
      setRoom(null);
      setStatus("disabled");
      setPeers([]);
    };
  }, [enabled, roomId, product, identityId, identityName, identityColor]);

  return { room, status, remotePeers: peers };
}
