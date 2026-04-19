"use client";

import type { ReactNode } from "react";
import { PresenceStack, type PresencePeer } from "@officeai/ui";
import type { AwarenessState } from "@officeai/realtime";
import type { RealtimeRoomState } from "./useRealtimeRoom";

interface PresenceSlotProps {
  readonly state: RealtimeRoomState;
}

/**
 * Top-bar adapter: turns a {@link RealtimeRoomState} into the
 * pure-display props expected by the {@link PresenceStack} primitive.
 * Keeps the realtime/Yjs types out of `@officeai/ui`.
 */
export function PresenceSlot({ state }: PresenceSlotProps): ReactNode {
  if (state.status === "disabled") return null;
  const room = state.room;
  const self: PresencePeer | undefined = room
    ? {
        id: room.identity.id,
        name: room.identity.name,
        color: room.identity.color,
      }
    : undefined;
  const peers: PresencePeer[] = state.remotePeers.map((p) => awarenessToPeer(p.state));
  return <PresenceStack peers={peers} self={self} status={state.status} />;
}

function awarenessToPeer(state: AwarenessState): PresencePeer {
  return {
    id: state.user.id,
    name: state.user.name,
    color: state.user.color,
    product: state.product,
    lastSeen: state.lastSeen,
  };
}
