"use client";

import { useEffect, useRef } from "react";
import type { CommandLite } from "@officeai/core";
import type { RoomClient } from "./RoomClient";

/**
 * Generic shape that all three product agents satisfy.
 *
 * Intentionally loose so this hook stays product-agnostic — it
 * inspects only the bits it needs (`subscribe` to learn about local
 * mutations, `dispatch` to apply remote ones). The `Mutation` shape
 * coming back from `subscribe` is also kept opaque for the same
 * reason; we only ever read `command.{type,payload,source,agentId}`.
 */
export interface BroadcastableAgent {
  subscribe(
    listener: (
      snapshot: unknown,
      mutation: {
        readonly status: string;
        readonly command: {
          readonly type: string;
          readonly payload: unknown;
          readonly source: string;
          readonly agentId?: string;
        };
      }
    ) => void
  ): () => void;
  dispatch(command: CommandLite): Promise<unknown> | unknown;
}

export interface UseCommandBroadcastOptions {
  /** May be `null` while the editor is bootstrapping or collab is off. */
  readonly agent: BroadcastableAgent | null;
  readonly room: RoomClient | null;
  /**
   * Optional filter — return false to skip a command that should
   * stay local (e.g. ephemeral selection markers). When omitted we
   * broadcast every approved mutation.
   */
  readonly shouldBroadcast?: (commandType: string) => boolean;
}

/**
 * Wire a {@link BroadcastableAgent} bidirectionally to a
 * {@link RoomClient}.
 *
 * - Local approved mutations → enqueued onto the shared command log.
 * - Remote command-log inserts → dispatched on the local agent with
 *   `source: "system"` so the bus doesn't classify them as pending
 *   agent suggestions.
 *
 * Echo prevention is handled by the room client (it tags every
 * envelope with our peer id and skips matches in the inbound path).
 * Re-broadcast loops are additionally guarded here by an
 * `applyingRemote` ref: when we dispatch a remote command, the
 * resulting `subscribe` callback is dropped so the same envelope
 * isn't pushed back onto the log.
 */
export function useCommandBroadcast(opts: UseCommandBroadcastOptions): void {
  const { agent, room, shouldBroadcast } = opts;
  const applyingRemoteRef = useRef(0);

  useEffect(() => {
    if (!agent || !room) return;

    const unsubLocal = agent.subscribe((_snap, mutation) => {
      if (applyingRemoteRef.current > 0) return;
      // Only mirror approved mutations — agent suggestions / rejections
      // stay client-local until the human accepts them.
      if (mutation.status !== "approved") return;
      const t = mutation.command.type;
      if (shouldBroadcast && !shouldBroadcast(t)) return;
      const cmd: CommandLite = {
        type: t,
        payload: mutation.command.payload,
        ...(mutation.command.source ? { source: mutation.command.source as CommandLite["source"] } : {}),
        ...(mutation.command.agentId ? { agentId: mutation.command.agentId } : {}),
      };
      try {
        room.broadcastCommand(cmd);
      } catch (err) {
        console.warn("[realtime] broadcast failed for", t, err);
      }
    });

    const unsubRemote = room.onRemoteCommand((command) => {
      applyingRemoteRef.current += 1;
      const result = agent.dispatch({ ...command, source: "system" });
      Promise.resolve(result)
        .catch((err) => {
          console.warn("[realtime] remote command rejected:", command.type, err);
        })
        .finally(() => {
          applyingRemoteRef.current -= 1;
        });
    });

    return () => {
      unsubLocal();
      unsubRemote();
    };
  }, [agent, room, shouldBroadcast]);
}
