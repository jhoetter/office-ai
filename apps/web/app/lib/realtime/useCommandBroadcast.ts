"use client";

import { useEffect, useRef } from "react";
import type { CommandLite } from "@officeai/core";
import type { RoomClient } from "./RoomClient";

/**
 * Result returned by `applyCommand` that this hook actually reads.
 * Mirrors `Mutation<TSnapshot>` from `@officeai/core` but kept
 * structural so the hook doesn't need to import the heavy generic.
 */
export interface BroadcastableMutationResult {
  readonly status: string;
  readonly rejection?: { readonly code: string; readonly message?: string };
}

/**
 * Generic shape that all three product agents satisfy.
 *
 * Intentionally loose so this hook stays product-agnostic — it
 * inspects only the bits it needs (`subscribe` to learn about local
 * mutations, `applyCommand` to apply remote ones). The `Mutation`
 * shape coming back from `subscribe` is also kept opaque for the
 * same reason; we only ever read `command.{type,payload,source,agentId}`.
 *
 * NOTE: the method is `applyCommand` (not `dispatch`) — every
 * product agent (`packages/{docx,xlsx,pptx,pdf}/src/agent/agent.ts`)
 * exposes that name. An earlier iteration named it `dispatch`, which
 * matched the underlying `CommandBus.dispatch` but NOT the agent
 * facade — the resulting `agent.dispatch is not a function` runtime
 * error silently broke every remote command (so users saw presence
 * avatars but no live edits). The interface name now matches the
 * agent so a future rename is caught at compile time.
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
  applyCommand(command: CommandLite): Promise<BroadcastableMutationResult> | BroadcastableMutationResult;
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
      let result: Promise<BroadcastableMutationResult> | BroadcastableMutationResult;
      try {
        result = agent.applyCommand({ ...command, source: "system" });
      } catch (err) {
        applyingRemoteRef.current -= 1;
        console.warn("[realtime] remote command threw synchronously:", command.type, err);
        return;
      }
      Promise.resolve(result)
        .then((mutation) => {
          // The bus returns `status: "rejected"` (with a `rejection`
          // payload) instead of throwing. Surface it so the dev
          // console reads "remote command rejected: …" the same way
          // it would for a local rejection.
          if (mutation && mutation.status === "rejected") {
            const code = mutation.rejection?.code ?? "unknown";
            const msg = mutation.rejection?.message ?? "";
            console.warn("[realtime] remote command rejected:", command.type, code, msg);
          }
        })
        .catch((err) => {
          console.warn("[realtime] remote command threw:", command.type, err);
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
