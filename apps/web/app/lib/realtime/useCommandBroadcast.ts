"use client";

import { useEffect, useRef } from "react";
import type { CommandLite } from "@officeai/core";
import type { RoomClient } from "./RoomClient";

/**
 * Generic shape that all three product agents satisfy.
 *
 * Intentionally loose so this hook stays product-agnostic — it
 * inspects only the bits it needs (`subscribe` to learn about local
 * mutations, `applyCommand` to apply remote ones). The `Mutation`
 * shape coming back from `subscribe` is also kept opaque for the
 * same reason; we only ever read `command.{type,payload,source,agentId}`.
 *
 * IMPORTANT: the method MUST be `applyCommand`, not `dispatch`. The
 * three product agents (`DocxAgent` / `XlsxAgent` / `PptxAgent`) all
 * expose `applyCommand`; an earlier version of this hook used
 * `dispatch` and silently broke every cross-tab apply at runtime
 * because the cast at the call-site hid the type mismatch.
 * Regression test: `useCommandBroadcast.test.ts`.
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
  applyCommand(
    command: CommandLite
  ): Promise<{ readonly rejection?: { readonly code: string; readonly message?: string } } | unknown>;
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
    return wireBroadcast({
      agent,
      room,
      shouldBroadcast,
      applyingRemote: applyingRemoteRef,
    });
  }, [agent, room, shouldBroadcast]);
}

/**
 * Pure (React-free) wiring: subscribes the local agent to the room
 * and the room to the local agent. Returns an unsubscribe function.
 *
 * Extracted from `useCommandBroadcast` so the contract — including
 * the apply-not-dispatch invariant — is testable without a React
 * testing harness.
 */
export function wireBroadcast(args: {
  readonly agent: BroadcastableAgent;
  readonly room: RoomClient;
  readonly shouldBroadcast?: (commandType: string) => boolean;
  readonly applyingRemote: { current: number };
}): () => void {
  const { agent, room, shouldBroadcast, applyingRemote } = args;

  const unsubLocal = agent.subscribe((_snap, mutation) => {
    if (applyingRemote.current > 0) return;
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
    applyingRemote.current += 1;
    const result = agent.applyCommand({ ...command, source: "system" });
    Promise.resolve(result)
      .then((mutation) => {
        // The command bus does NOT throw on rejection — it returns
        // `{ rejection: { code, message } }`. Surface those at warn
        // level so a peer's bad command doesn't fail silently while
        // every subsequent one applies cleanly.
        const r = (mutation as { rejection?: { code: string; message?: string } } | undefined)?.rejection;
        if (r) {
          console.warn("[realtime] remote command rejected:", command.type, r.code, r.message ?? "");
        }
      })
      .catch((err) => {
        console.warn("[realtime] remote command threw:", command.type, err);
      })
      .finally(() => {
        applyingRemote.current -= 1;
      });
  });

  return () => {
    unsubLocal();
    unsubRemote();
  };
}
