"use client";

import * as Y from "yjs";
import { WebsocketProvider } from "y-websocket";
import {
  COMMAND_LOG_KEY,
  decodeCommand,
  encodeCommand,
  envelopeToCommandLite,
  generateAnonymousIdentity,
  isOurEcho,
  type AnonymousIdentity,
  type AwarenessState,
  type CommandEnvelope,
} from "@officeai/realtime";
import type { CommandLite } from "@officeai/core";

export type ProductKind = "docx" | "xlsx" | "pptx";

export interface RoomClientOptions {
  readonly url: string;
  readonly roomId: string;
  readonly product: ProductKind;
}

export interface RoomClient {
  readonly identity: AnonymousIdentity;
  readonly peerId: string;
  /** Append a typed command to the shared log. Echoes back through `onRemoteCommand` are suppressed by `peerId` match. */
  broadcastCommand(command: CommandLite): void;
  /** Subscribe to remote commands. Returns an unsubscribe handle. */
  onRemoteCommand(listener: (command: CommandLite, env: CommandEnvelope) => void): () => void;
  /** Replace our published awareness state. */
  setAwareness(state: Partial<AwarenessState>): void;
  /** Subscribe to awareness changes (other peers added / updated / removed). */
  onAwareness(listener: () => void): () => void;
  /** All currently visible remote awareness states (excluding self). */
  getRemoteStates(): ReadonlyArray<{ clientId: number; state: AwarenessState }>;
  /** Connection lifecycle ("disconnected" | "connecting" | "connected"). */
  getStatus(): "disconnected" | "connecting" | "connected";
  onStatus(listener: (status: "disconnected" | "connecting" | "connected") => void): () => void;
  /** Tear everything down. Idempotent. */
  destroy(): void;
}

const STORAGE_KEY = "officeai.peerId";

function loadOrMintPeerId(): string {
  if (typeof window === "undefined") {
    return Math.random().toString(36).slice(2);
  }
  try {
    const existing = window.sessionStorage.getItem(STORAGE_KEY);
    if (existing) return existing;
    const next =
      typeof crypto !== "undefined" && "randomUUID" in crypto
        ? crypto.randomUUID()
        : Math.random().toString(36).slice(2);
    window.sessionStorage.setItem(STORAGE_KEY, next);
    return next;
  } catch {
    return Math.random().toString(36).slice(2);
  }
}

/**
 * Browser-side Yjs room client. Creates a Y.Doc + WebsocketProvider
 * pair, exposes a typed command log + awareness API, and is shared
 * via a small reference-count cache so two editors targeting the
 * same `roomId` (e.g. side-by-side panels in tests) don't fight
 * over the websocket.
 */
class RoomClientImpl implements RoomClient {
  readonly identity: AnonymousIdentity;
  readonly peerId: string;

  private readonly doc: Y.Doc;
  private readonly provider: WebsocketProvider;
  private readonly log: Y.Array<unknown>;
  private readonly product: ProductKind;
  private status: "disconnected" | "connecting" | "connected" = "disconnected";
  private readonly statusListeners = new Set<(s: "disconnected" | "connecting" | "connected") => void>();
  private readonly commandListeners = new Set<(command: CommandLite, env: CommandEnvelope) => void>();
  private readonly awarenessListeners = new Set<() => void>();
  private seq = 0;
  private destroyed = false;
  private readonly observeFn: (event: Y.YArrayEvent<unknown>) => void;
  private readonly statusFn: (e: { status: "disconnected" | "connecting" | "connected" }) => void;
  private readonly awarenessFn: () => void;

  constructor(opts: RoomClientOptions) {
    this.product = opts.product;
    this.peerId = loadOrMintPeerId();
    this.identity = generateAnonymousIdentity(this.peerId);
    this.doc = new Y.Doc();
    this.log = this.doc.getArray<unknown>(COMMAND_LOG_KEY);

    this.provider = new WebsocketProvider(opts.url, opts.roomId, this.doc, {
      connect: true,
    });

    this.observeFn = (event: Y.YArrayEvent<unknown>): void => {
      // We only care about *inserts*. Deletes / moves don't happen
      // in our append-only log today.
      let inserted: ReadonlyArray<unknown> | null = null;
      for (const change of event.changes.delta) {
        if (change.insert && Array.isArray(change.insert)) {
          inserted = change.insert as ReadonlyArray<unknown>;
        }
      }
      if (!inserted) return;
      for (const raw of inserted) {
        const env = decodeCommand(raw);
        if (!env) continue;
        if (isOurEcho(env, this.peerId)) continue;
        const lite = envelopeToCommandLite(env);
        for (const l of this.commandListeners) {
          try {
            l(lite, env);
          } catch (err) {
            console.error("[realtime] remote command listener threw:", err);
          }
        }
      }
    };
    this.log.observe(this.observeFn);

    this.statusFn = (e): void => {
      this.status = e.status;
      for (const l of this.statusListeners) {
        try {
          l(e.status);
        } catch {
          /* noop */
        }
      }
    };
    // y-websocket fires "status" with `{ status: "connected" | … }`.
    (
      this.provider as unknown as {
        on: (e: string, fn: (data: { status: "disconnected" | "connecting" | "connected" }) => void) => void;
      }
    ).on("status", this.statusFn);

    this.awarenessFn = (): void => {
      for (const l of this.awarenessListeners) {
        try {
          l();
        } catch {
          /* noop */
        }
      }
    };
    this.provider.awareness.on("change", this.awarenessFn);

    this.provider.awareness.setLocalState({
      user: this.identity,
      product: this.product,
      lastSeen: Date.now(),
    } satisfies AwarenessState);
  }

  broadcastCommand(command: CommandLite): void {
    if (this.destroyed) return;
    const envelope: CommandEnvelope = {
      peerId: this.peerId,
      seq: ++this.seq,
      command: {
        type: command.type,
        payload: command.payload,
        ...(command.source ? { source: command.source } : {}),
        ...(command.agentId ? { agentId: command.agentId } : {}),
      },
    };
    let encoded: CommandEnvelope;
    try {
      encoded = encodeCommand(envelope);
    } catch (err) {
      console.warn("[realtime] dropping non-serializable command:", command.type, err);
      return;
    }
    this.log.push([encoded]);
  }

  onRemoteCommand(listener: (command: CommandLite, env: CommandEnvelope) => void): () => void {
    this.commandListeners.add(listener);
    return () => this.commandListeners.delete(listener);
  }

  setAwareness(state: Partial<AwarenessState>): void {
    if (this.destroyed) return;
    const current = this.provider.awareness.getLocalState() as AwarenessState | null;
    const base: AwarenessState = current ?? {
      user: this.identity,
      product: this.product,
      lastSeen: Date.now(),
    };
    const next: AwarenessState = {
      ...base,
      ...state,
      user: state.user ?? base.user,
      product: state.product ?? base.product,
      lastSeen: Date.now(),
    };
    this.provider.awareness.setLocalState(next);
  }

  onAwareness(listener: () => void): () => void {
    this.awarenessListeners.add(listener);
    return () => this.awarenessListeners.delete(listener);
  }

  getRemoteStates(): ReadonlyArray<{ clientId: number; state: AwarenessState }> {
    const states = this.provider.awareness.getStates();
    const out: { clientId: number; state: AwarenessState }[] = [];
    const ourClientId = this.doc.clientID;
    for (const [clientId, state] of states.entries()) {
      if (clientId === ourClientId) continue;
      if (!state || typeof state !== "object") continue;
      const candidate = state as AwarenessState;
      if (!candidate.user || typeof candidate.user.id !== "string") continue;
      out.push({ clientId, state: candidate });
    }
    return out;
  }

  getStatus(): "disconnected" | "connecting" | "connected" {
    return this.status;
  }

  onStatus(listener: (s: "disconnected" | "connecting" | "connected") => void): () => void {
    this.statusListeners.add(listener);
    return () => this.statusListeners.delete(listener);
  }

  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    try {
      this.log.unobserve(this.observeFn);
    } catch {
      /* noop */
    }
    try {
      this.provider.awareness.off("change", this.awarenessFn);
    } catch {
      /* noop */
    }
    try {
      this.provider.awareness.setLocalState(null);
    } catch {
      /* noop */
    }
    try {
      (
        this.provider as unknown as {
          off: (
            e: string,
            fn: (data: { status: "disconnected" | "connecting" | "connected" }) => void
          ) => void;
        }
      ).off?.("status", this.statusFn);
    } catch {
      /* noop */
    }
    try {
      this.provider.destroy();
    } catch {
      /* noop */
    }
    try {
      this.doc.destroy();
    } catch {
      /* noop */
    }
    this.commandListeners.clear();
    this.awarenessListeners.clear();
    this.statusListeners.clear();
  }
}

export function createRoomClient(opts: RoomClientOptions): RoomClient {
  return new RoomClientImpl(opts);
}
