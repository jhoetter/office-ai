/**
 * Yjs room publishing for `apply` subcommands.
 *
 * Why this exists: when the office-agent CLI runs inside the hof-os
 * sub-agent sandbox, the user is watching the same document in their
 * browser through `@officeai/react-editors` joined to room
 * `hofos/asset/<key>`. Saving the result to S3 and bumping a presigned
 * URL makes the editor full-remount on every checkpoint — a janky
 * "F5" flash that breaks the "watching a coworker edit" feel.
 *
 * The fix is to push the same `CommandLite` envelopes the agent just
 * applied locally onto the room's shared `Y.Array<CommandEnvelope>`
 * (the substrate `useCommandBroadcast` already drains in the editor).
 * That way browser peers apply the commands in-place via
 * `agent.applyCommand({ ..., source: "system" })` and the document
 * morphs without unmounting.
 *
 * This module is intentionally a thin Node-side equivalent of
 * `apps/web/app/lib/realtime/RoomClient.ts` — different tradeoffs:
 *
 *   - Single-shot lifetime: connect → wait for sync → push → close.
 *     We don't track remote awareness or maintain a persistent Y.Doc
 *     because the CLI only writes; it doesn't read.
 *
 *   - Optional `--clear-room-after`: with the file save complete, the
 *     baseline OOXML on S3 already encodes every command in the room.
 *     Late-joining browsers refetch S3 + only need fresh commands
 *     after the save, so wiping the Y.Array keeps the invariant
 *     "room.commands == ops not yet baked into latest S3 save". Without
 *     this, a peer joining 5 minutes after the agent finished would
 *     re-apply every historical command on top of an already-current
 *     document.
 */
import { setTimeout as delay } from "node:timers/promises";
import type { Command } from "commander";
import * as Y from "yjs";
import { WebsocketProvider } from "y-websocket";
import { WebSocket as NodeWebSocket } from "ws";
import type { CommandLite, CommandSource } from "@officeai/core";

/**
 * Wire constants duplicated from `@officeai/realtime` so the agent
 * package stays NPM-publishable without taking a runtime dep on the
 * private workspace package. If the codec ever evolves, update both
 * sites in lock-step (there is a vitest contract test in
 * @officeai/realtime that pins the on-wire shape).
 */
const COMMAND_LOG_KEY = "commands";

/** On-wire envelope. Matches `@officeai/realtime` `CommandEnvelope`. */
interface CommandEnvelope<TPayload = unknown> {
  readonly peerId: string;
  readonly seq: number;
  readonly command: {
    readonly type: string;
    readonly payload: TPayload;
    readonly source?: CommandSource;
    readonly agentId?: string;
  };
}

/**
 * Subset of `@officeai/realtime` `AwarenessState` that the CLI
 * publishes. Cursor is opaque (per-product schema) so we accept any
 * JSON object here; the editor's product-specific awareness consumer
 * picks out the fields it cares about.
 */
interface AwarenessUserState {
  readonly id: string;
  readonly name: string;
  readonly color?: string;
}
interface AwarenessState {
  readonly user: AwarenessUserState;
  readonly product?: "docx" | "xlsx" | "pptx" | "pdf";
  readonly cursor?: Record<string, unknown>;
  readonly lastSeen?: number;
}

function encodeCommand<TPayload>(envelope: CommandEnvelope<TPayload>): CommandEnvelope<TPayload> {
  const json = JSON.stringify(envelope);
  return JSON.parse(json) as CommandEnvelope<TPayload>;
}

/**
 * Default agent-cursor color: `bit-orange` from the hofOS palette
 * (`AGENTS.md` brand section). Picked so the editor's PresenceStack
 * visually distinguishes the office-agent from anonymous human peers
 * (which use the FNV-1a-derived peer palette).
 */
const DEFAULT_AGENT_COLOR = "#F4A51C";
const DEFAULT_AGENT_NAME = "Office Agent";

export interface RealtimeFlags {
  readonly room?: string;
  readonly realtimeUrl?: string;
  readonly agentName?: string;
  readonly agentColor?: string;
  readonly agentCursor?: string;
  readonly clearRoomAfter?: boolean;
}

export interface RealtimePublishOptions extends RealtimeFlags {
  /** Logical product surface — drives the awareness `product` field. */
  readonly product: "docx" | "xlsx" | "pptx";
  /**
   * Stable identity for this CLI invocation; reused as both the
   * awareness `user.id` and the embedded `agentId` on every command
   * envelope so the editor can attribute changes back to a specific
   * agent run if needed.
   */
  readonly agentId?: string;
}

/**
 * Add the realtime opt-in flags to an `apply` / `apply-file` commander
 * command. All fields are optional: when `--room` and a resolved
 * realtime URL are both absent the publish step becomes a no-op,
 * preserving the legacy "write file, that's it" behaviour.
 *
 * The realtime URL falls back to `OAI_REALTIME_URL` env so the
 * sub-agent host can pin the relay endpoint once on the sandbox
 * environment instead of asking the LLM to thread it through every
 * apply call.
 */
export function attachRealtimeFlags(cmd: Command): Command {
  return cmd
    .option(
      "--room <id>",
      "Yjs room id to publish into (e.g. hofos/asset/<encoded-key>). " +
        "Falls back to OAI_ROOM_ID env. When absent, realtime publishing is skipped."
    )
    .option(
      "--realtime-url <wsUrl>",
      "Yjs websocket relay URL (ws:// or wss://). Falls back to OAI_REALTIME_URL env. " +
        "When absent, realtime publishing is skipped."
    )
    .option(
      "--agent-name <label>",
      `Display name for the agent in the editor's PresenceStack (default "${DEFAULT_AGENT_NAME}").`
    )
    .option(
      "--agent-color <hex>",
      `Hex color for the agent's awareness avatar (default ${DEFAULT_AGENT_COLOR}).`
    )
    .option(
      "--agent-cursor <json>",
      "Optional awareness cursor payload (JSON). Drives the editor's section-pointer overlay."
    )
    .option(
      "--clear-room-after",
      "After publishing, transactionally clear the Y.Array<CommandEnvelope>. Use when the " +
        "file save has already baked every published command into the canonical artifact " +
        "(see hof-os office sub-agent).",
      false
    );
}

/** Resolve the effective room + url from flags + env. */
function resolveTarget(opts: RealtimeFlags): { room: string; url: string } | null {
  const room = (opts.room ?? process.env.OAI_ROOM_ID ?? "").trim();
  const url = (opts.realtimeUrl ?? process.env.OAI_REALTIME_URL ?? "").trim();
  if (!room || !url) return null;
  return { room, url };
}

/**
 * Best-effort cursor parser. Bad JSON is silently dropped (logged to
 * stderr) so a malformed `--agent-cursor` never aborts the apply.
 */
function parseCursor(raw: string | undefined): Record<string, unknown> | null {
  if (!raw || !raw.trim()) return null;
  try {
    const parsed = JSON.parse(raw);
    if (parsed && typeof parsed === "object") return parsed as Record<string, unknown>;
  } catch (err) {
    process.stderr.write(`[realtime] ignoring invalid --agent-cursor JSON: ${(err as Error).message}\n`);
  }
  return null;
}

interface PublishResult {
  readonly published: number;
  readonly room: string;
  readonly url: string;
  readonly cleared: boolean;
}

/**
 * Connect to `<url>/<room>`, wait for the initial sync round-trip,
 * append `commands` to the shared `Y.Array<CommandEnvelope>`, set
 * awareness state once, then close.
 *
 * Returns `null` when realtime publishing is not configured (no room
 * and/or url) so callers can branch on the legacy path silently.
 *
 * Failures connect/sync are reported on stderr but never throw — a
 * down realtime relay must NOT break the file-write path.
 */
export async function publishCommandsToRealtime(
  opts: RealtimePublishOptions,
  commands: ReadonlyArray<CommandLite>
): Promise<PublishResult | null> {
  const target = resolveTarget(opts);
  if (!target) return null;
  if (commands.length === 0) {
    return { published: 0, room: target.room, url: target.url, cleared: false };
  }

  // y-websocket polls global `WebSocket`; Node 22 has it natively but
  // 20 doesn't, so install a `ws`-backed shim before constructing the
  // provider. Idempotent — overrides only when missing.
  const g = globalThis as unknown as { WebSocket?: unknown };
  if (typeof g.WebSocket === "undefined") {
    g.WebSocket = NodeWebSocket as unknown as typeof WebSocket;
  }

  const doc = new Y.Doc();
  const log = doc.getArray<unknown>(COMMAND_LOG_KEY);
  const provider = new WebsocketProvider(target.url, target.room, doc, {
    connect: true,
    WebSocketPolyfill: NodeWebSocket as unknown as typeof WebSocket,
  });

  const peerId = `office-agent-cli/${opts.agentId ?? "anon"}/${process.pid}/${Date.now().toString(36)}`;
  const agentId = opts.agentId ?? "office-agent-cli";
  const name = (opts.agentName ?? DEFAULT_AGENT_NAME).trim() || DEFAULT_AGENT_NAME;
  const color = (opts.agentColor ?? DEFAULT_AGENT_COLOR).trim() || DEFAULT_AGENT_COLOR;
  const cursor = parseCursor(opts.agentCursor);

  try {
    await waitForSync(provider, 5000);

    const awarenessPayload: AwarenessState = {
      user: { id: agentId, name, color },
      product: opts.product,
      lastSeen: Date.now(),
      ...(cursor ? { cursor: cursor as AwarenessState["cursor"] } : {}),
    };
    provider.awareness.setLocalState(awarenessPayload);

    const envelopes: CommandEnvelope[] = commands.map((cmd, i) => {
      const envelope: CommandEnvelope = {
        peerId,
        seq: i + 1,
        command: {
          type: cmd.type,
          payload: cmd.payload,
          source: cmd.source ?? "agent",
          agentId: cmd.agentId ?? agentId,
        },
      };
      return encodeCommand(envelope);
    });

    doc.transact(() => {
      log.push(envelopes);
    }, peerId);

    // Give the provider a tick to flush over the wire — `push` is
    // synchronous against the local Yjs structure, but the underlying
    // websocket frame is queued on the next macrotask.
    await delay(150);

    let cleared = false;
    if (opts.clearRoomAfter) {
      doc.transact(() => {
        log.delete(0, log.length);
      }, peerId);
      await delay(75);
      cleared = true;
    }

    provider.awareness.setLocalState(null);
    await delay(50);

    return {
      published: envelopes.length,
      room: target.room,
      url: target.url,
      cleared,
    };
  } catch (err) {
    process.stderr.write(
      `[realtime] publish to ${target.url}/${target.room} failed: ${(err as Error).message}\n`
    );
    return null;
  } finally {
    try {
      provider.destroy();
    } catch {
      /* noop */
    }
    try {
      doc.destroy();
    } catch {
      /* noop */
    }
  }
}

/**
 * Resolve when the y-websocket provider reports `synced` (its
 * canonical "we've round-tripped a sync step1/2 with the relay"
 * event). Falls back to a hard timeout so a wedged relay never hangs
 * the CLI indefinitely.
 */
function waitForSync(provider: WebsocketProvider, timeoutMs: number): Promise<void> {
  if (provider.synced) return Promise.resolve();
  return new Promise<void>((resolve, reject) => {
    const timer = setTimeout(() => {
      provider.off("sync", onSync);
      provider.off("status", onStatus);
      reject(new Error(`timed out after ${timeoutMs}ms waiting for initial sync`));
    }, timeoutMs);

    const cleanup = (): void => {
      clearTimeout(timer);
      provider.off("sync", onSync);
      provider.off("status", onStatus);
    };

    const onSync = (synced: boolean): void => {
      if (synced) {
        cleanup();
        resolve();
      }
    };
    const onStatus = (e: { status: "disconnected" | "connecting" | "connected" }): void => {
      // y-websocket emits a `sync(true)` shortly after `connected`;
      // we wait for that to be sure the room has loaded its baseline
      // state vector before we push, so our envelopes don't race the
      // server's syncStep2 reply (which would land out-of-order in
      // late-joining browsers).
      if (e.status === "disconnected") {
        cleanup();
        reject(new Error("websocket disconnected before initial sync"));
      }
    };

    provider.on("sync", onSync);
    provider.on("status", onStatus);
  });
}
