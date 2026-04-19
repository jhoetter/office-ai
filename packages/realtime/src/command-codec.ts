import type { CommandLite, CommandSource } from "@officeai/core";

/**
 * Y.Doc map key under which the shared command log lives. Every peer
 * appends to and reads from the same `Y.Array<CommandEnvelope>`.
 */
export const COMMAND_LOG_KEY = "commands";

/**
 * Subset of `CommandLite` we transmit. Payloads must be JSON-safe;
 * the codec asserts on serialize and throws loudly so unprintable
 * refs don't silently desync peers.
 */
export interface SerializableCommand<TPayload = unknown> {
  readonly type: string;
  readonly payload: TPayload;
  readonly source?: CommandSource;
  readonly agentId?: string;
}

/**
 * On-the-wire envelope. `peerId` is the originating peer's id so the
 * receiver can suppress its own echo (Y.Array fires `observe` for
 * locally-inserted entries too).
 */
export interface CommandEnvelope<TPayload = unknown> {
  readonly peerId: string;
  readonly seq: number;
  readonly command: SerializableCommand<TPayload>;
}

/**
 * Serialize a command for the wire. Throws synchronously when the
 * payload contains values that wouldn't round-trip via JSON (functions,
 * cycles, bigints, …) so a buggy command never makes it onto the
 * shared log where it would corrupt every peer's state.
 */
export function encodeCommand<TPayload>(envelope: CommandEnvelope<TPayload>): CommandEnvelope<TPayload> {
  const json = JSON.stringify(envelope);
  if (typeof json !== "string") {
    throw new Error(
      `realtime: command envelope for type "${envelope.command.type}" is not JSON-serializable`
    );
  }
  return JSON.parse(json) as CommandEnvelope<TPayload>;
}

/**
 * Inverse of {@link encodeCommand}. Defensive: returns `null` for
 * malformed entries (e.g. an older client wrote a different shape) so
 * the receiver can skip rather than crash.
 */
export function decodeCommand<TPayload>(raw: unknown): CommandEnvelope<TPayload> | null {
  if (!raw || typeof raw !== "object") return null;
  const obj = raw as Record<string, unknown>;
  if (typeof obj.peerId !== "string") return null;
  if (typeof obj.seq !== "number") return null;
  const command = obj.command as Record<string, unknown> | undefined;
  if (!command || typeof command !== "object") return null;
  if (typeof command.type !== "string") return null;
  return obj as unknown as CommandEnvelope<TPayload>;
}

/**
 * Returns true iff this envelope was minted by us (and therefore
 * should not be re-dispatched against the local agent).
 */
export function isOurEcho(envelope: CommandEnvelope, ourPeerId: string): boolean {
  return envelope.peerId === ourPeerId;
}

/**
 * Convert a serializable command back into the loose `CommandLite`
 * accepted by every agent in the workspace. Mirrors the contract in
 * `@officeai/core` — `source` defaults to `"system"` when missing,
 * matching the bus's `normalize()` behaviour.
 */
export function envelopeToCommandLite(envelope: CommandEnvelope): CommandLite {
  const { command } = envelope;
  return {
    type: command.type,
    payload: command.payload,
    ...(command.source ? { source: command.source } : {}),
    ...(command.agentId ? { agentId: command.agentId } : {}),
  };
}
