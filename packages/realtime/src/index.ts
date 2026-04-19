/**
 * `@officeai/realtime` — multi-user collaboration substrate.
 *
 * Headless, React-free package that owns:
 *
 * - The Yjs room shape (`Y.Doc` + a `Y.Array<SerializedCommand>`
 *   command log keyed by `commands`).
 * - The serialization codec for `CommandLite` envelopes so every
 *   typed command from any agent flows over the wire deterministically.
 * - The awareness payload (anonymous identity + cursor / selection)
 *   shared across DOCX / XLSX / PPTX editors.
 *
 * The "commands are the only mutation path" invariant from the core
 * spec is what makes this work: a peer applies a remote command via
 * its own `agent.dispatch(...)` and the local model converges by
 * construction. The single Y.Array gives total ordering across peers
 * for free; awareness gives Google-Docs-style presence.
 */
export {
  COMMAND_LOG_KEY,
  encodeCommand,
  decodeCommand,
  envelopeToCommandLite,
  isOurEcho,
  type CommandEnvelope,
  type SerializableCommand,
} from "./command-codec";
export {
  generateAnonymousIdentity,
  colorForPeer,
  ANONYMOUS_NAME_POOL,
  PRESENCE_PALETTE,
  type AnonymousIdentity,
} from "./identity";
export {
  type AwarenessState,
  type DocxCursor,
  type XlsxSelection,
  type PptxSelection,
  type RemotePresence,
} from "./awareness-types";
