/**
 * Anonymous-identity generator. Modelled on Google Docs' "Anonymous
 * Quokka" pattern: every unauthenticated peer gets a friendly two-word
 * handle plus a deterministic accent colour derived from their peer
 * id. Pure, no globals, safe to call from server tests.
 */

/**
 * Pool of friendly adjective + animal combos. Kept short and tasteful
 * — the goal is "feels familiar" not "infinite uniqueness".
 */
export const ANONYMOUS_NAME_POOL: ReadonlyArray<readonly [string, string]> = [
  ["Quick", "Quokka"],
  ["Fuzzy", "Otter"],
  ["Bright", "Finch"],
  ["Calm", "Heron"],
  ["Brave", "Lynx"],
  ["Swift", "Marten"],
  ["Gentle", "Doe"],
  ["Curious", "Tapir"],
  ["Witty", "Magpie"],
  ["Cosy", "Hare"],
  ["Lively", "Wren"],
  ["Wise", "Owl"],
  ["Kind", "Beaver"],
  ["Sunny", "Squirrel"],
  ["Mellow", "Vole"],
  ["Nimble", "Stoat"],
  ["Plucky", "Puffin"],
  ["Snug", "Pika"],
  ["Bold", "Badger"],
  ["Merry", "Mouse"],
];

/**
 * Pleasant, accessible-on-light-and-dark colour ramp. Used to tint
 * remote cursors / selections / avatars per peer.
 */
export const PRESENCE_PALETTE: ReadonlyArray<string> = [
  "#ef4444", // red
  "#f97316", // orange
  "#eab308", // yellow
  "#22c55e", // green
  "#14b8a6", // teal
  "#3b82f6", // blue
  "#8b5cf6", // violet
  "#ec4899", // pink
  "#06b6d4", // cyan
  "#84cc16", // lime
];

/**
 * The opaque per-tab identity surfaced in awareness. `id` is unique
 * per session (a UUID-ish random string); `name` is rendered in the
 * presence avatar tooltip; `color` paints the remote cursor / outline.
 */
export interface AnonymousIdentity {
  readonly id: string;
  readonly name: string;
  readonly color: string;
}

/**
 * 32-bit FNV-1a hash. Stable across machines (no Math.random) so the
 * same `peerId` always picks the same name + colour even after a
 * server restart, which makes ghost cursors readable across reconnects.
 */
function fnv1a(input: string): number {
  let h = 0x811c9dc5;
  for (let i = 0; i < input.length; i++) {
    h ^= input.charCodeAt(i);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return h >>> 0;
}

export function colorForPeer(peerId: string): string {
  const idx = fnv1a(peerId) % PRESENCE_PALETTE.length;
  return PRESENCE_PALETTE[idx]!;
}

function nameForPeer(peerId: string): string {
  const idx = fnv1a(`name:${peerId}`) % ANONYMOUS_NAME_POOL.length;
  const [adj, animal] = ANONYMOUS_NAME_POOL[idx]!;
  return `${adj} ${animal}`;
}

/**
 * Mint an identity for a freshly-opened tab. `peerId` should be a
 * sufficiently random string (e.g. `crypto.randomUUID()` on the
 * client) so distinct tabs never collide on name / colour by accident.
 */
export function generateAnonymousIdentity(peerId: string): AnonymousIdentity {
  return {
    id: peerId,
    name: nameForPeer(peerId),
    color: colorForPeer(peerId),
  };
}
