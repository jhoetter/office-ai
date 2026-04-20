/**
 * Anonymous-identity generator. Modelled on Google Docs' "Anonymous
 * Quokka" pattern: every unauthenticated peer gets a friendly two-word
 * handle plus a deterministic accent colour derived from their peer
 * id. Pure, no globals, safe to call from server tests.
 */

/**
 * Adjective + animal pool, modelled on Google Docs' "Anonymous
 * Quokka" pattern. We mix and match independently rather than
 * pre-pairing so the effective pool is `adjectives.length *
 * animals.length` (currently 40 * 40 = 1600 unique combinations) —
 * plenty of headroom before the birthday-paradox bites for a
 * realistic concurrent-user count.
 */
export const ANONYMOUS_ADJECTIVES: ReadonlyArray<string> = [
  "Quick", "Fuzzy", "Bright", "Calm", "Brave", "Swift", "Gentle", "Curious",
  "Witty", "Cosy", "Lively", "Wise", "Kind", "Sunny", "Mellow", "Nimble",
  "Plucky", "Snug", "Bold", "Merry", "Faithful", "Cheerful", "Jolly", "Breezy",
  "Spirited", "Dapper", "Earnest", "Friendly", "Glowing", "Honest", "Joyful",
  "Lucky", "Noble", "Peppy", "Radiant", "Sincere", "Tender", "Upbeat", "Vivid",
  "Zesty",
];

export const ANONYMOUS_ANIMALS: ReadonlyArray<string> = [
  "Quokka", "Otter", "Finch", "Heron", "Lynx", "Marten", "Doe", "Tapir",
  "Magpie", "Hare", "Wren", "Owl", "Beaver", "Squirrel", "Vole", "Stoat",
  "Puffin", "Pika", "Badger", "Mouse", "Dog", "Cat", "Fox", "Wolf",
  "Panda", "Koala", "Rabbit", "Hedgehog", "Raccoon", "Sparrow", "Robin",
  "Falcon", "Dolphin", "Seal", "Penguin", "Capybara", "Llama", "Alpaca",
  "Manatee", "Narwhal",
];

/**
 * Backwards-compatible flattened pool for code that wanted to pick a
 * "named pair". New callers should use `nameForPeer` directly so we
 * get the full cross-product space.
 */
export const ANONYMOUS_NAME_POOL: ReadonlyArray<readonly [string, string]> =
  ANONYMOUS_ADJECTIVES.map((adj) => [adj, ANONYMOUS_ANIMALS[0]!] as const);

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
  // Two independent hashes so the adjective and animal slots are
  // chosen independently — gives us the full `40 * 40 = 1600`
  // combination space rather than the 40-entry diagonal.
  const adjIdx = fnv1a(`adj:${peerId}`) % ANONYMOUS_ADJECTIVES.length;
  const animalIdx = fnv1a(`animal:${peerId}`) % ANONYMOUS_ANIMALS.length;
  return `${ANONYMOUS_ADJECTIVES[adjIdx]} ${ANONYMOUS_ANIMALS[animalIdx]}`;
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
