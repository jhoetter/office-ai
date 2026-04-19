import { describe, expect, it } from "vitest";
import {
  ANONYMOUS_NAME_POOL,
  PRESENCE_PALETTE,
  colorForPeer,
  generateAnonymousIdentity,
} from "./identity";

describe("identity", () => {
  it("name pool and palette are non-empty", () => {
    expect(ANONYMOUS_NAME_POOL.length).toBeGreaterThan(0);
    expect(PRESENCE_PALETTE.length).toBeGreaterThan(0);
  });

  it("colorForPeer is deterministic and uses the palette", () => {
    const c1 = colorForPeer("peer-1");
    const c2 = colorForPeer("peer-1");
    expect(c1).toBe(c2);
    expect(PRESENCE_PALETTE.includes(c1)).toBe(true);
  });

  it("generateAnonymousIdentity produces stable name + color for a peerId", () => {
    const a = generateAnonymousIdentity("peer-A");
    const b = generateAnonymousIdentity("peer-A");
    expect(a).toEqual(b);
    expect(a.name.split(" ").length).toBe(2);
  });

  it("different peerIds usually produce different names", () => {
    const names = new Set<string>();
    for (let i = 0; i < 20; i++) {
      names.add(generateAnonymousIdentity(`peer-${i}`).name);
    }
    expect(names.size).toBeGreaterThan(1);
  });
});
