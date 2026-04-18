import { describe, expect, it } from "vitest";
import { listRegisteredFunctions } from "./registered-functions.js";

describe("listRegisteredFunctions", () => {
  it("returns a non-empty alphabetically sorted catalogue", () => {
    const list = listRegisteredFunctions();
    expect(list.length).toBeGreaterThan(50);
    for (let i = 1; i < list.length; i++) {
      expect(list[i - 1].name <= list[i].name).toBe(true);
    }
  });

  it("includes a representative function from each P0 category", () => {
    const list = listRegisteredFunctions();
    const byName = (n: string) => list.find((f) => f.name === n);
    expect(byName("SUM")?.category).toBe("math");
    expect(byName("IF")?.category).toBe("logic");
    expect(byName("ISBLANK")?.category).toBe("info");
    expect(byName("VLOOKUP")?.category).toBe("lookup");
    expect(byName("CONCAT")?.category).toBe("text");
  });

  it("captures arity (min/max) for variadic functions", () => {
    const sum = listRegisteredFunctions().find((f) => f.name === "SUM");
    expect(sum).toBeDefined();
    expect(sum!.arity.min).toBeGreaterThanOrEqual(1);
    // SUM has unbounded args; the registry uses Number.POSITIVE_INFINITY.
    expect(sum!.arity.max).toBeGreaterThanOrEqual(sum!.arity.min);
  });

  it("flags volatile functions", () => {
    const list = listRegisteredFunctions();
    const now = list.find((f) => f.name === "NOW");
    if (now) expect(now.volatile).toBe(true);
    const rand = list.find((f) => f.name === "RAND");
    if (rand) expect(rand.volatile).toBe(true);
  });
});
