import { describe, expect, it } from "vitest";
import {
  decodeCommand,
  encodeCommand,
  envelopeToCommandLite,
  isOurEcho,
  type CommandEnvelope,
} from "./command-codec";

describe("command-codec", () => {
  it("round-trips a JSON-safe envelope", () => {
    const env: CommandEnvelope = {
      peerId: "peer-a",
      seq: 42,
      command: {
        type: "xlsx:set-cell-value",
        payload: { sheet: "Sheet1", cell: "A1", value: { kind: "string", text: "hi" } },
      },
    };
    const encoded = encodeCommand(env);
    expect(encoded).toEqual(env);
    const decoded = decodeCommand(JSON.parse(JSON.stringify(encoded)));
    expect(decoded).toEqual(env);
  });

  it("returns null for malformed payloads", () => {
    expect(decodeCommand(null)).toBeNull();
    expect(decodeCommand({})).toBeNull();
    expect(decodeCommand({ peerId: 12, seq: 1, command: { type: "x" } })).toBeNull();
    expect(decodeCommand({ peerId: "p", seq: "1", command: { type: "x" } })).toBeNull();
    expect(decodeCommand({ peerId: "p", seq: 1, command: {} })).toBeNull();
  });

  it("isOurEcho only fires for the originating peer", () => {
    const env: CommandEnvelope = {
      peerId: "peer-a",
      seq: 1,
      command: { type: "noop", payload: {} },
    };
    expect(isOurEcho(env, "peer-a")).toBe(true);
    expect(isOurEcho(env, "peer-b")).toBe(false);
  });

  it("envelopeToCommandLite preserves source / agentId when present", () => {
    const lite = envelopeToCommandLite({
      peerId: "peer-x",
      seq: 7,
      command: {
        type: "docx:insert-text",
        payload: { at: "p:0", text: "hello" },
        source: "human",
        agentId: undefined,
      },
    });
    expect(lite.type).toBe("docx:insert-text");
    expect(lite.payload).toEqual({ at: "p:0", text: "hello" });
    expect(lite.source).toBe("human");
  });

  it("rejects payloads with non-JSON values when encoding", () => {
    const env: CommandEnvelope = {
      peerId: "p",
      seq: 1,
      command: {
        type: "broken",
        // BigInt is not JSON-serializable.
        payload: { big: BigInt(1) },
      },
    };
    expect(() => encodeCommand(env)).toThrow();
  });
});
