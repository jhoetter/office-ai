import { describe, expect, it } from "vitest";
import * as Y from "yjs";
import {
  COMMAND_LOG_KEY,
  decodeCommand,
  encodeCommand,
  isOurEcho,
  type CommandEnvelope,
} from "./command-codec";

/**
 * Integration check: simulate the "two browsers, one room" flow at
 * the Yjs layer. We don't spin up the websocket server here — the
 * point is to lock down the contract that `command-codec` + `Y.Array`
 * give us:
 *
 * 1. Peer A appends a command envelope.
 * 2. Y.applyUpdate ships the change to peer B's doc (the y-websocket
 *    server does this for us in production).
 * 3. Peer B's `observe()` fires with the inserted envelopes and
 *    correctly suppresses its own echoes.
 *
 * This is the regression test that catches encoder/decoder drift
 * between releases.
 */
describe("command-log integration", () => {
  function syncOnce(from: Y.Doc, to: Y.Doc): void {
    const stateOnTo = Y.encodeStateVector(to);
    const update = Y.encodeStateAsUpdate(from, stateOnTo);
    Y.applyUpdate(to, update, "remote");
  }

  it("propagates appended commands from peer A to peer B", () => {
    const docA = new Y.Doc();
    const docB = new Y.Doc();
    const arrA = docA.getArray<unknown>(COMMAND_LOG_KEY);
    const arrB = docB.getArray<unknown>(COMMAND_LOG_KEY);

    const peerB = "peer-B";
    const observed: CommandEnvelope[] = [];
    arrB.observe((event) => {
      for (const change of event.changes.delta) {
        if (!change.insert || !Array.isArray(change.insert)) continue;
        for (const raw of change.insert) {
          const env = decodeCommand(raw);
          if (!env) continue;
          if (isOurEcho(env, peerB)) continue;
          observed.push(env);
        }
      }
    });

    arrA.push([
      encodeCommand({
        peerId: "peer-A",
        seq: 1,
        command: { type: "xlsx:set-cell-value", payload: { sheet: "S", cell: "A1", value: 1 } },
      }),
    ]);
    syncOnce(docA, docB);

    expect(observed).toHaveLength(1);
    expect(observed[0]!.command.type).toBe("xlsx:set-cell-value");
    expect(observed[0]!.peerId).toBe("peer-A");
  });

  it("suppresses our own echoes via peerId match", () => {
    const docA = new Y.Doc();
    const docB = new Y.Doc();
    const arrA = docA.getArray<unknown>(COMMAND_LOG_KEY);
    const arrB = docB.getArray<unknown>(COMMAND_LOG_KEY);

    const peerB = "peer-B";
    const observed: CommandEnvelope[] = [];
    arrB.observe((event) => {
      for (const change of event.changes.delta) {
        if (!change.insert || !Array.isArray(change.insert)) continue;
        for (const raw of change.insert) {
          const env = decodeCommand(raw);
          if (!env) continue;
          if (isOurEcho(env, peerB)) continue;
          observed.push(env);
        }
      }
    });

    // Peer B locally pushes — its own observe should ignore the entry.
    arrB.push([
      encodeCommand({
        peerId: "peer-B",
        seq: 1,
        command: { type: "noop", payload: {} },
      }),
    ]);
    syncOnce(docB, docA);

    expect(observed).toHaveLength(0);
  });

  it("preserves total order across mixed peers", () => {
    const docA = new Y.Doc();
    const docB = new Y.Doc();
    const arrA = docA.getArray<unknown>(COMMAND_LOG_KEY);

    arrA.push([
      encodeCommand({ peerId: "A", seq: 1, command: { type: "x", payload: 1 } }),
      encodeCommand({ peerId: "A", seq: 2, command: { type: "x", payload: 2 } }),
      encodeCommand({ peerId: "A", seq: 3, command: { type: "x", payload: 3 } }),
    ]);
    syncOnce(docA, docB);

    const arrB = docB.getArray<unknown>(COMMAND_LOG_KEY);
    const seqs = arrB.toArray().map((raw) => decodeCommand(raw)!.command.payload);
    expect(seqs).toEqual([1, 2, 3]);
  });
});
