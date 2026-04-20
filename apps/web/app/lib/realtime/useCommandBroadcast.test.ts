import { describe, expect, it, vi } from "vitest";
import type { CommandLite } from "@officeai/core";
import type { CommandEnvelope } from "@officeai/realtime";
import type { RoomClient } from "./RoomClient";
import { wireBroadcast, type BroadcastableAgent } from "./useCommandBroadcast";

/**
 * Regression tests for the broadcast wiring. The contract is small
 * but easy to break: the wrong method name on the agent fails
 * silently in production because every cast hides the mismatch.
 *
 * We test `wireBroadcast` (the React-free helper) rather than the
 * `useCommandBroadcast` hook so we don't need a React testing
 * harness in this app's test rig.
 */

interface FakeRoom {
  readonly client: RoomClient;
  emitRemote: (cmd: CommandLite) => void;
  readonly broadcastSpy: ReturnType<typeof vi.fn>;
}

function makeFakeRoom(): FakeRoom {
  let remoteListener: ((cmd: CommandLite, env: CommandEnvelope) => void) | null = null;
  const broadcastSpy = vi.fn();
  const client = {
    identity: { id: "self", name: "self", color: "#000" },
    peerId: "self",
    broadcastCommand: broadcastSpy,
    onRemoteCommand: (l: (cmd: CommandLite, env: CommandEnvelope) => void) => {
      remoteListener = l;
      return () => {
        remoteListener = null;
      };
    },
    setAwareness: vi.fn(),
    onAwareness: () => () => {},
    getRemoteStates: () => [],
    getStatus: () => "connected" as const,
    onStatus: () => () => {},
    destroy: vi.fn(),
  } as unknown as RoomClient;

  return {
    client,
    broadcastSpy,
    emitRemote: (cmd: CommandLite) => {
      remoteListener?.(cmd, {
        peerId: "peer-2",
        seq: 1,
        command: { type: cmd.type, payload: cmd.payload },
      });
    },
  };
}

interface FakeAgent {
  readonly agent: BroadcastableAgent;
  emitMutation: (m: {
    readonly status: string;
    readonly command: { readonly type: string; readonly payload: unknown; readonly source: string };
  }) => void;
}

function makeFakeAgent(applyImpl: BroadcastableAgent["applyCommand"]): FakeAgent {
  let listener: Parameters<BroadcastableAgent["subscribe"]>[0] | null = null;
  const agent: BroadcastableAgent = {
    subscribe(l) {
      listener = l;
      return () => {
        listener = null;
      };
    },
    applyCommand: applyImpl,
  };
  return {
    agent,
    emitMutation: (m) => {
      listener?.(null, m);
    },
  };
}

describe("wireBroadcast", () => {
  it("uses agent.applyCommand (NOT dispatch) for inbound remote commands", async () => {
    const apply = vi.fn().mockResolvedValue({});
    const { agent } = makeFakeAgent(apply);
    const fake = makeFakeRoom();
    const ref = { current: 0 };

    const unsub = wireBroadcast({ agent, room: fake.client, applyingRemote: ref });
    fake.emitRemote({
      type: "docx:insert-text",
      payload: { at: { paragraph: 0, run: 0, offset: 0 }, text: "hi" },
    });

    await Promise.resolve();
    await Promise.resolve();
    expect(apply).toHaveBeenCalledTimes(1);
    expect(apply.mock.calls[0]?.[0]).toMatchObject({ type: "docx:insert-text", source: "system" });
    unsub();
  });

  it("warns (but does not throw) when a remote command rejects", async () => {
    const apply = vi.fn().mockResolvedValue({ rejection: { code: "invalid-payload", message: "boom" } });
    const { agent } = makeFakeAgent(apply);
    const fake = makeFakeRoom();
    const ref = { current: 0 };
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});

    const unsub = wireBroadcast({ agent, room: fake.client, applyingRemote: ref });
    fake.emitRemote({ type: "docx:insert-text", payload: {} });

    await Promise.resolve();
    await Promise.resolve();
    expect(warn).toHaveBeenCalledWith(
      expect.stringContaining("remote command rejected"),
      "docx:insert-text",
      "invalid-payload",
      "boom"
    );
    warn.mockRestore();
    unsub();
  });

  it("ignores local mutations dispatched while applying a remote command (no echo loop)", async () => {
    const resolvers: Array<(v: unknown) => void> = [];
    const apply: BroadcastableAgent["applyCommand"] = vi.fn(() => {
      return new Promise<unknown>((res) => {
        resolvers.push(res);
      });
    });
    const fakeAgent = makeFakeAgent(apply);
    const fake = makeFakeRoom();
    const ref = { current: 0 };

    const unsub = wireBroadcast({ agent: fakeAgent.agent, room: fake.client, applyingRemote: ref });
    fake.emitRemote({ type: "docx:insert-text", payload: {} });

    fakeAgent.emitMutation({
      status: "approved",
      command: { type: "docx:insert-text", payload: {}, source: "system" },
    });
    expect(fake.broadcastSpy).not.toHaveBeenCalled();

    resolvers[0]?.({});
    await Promise.resolve();
    await Promise.resolve();
    unsub();
  });

  it("broadcasts approved local mutations onto the room log", async () => {
    const apply = vi.fn().mockResolvedValue({});
    const fakeAgent = makeFakeAgent(apply);
    const fake = makeFakeRoom();
    const ref = { current: 0 };

    const unsub = wireBroadcast({ agent: fakeAgent.agent, room: fake.client, applyingRemote: ref });
    fakeAgent.emitMutation({
      status: "approved",
      command: { type: "xlsx:set-cell-value", payload: { sheet: "S", ref: "A1", value: 1 }, source: "human" },
    });
    expect(fake.broadcastSpy).toHaveBeenCalledTimes(1);
    expect(fake.broadcastSpy.mock.calls[0]?.[0]).toMatchObject({
      type: "xlsx:set-cell-value",
      source: "human",
    });
    unsub();
  });

  it("skips mutations that the shouldBroadcast filter rejects", async () => {
    const apply = vi.fn().mockResolvedValue({});
    const fakeAgent = makeFakeAgent(apply);
    const fake = makeFakeRoom();
    const ref = { current: 0 };

    const unsub = wireBroadcast({
      agent: fakeAgent.agent,
      room: fake.client,
      applyingRemote: ref,
      shouldBroadcast: (t) => t !== "xlsx:set-selection",
    });
    fakeAgent.emitMutation({
      status: "approved",
      command: { type: "xlsx:set-selection", payload: {}, source: "human" },
    });
    expect(fake.broadcastSpy).not.toHaveBeenCalled();
    unsub();
  });
});
