import { describe, expect, it } from "vitest";
import { CommandBus } from "./bus.js";
import { CommandError, type CommandHandler, type Mutation } from "./types.js";
import type { DocumentSnapshot } from "../types/document.js";

interface ToySnapshot extends DocumentSnapshot<{ value: number }> {
  format: "docx";
}

const initial: ToySnapshot = {
  format: "docx",
  revision: 0,
  root: { value: 0 },
  partHashes: {},
};

const incHandler: CommandHandler<{ by: number }, ToySnapshot> = {
  type: "toy:inc",
  apply(snapshot, payload) {
    return {
      next: {
        ...snapshot,
        revision: snapshot.revision + 1,
        root: { value: snapshot.root.value + payload.by },
      },
      diff: {
        format: "docx",
        fromRevision: snapshot.revision,
        toRevision: snapshot.revision + 1,
        changes: [
          {
            kind: "node-updated",
            nodeId: "root",
            path: ["root", "value"],
            field: "value",
            summary: `+${payload.by}`,
          },
        ],
      },
    };
  },
};

describe("CommandBus", () => {
  it("dispatches a human command and approves immediately", async () => {
    const bus = new CommandBus<ToySnapshot>(initial);
    bus.register(incHandler);
    const m = await bus.dispatch({ type: "toy:inc", payload: { by: 5 }, source: "human" });
    expect(m.status).toBe("approved");
    expect(bus.getApproved().root.value).toBe(5);
    expect(bus.getWorking().root.value).toBe(5);
    expect(bus.getPending()).toHaveLength(0);
  });

  it("queues an agent command as pending; approve collapses into approved", async () => {
    const bus = new CommandBus<ToySnapshot>(initial);
    bus.register(incHandler);
    const m: Mutation<ToySnapshot> = await bus.dispatch({
      type: "toy:inc",
      payload: { by: 3 },
      source: "agent",
      agentId: "a1",
    });
    expect(m.status).toBe("pending");
    expect(bus.getApproved().root.value).toBe(0);
    expect(bus.getWorking().root.value).toBe(3);
    bus.approveMutation(m.id);
    expect(bus.getApproved().root.value).toBe(3);
    expect(bus.getPending()).toHaveLength(0);
  });

  it("reject removes a pending mutation and rebases working", async () => {
    const bus = new CommandBus<ToySnapshot>(initial);
    bus.register(incHandler);
    const m1 = await bus.dispatch({ type: "toy:inc", payload: { by: 2 }, source: "agent" });
    const m2 = await bus.dispatch({ type: "toy:inc", payload: { by: 4 }, source: "agent" });
    expect(bus.getWorking().root.value).toBe(6);
    bus.rejectMutation(m1.id);
    expect(bus.getWorking().root.value).toBe(4);
    expect(bus.getPending()).toHaveLength(1);
    expect(bus.getPending()[0].id).toBe(m2.id);
  });

  it("rolls back approved snapshots", async () => {
    const bus = new CommandBus<ToySnapshot>(initial);
    bus.register(incHandler);
    await bus.dispatch({ type: "toy:inc", payload: { by: 1 }, source: "human" });
    await bus.dispatch({ type: "toy:inc", payload: { by: 1 }, source: "human" });
    expect(bus.getApproved().revision).toBe(2);
    bus.rollback(0);
    expect(bus.getApproved().revision).toBe(0);
    expect(bus.getApproved().root.value).toBe(0);
  });

  it("rejects when no handler is registered", async () => {
    const bus = new CommandBus<ToySnapshot>(initial);
    await expect(bus.dispatch({ type: "missing", payload: {}, source: "human" })).rejects.toBeInstanceOf(
      CommandError
    );
  });

  it("captures handler errors as a rejected mutation", async () => {
    const bus = new CommandBus<ToySnapshot>(initial);
    bus.register({
      type: "boom",
      apply() {
        throw new CommandError("kaboom", "no");
      },
    });
    const m = await bus.dispatch({ type: "boom", payload: {}, source: "human" });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("kaboom");
  });

  it("notifies subscribers after each mutation", async () => {
    const bus = new CommandBus<ToySnapshot>(initial);
    bus.register(incHandler);
    const seen: number[] = [];
    bus.subscribe((s) => seen.push(s.root.value));
    await bus.dispatch({ type: "toy:inc", payload: { by: 1 }, source: "human" });
    await bus.dispatch({ type: "toy:inc", payload: { by: 2 }, source: "human" });
    expect(seen).toEqual([1, 3]);
  });

  describe("undo / redo", () => {
    it("undoes the most recent approved mutation", async () => {
      const bus = new CommandBus<ToySnapshot>(initial);
      bus.register(incHandler);
      await bus.dispatch({ type: "toy:inc", payload: { by: 5 }, source: "human" });
      await bus.dispatch({ type: "toy:inc", payload: { by: 3 }, source: "human" });
      expect(bus.getApproved().root.value).toBe(8);
      expect(bus.canUndo()).toBe(true);
      const undone = bus.undo();
      expect(undone?.status).toBe("undone");
      expect(bus.getApproved().root.value).toBe(5);
      expect(bus.canRedo()).toBe(true);
    });

    it("redoes by re-applying against the current approved snapshot", async () => {
      const bus = new CommandBus<ToySnapshot>(initial);
      bus.register(incHandler);
      await bus.dispatch({ type: "toy:inc", payload: { by: 5 }, source: "human" });
      bus.undo();
      expect(bus.getApproved().root.value).toBe(0);
      const redone = bus.redo();
      expect(redone?.status).toBe("approved");
      expect(bus.getApproved().root.value).toBe(5);
      expect(bus.canRedo()).toBe(false);
    });

    it("clears the redo stack when a new authored mutation lands", async () => {
      const bus = new CommandBus<ToySnapshot>(initial);
      bus.register(incHandler);
      await bus.dispatch({ type: "toy:inc", payload: { by: 5 }, source: "human" });
      bus.undo();
      expect(bus.canRedo()).toBe(true);
      await bus.dispatch({ type: "toy:inc", payload: { by: 7 }, source: "human" });
      expect(bus.canRedo()).toBe(false);
      expect(bus.getApproved().root.value).toBe(7);
    });

    it("returns null when nothing to undo / redo", () => {
      const bus = new CommandBus<ToySnapshot>(initial);
      bus.register(incHandler);
      expect(bus.canUndo()).toBe(false);
      expect(bus.canRedo()).toBe(false);
      expect(bus.undo()).toBeNull();
      expect(bus.redo()).toBeNull();
    });

    it("supports multi-step undo + redo round-trips", async () => {
      const bus = new CommandBus<ToySnapshot>(initial);
      bus.register(incHandler);
      for (const v of [1, 2, 3, 4]) {
        await bus.dispatch({ type: "toy:inc", payload: { by: v }, source: "human" });
      }
      expect(bus.getApproved().root.value).toBe(10);
      bus.undo(); bus.undo(); bus.undo();
      expect(bus.getApproved().root.value).toBe(1);
      bus.redo(); bus.redo();
      expect(bus.getApproved().root.value).toBe(6);
    });
  });
});
