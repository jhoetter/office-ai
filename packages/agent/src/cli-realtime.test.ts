/**
 * Realtime publishing contract test.
 *
 * Spins up an in-process Yjs websocket relay (mirrors the production
 * `apps/realtime-server` server) and verifies that running a `docx
 * apply` with `--room` + `--realtime-url` causes the supplied
 * commands to land on a connected subscriber's `Y.Array<commands>` —
 * the exact substrate the editor's `useCommandBroadcast` drains.
 *
 * Boundary: the test pins the on-wire envelope shape (peerId / seq /
 * command{type,payload,source,agentId}) so a future codec change in
 * `@officeai/realtime` would break this test loudly, prompting an
 * explicit update of the inlined codec in `cli-realtime.ts`.
 */
import { mkdtempSync, writeFileSync, readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve as resolvePath } from "node:path";
import { fileURLToPath } from "node:url";
import http, { type Server } from "node:http";
import type { AddressInfo } from "node:net";
import { WebSocketServer, WebSocket as NodeWebSocket } from "ws";
import * as Y from "yjs";
import { WebsocketProvider } from "y-websocket";
import * as syncProtocol from "y-protocols/sync";
import * as awarenessProtocol from "y-protocols/awareness";
import { encoding, decoding } from "lib0";
import { Document, Packer, Paragraph, TextRun, HeadingLevel } from "docx";
import { afterAll, afterEach, beforeAll, describe, expect, it } from "vitest";
import { runCli } from "./cli.js";

const here = dirname(fileURLToPath(import.meta.url));
void here;

const MESSAGE_SYNC = 0;
const MESSAGE_AWARENESS = 1;

class CapturedStream {
  chunks: string[] = [];
  write(s: string | Uint8Array): boolean {
    this.chunks.push(typeof s === "string" ? s : Buffer.from(s).toString("utf8"));
    return true;
  }
  text(): string {
    return this.chunks.join("");
  }
}

function makeIO() {
  const stdout = new CapturedStream();
  const stderr = new CapturedStream();
  return {
    io: {
      stdout: stdout as unknown as NodeJS.WritableStream,
      stderr: stderr as unknown as NodeJS.WritableStream,
    },
    stdout,
    stderr,
  };
}

interface Room {
  doc: Y.Doc;
  awareness: awarenessProtocol.Awareness;
  sockets: Set<NodeWebSocket>;
}

/**
 * Minimal in-process relay equivalent to apps/realtime-server. Lives
 * here so we don't have to wire a dev-dep on the workspace's relay
 * server package (it's `private`).
 */
async function startRelay(): Promise<{
  url: string;
  close: () => Promise<void>;
  rooms: Map<string, Room>;
}> {
  const rooms = new Map<string, Room>();
  const getRoom = (id: string): Room => {
    let room = rooms.get(id);
    if (!room) {
      const doc = new Y.Doc();
      const awareness = new awarenessProtocol.Awareness(doc);
      awareness.setLocalState(null);
      room = { doc, awareness, sockets: new Set() };
      doc.on("update", (update: Uint8Array, origin: unknown) => {
        const enc = encoding.createEncoder();
        encoding.writeVarUint(enc, MESSAGE_SYNC);
        syncProtocol.writeUpdate(enc, update);
        const buf = encoding.toUint8Array(enc);
        for (const ws of room!.sockets) {
          if (ws !== origin && ws.readyState === ws.OPEN) ws.send(buf);
        }
      });
      awareness.on(
        "update",
        (
          { added, updated, removed }: { added: number[]; updated: number[]; removed: number[] },
          origin: unknown
        ) => {
          const changed = added.concat(updated, removed);
          const enc = encoding.createEncoder();
          encoding.writeVarUint(enc, MESSAGE_AWARENESS);
          encoding.writeVarUint8Array(enc, awarenessProtocol.encodeAwarenessUpdate(awareness, changed));
          const buf = encoding.toUint8Array(enc);
          for (const ws of room!.sockets) {
            if (ws !== origin && ws.readyState === ws.OPEN) ws.send(buf);
          }
        }
      );
      rooms.set(id, room);
    }
    return room;
  };

  const httpServer: Server = http.createServer((_req, res) => {
    res.writeHead(200);
    res.end("ok");
  });
  const wss = new WebSocketServer({ server: httpServer });
  wss.on("connection", (ws, request) => {
    ws.binaryType = "arraybuffer";
    const url = new URL(request.url ?? "/", "http://localhost");
    const roomId = decodeURIComponent(url.pathname.replace(/^\//, "")) || "default";
    const room = getRoom(roomId);
    room.sockets.add(ws);
    {
      const enc = encoding.createEncoder();
      encoding.writeVarUint(enc, MESSAGE_SYNC);
      syncProtocol.writeSyncStep1(enc, room.doc);
      ws.send(encoding.toUint8Array(enc));
    }
    ws.on("message", (data: ArrayBuffer | Buffer) => {
      const bytes = data instanceof ArrayBuffer ? new Uint8Array(data) : new Uint8Array(data);
      const dec = decoding.createDecoder(bytes);
      const t = decoding.readVarUint(dec);
      if (t === MESSAGE_SYNC) {
        const enc = encoding.createEncoder();
        encoding.writeVarUint(enc, MESSAGE_SYNC);
        syncProtocol.readSyncMessage(dec, enc, room.doc, ws);
        if (encoding.length(enc) > 1) ws.send(encoding.toUint8Array(enc));
      } else if (t === MESSAGE_AWARENESS) {
        awarenessProtocol.applyAwarenessUpdate(room.awareness, decoding.readVarUint8Array(dec), ws);
      }
    });
    ws.on("close", () => {
      room.sockets.delete(ws);
      awarenessProtocol.removeAwarenessStates(room.awareness, [room.doc.clientID], null);
    });
  });

  return new Promise((resolve) => {
    httpServer.listen(0, "127.0.0.1", () => {
      const port = (httpServer.address() as AddressInfo).port;
      resolve({
        url: `ws://127.0.0.1:${port}`,
        rooms,
        close: () =>
          new Promise<void>((res2) => {
            for (const ws of wss.clients) {
              try {
                ws.close();
              } catch {
                /* noop */
              }
            }
            httpServer.close(() => res2());
          }),
      });
    });
  });
}

async function makeDocxFixture(): Promise<string> {
  const dir = mkdtempSync(join(tmpdir(), "office-agent-realtime-"));
  const doc = new Document({
    creator: "test",
    sections: [
      {
        properties: {},
        children: [
          new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Hello")] }),
          new Paragraph({ children: [new TextRun("first paragraph body")] }),
        ],
      },
    ],
  });
  const buf = await Packer.toBuffer(doc);
  const fixture = join(dir, "fixture.docx");
  writeFileSync(fixture, buf);
  return fixture;
}

describe("docx apply --room + --realtime-url", () => {
  let relay: Awaited<ReturnType<typeof startRelay>>;

  beforeAll(async () => {
    relay = await startRelay();
  });

  afterAll(async () => {
    await relay.close();
  });

  afterEach(async () => {
    // Clear all rooms between cases so envelopes from one test don't
    // bleed into the next subscriber.
    for (const [, room] of relay.rooms) {
      room.doc.transact(() => {
        const arr = room.doc.getArray<unknown>("commands");
        arr.delete(0, arr.length);
      });
    }
  });

  it("publishes each applied command onto the room's commands Y.Array", async () => {
    const fixture = await makeDocxFixture();
    const dir = mkdtempSync(join(tmpdir(), "office-agent-realtime-out-"));
    const cmdsPath = join(dir, "commands.json");
    writeFileSync(
      cmdsPath,
      JSON.stringify({
        commands: [
          {
            type: "docx:insert-text",
            payload: { at: { paragraph: 1, run: 0, offset: 0 }, text: "Live " },
          },
          {
            type: "docx:set-paragraph-style",
            payload: { at: { paragraph: 1 }, style: "Heading2" },
          },
        ],
      })
    );

    const roomId = `test-room-${Date.now()}`;
    const subDoc = new Y.Doc();
    const subProvider = new WebsocketProvider(relay.url, roomId, subDoc, {
      connect: true,
      WebSocketPolyfill: NodeWebSocket as unknown as typeof WebSocket,
    });
    await new Promise<void>((resolve) => subProvider.once("sync", () => resolve()));

    const subLog = subDoc.getArray<unknown>("commands");
    const seen: unknown[] = [];
    subLog.observe(() => {
      seen.length = 0;
      subLog.toArray().forEach((v) => seen.push(v));
    });

    const out = join(dir, "out.docx");
    const { io, stdout } = makeIO();
    const code = await runCli(
      [
        "docx",
        "apply",
        "--file",
        fixture,
        "--commands",
        cmdsPath,
        "--out",
        out,
        "--room",
        roomId,
        "--realtime-url",
        relay.url,
        "--agent-name",
        "Test Office Agent",
        "--agent-color",
        "#abcdef",
      ],
      io
    );
    expect(code).toBe(0);

    const parsed = JSON.parse(stdout.text());
    expect(parsed.realtime).toBeDefined();
    expect(parsed.realtime.published).toBe(2);
    expect(parsed.realtime.room).toBe(roomId);

    // Wait briefly for the subscriber's observer to fire.
    await new Promise((r) => setTimeout(r, 200));
    expect(seen).toHaveLength(2);
    const first = seen[0] as {
      peerId: string;
      seq: number;
      command: { type: string; payload: unknown; source?: string; agentId?: string };
    };
    expect(first.command.type).toBe("docx:insert-text");
    expect(first.command.source).toBe("agent");
    expect(first.command.agentId).toBe("office-agent-cli");
    expect(typeof first.peerId).toBe("string");
    expect(first.seq).toBe(1);

    subProvider.destroy();
    subDoc.destroy();
  }, 15000);

  it("--clear-room-after empties the Y.Array post-publish", async () => {
    const fixture = await makeDocxFixture();
    const dir = mkdtempSync(join(tmpdir(), "office-agent-realtime-clear-"));
    const cmdsPath = join(dir, "commands.json");
    writeFileSync(
      cmdsPath,
      JSON.stringify({
        commands: [
          {
            type: "docx:insert-text",
            payload: { at: { paragraph: 1, run: 0, offset: 0 }, text: "X " },
          },
        ],
      })
    );

    const roomId = `test-room-clear-${Date.now()}`;
    const out = join(dir, "out.docx");
    const { io, stdout } = makeIO();
    const code = await runCli(
      [
        "docx",
        "apply",
        "--file",
        fixture,
        "--commands",
        cmdsPath,
        "--out",
        out,
        "--room",
        roomId,
        "--realtime-url",
        relay.url,
        "--clear-room-after",
      ],
      io
    );
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text());
    expect(parsed.realtime.cleared).toBe(true);

    // Late-joining subscriber should see an empty room (the
    // canonical "fresh ops since last save" invariant).
    const lateDoc = new Y.Doc();
    const lateProvider = new WebsocketProvider(relay.url, roomId, lateDoc, {
      connect: true,
      WebSocketPolyfill: NodeWebSocket as unknown as typeof WebSocket,
    });
    await new Promise<void>((resolve) => lateProvider.once("sync", () => resolve()));
    await new Promise((r) => setTimeout(r, 100));
    const lateLog = lateDoc.getArray<unknown>("commands");
    expect(lateLog.toArray()).toHaveLength(0);

    lateProvider.destroy();
    lateDoc.destroy();

    // Sanity: making sure the file write still happened.
    const bytes = readFileSync(out);
    expect(bytes.length).toBeGreaterThan(0);
  }, 15000);

  it("absent --room/--realtime-url is a no-op (legacy behaviour preserved)", async () => {
    const fixture = await makeDocxFixture();
    const dir = mkdtempSync(join(tmpdir(), "office-agent-realtime-noop-"));
    const cmdsPath = join(dir, "commands.json");
    writeFileSync(
      cmdsPath,
      JSON.stringify({
        commands: [
          {
            type: "docx:insert-text",
            payload: { at: { paragraph: 1, run: 0, offset: 0 }, text: "Y " },
          },
        ],
      })
    );
    const out = join(dir, "out.docx");
    const { io, stdout } = makeIO();
    const code = await runCli(["docx", "apply", "--file", fixture, "--commands", cmdsPath, "--out", out], io);
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text());
    expect(parsed.realtime).toBeUndefined();
    expect(parsed.wrote).toBe(out);
  });
});
