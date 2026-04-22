/**
 * Self-hosted Yjs websocket server.
 *
 * One in-process map of `roomId → { doc, sockets }`. Speaks the
 * standard Yjs sync + awareness protocol so any `y-websocket`
 * client (browser or node) can connect.
 *
 * Boot via `pnpm --filter @officeai/realtime-server dev` (port
 * defaults to 1234, override with `OAI_RT_PORT`). The Next.js dev
 * host expects `ws://localhost:1234` — see `apps/web/app/lib/realtime/`.
 *
 * Designed for localhost demos and small teams. Persistence is
 * intentionally absent: when the last peer leaves a room the doc
 * stays warm in memory until the process exits, then resets. The
 * real persistence path is the OOXML file the user explicitly saves.
 */
import http from "node:http";
import { WebSocketServer, type WebSocket } from "ws";
import * as Y from "yjs";
import * as syncProtocol from "y-protocols/sync";
import * as awarenessProtocol from "y-protocols/awareness";
import { encoding, decoding } from "lib0";

const PORT = Number(process.env.OAI_RT_PORT ?? 1234);
const HOST = process.env.OAI_RT_HOST ?? "localhost";

/** Yjs sync protocol message types (see y-protocols README). */
const MESSAGE_SYNC = 0;
const MESSAGE_AWARENESS = 1;

interface Room {
  readonly id: string;
  readonly doc: Y.Doc;
  readonly awareness: awarenessProtocol.Awareness;
  readonly sockets: Set<WebSocket>;
}

const rooms = new Map<string, Room>();

function getRoom(id: string): Room {
  let room = rooms.get(id);
  if (!room) {
    const doc = new Y.Doc();
    const awareness = new awarenessProtocol.Awareness(doc);
    awareness.setLocalState(null);
    room = { id, doc, awareness, sockets: new Set() };

    doc.on("update", (update: Uint8Array, origin: unknown) => {
      const enc = encoding.createEncoder();
      encoding.writeVarUint(enc, MESSAGE_SYNC);
      syncProtocol.writeUpdate(enc, update);
      const buf = encoding.toUint8Array(enc);
      for (const ws of room!.sockets) {
        if (ws !== origin && ws.readyState === ws.OPEN) {
          ws.send(buf);
        }
      }
    });

    awareness.on(
      "update",
      (
        { added, updated, removed }: { added: number[]; updated: number[]; removed: number[] },
        origin: unknown
      ) => {
        const changedClients = added.concat(updated, removed);
        const enc = encoding.createEncoder();
        encoding.writeVarUint(enc, MESSAGE_AWARENESS);
        encoding.writeVarUint8Array(enc, awarenessProtocol.encodeAwarenessUpdate(awareness, changedClients));
        const buf = encoding.toUint8Array(enc);
        for (const ws of room!.sockets) {
          if (ws !== origin && ws.readyState === ws.OPEN) {
            ws.send(buf);
          }
        }
      }
    );

    rooms.set(id, room);
  }
  return room;
}

function handleMessage(room: Room, ws: WebSocket, message: Uint8Array): void {
  const dec = decoding.createDecoder(message);
  const messageType = decoding.readVarUint(dec);
  switch (messageType) {
    case MESSAGE_SYNC: {
      const enc = encoding.createEncoder();
      encoding.writeVarUint(enc, MESSAGE_SYNC);
      // `readSyncMessage` writes any reply it produced into `enc`.
      syncProtocol.readSyncMessage(dec, enc, room.doc, ws);
      // Only send if there's actually a reply (length > 1 because
      // we already wrote the messageType byte).
      if (encoding.length(enc) > 1) {
        ws.send(encoding.toUint8Array(enc));
      }
      break;
    }
    case MESSAGE_AWARENESS: {
      awarenessProtocol.applyAwarenessUpdate(room.awareness, decoding.readVarUint8Array(dec), ws);
      break;
    }
    default:
      // Unknown message type — ignore.
      break;
  }
}

function setupConnection(ws: WebSocket, request: http.IncomingMessage): void {
  ws.binaryType = "arraybuffer";
  const url = new URL(request.url ?? "/", `http://${request.headers.host ?? "localhost"}`);
  // Path is treated as the room id (matches y-websocket convention).
  const roomId = decodeURIComponent(url.pathname.replace(/^\//, "")) || "default";
  const room = getRoom(roomId);
  room.sockets.add(ws);

  // Initial sync: send our state vector and any local awareness state.
  {
    const enc = encoding.createEncoder();
    encoding.writeVarUint(enc, MESSAGE_SYNC);
    syncProtocol.writeSyncStep1(enc, room.doc);
    ws.send(encoding.toUint8Array(enc));

    const awarenessStates = room.awareness.getStates();
    if (awarenessStates.size > 0) {
      const enc2 = encoding.createEncoder();
      encoding.writeVarUint(enc2, MESSAGE_AWARENESS);
      encoding.writeVarUint8Array(
        enc2,
        awarenessProtocol.encodeAwarenessUpdate(room.awareness, Array.from(awarenessStates.keys()))
      );
      ws.send(encoding.toUint8Array(enc2));
    }
  }

  ws.on("message", (data: ArrayBuffer | Buffer) => {
    const bytes = data instanceof ArrayBuffer ? new Uint8Array(data) : new Uint8Array(data);
    try {
      handleMessage(room, ws, bytes);
    } catch (err) {
      console.error(`[realtime] message error in room ${roomId}:`, err);
    }
  });

  ws.on("close", () => {
    room.sockets.delete(ws);
    awarenessProtocol.removeAwarenessStates(room.awareness, [room.doc.clientID], null);
    // Don't garbage-collect the doc — keep it warm so a refreshing
    // tab finds the same state. A simple LRU eviction can come later.
  });

  ws.on("error", (err: unknown) => {
    console.error(`[realtime] socket error in room ${roomId}:`, err);
  });
}

const httpServer = http.createServer((req, res) => {
  if (req.url === "/health") {
    res.writeHead(200, { "Content-Type": "application/json" });
    res.end(
      JSON.stringify({
        ok: true,
        rooms: rooms.size,
        clients: Array.from(rooms.values()).reduce((n, r) => n + r.sockets.size, 0),
      })
    );
    return;
  }
  res.writeHead(200, { "Content-Type": "text/plain" });
  res.end("office-ai realtime server\n");
});

const wss = new WebSocketServer({ server: httpServer });
wss.on("connection", setupConnection);

httpServer.listen(PORT, HOST, () => {
  console.log(`[realtime] listening on ws://${HOST}:${PORT}`);
});

const shutdown = (): void => {
  console.log("[realtime] shutting down");
  for (const ws of wss.clients) {
    try {
      ws.close();
    } catch {
      /* noop */
    }
  }
  httpServer.close(() => process.exit(0));
};
process.on("SIGINT", shutdown);
process.on("SIGTERM", shutdown);
