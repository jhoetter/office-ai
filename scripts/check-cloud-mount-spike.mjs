#!/usr/bin/env node
import { createServer } from "node:http";
import { readFileSync } from "node:fs";
import { relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import {
  loadCloudObjectBytes,
  saveCloudObjectBytes,
} from "../examples/sonaloop-cloud-mount-spike/host/src/cloud-object-store.mjs";

const ROOT = resolve(fileURLToPath(import.meta.url), "..", "..");
const FIXTURE_PATH = resolve(ROOT, "fixtures/xlsx/synthetic/01-single-sheet-numbers.xlsx");
const APP_PATH = resolve(ROOT, "examples/sonaloop-cloud-mount-spike/host/src/App.tsx");

function rel(path) {
  return relative(ROOT, path).replace(/\\/g, "/");
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

function readBody(req) {
  return new Promise((resolveRead, reject) => {
    const chunks = [];
    req.on("data", (chunk) => chunks.push(Buffer.from(chunk)));
    req.on("error", reject);
    req.on("end", () => resolveRead(Buffer.concat(chunks)));
  });
}

function startPresignServer(initialBytes) {
  const objectKey = "workspaces/demo/assets/revenue-model.xlsx";
  let stored = Buffer.from(initialBytes);
  let etag = '"rev-1"';
  let writes = 0;

  const server = createServer(async (req, res) => {
    try {
      const url = new URL(req.url ?? "/", "http://127.0.0.1");

      if (req.method === "GET" && url.pathname === "/api/office-ai/presign-get") {
        assert(url.searchParams.get("key") === objectKey, "GET presign received unexpected object key");
        res.setHeader("content-type", "application/json");
        res.end(
          JSON.stringify({
            url: `${baseUrl}/objects/${encodeURIComponent(objectKey)}?signature=get`,
            headers: { "x-office-ai-demo-read": "1" },
            etag,
            filename: "revenue-model.xlsx",
          })
        );
        return;
      }

      if (req.method === "POST" && url.pathname === "/api/office-ai/presign-put") {
        const body = JSON.parse((await readBody(req)).toString("utf8"));
        assert(body.key === objectKey, "PUT presign received unexpected object key");
        assert(
          body.contentType === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          "PUT presign received unexpected content type"
        );
        assert(body.ifMatch === etag, "PUT presign did not receive the current etag");
        res.setHeader("content-type", "application/json");
        res.end(
          JSON.stringify({
            url: `${baseUrl}/objects/${encodeURIComponent(objectKey)}?signature=put`,
            headers: { "x-office-ai-demo-write": "1" },
          })
        );
        return;
      }

      if (req.method === "GET" && url.pathname.startsWith("/objects/")) {
        assert(url.searchParams.get("signature") === "get", "object GET missed presigned marker");
        assert(req.headers["x-office-ai-demo-read"] === "1", "object GET missed presigned header");
        res.setHeader("content-type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        res.setHeader("etag", etag);
        res.end(stored);
        return;
      }

      if (req.method === "PUT" && url.pathname.startsWith("/objects/")) {
        assert(url.searchParams.get("signature") === "put", "object PUT missed presigned marker");
        assert(req.headers["x-office-ai-demo-write"] === "1", "object PUT missed presigned header");
        stored = await readBody(req);
        writes += 1;
        etag = `"rev-${writes + 1}"`;
        res.statusCode = 200;
        res.setHeader("etag", etag);
        res.end("ok");
        return;
      }

      res.statusCode = 404;
      res.end("not found");
    } catch (err) {
      res.statusCode = 500;
      res.end(err instanceof Error ? err.message : String(err));
    }
  });

  let baseUrl = "";

  return new Promise((resolveStart, reject) => {
    server.on("error", reject);
    server.listen(0, "127.0.0.1", () => {
      const address = server.address();
      assert(address && typeof address === "object", "server did not bind to a TCP port");
      baseUrl = `http://127.0.0.1:${address.port}`;
      resolveStart({
        baseUrl,
        objectKey,
        close: () =>
          new Promise((resolveClose, rejectClose) =>
            server.close((err) => (err ? rejectClose(err) : resolveClose()))
          ),
        getStored: () => stored,
        getWrites: () => writes,
      });
    });
  });
}

async function main() {
  const fixture = readFileSync(FIXTURE_PATH);
  const appSource = readFileSync(APP_PATH, "utf8");

  assert(!/<iframe\b/i.test(appSource), `${rel(APP_PATH)} must not use an iframe`);
  assert(
    appSource.includes("@officeai/react-editors/components/xlsx"),
    `${rel(APP_PATH)} must mount the XLSX editor package entrypoint`
  );
  assert(
    appSource.includes("initialBytes={cloudDocument.bytes}"),
    `${rel(APP_PATH)} must pass host-loaded bytes`
  );
  assert(
    appSource.includes("onSave={handleSave}"),
    `${rel(APP_PATH)} must route Save through the host callback`
  );
  assert(
    appSource.includes("room={null}"),
    `${rel(APP_PATH)} must explicitly avoid implicit realtime rooms in the spike`
  );
  assert(
    appSource.includes("hideLocalFileOpen"),
    `${rel(APP_PATH)} must hide local file-open affordances in cloud mode`
  );

  const server = await startPresignServer(fixture);
  try {
    const endpoints = {
      presignGetUrl: `${server.baseUrl}/api/office-ai/presign-get`,
      presignPutUrl: `${server.baseUrl}/api/office-ai/presign-put`,
    };
    const loaded = await loadCloudObjectBytes({ ...endpoints, objectKey: server.objectKey });
    assert(
      Buffer.compare(Buffer.from(loaded.bytes), fixture) === 0,
      "presigned GET did not return the fixture bytes"
    );
    assert(loaded.etag === '"rev-1"', "presigned GET did not propagate the object revision");
    assert(loaded.filename === "revenue-model.xlsx", "presigned GET did not propagate the object filename");

    const saved = Buffer.concat([fixture, Buffer.from([0x0a])]);
    const result = await saveCloudObjectBytes({
      ...endpoints,
      objectKey: server.objectKey,
      bytes: new Uint8Array(saved),
      mime: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      filename: "revenue-model.xlsx",
      etag: loaded.etag,
    });

    assert(server.getWrites() === 1, "presigned PUT was not called exactly once");
    assert(Buffer.compare(server.getStored(), saved) === 0, "presigned PUT did not persist the edited bytes");
    assert(result.etag === '"rev-2"', "presigned PUT did not return the next object revision");
  } finally {
    await server.close();
  }

  console.log("cloud-mount-spike: OK");
  console.log(`  fixture: ${rel(FIXTURE_PATH)}`);
  console.log(`  host: ${rel(APP_PATH)}`);
}

main().catch((err) => {
  console.error("cloud-mount-spike: FAILED");
  console.error(err instanceof Error ? (err.stack ?? err.message) : String(err));
  process.exit(1);
});
