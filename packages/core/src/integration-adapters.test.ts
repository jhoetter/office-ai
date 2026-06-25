import { describe, expect, it } from "vitest";
import { createLocalIntegrationAdapters, createOfficeAiEvent } from "./integration-adapters.js";

describe("integration adapter boundary", () => {
  it("provides a standalone local reference adapter set", async () => {
    const adapters = createLocalIntegrationAdapters({
      actor: { id: "user-1", displayName: "Ada" },
      webBaseUrl: "http://localhost:3100",
      mcpCommand: "pnpm",
      mcpArgs: ["office-agent", "mcp"],
    });

    const path = adapters.storage.join(adapters.storage.root, "sessions", "s1", "meta.json");
    await adapters.storage.writeBytesAtomic(path, new TextEncoder().encode('{"ok":true}'));

    expect(adapters.storage.capabilities).toEqual({
      atomicWrite: true,
      localPaths: false,
      locks: "none",
      watch: false,
    });
    expect(await adapters.storage.exists(path)).toBe(true);
    expect(
      await adapters.storage.list(adapters.storage.join(adapters.storage.root, "sessions", "s1"))
    ).toEqual(["meta.json"]);
    expect(new TextDecoder().decode(await adapters.storage.readBytes(path))).toBe('{"ok":true}');

    const imported = await adapters.assets.importAsset({
      format: "docx",
      bytes: new Uint8Array([1, 2, 3]),
      source: { kind: "generated", label: "unit-test" },
    });
    expect(imported.mediaType).toBe(
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    expect([...new Uint8Array(await adapters.assets.readAssetBytes(imported.id))]).toEqual([1, 2, 3]);

    const actor = await adapters.identity.getCurrentActor();
    await adapters.events.emit(createOfficeAiEvent("document.exported", { actor: actor ?? undefined }));
    expect(adapters.events.events).toHaveLength(1);
    expect(adapters.events.events[0].actor?.id).toBe("user-1");

    await expect(adapters.ui.getEmbedding({ format: "docx", documentId: "doc-1" })).resolves.toMatchObject({
      mode: "web-url",
      url: "http://localhost:3100/sessions/doc-1",
    });
    await expect(adapters.mcp.describeHost()).resolves.toMatchObject({
      transport: "stdio",
      command: "pnpm",
      args: ["office-agent", "mcp"],
    });
  });
});
