import { mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { Document, HeadingLevel, Packer, Paragraph, TextRun } from "docx";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { createMcpServer, __resetMcpSessionsForTests } from "./mcp.js";

async function makeFixture(): Promise<string> {
  const dir = mkdtempSync(join(tmpdir(), "mcp-fixture-"));
  const doc = new Document({
    creator: "test",
    sections: [
      {
        properties: {},
        children: [
          new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Hello")] }),
          new Paragraph({ children: [new TextRun("first paragraph body")] }),
          new Paragraph({ children: [new TextRun("second paragraph body")] }),
        ],
      },
    ],
  });
  const buf = await Packer.toBuffer(doc);
  const inputPath = join(dir, "in.docx");
  writeFileSync(inputPath, buf);
  return inputPath;
}

async function makeClient(): Promise<Client> {
  const server = createMcpServer();
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  const client = new Client({ name: "officeai-test", version: "0.0.0" }, { capabilities: {} });
  await Promise.all([client.connect(clientTransport), server.connect(serverTransport)]);
  return client;
}

function structured(result: unknown): Record<string, unknown> {
  if (!result || typeof result !== "object") throw new Error("expected structured tool result");
  const r = result as { structuredContent?: unknown; content?: unknown; isError?: boolean };
  if (r.isError) {
    throw new Error(`tool returned error: ${JSON.stringify(r.content)}`);
  }
  if (r.structuredContent && typeof r.structuredContent === "object") {
    return r.structuredContent as Record<string, unknown>;
  }
  if (Array.isArray(r.content) && r.content.length > 0) {
    const first = r.content[0] as { text?: string };
    if (first.text) return JSON.parse(first.text);
  }
  throw new Error("tool result had neither structuredContent nor parseable text");
}

describe("OfficeAI MCP server", () => {
  beforeEach(() => __resetMcpSessionsForTests());
  afterEach(() => __resetMcpSessionsForTests());

  it("lists every registered docx_* and pptx_* tool", async () => {
    const client = await makeClient();
    const list = await client.listTools();
    const names = list.tools.map((t) => t.name).sort();
    expect(names).toEqual([
      "docx_apply_command",
      "docx_diff",
      "docx_get_text",
      "docx_inspect",
      "docx_load",
      "docx_save",
      "docx_search",
      "pptx_apply_command",
      "pptx_diff",
      "pptx_get_text",
      "pptx_inspect",
      "pptx_load",
      "pptx_save",
      "pptx_search",
    ]);
  });

  it("docx_load returns a handle and an inspection summary", async () => {
    const client = await makeClient();
    const path = await makeFixture();
    const out = structured(await client.callTool({ name: "docx_load", arguments: { path } }));
    expect(typeof out.handle).toBe("string");
    const summary = out.summary as { paragraphs: number; tables: number; comments: number };
    expect(summary.paragraphs).toBe(3);
    expect(summary.comments).toBe(0);
    expect(summary.tables).toBe(0);
  });

  it("docx_inspect returns counts and parts list", async () => {
    const client = await makeClient();
    const path = await makeFixture();
    const loaded = structured(await client.callTool({ name: "docx_load", arguments: { path } }));
    const handle = loaded.handle as string;
    const insp = structured(await client.callTool({ name: "docx_inspect", arguments: { handle } }));
    expect(insp.format).toBe("docx");
    expect(insp.paragraphs).toBe(3);
    expect(Array.isArray(insp.parts)).toBe(true);
    expect((insp.parts as string[]).some((p) => p.includes("document.xml"))).toBe(true);
  });

  it("docx_get_text supports markdown, json, and text formats", async () => {
    const client = await makeClient();
    const path = await makeFixture();
    const loaded = structured(await client.callTool({ name: "docx_load", arguments: { path } }));
    const handle = loaded.handle as string;

    const md = structured(
      await client.callTool({ name: "docx_get_text", arguments: { handle, format: "markdown" } })
    );
    expect(typeof md.content).toBe("string");
    expect(md.content as string).toContain("# Hello");

    const json = structured(
      await client.callTool({ name: "docx_get_text", arguments: { handle, format: "json" } })
    );
    expect(json.format).toBe("docx");
    expect(Array.isArray(json.paragraphs)).toBe(true);
    expect((json.paragraphs as Array<{ text: string }>).length).toBe(3);

    const text = structured(
      await client.callTool({ name: "docx_get_text", arguments: { handle, format: "text" } })
    );
    expect(text.content as string).toContain("first paragraph body");
    expect(text.content as string).not.toContain("#");
  });

  it("docx_search returns matches", async () => {
    const client = await makeClient();
    const path = await makeFixture();
    const loaded = structured(await client.callTool({ name: "docx_load", arguments: { path } }));
    const handle = loaded.handle as string;
    const r = structured(
      await client.callTool({
        name: "docx_search",
        arguments: { handle, query: "paragraph" },
      })
    );
    const matches = r.matches as Array<{ paragraphIndex: number; match: string }>;
    expect(matches.length).toBeGreaterThanOrEqual(2);
    expect(matches[0].match.toLowerCase()).toBe("paragraph");
  });

  it("docx_apply_command can insert text and docx_save persists it", async () => {
    const client = await makeClient();
    const path = await makeFixture();
    const loaded = structured(await client.callTool({ name: "docx_load", arguments: { path } }));
    const handle = loaded.handle as string;

    const apply = structured(
      await client.callTool({
        name: "docx_apply_command",
        arguments: {
          handle,
          type: "docx:insert-text",
          payload: { at: { paragraph: 1, run: 0, offset: 0 }, text: "MCP " },
        },
      })
    );
    expect((apply.mutation as { status: string }).status).toBe("approved");

    const dir = mkdtempSync(join(tmpdir(), "mcp-out-"));
    const out = join(dir, "out.docx");
    const saved = structured(
      await client.callTool({ name: "docx_save", arguments: { handle, out_path: out } })
    );
    expect(saved.wrote).toBe(out);
    const reloaded = readFileSync(out);
    expect(reloaded.byteLength).toBeGreaterThan(0);
  });

  it("end-to-end comment lifecycle via docx_apply_command", async () => {
    const client = await makeClient();
    const path = await makeFixture();
    const loaded = structured(await client.callTool({ name: "docx_load", arguments: { path } }));
    const handle = loaded.handle as string;

    await client.callTool({
      name: "docx_apply_command",
      arguments: {
        handle,
        type: "docx:add-comment",
        payload: {
          range: {
            start: { paragraph: 1, run: 0, offset: 0 },
            end: { paragraph: 1, run: 0, offset: 5 },
          },
          text: "rephrase?",
          author: "AI",
          initials: "AI",
        },
      },
    });

    const inspectAfterAdd = structured(
      await client.callTool({ name: "docx_inspect", arguments: { handle } })
    );
    expect(inspectAfterAdd.comments).toBe(1);

    const projection = structured(
      await client.callTool({ name: "docx_get_text", arguments: { handle, format: "json" } })
    );
    const comments = projection.comments as Array<{ id: string; resolved?: boolean; parentId?: string }>;
    expect(comments).toHaveLength(1);
    const commentId = comments[0].id;

    await client.callTool({
      name: "docx_apply_command",
      arguments: {
        handle,
        type: "docx:resolve-comment",
        payload: { commentId },
      },
    });
    const afterResolve = structured(
      await client.callTool({ name: "docx_get_text", arguments: { handle, format: "json" } })
    );
    expect((afterResolve.comments as Array<{ resolved?: boolean }>)[0].resolved).toBe(true);

    await client.callTool({
      name: "docx_apply_command",
      arguments: {
        handle,
        type: "docx:reply-comment",
        payload: { parentId: commentId, text: "ack", author: "Bob" },
      },
    });
    const afterReply = structured(
      await client.callTool({ name: "docx_get_text", arguments: { handle, format: "json" } })
    );
    const replyComments = afterReply.comments as Array<{ id: string; parentId?: string }>;
    expect(replyComments).toHaveLength(2);
    expect(replyComments.find((c) => c.id !== commentId)?.parentId).toBe(commentId);

    await client.callTool({
      name: "docx_apply_command",
      arguments: {
        handle,
        type: "docx:delete-comment",
        payload: { commentId },
      },
    });
    const afterDelete = structured(await client.callTool({ name: "docx_inspect", arguments: { handle } }));
    expect(afterDelete.comments).toBe(0);
  });

  it("docx_diff reports changes between two handles", async () => {
    const client = await makeClient();
    const path = await makeFixture();
    const a = structured(await client.callTool({ name: "docx_load", arguments: { path } }));
    const b = structured(await client.callTool({ name: "docx_load", arguments: { path } }));

    await client.callTool({
      name: "docx_apply_command",
      arguments: {
        handle: b.handle,
        type: "docx:insert-text",
        payload: { at: { paragraph: 1, run: 0, offset: 0 }, text: "DIFF " },
      },
    });

    const diff = structured(
      await client.callTool({
        name: "docx_diff",
        arguments: { before: a.handle, after: b.handle },
      })
    );
    const paragraphs = diff.paragraphs as { added: number; removed: number; modified: number };
    expect(paragraphs.modified).toBe(1);
  });

  it("returns an error for unknown handles", async () => {
    const client = await makeClient();
    const r = await client.callTool({
      name: "docx_inspect",
      arguments: { handle: "nope" },
    });
    expect(r.isError).toBe(true);
  });
});

describe("OfficeAI MCP server — PPTX tools", () => {
  beforeEach(() => __resetMcpSessionsForTests());
  afterEach(() => __resetMcpSessionsForTests());

  const PPTX_FIXTURE = resolve(
    fileURLToPath(new URL("../../../fixtures/pptx/synthetic/03-title-and-content.pptx", import.meta.url))
  );

  it("pptx_load returns a handle and an inspection summary", async () => {
    const client = await makeClient();
    const out = structured(
      await client.callTool({ name: "pptx_load", arguments: { path: PPTX_FIXTURE } })
    );
    expect(typeof out.handle).toBe("string");
    const summary = out.summary as {
      slides: number;
      shapeCounts: { text: number };
      slideSize: { cxEmu: number; cyEmu: number };
    };
    expect(summary.slides).toBe(1);
    expect(summary.shapeCounts.text).toBeGreaterThanOrEqual(1);
    expect(summary.slideSize.cxEmu).toBeGreaterThan(0);
  });

  it("pptx_get_text supports markdown / json / text", async () => {
    const client = await makeClient();
    const loaded = structured(
      await client.callTool({ name: "pptx_load", arguments: { path: PPTX_FIXTURE } })
    );
    const handle = loaded.handle as string;

    const md = structured(
      await client.callTool({ name: "pptx_get_text", arguments: { handle, format: "markdown" } })
    );
    expect((md.content as string).startsWith("# Presentation")).toBe(true);

    const json = structured(
      await client.callTool({ name: "pptx_get_text", arguments: { handle, format: "json" } })
    );
    expect(json.format).toBe("pptx");
    expect(Array.isArray(json.slides)).toBe(true);

    const text = structured(
      await client.callTool({ name: "pptx_get_text", arguments: { handle, format: "text" } })
    );
    expect(typeof text.content).toBe("string");
  });

  it("pptx_apply_command + pptx_save round-trips an edit through disk", async () => {
    const client = await makeClient();
    const loaded = structured(
      await client.callTool({ name: "pptx_load", arguments: { path: PPTX_FIXTURE } })
    );
    const handle = loaded.handle as string;

    const apply = structured(
      await client.callTool({
        name: "pptx_apply_command",
        arguments: { handle, type: "pptx:add-slide", payload: {} },
      })
    );
    expect((apply.mutation as { status: string }).status).toBe("approved");

    const dir = mkdtempSync(join(tmpdir(), "mcp-pptx-out-"));
    const out = join(dir, "out.pptx");
    const saved = structured(
      await client.callTool({ name: "pptx_save", arguments: { handle, out_path: out } })
    );
    expect(saved.wrote).toBe(out);
    expect(readFileSync(out).byteLength).toBeGreaterThan(0);
  });

  it("pptx_diff reports slide-level changes between two handles", async () => {
    const client = await makeClient();
    const a = structured(await client.callTool({ name: "pptx_load", arguments: { path: PPTX_FIXTURE } }));
    const b = structured(await client.callTool({ name: "pptx_load", arguments: { path: PPTX_FIXTURE } }));

    await client.callTool({
      name: "pptx_apply_command",
      arguments: { handle: b.handle, type: "pptx:add-slide", payload: {} },
    });

    const diff = structured(
      await client.callTool({ name: "pptx_diff", arguments: { before: a.handle, after: b.handle } })
    );
    expect((diff.slides as { added: number }).added).toBe(1);
  });

  it("pptx_search returns text matches", async () => {
    const client = await makeClient();
    const loaded = structured(
      await client.callTool({ name: "pptx_load", arguments: { path: PPTX_FIXTURE } })
    );
    const r = structured(
      await client.callTool({
        name: "pptx_search",
        arguments: { handle: loaded.handle, query: "Title" },
      })
    );
    expect(Array.isArray(r.matches)).toBe(true);
  });

  it("returns an error for unknown pptx handles", async () => {
    const client = await makeClient();
    const r = await client.callTool({ name: "pptx_inspect", arguments: { handle: "nope" } });
    expect(r.isError).toBe(true);
  });
});
