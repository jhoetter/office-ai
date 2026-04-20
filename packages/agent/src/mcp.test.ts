import { mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve, resolve as resolvePath } from "node:path";
import { fileURLToPath } from "node:url";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { Document, HeadingLevel, Packer, Paragraph, TextRun } from "docx";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { XlsxAgent } from "@officeai/xlsx";
import { createMcpServer, __resetMcpSessionsForTests } from "./mcp.js";

const here = dirname(fileURLToPath(import.meta.url));
const xlsxFixtures = resolvePath(here, "../../../fixtures/xlsx/synthetic");

function copyXlsxFixture(name: string, dest: string): string {
  const target = join(dest, name);
  writeFileSync(target, readFileSync(resolvePath(xlsxFixtures, name)));
  return target;
}

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

  it("lists every registered docx_*, xlsx_*, pptx_* and pdf_* tool", async () => {
    const client = await makeClient();
    const list = await client.listTools();
    const names = list.tools.map((t) => t.name).sort();
    expect(names).toEqual([
      "docx_apply_command",
      "docx_approve",
      "docx_diff",
      "docx_get_page_text",
      "docx_get_pages",
      "docx_get_text",
      "docx_inspect",
      "docx_list_pending",
      "docx_load",
      "docx_reject",
      "docx_save",
      "docx_search",
      "pdf_add_page_numbers",
      "pdf_add_watermark",
      "pdf_annotations",
      "pdf_delete_pages",
      "pdf_export_markdown",
      "pdf_extract_pages",
      "pdf_fill_form",
      "pdf_flatten_form",
      "pdf_list_form_fields",
      "pdf_load",
      "pdf_merge",
      "pdf_metadata",
      "pdf_outline",
      "pdf_read_page",
      "pdf_reorder_pages",
      "pdf_reset_form",
      "pdf_rotate_pages",
      "pdf_search",
      "pdf_set_metadata",
      "pptx_apply_command",
      "pptx_diff",
      "pptx_get_text",
      "pptx_inspect",
      "pptx_load",
      "pptx_save",
      "pptx_search",
      "xlsx_add_comment",
      "xlsx_add_sheet",
      "xlsx_apply_command",
      "xlsx_approve",
      "xlsx_delete_column",
      "xlsx_delete_row",
      "xlsx_diff",
      "xlsx_get_text",
      "xlsx_insert_column",
      "xlsx_insert_row",
      "xlsx_inspect",
      "xlsx_list_pending",
      "xlsx_list_sheets",
      "xlsx_load",
      "xlsx_merge",
      "xlsx_redo",
      "xlsx_reject",
      "xlsx_rename_sheet",
      "xlsx_save",
      "xlsx_search",
      "xlsx_set_cell",
      "xlsx_set_format",
      "xlsx_set_formula",
      "xlsx_set_range",
      "xlsx_undo",
      "xlsx_unmerge",
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

  it("docx_get_pages returns at least one doc-start page for a single-page doc", async () => {
    const client = await makeClient();
    const path = await makeFixture();
    const loaded = structured(await client.callTool({ name: "docx_load", arguments: { path } }));
    const handle = loaded.handle as string;
    const out = structured(await client.callTool({ name: "docx_get_pages", arguments: { handle } }));
    const pages = out.pages as Array<{
      pageNumber: number;
      trigger: string;
      preview: string;
    }>;
    expect(pages.length).toBeGreaterThanOrEqual(1);
    expect(pages[0].pageNumber).toBe(1);
    expect(pages[0].trigger).toBe("doc-start");
    expect(out.total).toBe(pages.length);
  });

  it("docx_get_page_text returns markdown for an in-range page and errors for out-of-range", async () => {
    const client = await makeClient();
    const path = await makeFixture();
    const loaded = structured(await client.callTool({ name: "docx_load", arguments: { path } }));
    const handle = loaded.handle as string;
    const md = structured(
      await client.callTool({ name: "docx_get_page_text", arguments: { handle, page: 1 } })
    );
    expect(md.format).toBe("markdown");
    expect(typeof md.content).toBe("string");
    expect(md.content as string).toContain("# Hello");

    const errResult = (await client.callTool({
      name: "docx_get_page_text",
      arguments: { handle, page: 999 },
    })) as { isError?: boolean; content?: Array<{ text: string }> };
    expect(errResult.isError).toBe(true);
    const text = errResult.content?.[0]?.text ?? "";
    expect(text).toContain("out-of-range");
  });

  it("docx_get_text with with_page_sections injects page anchors", async () => {
    const client = await makeClient();
    const path = await makeFixture();
    const loaded = structured(await client.callTool({ name: "docx_load", arguments: { path } }));
    const handle = loaded.handle as string;
    const md = structured(
      await client.callTool({
        name: "docx_get_text",
        arguments: { handle, format: "markdown", with_page_sections: true },
      })
    );
    expect(md.content as string).toContain("<!-- page 1 -->");
    expect(md.content as string).toContain("## Page 1");
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

  it("pending review flow: apply auto_approve=false, list, then approve & reject", async () => {
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
          payload: { at: { paragraph: 1, run: 0, offset: 0 }, text: "PENDING " },
          auto_approve: false,
        },
      })
    );
    expect((apply.mutation as { status: string }).status).toBe("pending");
    const mutationId = (apply.mutation as { id: string }).id;

    const list1 = structured(await client.callTool({ name: "docx_list_pending", arguments: { handle } }));
    expect(list1.pending as Array<{ id: string }>).toHaveLength(1);
    expect((list1.pending as Array<{ id: string }>)[0].id).toBe(mutationId);

    const approve = structured(
      await client.callTool({
        name: "docx_approve",
        arguments: { handle, mutation_id: mutationId },
      })
    );
    expect(approve.approved).toBe(mutationId);

    const list2 = structured(await client.callTool({ name: "docx_list_pending", arguments: { handle } }));
    expect(list2.pending as unknown[]).toHaveLength(0);

    const apply2 = structured(
      await client.callTool({
        name: "docx_apply_command",
        arguments: {
          handle,
          type: "docx:insert-text",
          payload: { at: { paragraph: 1, run: 0, offset: 0 }, text: "REJECTME " },
          auto_approve: false,
        },
      })
    );
    const id2 = (apply2.mutation as { id: string }).id;
    const reject = structured(
      await client.callTool({
        name: "docx_reject",
        arguments: { handle, mutation_id: id2, reason: "no thanks" },
      })
    );
    expect(reject.rejected).toBe(id2);
    expect(reject.reason).toBe("no thanks");

    const list3 = structured(await client.callTool({ name: "docx_list_pending", arguments: { handle } }));
    expect(list3.pending as unknown[]).toHaveLength(0);
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
  const PPTX_TABLE_FIXTURE = resolve(
    fileURLToPath(new URL("../../../fixtures/pptx/synthetic/06-with-table.pptx", import.meta.url))
  );
  const PPTX_CHART_FIXTURE = resolve(
    fileURLToPath(new URL("../../../fixtures/pptx/synthetic/09-with-chart.pptx", import.meta.url))
  );
  const PPTX_ANIM_FIXTURE = resolve(
    fileURLToPath(new URL("../../../fixtures/pptx/synthetic/10-with-anim.pptx", import.meta.url))
  );

  it("pptx_load returns a handle and an inspection summary", async () => {
    const client = await makeClient();
    const out = structured(await client.callTool({ name: "pptx_load", arguments: { path: PPTX_FIXTURE } }));
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

  it("pptx_apply_command dispatches typed table commands", async () => {
    const client = await makeClient();
    const loaded = structured(
      await client.callTool({ name: "pptx_load", arguments: { path: PPTX_TABLE_FIXTURE } })
    );
    const handle = loaded.handle as string;

    const json = structured(
      await client.callTool({ name: "pptx_get_text", arguments: { handle, format: "json" } })
    ) as unknown as {
      slides: Array<{ index: number; shapes: Array<{ id: string; kind: string }> }>;
    };
    let slideIndex = -1;
    let tableId = "";
    for (const slide of json.slides) {
      const t = slide.shapes.find((s) => s.kind === "table");
      if (t) {
        slideIndex = slide.index;
        tableId = t.id;
        break;
      }
    }
    expect(tableId).not.toBe("");

    const apply = structured(
      await client.callTool({
        name: "pptx_apply_command",
        arguments: {
          handle,
          type: "pptx:table-set-cell-text",
          payload: { slideIndex, shapeId: tableId, row: 0, column: 0, text: "MCP Cell" },
        },
      })
    );
    expect((apply.mutation as { status: string }).status).toBe("approved");

    const after = structured(
      await client.callTool({ name: "pptx_get_text", arguments: { handle, format: "text" } })
    );
    expect((after.content as string).includes("MCP Cell")).toBe(true);
  });

  it("pptx_apply_command dispatches typed chart commands and json projection surfaces them", async () => {
    const client = await makeClient();
    const loaded = structured(
      await client.callTool({ name: "pptx_load", arguments: { path: PPTX_CHART_FIXTURE } })
    );
    const handle = loaded.handle as string;

    const summary = loaded.summary as { shapeCounts: { chart: number } };
    expect(summary.shapeCounts.chart).toBeGreaterThanOrEqual(1);

    const json = structured(
      await client.callTool({ name: "pptx_get_text", arguments: { handle, format: "json" } })
    ) as unknown as {
      slides: Array<{
        index: number;
        shapes: Array<{ id: string; kind: string; chart?: { chartType: string } }>;
      }>;
    };
    let slideIndex = -1;
    let chartId = "";
    for (const slide of json.slides) {
      const c = slide.shapes.find((s) => s.kind === "chart");
      if (c) {
        slideIndex = slide.index;
        chartId = c.id;
        break;
      }
    }
    expect(chartId).not.toBe("");

    const apply = structured(
      await client.callTool({
        name: "pptx_apply_command",
        arguments: {
          handle,
          type: "pptx:set-chart-title",
          payload: { slideIndex, shapeId: chartId, title: "MCP Chart Title" },
        },
      })
    );
    expect((apply.mutation as { status: string }).status).toBe("approved");

    const apply2 = structured(
      await client.callTool({
        name: "pptx_apply_command",
        arguments: {
          handle,
          type: "pptx:set-chart-type",
          payload: { slideIndex, shapeId: chartId, chartType: "pie" },
        },
      })
    );
    expect((apply2.mutation as { status: string }).status).toBe("approved");

    const md = structured(
      await client.callTool({
        name: "pptx_get_text",
        arguments: { handle, format: "markdown", slide: slideIndex },
      })
    );
    expect((md.content as string).includes("MCP Chart Title")).toBe(true);

    const after = structured(
      await client.callTool({ name: "pptx_get_text", arguments: { handle, format: "json" } })
    ) as unknown as {
      slides: Array<{ shapes: Array<{ kind: string; chart?: { chartType: string; title?: string } }> }>;
    };
    const ch = after.slides.flatMap((s) => s.shapes).find((sh) => sh.kind === "chart");
    expect(ch?.chart?.chartType).toBe("pie");
    expect(ch?.chart?.title).toBe("MCP Chart Title");
  });

  it("pptx_apply_command dispatches typed animation commands and json projection surfaces them", async () => {
    const client = await makeClient();
    const loaded = structured(
      await client.callTool({ name: "pptx_load", arguments: { path: PPTX_ANIM_FIXTURE } })
    );
    const handle = loaded.handle as string;

    const summary = loaded.summary as { animations: number; transitions: number };
    expect(summary.animations).toBeGreaterThan(0);
    expect(summary.transitions).toBeGreaterThan(0);

    const json = structured(
      await client.callTool({ name: "pptx_get_text", arguments: { handle, format: "json" } })
    ) as unknown as {
      slides: Array<{
        index: number;
        transition?: { kind: string };
        animations?: Array<{ id: string; effect: string; targetCNvPrId: number }>;
      }>;
    };
    const slide0 = json.slides[0]!;
    expect(slide0.transition?.kind).toBe("fade");
    expect(slide0.animations?.length).toBe(2);

    const apply = structured(
      await client.callTool({
        name: "pptx_apply_command",
        arguments: {
          handle,
          type: "pptx:set-slide-transition",
          payload: { slideIndex: 0, kind: "push", speed: "fast" },
        },
      })
    );
    expect((apply.mutation as { status: string }).status).toBe("approved");

    const dropId = slide0.animations![0]!.id;
    const apply2 = structured(
      await client.callTool({
        name: "pptx_apply_command",
        arguments: {
          handle,
          type: "pptx:remove-shape-animation",
          payload: { slideIndex: 0, animationId: dropId },
        },
      })
    );
    expect((apply2.mutation as { status: string }).status).toBe("approved");

    const after = structured(
      await client.callTool({ name: "pptx_get_text", arguments: { handle, format: "json" } })
    ) as unknown as {
      slides: Array<{
        transition?: { kind: string; speed?: string };
        animations?: Array<{ effect: string }>;
      }>;
    };
    expect(after.slides[0]!.transition?.kind).toBe("push");
    expect(after.slides[0]!.transition?.speed).toBe("fast");
    expect(after.slides[0]!.animations?.length).toBe(1);

    const md = structured(
      await client.callTool({
        name: "pptx_get_text",
        arguments: { handle, format: "markdown", slide: 0 },
      })
    );
    expect((md.content as string).includes("transition")).toBe(true);
    expect((md.content as string).includes("push")).toBe(true);
  });

  it("returns an error for unknown pptx handles", async () => {
    const client = await makeClient();
    const r = await client.callTool({ name: "pptx_inspect", arguments: { handle: "nope" } });
    expect(r.isError).toBe(true);
  });
});

describe("OfficeAI MCP server — xlsx tools", () => {
  beforeEach(() => __resetMcpSessionsForTests());
  afterEach(() => __resetMcpSessionsForTests());

  async function loadFixture(client: Client, name: string): Promise<{ handle: string; sheet: string }> {
    const dir = mkdtempSync(join(tmpdir(), "mcp-xlsx-"));
    const path = copyXlsxFixture(name, dir);
    const out = structured(await client.callTool({ name: "xlsx_load", arguments: { path } }));
    const summary = out.summary as { sheets: Array<{ name: string; kind: string }> };
    const sheet = summary.sheets.find((s) => s.kind === "worksheet")!.name;
    return { handle: out.handle as string, sheet };
  }

  it("xlsx_load returns a handle and an inspection summary", async () => {
    const client = await makeClient();
    const dir = mkdtempSync(join(tmpdir(), "mcp-xlsx-load-"));
    const path = copyXlsxFixture("01-single-sheet-numbers.xlsx", dir);
    const out = structured(await client.callTool({ name: "xlsx_load", arguments: { path } }));
    expect(typeof out.handle).toBe("string");
    expect(out.path).toBe(path);
    const summary = out.summary as { format: string; sheets: unknown[] };
    expect(summary.format).toBe("xlsx");
    expect(Array.isArray(summary.sheets)).toBe(true);
  });

  it("xlsx_apply_command with xlsx:set-cell-value updates the snapshot", async () => {
    const client = await makeClient();
    const { handle, sheet } = await loadFixture(client, "01-single-sheet-numbers.xlsx");
    const apply = structured(
      await client.callTool({
        name: "xlsx_apply_command",
        arguments: {
          handle,
          type: "xlsx:set-cell-value",
          payload: { sheet, ref: "AA50", value: "via-apply-command" },
        },
      })
    );
    expect((apply.mutation as { status: string }).status).toBe("approved");

    const projection = structured(
      await client.callTool({
        name: "xlsx_get_text",
        arguments: { handle, format: "json", sheet, range: "AA50:AA50" },
      })
    );
    expect((projection.cells as Array<{ value: unknown }>)[0].value).toBe("via-apply-command");
  });

  it("xlsx_set_cell convenience tool produces an equivalent mutation", async () => {
    const client = await makeClient();
    const { handle, sheet } = await loadFixture(client, "01-single-sheet-numbers.xlsx");
    const apply = structured(
      await client.callTool({
        name: "xlsx_set_cell",
        arguments: { handle, sheet, ref: "AB1", value: 99 },
      })
    );
    expect((apply.mutation as { status: string }).status).toBe("approved");

    const projection = structured(
      await client.callTool({
        name: "xlsx_get_text",
        arguments: { handle, format: "json", sheet, range: "AB1:AB1" },
      })
    );
    expect((projection.cells as Array<{ value: unknown }>)[0].value).toBe(99);
  });

  it("xlsx_undo reverts the last mutation; xlsx_redo reapplies it", async () => {
    const client = await makeClient();
    const { handle, sheet } = await loadFixture(client, "01-single-sheet-numbers.xlsx");

    // Establish a known "after" value so the round-trip is observable.
    await client.callTool({
      name: "xlsx_set_cell",
      arguments: { handle, sheet, ref: "AD1", value: 42 },
    });
    const before = structured(
      await client.callTool({
        name: "xlsx_get_text",
        arguments: { handle, format: "json", sheet, range: "AD1:AD1" },
      })
    );
    expect((before.cells as Array<{ value: unknown }>)[0].value).toBe(42);

    const undone = structured(await client.callTool({ name: "xlsx_undo", arguments: { handle } }));
    expect(undone.did_undo).toBe(true);
    expect(undone.can_redo).toBe(true);
    const after = structured(
      await client.callTool({
        name: "xlsx_get_text",
        arguments: { handle, format: "json", sheet, range: "AD1:AD1" },
      })
    );
    // Cell is now empty; the projection drops the cell entirely.
    expect((after.cells as Array<unknown>).length).toBe(0);

    const redone = structured(await client.callTool({ name: "xlsx_redo", arguments: { handle } }));
    expect(redone.did_redo).toBe(true);
    const restored = structured(
      await client.callTool({
        name: "xlsx_get_text",
        arguments: { handle, format: "json", sheet, range: "AD1:AD1" },
      })
    );
    expect((restored.cells as Array<{ value: unknown }>)[0].value).toBe(42);

    // Empty undo / redo are no-ops, not errors.
    await client.callTool({ name: "xlsx_redo", arguments: { handle } });
    const noopRedo = structured(await client.callTool({ name: "xlsx_redo", arguments: { handle } }));
    expect(noopRedo.did_redo).toBe(false);
  });

  it("xlsx_save writes the modified workbook to disk", async () => {
    const client = await makeClient();
    const { handle, sheet } = await loadFixture(client, "01-single-sheet-numbers.xlsx");
    await client.callTool({
      name: "xlsx_set_cell",
      arguments: { handle, sheet, ref: "AC1", value: "saved!" },
    });
    const dir = mkdtempSync(join(tmpdir(), "mcp-xlsx-save-"));
    const out = join(dir, "out.xlsx");
    const saved = structured(
      await client.callTool({ name: "xlsx_save", arguments: { handle, out_path: out } })
    );
    expect(saved.wrote).toBe(out);
    const reloaded = await XlsxAgent.fromBuffer(readFileSync(out));
    const reloadedSheet = reloaded.getSnapshot().root.sheets.find((s) => s.name === sheet)!;
    const matched = [...reloadedSheet.cells.values()].find((c) => c.value === "saved!");
    expect(matched).toBeDefined();
  });

  it("xlsx_diff works for both handle-pair and disk modes", async () => {
    const client = await makeClient();
    const dir = mkdtempSync(join(tmpdir(), "mcp-xlsx-diff-"));
    const path = copyXlsxFixture("01-single-sheet-numbers.xlsx", dir);
    const a = structured(await client.callTool({ name: "xlsx_load", arguments: { path } }));
    const b = structured(await client.callTool({ name: "xlsx_load", arguments: { path } }));
    const sheet = (a.summary as { sheets: Array<{ name: string; kind: string }> }).sheets.find(
      (s) => s.kind === "worksheet"
    )!.name;

    await client.callTool({
      name: "xlsx_set_cell",
      arguments: { handle: b.handle, sheet, ref: "AD1", value: "diff" },
    });

    const handlePairDiff = structured(
      await client.callTool({
        name: "xlsx_diff",
        arguments: { before: a.handle, after: b.handle },
      })
    );
    expect(handlePairDiff.format).toBe("xlsx");
    expect((handlePairDiff.changes as unknown[]).length).toBeGreaterThanOrEqual(1);

    const diskDiff = structured(
      await client.callTool({
        name: "xlsx_diff",
        arguments: { handle: b.handle, against: "disk" },
      })
    );
    expect(diskDiff.format).toBe("xlsx");
    expect((diskDiff.changes as unknown[]).length).toBeGreaterThanOrEqual(1);
  });

  it("xlsx pending review flow: list, approve, reject", async () => {
    const client = await makeClient();
    const { handle, sheet } = await loadFixture(client, "01-single-sheet-numbers.xlsx");

    const apply = structured(
      await client.callTool({
        name: "xlsx_apply_command",
        arguments: {
          handle,
          type: "xlsx:set-cell-value",
          payload: { sheet, ref: "AE1", value: "pending" },
          auto_approve: false,
        },
      })
    );
    expect((apply.mutation as { status: string }).status).toBe("pending");
    const id = (apply.mutation as { id: string }).id;

    const list1 = structured(await client.callTool({ name: "xlsx_list_pending", arguments: { handle } }));
    expect((list1.pending as Array<{ id: string }>).map((p) => p.id)).toContain(id);

    const approve = structured(
      await client.callTool({ name: "xlsx_approve", arguments: { handle, mutation_id: id } })
    );
    expect(approve.approved).toBe(id);

    const list2 = structured(await client.callTool({ name: "xlsx_list_pending", arguments: { handle } }));
    expect(list2.pending as unknown[]).toHaveLength(0);

    const apply2 = structured(
      await client.callTool({
        name: "xlsx_apply_command",
        arguments: {
          handle,
          type: "xlsx:set-cell-value",
          payload: { sheet, ref: "AF1", value: "REJECTME" },
          auto_approve: false,
        },
      })
    );
    const id2 = (apply2.mutation as { id: string }).id;
    const rej = structured(
      await client.callTool({
        name: "xlsx_reject",
        arguments: { handle, mutation_id: id2, reason: "nope" },
      })
    );
    expect(rej.rejected).toBe(id2);
    expect(rej.reason).toBe("nope");
  });

  it("xlsx_list_sheets returns sheets in tab order", async () => {
    const client = await makeClient();
    const dir = mkdtempSync(join(tmpdir(), "mcp-xlsx-list-"));
    const path = copyXlsxFixture("02-multi-sheet.xlsx", dir);
    const loaded = structured(await client.callTool({ name: "xlsx_load", arguments: { path } }));
    const list = structured(
      await client.callTool({ name: "xlsx_list_sheets", arguments: { handle: loaded.handle } })
    );
    const sheets = list.sheets as Array<{ name: string; index: number }>;
    expect(sheets.length).toBeGreaterThan(1);
    const indexes = sheets.map((s) => s.index);
    expect(indexes).toEqual([...indexes].sort((a, b) => a - b));
  });

  it("returns an error for unknown xlsx handles", async () => {
    const client = await makeClient();
    const r = await client.callTool({
      name: "xlsx_inspect",
      arguments: { handle: "nope" },
    });
    expect(r.isError).toBe(true);
  });
});
