import { existsSync, mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { describe, expect, it } from "vitest";
import { fixturePath, requiredMatrixFixture } from "../fixture-matrix.js";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "../..");
const cliPath = resolve(repoRoot, "packages/agent/dist/cli.js");

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

describe("office-agent mcp stdio", () => {
  it("imports, projects and exports a DOCX through the canonical tools", async () => {
    if (!existsSync(cliPath)) {
      throw new Error(`Missing ${cliPath}; run pnpm --filter @officeai/agent build before this smoke.`);
    }
    const dataDir = mkdtempSync(join(tmpdir(), "officeai-stdio-data-"));
    const outDir = mkdtempSync(join(tmpdir(), "officeai-stdio-out-"));
    const childEnv: Record<string, string> = {};
    for (const [key, value] of Object.entries(process.env)) {
      if (value !== undefined) childEnv[key] = value;
    }
    childEnv.OFFICEAI_DATA_DIR = dataDir;

    const transport = new StdioClientTransport({
      command: process.execPath,
      args: [cliPath, "mcp"],
      cwd: repoRoot,
      env: childEnv,
      stderr: "pipe",
    });
    const client = new Client({ name: "officeai-stdio-smoke", version: "0.0.0" }, { capabilities: {} });

    try {
      await client.connect(transport);
      const session = structured(
        await client.callTool({ name: "create_session", arguments: { title: "stdio smoke" } })
      );
      const sessionId = session.sessionId as string;
      const fixture = requiredMatrixFixture("docx", {
        complexity: "simple",
        expectedBehavior: "import",
      });
      const imported = structured(
        await client.callTool({
          name: "import_document",
          arguments: { session_id: sessionId, path: fixturePath(fixture) },
        })
      );
      const document = imported.document as { documentId: string; sessionId: string; format: string };
      expect(document.sessionId).toBe(sessionId);
      expect(document.format).toBe("docx");

      const projection = structured(
        await client.callTool({
          name: "get_document_projection",
          arguments: { document_id: document.documentId, projection: "markdown" },
        })
      );
      expect(typeof projection.content).toBe("string");
      expect((projection.content as string).length).toBeGreaterThan(20);

      const out = join(outDir, "out.docx");
      const exported = structured(
        await client.callTool({
          name: "export_document",
          arguments: { document_id: document.documentId, out_path: out },
        })
      );
      expect((exported.exported as { bytes?: number }).bytes ?? 0).toBeGreaterThan(0);
    } finally {
      await client.close();
      rmSync(dataDir, { recursive: true, force: true });
      rmSync(outDir, { recursive: true, force: true });
    }
  }, 20_000);
});
