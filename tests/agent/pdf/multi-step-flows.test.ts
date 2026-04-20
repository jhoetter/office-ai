import { mkdtempSync, readFileSync } from "node:fs";
import { readFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { runCli } from "@officeai/agent";
import { PdfAgent } from "@officeai/pdf";

/**
 * `tests/agent/pdf` is intentionally complementary to
 * `packages/agent/src/pdf-cli.test.ts`. Those tests pin individual
 * subcommands; the tests here exercise multi-step CLI flows like
 * "fill-form → flatten-form → re-list yields no fillable widgets" and
 * "rotate → set-metadata → re-parse preserves both edits".
 */

const FIXTURE_DIR = resolve(__dirname, "../../../fixtures/pdf");
const TMP = mkdtempSync(join(tmpdir(), "pdf-agent-flows-"));

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

function makeIO(): {
  io: { stdout: NodeJS.WritableStream; stderr: NodeJS.WritableStream };
  stdout: CapturedStream;
  stderr: CapturedStream;
} {
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

beforeAll(() => {
  process.env.OFFICEAI_DETERMINISTIC_IDS = "1";
});
afterAll(() => {
  delete process.env.OFFICEAI_DETERMINISTIC_IDS;
});

const fixture = (name: string): string => join(FIXTURE_DIR, name);

async function readBytes(path: string): Promise<Uint8Array> {
  const buf = await readFile(path);
  return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
}

describe("office-agent pdf — multi-step flows", () => {
  it("fill-form → flatten-form → re-list yields no fillable widgets", async () => {
    const filled = join(TMP, "fill-flat-1.pdf");
    const flat = join(TMP, "fill-flat-2.pdf");

    {
      const { io, stdout } = makeIO();
      const code = await runCli(
        [
          "pdf",
          "fill-form",
          fixture("acroform-basic.pdf"),
          "--values",
          JSON.stringify({
            "first.name": "Grace",
            agree: true,
            country: "JP",
            plan: "pro",
          }),
          "--out",
          filled,
        ],
        io,
      );
      expect(code).toBe(0);
      const env = JSON.parse(stdout.text()) as { schema: string; summary: string };
      expect(env.schema).toBe("office-agent/pdf-fill-form@1");
      expect(env.summary).toMatch(/filled 4/);
    }

    // Verify the values landed before flattening.
    {
      const { io, stdout } = makeIO();
      const code = await runCli(["pdf", "list-form-fields", filled], io);
      expect(code).toBe(0);
      const env = JSON.parse(stdout.text()) as {
        fields: ReadonlyArray<{ name: string; value?: unknown }>;
      };
      const byName = Object.fromEntries(env.fields.map((f) => [f.name, f] as const));
      expect(byName["first.name"]?.value).toBe("Grace");
      expect(byName["agree"]?.value).toBe(true);
      expect(byName["country"]?.value).toEqual(["JP"]);
      expect(byName["plan"]?.value).toBe("pro");
    }

    {
      const { io, stdout } = makeIO();
      const code = await runCli(["pdf", "flatten-form", filled, "--out", flat], io);
      expect(code).toBe(0);
      const env = JSON.parse(stdout.text()) as { schema: string; bytes: number };
      expect(env.schema).toBe("office-agent/pdf-flatten-form@1");
      expect(env.bytes).toBeGreaterThan(0);
    }

    {
      const { io, stdout } = makeIO();
      const code = await runCli(["pdf", "list-form-fields", flat], io);
      expect(code).toBe(0);
      const env = JSON.parse(stdout.text()) as { fields: ReadonlyArray<unknown> };
      expect(env.fields).toHaveLength(0);
    }

    // Final round-trip: bytes still parse, page count unchanged.
    const reparsed = await PdfAgent.fromBuffer(await readBytes(flat));
    expect(reparsed.getSnapshot().root.pages).toHaveLength(1);
  });

  it("rotate-pages → set-metadata → re-parse preserves both edits", async () => {
    const rotated = join(TMP, "rot-meta-1.pdf");
    const titled = join(TMP, "rot-meta-2.pdf");

    {
      const { io } = makeIO();
      const code = await runCli(
        [
          "pdf",
          "rotate-pages",
          fixture("simple-text-3page.pdf"),
          "--pages",
          "1,3",
          "--delta",
          "180",
          "--out",
          rotated,
        ],
        io,
      );
      expect(code).toBe(0);
    }

    {
      const { io } = makeIO();
      const code = await runCli(
        [
          "pdf",
          "set-metadata",
          rotated,
          "--title",
          "Multi-Step",
          "--subject",
          "rotate then patch",
          "--out",
          titled,
        ],
        io,
      );
      expect(code).toBe(0);
    }

    const reparsed = await PdfAgent.fromBuffer(await readBytes(titled));
    expect(reparsed.getSnapshot().root.metadata.title).toBe("Multi-Step");
    expect(reparsed.getSnapshot().root.metadata.subject).toBe("rotate then patch");
    expect(reparsed.getSnapshot().root.pages.map((p) => p.rotation)).toEqual([180, 0, 180]);
  });

  it("split → merge round-trip restores the original page count", async () => {
    const prefix = join(TMP, "sp");
    {
      const { io, stdout } = makeIO();
      const code = await runCli(
        ["pdf", "split", fixture("simple-text-3page.pdf"), "--at", "2", "--out-prefix", prefix],
        io,
      );
      expect(code).toBe(0);
      const env = JSON.parse(stdout.text()) as { outputs: ReadonlyArray<{ out: string }> };
      expect(env.outputs).toHaveLength(2);
    }

    const merged = join(TMP, "sp-merged.pdf");
    const { io, stdout } = makeIO();
    const code = await runCli(
      ["pdf", "merge", `${prefix}-001.pdf`, `${prefix}-002.pdf`, "--out", merged],
      io,
    );
    expect(code).toBe(0);
    const env = JSON.parse(stdout.text()) as { schema: string; bytes: number };
    expect(env.schema).toBe("office-agent/pdf-merge@1");
    expect(env.bytes).toBeGreaterThan(0);

    const reparsed = await PdfAgent.fromBuffer(new Uint8Array(readFileSync(merged)));
    expect(reparsed.getSnapshot().root.pages).toHaveLength(3);
  });

  it("watermark → page-numbers → re-parse keeps page count and emits valid bytes", async () => {
    const wm = join(TMP, "wm-1.pdf");
    const numbered = join(TMP, "wm-2.pdf");

    {
      const { io } = makeIO();
      const code = await runCli(
        [
          "pdf",
          "add-watermark",
          fixture("simple-text-3page.pdf"),
          "--text",
          "CONFIDENTIAL",
          "--out",
          wm,
        ],
        io,
      );
      expect(code).toBe(0);
    }

    {
      const { io } = makeIO();
      const code = await runCli(
        ["pdf", "add-page-numbers", wm, "--position", "bottom-right", "--out", numbered],
        io,
      );
      expect(code).toBe(0);
    }

    const reparsed = await PdfAgent.fromBuffer(await readBytes(numbered));
    expect(reparsed.getSnapshot().root.pages).toHaveLength(3);
  });

  it("read-metadata followed by an out-of-range read-page produces a structured error", async () => {
    {
      const { io, stdout } = makeIO();
      const code = await runCli(
        ["pdf", "read-metadata", fixture("simple-text-1page.pdf")],
        io,
      );
      expect(code).toBe(0);
      const env = JSON.parse(stdout.text()) as { pageCount: number };
      expect(env.pageCount).toBe(1);
    }
    {
      const { io, stderr } = makeIO();
      const code = await runCli(
        ["pdf", "read-page", fixture("simple-text-1page.pdf"), "--page", "5"],
        io,
      );
      expect(code).toBe(1);
      const lines = stderr.text().trim().split("\n").filter(Boolean);
      const env = JSON.parse(lines[lines.length - 1]) as { error: string };
      expect(env.error).toBe("page-out-of-range");
    }
  });
});
