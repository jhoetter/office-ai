import { mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { runCli } from "@officeai/agent";
import { PdfAgent } from "@officeai/pdf";
import { FIXTURE_DIR, isPdfBytes, loadFixture } from "./helpers.js";

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

const TMP = mkdtempSync(join(tmpdir(), "pdf-roundtrip-cli-"));

beforeAll(() => {
  process.env.OFFICEAI_DETERMINISTIC_IDS = "1";
});
afterAll(() => {
  delete process.env.OFFICEAI_DETERMINISTIC_IDS;
});

function fixturePath(name: string): string {
  return join(FIXTURE_DIR, name);
}

/**
 * CLI-level roundtrip: drive the `office-agent pdf-*` surface against
 * the on-disk fixture corpus, asserting JSON envelope shape, output
 * bytes > 0, and re-parse round-trip.
 */
describe("PDF roundtrip — CLI envelope shape & re-parse", () => {
  it("read-metadata matches metadata-rich.pdf's Info dict", async () => {
    const { io, stdout } = makeIO();
    const code = await runCli(["pdf", "read-metadata", fixturePath("metadata-rich.pdf")], io);
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text()) as {
      schema: string;
      title?: string;
      author?: string;
      pageCount: number;
    };
    expect(parsed.schema).toBe("office-agent/pdf-read-metadata@1");
    expect(parsed.title).toBe("Metadata-Rich Fixture");
    expect(parsed.author).toBe("Office AI Night Shift");
    expect(parsed.pageCount).toBe(1);
  });

  it("read-outline returns the 3-chapter tree from with-outline.pdf", async () => {
    const { io, stdout } = makeIO();
    const code = await runCli(["pdf", "read-outline", fixturePath("with-outline.pdf")], io);
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text()) as {
      schema: string;
      outline: ReadonlyArray<{ title: string; children: ReadonlyArray<unknown> }>;
    };
    expect(parsed.schema).toBe("office-agent/pdf-read-outline@1");
    expect(parsed.outline).toHaveLength(3);
    expect(parsed.outline[1].children).toHaveLength(2);
  });

  it("read-annotations enumerates the link annotation in with-link-annot.pdf", async () => {
    const { io, stdout } = makeIO();
    const code = await runCli(
      ["pdf", "read-annotations", fixturePath("with-link-annot.pdf")],
      io,
    );
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text()) as {
      schema: string;
      annotations: ReadonlyArray<{ kind: string; url?: string }>;
    };
    expect(parsed.schema).toBe("office-agent/pdf-read-annotations@1");
    expect(parsed.annotations).toHaveLength(1);
    expect(parsed.annotations[0].kind).toBe("link");
    expect(parsed.annotations[0].url).toMatch(/^https:\/\/cursor\.com\/?$/);
  });

  it("rotate-pages writes rotated bytes that re-parse with the new rotation", async () => {
    const out = join(TMP, "cli-rot.pdf");
    const { io, stdout } = makeIO();
    const code = await runCli(
      [
        "pdf",
        "rotate-pages",
        fixturePath("simple-text-3page.pdf"),
        "--pages",
        "2",
        "--delta",
        "90",
        "--out",
        out,
      ],
      io,
    );
    expect(code).toBe(0);
    const env = JSON.parse(stdout.text()) as { schema: string; bytes: number };
    expect(env.schema).toBe("office-agent/pdf-rotate-pages@1");
    expect(env.bytes).toBeGreaterThan(0);
    const written = new Uint8Array(readFileSync(out));
    expect(isPdfBytes(written)).toBe(true);
    const reparsed = await PdfAgent.fromBuffer(written);
    expect(reparsed.getSnapshot().root.pages.map((p) => p.rotation)).toEqual([0, 90, 0]);
  });

  it("set-metadata via the CLI persists across re-parse", async () => {
    const out = join(TMP, "cli-meta.pdf");
    const { io, stdout } = makeIO();
    const code = await runCli(
      [
        "pdf",
        "set-metadata",
        fixturePath("simple-text-1page.pdf"),
        "--title",
        "From CLI",
        "--author",
        "Roundtrip",
        "--out",
        out,
      ],
      io,
    );
    expect(code).toBe(0);
    const env = JSON.parse(stdout.text()) as { schema: string };
    expect(env.schema).toBe("office-agent/pdf-set-metadata@1");
    const reparsed = await PdfAgent.fromBuffer(new Uint8Array(readFileSync(out)));
    expect(reparsed.getSnapshot().root.metadata.title).toBe("From CLI");
    expect(reparsed.getSnapshot().root.metadata.author).toBe("Roundtrip");
  });

  it("export-markdown without --out streams a non-empty markdown blob to stdout", async () => {
    const { io, stdout } = makeIO();
    const code = await runCli(["pdf", "export-markdown", fixturePath("with-outline.pdf")], io);
    expect(code).toBe(0);
    const txt = stdout.text();
    expect(txt).toMatch(/Outline/);
    expect(txt).toMatch(/Page 1/);
  });

  // Drives writes for the audit's file-output assertion via a real
  // CLI flow rather than in-process pdf-edit.
  it("add-watermark CLI writes valid PDF bytes", async () => {
    const out = join(TMP, "cli-wm.pdf");
    const { io, stdout } = makeIO();
    const code = await runCli(
      [
        "pdf",
        "add-watermark",
        fixturePath("simple-text-1page.pdf"),
        "--text",
        "DRAFT",
        "--out",
        out,
      ],
      io,
    );
    expect(code).toBe(0);
    const env = JSON.parse(stdout.text()) as { schema: string; bytes: number };
    expect(env.schema).toBe("office-agent/pdf-add-watermark@1");
    expect(env.bytes).toBeGreaterThan(0);
    const written = new Uint8Array(readFileSync(out));
    expect(isPdfBytes(written)).toBe(true);
    const reparsed = await PdfAgent.fromBuffer(written);
    expect(reparsed.getSnapshot().root.pages).toHaveLength(1);
    // Touch loadFixture so the import isn't dead — used by helpers
    // consumers as the "give me the canonical bytes" entry point.
    const baseline = await loadFixture("simple-text-1page.pdf");
    expect(baseline.length).toBeGreaterThan(0);
  });
});
