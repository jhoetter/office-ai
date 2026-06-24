import { mkdtempSync, writeFileSync, readFileSync, mkdirSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve as resolvePath } from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { Document, Packer, Paragraph, TextRun, HeadingLevel } from "docx";
import { runCli } from "./cli.js";
import { DocxAgent } from "@officeai/docx";
import { XlsxAgent } from "@officeai/xlsx";

const here = dirname(fileURLToPath(import.meta.url));
const xlsxFixtures = resolvePath(here, "../../../fixtures/xlsx/synthetic");

function copyFixture(name: string, dest: string): string {
  const src = resolvePath(xlsxFixtures, name);
  const target = join(dest, name);
  writeFileSync(target, readFileSync(src));
  return target;
}

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

async function makeFixture(): Promise<string> {
  const dir = mkdtempSync(join(tmpdir(), "office-agent-cli-"));
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

/**
 * Hand-rolled tracked-changes fixture: paragraph 0 starts with the literal
 * "Original " followed by a `<w:ins w:id="42">INSERTED</w:ins>` wrapper.
 * Hand-rolling the OOXML keeps the test independent of whatever the `docx`
 * library happens to emit for tracked changes today.
 */
async function makeTrackedChangesFixture(): Promise<string> {
  const dir = mkdtempSync(join(tmpdir(), "office-agent-tc-"));
  const z = new JSZip();
  z.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`
  );
  z.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );
  z.file(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`
  );
  z.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">Original </w:t></w:r>
      <w:ins w:id="42" w:author="A" w:date="2024-01-01T00:00:00Z">
        <w:r><w:t>INSERTED</w:t></w:r>
      </w:ins>
    </w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
  </w:body>
</w:document>`
  );
  const buf = (await z.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  })) as Buffer;
  const inputPath = join(dir, "tc.docx");
  writeFileSync(inputPath, buf);
  return inputPath;
}

/**
 * 1×1 transparent PNG (the smallest valid PNG by byte count). Good enough
 * for `docx insert-image` round-trip tests — the handler validates magic
 * bytes for PNG via the mime type hint, not by re-decoding the raster.
 */
const TINY_PNG_BYTES = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52, 0x00, 0x00,
  0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4, 0x89, 0x00, 0x00, 0x00,
  0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d,
  0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
]);

describe("office-agent CLI", () => {
  it("read prints markdown", async () => {
    const input = await makeFixture();
    const { io, stdout } = makeIO();
    const code = await runCli(["read", "-i", input], io);
    expect(code).toBe(0);
    expect(stdout.text()).toContain("# Hello");
    expect(stdout.text()).toContain("first paragraph body");
  });

  it("search prints JSON matches", async () => {
    const input = await makeFixture();
    const { io, stdout } = makeIO();
    const code = await runCli(["search", "-i", input, "-q", "paragraph"], io);
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text());
    expect(Array.isArray(parsed)).toBe(true);
    expect(parsed.length).toBeGreaterThanOrEqual(2);
  });

  it("doctor prints a machine-readable runtime report", async () => {
    const dir = mkdtempSync(join(tmpdir(), "office-agent-doctor-"));
    const { io, stdout } = makeIO();
    const code = await runCli(["doctor", "--json", "--data-dir", dir], io);
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text()) as {
      schema: string;
      ok: boolean;
      checks: Array<{ code: string; status: string }>;
    };
    expect(parsed.schema).toBe("office-ai/doctor@1");
    expect(parsed.ok).toBe(true);
    expect(parsed.checks.map((check) => check.code)).toEqual(
      expect.arrayContaining(["node", "package-manager", "data-dir"])
    );
  });

  it("sessions migrate prints a machine-readable migration report", async () => {
    const dir = mkdtempSync(join(tmpdir(), "office-agent-migrate-"));
    const sessionDir = join(dir, "sessions", "session_legacy");
    mkdirSync(sessionDir, { recursive: true });
    writeFileSync(
      join(sessionDir, "session.json"),
      JSON.stringify({
        id: "session_legacy",
        title: "Legacy",
        createdAt: "2026-06-24T09:00:00.000Z",
        updatedAt: "2026-06-24T09:00:00.000Z",
        documentIds: [],
      })
    );
    const { io, stdout } = makeIO();
    const code = await runCli(["sessions", "migrate", "--json", "--data-dir", dir], io);
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text()) as {
      schema: string;
      migrations: Array<{ kind: string; id: string; backupPath: string }>;
    };
    expect(parsed.schema).toBe("office-ai/session-store-migration@1");
    expect(parsed.migrations).toEqual([
      expect.objectContaining({
        kind: "session",
        id: "session_legacy",
        backupPath: expect.stringContaining("backups"),
      }),
    ]);
  });

  it("insert-text writes a modified file that re-reads with the new text", async () => {
    const input = await makeFixture();
    const dir = mkdtempSync(join(tmpdir(), "office-agent-out-"));
    const output = join(dir, "out.docx");
    const { io } = makeIO();
    const code = await runCli(
      ["insert-text", "-i", input, "-o", output, "--at", "paragraph:1/run:0/text:0", "--text", "EDIT — "],
      io
    );
    expect(code).toBe(0);
    const agent = await DocxAgent.fromBuffer(readFileSync(output));
    expect(agent.toMarkdown()).toContain("EDIT — first paragraph body");
  });

  it("comment writes a file with a comments part", async () => {
    const input = await makeFixture();
    const dir = mkdtempSync(join(tmpdir(), "office-agent-out-"));
    const output = join(dir, "out.docx");
    const { io } = makeIO();
    const code = await runCli(
      [
        "comment",
        "-i",
        input,
        "-o",
        output,
        "--range",
        "paragraph:1/text:0..5",
        "--text",
        "rephrase?",
        "--author",
        "Tester",
        "--initials",
        "T",
      ],
      io
    );
    expect(code).toBe(0);
    const agent = await DocxAgent.fromBuffer(readFileSync(output));
    expect(agent.getSnapshot().root.comments).toHaveLength(1);
    expect(agent.getSnapshot().root.comments[0].author).toBe("Tester");
  });

  it("apply runs a JSON command file", async () => {
    const input = await makeFixture();
    const dir = mkdtempSync(join(tmpdir(), "office-agent-apply-"));
    mkdirSync(dir, { recursive: true });
    const cmdsPath = join(dir, "commands.json");
    writeFileSync(
      cmdsPath,
      JSON.stringify({
        commands: [
          {
            type: "docx:insert-text",
            payload: { at: { paragraph: 1, run: 0, offset: 0 }, text: "JSON " },
          },
          {
            type: "docx:set-paragraph-style",
            payload: { at: { paragraph: 2 }, style: "Heading2" },
          },
        ],
      })
    );
    const output = join(dir, "out.docx");
    const { io, stdout } = makeIO();
    const code = await runCli(["apply", "-i", input, "-o", output, "-c", cmdsPath], io);
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text());
    expect(parsed.mutations).toHaveLength(2);
    const agent = await DocxAgent.fromBuffer(readFileSync(output));
    expect(agent.toMarkdown()).toContain("JSON first paragraph body");
  });

  it("apply with N>=3 commands does not duplicate insert-text (§G0 regression)", async () => {
    // Repro of the bug captured in docs/cli-gap-report.md §G0:
    // the previous `getPendingMutations().forEach(approveMutation)` pattern
    // mutated `this.pending` mid-iteration, leaving half the mutations
    // pending and re-applying them via `recomputeWorking`. A four-command
    // batch produced `BBB...` in paragraph 1 instead of `B...`.
    const input = await makeFixture();
    const dir = mkdtempSync(join(tmpdir(), "office-agent-g0-"));
    mkdirSync(dir, { recursive: true });
    const cmdsPath = join(dir, "commands.json");
    writeFileSync(
      cmdsPath,
      JSON.stringify({
        commands: [
          { type: "docx:insert-text", payload: { at: { paragraph: 0, offset: 0 }, text: "A" } },
          { type: "docx:insert-text", payload: { at: { paragraph: 1, offset: 0 }, text: "B" } },
          { type: "docx:insert-text", payload: { at: { paragraph: 2, offset: 0 }, text: "C" } },
        ],
      })
    );
    const output = join(dir, "out.docx");
    const { io } = makeIO();
    const code = await runCli(["apply", "-i", input, "-o", output, "-c", cmdsPath], io);
    expect(code).toBe(0);
    const md = (await DocxAgent.fromBuffer(readFileSync(output))).toMarkdown();
    expect(md).toContain("AHello");
    expect(md).toContain("Bfirst paragraph body");
    expect(md).toContain("Csecond paragraph body");
    expect(md).not.toMatch(/BB/);
    expect(md).not.toMatch(/CC/);
  });

  it("invalid selector reports an error", async () => {
    const input = await makeFixture();
    const dir = mkdtempSync(join(tmpdir(), "office-agent-bad-"));
    const output = join(dir, "out.docx");
    const { io, stderr } = makeIO();
    const code = await runCli(
      ["insert-text", "-i", input, "-o", output, "--at", "garbage:0", "--text", "x"],
      io
    );
    expect(code).toBe(64);
    expect(stderr.text()).toMatch(/selector error/i);
  });
});

describe("office-agent docx subcommand group", () => {
  it("docx inspect prints structural counts as JSON", async () => {
    const input = await makeFixture();
    const { io, stdout } = makeIO();
    const code = await runCli(["docx", "inspect", "--file", input], io);
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text());
    expect(parsed.format).toBe("docx");
    expect(parsed.paragraphs).toBe(3);
    expect(parsed.comments).toBe(0);
    expect(Array.isArray(parsed.parts)).toBe(true);
  });

  it("docx read --format markdown emits markdown by default", async () => {
    const input = await makeFixture();
    const { io, stdout } = makeIO();
    const code = await runCli(["docx", "read", "--file", input, "--format", "markdown"], io);
    expect(code).toBe(0);
    expect(stdout.text()).toContain("# Hello");
  });

  it("docx read --format json emits a paragraph projection", async () => {
    const input = await makeFixture();
    const { io, stdout } = makeIO();
    const code = await runCli(["docx", "read", "--file", input, "--format", "json"], io);
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text());
    expect(parsed.format).toBe("docx");
    expect(parsed.paragraphs).toHaveLength(3);
    expect(parsed.paragraphs[1].text).toBe("first paragraph body");
  });

  it("docx read --format text emits plain text", async () => {
    const input = await makeFixture();
    const { io, stdout } = makeIO();
    const code = await runCli(["docx", "read", "--file", input, "--format", "text"], io);
    expect(code).toBe(0);
    const text = stdout.text();
    expect(text).toContain("Hello");
    expect(text).toContain("first paragraph body");
    expect(text).not.toContain("#");
  });

  it("docx write accepts a section:N/paragraph:M selector", async () => {
    const input = await makeFixture();
    const dir = mkdtempSync(join(tmpdir(), "office-agent-docx-write-"));
    const output = join(dir, "out.docx");
    const { io } = makeIO();
    const code = await runCli(
      [
        "docx",
        "write",
        "--file",
        input,
        "--out",
        output,
        "--at",
        "section:0/paragraph:1/run:0/text:0",
        "--text",
        "EDIT — ",
      ],
      io
    );
    expect(code).toBe(0);
    const agent = await DocxAgent.fromBuffer(readFileSync(output));
    expect(agent.toMarkdown()).toContain("EDIT — first paragraph body");
  });

  it("docx style applies a paragraph style", async () => {
    const input = await makeFixture();
    const dir = mkdtempSync(join(tmpdir(), "office-agent-docx-style-"));
    const output = join(dir, "out.docx");
    const { io } = makeIO();
    const code = await runCli(
      ["docx", "style", "--file", input, "--out", output, "--at", "paragraph:2", "--style", "Heading2"],
      io
    );
    expect(code).toBe(0);
    const agent = await DocxAgent.fromBuffer(readFileSync(output));
    const p2 = agent.getSnapshot().root.body[2];
    if (p2.kind !== "paragraph") throw new Error("expected paragraph at index 2");
    expect(p2.properties.styleId).toBe("Heading2");
  });

  it("docx comment + resolve-comment + reply-comment + delete-comment lifecycle", async () => {
    const input = await makeFixture();
    const dir = mkdtempSync(join(tmpdir(), "office-agent-docx-cl-"));
    const step1 = join(dir, "step1.docx");
    const step2 = join(dir, "step2.docx");
    const step3 = join(dir, "step3.docx");
    const step4 = join(dir, "step4.docx");
    const { io, stdout } = makeIO();

    let code = await runCli(
      [
        "docx",
        "comment",
        "--file",
        input,
        "--out",
        step1,
        "--range",
        "paragraph:1/text:0..5",
        "--text",
        "rephrase?",
        "--author",
        "Tester",
      ],
      io
    );
    expect(code).toBe(0);

    let agent = await DocxAgent.fromBuffer(readFileSync(step1));
    expect(agent.getSnapshot().root.comments).toHaveLength(1);
    const commentId = agent.getSnapshot().root.comments[0].id;

    const r1 = makeIO();
    code = await runCli(
      ["docx", "resolve-comment", "--file", step1, "--out", step2, "--id", commentId],
      r1.io
    );
    expect(code).toBe(0);
    agent = await DocxAgent.fromBuffer(readFileSync(step2));
    expect(agent.getSnapshot().root.comments[0].resolved).toBe(true);

    const r2 = makeIO();
    code = await runCli(
      [
        "docx",
        "reply-comment",
        "--file",
        step2,
        "--out",
        step3,
        "--parent",
        commentId,
        "--text",
        "ack",
        "--author",
        "Bob",
      ],
      r2.io
    );
    expect(code).toBe(0);
    agent = await DocxAgent.fromBuffer(readFileSync(step3));
    expect(agent.getSnapshot().root.comments).toHaveLength(2);
    expect(agent.getSnapshot().root.comments[1].parentId).toBe(commentId);

    const r3 = makeIO();
    code = await runCli(
      ["docx", "delete-comment", "--file", step3, "--out", step4, "--id", commentId],
      r3.io
    );
    expect(code).toBe(0);
    agent = await DocxAgent.fromBuffer(readFileSync(step4));
    expect(agent.getSnapshot().root.comments).toHaveLength(0);

    void stdout;
  });

  it("docx format-range applies bold to a text range", async () => {
    const input = await makeFixture();
    const dir = mkdtempSync(join(tmpdir(), "office-agent-format-"));
    const output = join(dir, "out.docx");
    const { io } = makeIO();
    const code = await runCli(
      [
        "docx",
        "format-range",
        "--file",
        input,
        "--out",
        output,
        "--range",
        "paragraph:1/text:0..5",
        "--bold",
        "true",
      ],
      io
    );
    expect(code).toBe(0);
    const agent = await DocxAgent.fromBuffer(readFileSync(output));
    const p1 = agent.getSnapshot().root.body[1];
    if (p1.kind !== "paragraph") throw new Error("expected paragraph at index 1");
    const firstRun = p1.children.find((c) => c.kind === "run");
    if (!firstRun || firstRun.kind !== "run") throw new Error("expected a run");
    expect(firstRun.properties.bold).toBe(true);
  });

  it("docx format-range without any flag exits with usage error", async () => {
    const input = await makeFixture();
    const { io, stderr } = makeIO();
    const code = await runCli(
      ["docx", "format-range", "--file", input, "--range", "paragraph:1/text:0..5"],
      io
    );
    expect(code).toBe(64);
    expect(stderr.text()).toMatch(/at least one of/i);
  });

  it("docx insert-image inserts an inline image and adds a media part", async () => {
    const input = await makeFixture();
    const dir = mkdtempSync(join(tmpdir(), "office-agent-image-"));
    const output = join(dir, "out.docx");
    const imagePath = join(dir, "tiny.png");
    writeFileSync(imagePath, TINY_PNG_BYTES);
    const { io } = makeIO();
    const code = await runCli(
      [
        "docx",
        "insert-image",
        "--file",
        input,
        "--out",
        output,
        "--at",
        "paragraph:1",
        "--image",
        imagePath,
        "--width",
        "64",
        "--height",
        "64",
        "--alt",
        "tiny",
      ],
      io
    );
    expect(code).toBe(0);
    const reloaded = readFileSync(output);
    const z = await JSZip.loadAsync(reloaded);
    const mediaParts = Object.keys(z.files).filter((p) => p.startsWith("word/media/"));
    expect(mediaParts.length).toBeGreaterThanOrEqual(1);
  });

  it("docx insert-table inserts a 2×2 table at the position selector", async () => {
    const input = await makeFixture();
    const dir = mkdtempSync(join(tmpdir(), "office-agent-table-"));
    const output = join(dir, "out.docx");
    const { io } = makeIO();
    const code = await runCli(
      [
        "docx",
        "insert-table",
        "--file",
        input,
        "--out",
        output,
        "--at",
        "paragraph:0",
        "--rows",
        "2",
        "--cols",
        "3",
      ],
      io
    );
    expect(code).toBe(0);
    const agent = await DocxAgent.fromBuffer(readFileSync(output));
    const tables = agent.getSnapshot().root.body.filter((b) => b.kind === "table");
    expect(tables).toHaveLength(1);
    if (tables[0].kind !== "table") throw new Error("expected table");
    expect(tables[0].rows).toHaveLength(2);
    expect(tables[0].rows[0].cells).toHaveLength(3);
  });

  // `docx set-cell-text` is exercised at the handler level in
  // packages/docx/src/commands/tables.test.ts. The CLI shim cannot be
  // covered with a multi-step round-trip here because the parser mints
  // fresh `tableId`s on every load — there's no stable way to thread the
  // id between two CLI invocations until we either expose `docx inspect
  // --tables` or wire a deterministic id minter into the CLI.

  it("docx write --no-approve emits a pending mutation summary", async () => {
    const input = await makeFixture();
    const dir = mkdtempSync(join(tmpdir(), "office-agent-pending-"));
    const output = join(dir, "out.docx");
    const { io, stdout } = makeIO();
    const code = await runCli(
      [
        "docx",
        "write",
        "--file",
        input,
        "--out",
        output,
        "--at",
        "paragraph:1/run:0/text:0",
        "--text",
        "PENDING ",
        "--no-approve",
      ],
      io
    );
    expect(code).toBe(0);
    const summary = JSON.parse(stdout.text());
    expect(summary.mutation.status).toBe("pending");
    // The on-disk file is still written from the working snapshot, so the
    // edit lands in bytes — `--no-approve` only changes the bus state and
    // the printed summary status.
    const agent = await DocxAgent.fromBuffer(readFileSync(output));
    expect(agent.toMarkdown()).toContain("PENDING first paragraph body");
  });

  it("docx accept-change folds an inserted run into its parent paragraph", async () => {
    const input = await makeTrackedChangesFixture();
    const dir = mkdtempSync(join(tmpdir(), "office-agent-accept-"));
    const output = join(dir, "out.docx");
    const { io } = makeIO();
    const code = await runCli(["docx", "accept-change", "--file", input, "--out", output, "--id", "42"], io);
    expect(code).toBe(0);
    const agent = await DocxAgent.fromBuffer(readFileSync(output));
    const p0 = agent.getSnapshot().root.body[0];
    if (p0.kind !== "paragraph") throw new Error("expected paragraph");
    // After accept, no revision wrapper remains
    const wrappers = p0.children.filter((c) => c.kind === "revision");
    expect(wrappers).toHaveLength(0);
  });

  it("docx reject-change drops an inserted run entirely", async () => {
    const input = await makeTrackedChangesFixture();
    const dir = mkdtempSync(join(tmpdir(), "office-agent-reject-"));
    const output = join(dir, "out.docx");
    const { io } = makeIO();
    const code = await runCli(["docx", "reject-change", "--file", input, "--out", output, "--id", "42"], io);
    expect(code).toBe(0);
    const agent = await DocxAgent.fromBuffer(readFileSync(output));
    const p0 = agent.getSnapshot().root.body[0];
    if (p0.kind !== "paragraph") throw new Error("expected paragraph");
    const text = p0.children
      .filter((c) => c.kind === "run")
      .flatMap((c) => (c.kind === "run" ? c.children : []))
      .map((c) => (c.kind === "text" ? c.text : ""))
      .join("");
    expect(text).toBe("Original ");
    expect(text).not.toContain("INSERTED");
  });

  it("docx diff reports paragraph modifications", async () => {
    const input = await makeFixture();
    const dir = mkdtempSync(join(tmpdir(), "office-agent-docx-diff-"));
    const after = join(dir, "after.docx");
    const writeIo = makeIO();
    let code = await runCli(
      [
        "docx",
        "write",
        "--file",
        input,
        "--out",
        after,
        "--at",
        "paragraph:1/run:0/text:0",
        "--text",
        "DIFF ",
      ],
      writeIo.io
    );
    expect(code).toBe(0);

    const diffIo = makeIO();
    code = await runCli(["docx", "diff", "--before", input, "--after", after], diffIo.io);
    expect(code).toBe(0);
    const parsed = JSON.parse(diffIo.stdout.text());
    expect(parsed.paragraphs.modified).toBe(1);
  });
});

describe("office-agent xlsx subcommand group", () => {
  it("xlsx inspect prints structural summary as JSON", async () => {
    const dir = mkdtempSync(join(tmpdir(), "office-agent-xlsx-inspect-"));
    const input = copyFixture("02-multi-sheet.xlsx", dir);
    const { io, stdout } = makeIO();
    const code = await runCli(["xlsx", "inspect", "--file", input], io);
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text());
    expect(parsed.format).toBe("xlsx");
    expect(Array.isArray(parsed.sheets)).toBe(true);
    expect(parsed.sheets.length).toBeGreaterThan(1);
    expect(Array.isArray(parsed.parts)).toBe(true);
    expect(parsed.parts.some((p: string) => p.includes("workbook.xml"))).toBe(true);
  });

  it("xlsx read --format json projects the first worksheet's bounding box", async () => {
    const dir = mkdtempSync(join(tmpdir(), "office-agent-xlsx-read-"));
    const input = copyFixture("01-single-sheet-numbers.xlsx", dir);
    const { io, stdout } = makeIO();
    const code = await runCli(["xlsx", "read", "--file", input, "--format", "json"], io);
    expect(code).toBe(0);
    const parsed = JSON.parse(stdout.text());
    expect(parsed.format).toBe("xlsx");
    expect(typeof parsed.sheet).toBe("string");
    expect(Array.isArray(parsed.cells)).toBe(true);
    expect(parsed.cells.length).toBeGreaterThan(0);
  });

  it("xlsx read --format markdown emits a sheet header", async () => {
    const dir = mkdtempSync(join(tmpdir(), "office-agent-xlsx-md-"));
    const input = copyFixture("01-single-sheet-numbers.xlsx", dir);
    const { io, stdout } = makeIO();
    const code = await runCli(["xlsx", "read", "--file", input, "--format", "markdown"], io);
    expect(code).toBe(0);
    expect(stdout.text()).toMatch(/^## /m);
  });

  it("xlsx set-cell round-trips a literal value to disk", async () => {
    const dir = mkdtempSync(join(tmpdir(), "office-agent-xlsx-setcell-"));
    const input = copyFixture("01-single-sheet-numbers.xlsx", dir);
    const out = join(dir, "out.xlsx");
    const baseline = await XlsxAgent.fromBuffer(readFileSync(input));
    const sheetName = baseline.listSheets()[0]!.name;
    const { io, stdout } = makeIO();
    const code = await runCli(
      [
        "xlsx",
        "set-cell",
        "--file",
        input,
        "--out",
        out,
        "--sheet",
        sheetName,
        "--ref",
        "AA50",
        "--value",
        '"agent-edit"',
      ],
      io
    );
    expect(code).toBe(0);
    const summary = JSON.parse(stdout.text());
    expect(summary.wrote).toBe(out);
    expect(summary.mutation.status).toBe("approved");

    const reloaded = await XlsxAgent.fromBuffer(readFileSync(out));
    const sheet = reloaded.getSnapshot().root.sheets.find((s) => s.name === sheetName)!;
    const matched = [...sheet.cells.values()].find((c) => c.value === "agent-edit");
    expect(matched).toBeDefined();
    expect(matched!.row).toBe(49);
    expect(matched!.col).toBe(26);
  });

  it("xlsx add-sheet appends a worksheet that re-loads", async () => {
    const dir = mkdtempSync(join(tmpdir(), "office-agent-xlsx-add-"));
    const input = copyFixture("01-single-sheet-numbers.xlsx", dir);
    const out = join(dir, "out.xlsx");
    const { io } = makeIO();
    const code = await runCli(
      ["xlsx", "add-sheet", "--file", input, "--out", out, "--name", "Brand new sheet"],
      io
    );
    expect(code).toBe(0);
    const reloaded = await XlsxAgent.fromBuffer(readFileSync(out));
    const sheets = reloaded.listSheets().map((s) => s.name);
    expect(sheets).toContain("Brand new sheet");
  });

  it("xlsx diff between two on-disk files reports the cell change", async () => {
    const dir = mkdtempSync(join(tmpdir(), "office-agent-xlsx-diff-"));
    const before = copyFixture("01-single-sheet-numbers.xlsx", dir);
    const after = join(dir, "after.xlsx");
    const baseline = await XlsxAgent.fromBuffer(readFileSync(before));
    const sheetName = baseline.listSheets()[0]!.name;

    const writeIo = makeIO();
    let code = await runCli(
      [
        "xlsx",
        "set-cell",
        "--file",
        before,
        "--out",
        after,
        "--sheet",
        sheetName,
        "--ref",
        "Z99",
        "--value",
        "42",
      ],
      writeIo.io
    );
    expect(code).toBe(0);

    const diffIo = makeIO();
    code = await runCli(["xlsx", "diff", "--before", before, "--after", after], diffIo.io);
    expect(code).toBe(0);
    const diff = JSON.parse(diffIo.stdout.text());
    expect(diff.format).toBe("xlsx");
    expect(diff.changes.length).toBeGreaterThanOrEqual(1);
    const summaries = diff.changes.map((c: { summary?: string }) => c.summary ?? "");
    expect(summaries.some((s: string) => s.includes("Z99"))).toBe(true);
  });

  it("xlsx set-cell --no-approve reports a pending mutation summary", async () => {
    const dir = mkdtempSync(join(tmpdir(), "office-agent-xlsx-pending-"));
    const input = copyFixture("01-single-sheet-numbers.xlsx", dir);
    const out = join(dir, "out.xlsx");
    const baseline = await XlsxAgent.fromBuffer(readFileSync(input));
    const sheetName = baseline.listSheets()[0]!.name;
    const { io, stdout } = makeIO();
    const code = await runCli(
      [
        "xlsx",
        "set-cell",
        "--file",
        input,
        "--out",
        out,
        "--sheet",
        sheetName,
        "--ref",
        "AB1",
        "--value",
        "7",
        "--no-approve",
      ],
      io
    );
    expect(code).toBe(0);
    const summary = JSON.parse(stdout.text());
    expect(summary.mutation.status).toBe("pending");
  });
});
