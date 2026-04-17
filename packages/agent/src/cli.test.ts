import { mkdtempSync, writeFileSync, readFileSync, mkdirSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { Document, Packer, Paragraph, TextRun, HeadingLevel } from "docx";
import { runCli } from "./cli.js";
import { DocxAgent } from "@officeai/docx";

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

  it("xlsx subcommand reports unimplemented with non-zero exit", async () => {
    const { io, stderr } = makeIO();
    const code = await runCli(["xlsx"], io);
    expect(code).toBe(2);
    expect(stderr.text()).toContain("XLSX support is not yet implemented");
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
