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
