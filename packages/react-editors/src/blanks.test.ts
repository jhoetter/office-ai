/**
 * Smoke tests for `@officeai/react-editors/blanks`.
 *
 * The intent is to lock in the bytes-out contract embedding hosts
 * (hof-os) rely on for the "Create new" dropdown — each builder
 * returns a non-empty Uint8Array with the expected file-format
 * magic bytes, and the runtime dispatcher routes correctly.
 */
import { describe, it, expect } from "vitest";
import { makeBlankDocx, makeBlankXlsx, makeBlankPptx, makeBlankPdf, makeBlank } from "./blanks/index.js";
import {
  DOCX_MIME,
  XLSX_MIME,
  PPTX_MIME,
  PDF_MIME,
  DEFAULT_BLANK_FILENAME,
  MIME_BY_FORMAT,
  detectFormatFromFilename,
} from "./mime.js";

const ZIP_MAGIC = [0x50, 0x4b, 0x03, 0x04];
function isZip(bytes: Uint8Array): boolean {
  return ZIP_MAGIC.every((b, i) => bytes[i] === b);
}

describe("blank-file builders", () => {
  it("makeBlankDocx produces a non-empty OOXML zip", async () => {
    const bytes = await makeBlankDocx();
    expect(bytes).toBeInstanceOf(Uint8Array);
    expect(bytes.byteLength).toBeGreaterThan(0);
    expect(isZip(bytes)).toBe(true);
  });

  it("makeBlankXlsx produces a non-empty OOXML zip", async () => {
    const bytes = await makeBlankXlsx();
    expect(bytes.byteLength).toBeGreaterThan(0);
    expect(isZip(bytes)).toBe(true);
  });

  it("makeBlankPptx produces a non-empty OOXML zip", async () => {
    const bytes = await makeBlankPptx();
    expect(bytes.byteLength).toBeGreaterThan(0);
    expect(isZip(bytes)).toBe(true);
  });

  it("makeBlankPdf produces a non-empty PDF stream", async () => {
    const bytes = await makeBlankPdf();
    expect(bytes.byteLength).toBeGreaterThan(0);
    expect(Buffer.from(bytes.subarray(0, 4)).toString("ascii")).toBe("%PDF");
  });

  it("makeBlank dispatches to the right builder for every OfficeFormat", async () => {
    const docx = await makeBlank("docx");
    const xlsx = await makeBlank("xlsx");
    const pptx = await makeBlank("pptx");
    const pdf = await makeBlank("pdf");
    expect(isZip(docx)).toBe(true);
    expect(isZip(xlsx)).toBe(true);
    expect(isZip(pptx)).toBe(true);
    expect(Buffer.from(pdf.subarray(0, 4)).toString("ascii")).toBe("%PDF");
  });
});

describe("mime constants + helpers", () => {
  it("MIME_BY_FORMAT covers every OfficeFormat with the canonical IANA strings", () => {
    expect(MIME_BY_FORMAT.docx).toBe(DOCX_MIME);
    expect(MIME_BY_FORMAT.xlsx).toBe(XLSX_MIME);
    expect(MIME_BY_FORMAT.pptx).toBe(PPTX_MIME);
    expect(MIME_BY_FORMAT.pdf).toBe(PDF_MIME);
  });

  it("DEFAULT_BLANK_FILENAME pairs each format with a sensible default name", () => {
    expect(DEFAULT_BLANK_FILENAME.docx).toBe("Untitled.docx");
    expect(DEFAULT_BLANK_FILENAME.xlsx).toBe("Untitled.xlsx");
    expect(DEFAULT_BLANK_FILENAME.pptx).toBe("Untitled.pptx");
    expect(DEFAULT_BLANK_FILENAME.pdf).toBe("Untitled.pdf");
  });

  it("detectFormatFromFilename is case-insensitive and rejects unknown extensions", () => {
    expect(detectFormatFromFilename("expenses.xlsx")).toBe("xlsx");
    expect(detectFormatFromFilename("Plan.DOCX")).toBe("docx");
    expect(detectFormatFromFilename("deck.PPTX")).toBe("pptx");
    expect(detectFormatFromFilename("scan.pdf")).toBe("pdf");
    expect(detectFormatFromFilename("notes.txt")).toBe(null);
    expect(detectFormatFromFilename("no-extension")).toBe(null);
  });
});
