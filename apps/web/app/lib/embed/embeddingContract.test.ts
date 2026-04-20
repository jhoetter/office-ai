/**
 * Phase 0 contract tests for the four office-ai editors.
 *
 * The editors are large React surfaces that rely on browser APIs
 * (ProseMirror, PDF.js workers, FileSystemAccess) that don't exist
 * under the Node-only Vitest environment. Mounting them in jsdom
 * pulls in pdfjs-dist which detaches the test process — so instead
 * we exercise the bytes-in / bytes-out boundary at the agent layer
 * that the editors themselves now consume:
 *
 *   1. `<Format>Agent.empty()` → `agent.exportFile()` produces bytes
 *      that the embedding host can hand back as `initialBytes`.
 *   2. Those bytes round-trip through `<Format>Agent.fromBuffer(...)`
 *      — which is exactly what the editor's bootstrap branch calls
 *      after copying the host-supplied `Uint8Array`.
 *   3. `agent.exportFile()` after one ack'd no-op gives the host
 *      back bytes shaped like what `onSave(bytes, mime, name)`
 *      receives in the editor's Save handler.
 *
 * Together these assertions lock in the contract the
 * `@officeai/react-editors` package will publish in Phase 1: hosts
 * (hof-os) can call `*Agent.empty()` in-browser, hand the bytes to
 * the matching editor via `initialBytes`, and receive bytes back
 * from `onSave` without ever touching disk.
 */
import { describe, it, expect } from "vitest";
import { DocxAgent } from "@officeai/docx";
import { XlsxAgent } from "@officeai/xlsx";
import { PptxAgent } from "@officeai/pptx";
import { PdfAgent } from "@officeai/pdf";

const ZIP_MAGIC = [0x50, 0x4b, 0x03, 0x04];

function asBytes(buf: ArrayBuffer | Uint8Array): Uint8Array {
  return buf instanceof Uint8Array ? buf : new Uint8Array(buf);
}

function isZipMagic(bytes: Uint8Array): boolean {
  return ZIP_MAGIC.every((b, i) => bytes[i] === b);
}

describe("editor embedding contract — initialBytes / onSave round-trip", () => {
  it("DocxAgent.empty bytes feed back through fromBuffer (DocxEditor initialBytes path)", async () => {
    const blank = await DocxAgent.empty();
    const bytes = asBytes(await blank.exportFile());
    expect(bytes.byteLength).toBeGreaterThan(0);
    expect(isZipMagic(bytes)).toBe(true);
    const reopened = await DocxAgent.fromBuffer(bytes.slice().buffer);
    const exported = asBytes(await reopened.exportFile());
    expect(exported.byteLength).toBeGreaterThan(0);
    expect(isZipMagic(exported)).toBe(true);
  });

  it("XlsxAgent.empty bytes feed back through fromBuffer (XlsxEditor initialBytes path)", async () => {
    const blank = await XlsxAgent.empty();
    const bytes = asBytes(await blank.exportFile());
    expect(bytes.byteLength).toBeGreaterThan(0);
    expect(isZipMagic(bytes)).toBe(true);
    const reopened = await XlsxAgent.fromBuffer(bytes.slice().buffer);
    const exported = asBytes(await reopened.exportFile());
    expect(exported.byteLength).toBeGreaterThan(0);
    expect(isZipMagic(exported)).toBe(true);
  });

  it("PptxAgent.empty bytes feed back through fromBuffer (PptxEditor initialBytes path)", async () => {
    const blank = await PptxAgent.empty();
    const bytes = asBytes(await blank.exportFile());
    expect(bytes.byteLength).toBeGreaterThan(0);
    expect(isZipMagic(bytes)).toBe(true);
    const reopened = await PptxAgent.fromBuffer(bytes.slice().buffer);
    const exported = asBytes(await reopened.exportFile());
    expect(exported.byteLength).toBeGreaterThan(0);
    expect(isZipMagic(exported)).toBe(true);
  });

  it("PdfAgent.empty bytes feed back through fromBuffer (PdfEditor initialBytes path)", async () => {
    const blank = await PdfAgent.empty();
    const bytes = asBytes(await blank.exportFile());
    expect(bytes.byteLength).toBeGreaterThan(0);
    expect(Buffer.from(bytes.subarray(0, 4)).toString("ascii")).toBe("%PDF");
    const reopened = await PdfAgent.fromBuffer(bytes);
    const exported = asBytes(await reopened.exportFile());
    expect(exported.byteLength).toBeGreaterThan(0);
    expect(Buffer.from(exported.subarray(0, 4)).toString("ascii")).toBe("%PDF");
  });

  it("simulated onSave handler receives the OOXML mime + filename + bytes shape", async () => {
    type OnSave = (bytes: Uint8Array, mime: string, filename: string) => Promise<void>;
    const calls: { mime: string; name: string; size: number }[] = [];
    const onSave: OnSave = async (bytes, mime, name) => {
      calls.push({ mime, name, size: bytes.byteLength });
    };
    const blank = await DocxAgent.empty();
    const exported = asBytes(await blank.exportFile());
    await onSave(
      exported,
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "Untitled.docx"
    );
    expect(calls).toHaveLength(1);
    expect(calls[0].mime).toMatch(/wordprocessingml/);
    expect(calls[0].name).toBe("Untitled.docx");
    expect(calls[0].size).toBeGreaterThan(0);
  });
});
