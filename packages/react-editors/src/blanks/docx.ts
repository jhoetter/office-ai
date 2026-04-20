/**
 * In-browser blank-DOCX builder for embedding hosts.
 *
 * Wraps `DocxAgent.empty()` from `@officeai/docx` so callers don't
 * have to know about the agent layer — they just `await
 * makeBlankDocx()` and get bytes ready to upload to their own storage
 * (e.g. hof-os's presigned S3 PUT path). The result is the exact
 * same bytes the office-ai apps/web "New document" action produces.
 */
import { DocxAgent } from "@officeai/docx";

/**
 * Build a brand-new, single-empty-paragraph DOCX document and return
 * its raw bytes. Safe to call from the browser — no `node:fs` use,
 * no FileSystemAccess prompt, no network round-trip.
 */
export async function makeBlankDocx(): Promise<Uint8Array> {
  const agent = await DocxAgent.empty();
  const buf = await agent.exportFile();
  // `agent.exportFile()` is typed as ArrayBuffer for the OOXML
  // formats; normalise to Uint8Array so `new File([bytes], …)` and
  // presigned-PUT consumers don't have to special-case the type.
  return new Uint8Array(buf);
}
