/**
 * In-browser blank-PDF builder for embedding hosts.
 *
 * Wraps `PdfAgent.empty()` from `@officeai/pdf` — produces a
 * single-Letter-page blank PDF. Safe to call from the browser.
 */
import { PdfAgent } from "@officeai/pdf";

export async function makeBlankPdf(): Promise<Uint8Array> {
  const agent = await PdfAgent.empty();
  // PdfAgent.exportFile already returns Uint8Array.
  return await agent.exportFile();
}
