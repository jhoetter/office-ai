/**
 * In-browser blank-PPTX builder for embedding hosts.
 *
 * See `./docx.ts` for the rationale.
 */
import { PptxAgent } from "@officeai/pptx";

export async function makeBlankPptx(): Promise<Uint8Array> {
  const agent = await PptxAgent.empty();
  const buf = await agent.exportFile();
  return new Uint8Array(buf);
}
