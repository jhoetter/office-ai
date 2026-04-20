/**
 * In-browser blank-XLSX builder for embedding hosts.
 *
 * See `./docx.ts` for the rationale.
 */
import { XlsxAgent } from "@officeai/xlsx";

export async function makeBlankXlsx(): Promise<Uint8Array> {
  const agent = await XlsxAgent.empty();
  const buf = await agent.exportFile();
  return new Uint8Array(buf);
}
