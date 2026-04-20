/**
 * Aggregate barrel for `@officeai/react-editors/blanks`.
 *
 * Hosts that want a single import for all four blank-file builders
 * pull from here; hosts targeting one specific format should prefer
 * the per-format subpaths (`@officeai/react-editors/blanks/docx`
 * etc) so unused formats tree-shake out cleanly.
 */
import { makeBlankDocx } from "./docx.js";
import { makeBlankXlsx } from "./xlsx.js";
import { makeBlankPptx } from "./pptx.js";
import { makeBlankPdf } from "./pdf.js";
import type { OfficeFormat } from "../mime.js";

export { makeBlankDocx, makeBlankXlsx, makeBlankPptx, makeBlankPdf };

/**
 * Convenience dispatcher: pick the right blank-file builder for a
 * given OfficeFormat. Hosts use this when the format is only known
 * at runtime (e.g. a "Create new" dropdown that hands a string into
 * a single create handler).
 */
export async function makeBlank(format: OfficeFormat): Promise<Uint8Array> {
  switch (format) {
    case "docx":
      return makeBlankDocx();
    case "xlsx":
      return makeBlankXlsx();
    case "pptx":
      return makeBlankPptx();
    case "pdf":
      return makeBlankPdf();
    default: {
      const exhaustive: never = format;
      throw new Error(`Unknown OfficeFormat: ${String(exhaustive)}`);
    }
  }
}
