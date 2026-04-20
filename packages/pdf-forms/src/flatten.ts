import { PDFDocument } from "pdf-lib";

/**
 * Flatten the form so widgets become page content (non-editable).
 * Useful before sharing a filled-in PDF as the final answer.
 */
export const flattenForm = async (buffer: Uint8Array): Promise<Uint8Array> => {
  const pdf = await PDFDocument.load(buffer, { updateMetadata: false });
  const form = pdf.getForm();
  form.flatten();
  return pdf.save();
};
