import { PDFCheckBox, PDFDocument, PDFDropdown, PDFOptionList, PDFRadioGroup, PDFTextField } from "pdf-lib";

/** Reset every field to its default. */
export const resetForm = async (buffer: Uint8Array): Promise<Uint8Array> => {
  const pdf = await PDFDocument.load(buffer, { updateMetadata: false });
  const form = pdf.getForm();
  for (const field of form.getFields()) {
    if (field instanceof PDFTextField) {
      field.setText("");
    } else if (field instanceof PDFCheckBox) {
      field.uncheck();
    } else if (field instanceof PDFRadioGroup) {
      field.clear();
    } else if (field instanceof PDFDropdown || field instanceof PDFOptionList) {
      field.clear();
    }
  }
  return pdf.save();
};
