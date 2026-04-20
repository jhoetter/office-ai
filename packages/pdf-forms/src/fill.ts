import { PDFCheckBox, PDFDocument, PDFDropdown, PDFOptionList, PDFRadioGroup, PDFTextField } from "pdf-lib";

export interface FillFormOptions {
  readonly values: Readonly<Record<string, string | boolean | ReadonlyArray<string>>>;
  /** When true, flatten the form on save (no longer fillable). */
  readonly flatten?: boolean;
}

export const fillForm = async (buffer: Uint8Array, opts: FillFormOptions): Promise<Uint8Array> => {
  const pdf = await PDFDocument.load(buffer, { updateMetadata: false });
  const form = pdf.getForm();

  for (const [name, value] of Object.entries(opts.values)) {
    let field;
    try {
      field = form.getField(name);
    } catch {
      throw new Error(`pdf-forms/fill: field "${name}" not found`);
    }
    if (field instanceof PDFTextField) {
      if (typeof value !== "string") {
        throw new Error(`pdf-forms/fill: field "${name}" expects a string`);
      }
      const max = field.getMaxLength();
      if (typeof max === "number" && value.length > max) {
        throw new Error(`pdf-forms/fill: field "${name}" exceeds MaxLen ${max}`);
      }
      field.setText(value);
    } else if (field instanceof PDFCheckBox) {
      if (typeof value !== "boolean") {
        throw new Error(`pdf-forms/fill: field "${name}" expects a boolean`);
      }
      if (value) field.check();
      else field.uncheck();
    } else if (field instanceof PDFRadioGroup) {
      if (typeof value !== "string") {
        throw new Error(`pdf-forms/fill: field "${name}" expects a string`);
      }
      field.select(value);
    } else if (field instanceof PDFDropdown) {
      if (typeof value !== "string") {
        throw new Error(`pdf-forms/fill: field "${name}" expects a string`);
      }
      field.select(value);
    } else if (field instanceof PDFOptionList) {
      if (!Array.isArray(value) && typeof value !== "string") {
        throw new Error(`pdf-forms/fill: field "${name}" expects string or string[]`);
      }
      const selections = Array.isArray(value) ? value : [value as string];
      field.select(selections);
    } else {
      throw new Error(`pdf-forms/fill: unsupported field type for "${name}"`);
    }
  }

  if (opts.flatten === true) form.flatten();
  return pdf.save();
};
