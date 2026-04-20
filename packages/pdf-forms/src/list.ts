import {
  PDFCheckBox,
  PDFDocument,
  PDFDropdown,
  PDFOptionList,
  PDFRadioGroup,
  PDFSignature,
  PDFTextField,
} from "pdf-lib";

export interface ListedFormField {
  readonly name: string;
  readonly type: "text" | "checkbox" | "radio" | "dropdown" | "option-list" | "signature" | "unknown";
  readonly value?: string | boolean | ReadonlyArray<string>;
  readonly options?: ReadonlyArray<string>;
  readonly readOnly: boolean;
  readonly required: boolean;
  readonly maxLength?: number;
}

export const listFormFields = async (buffer: Uint8Array): Promise<ReadonlyArray<ListedFormField>> => {
  const pdf = await PDFDocument.load(buffer, { updateMetadata: false });
  const form = pdf.getForm();
  return form.getFields().map((field) => {
    const name = field.getName();
    const readOnly = field.isReadOnly();
    const required = field.isRequired();
    if (field instanceof PDFTextField) {
      const out: ListedFormField = {
        name,
        type: "text",
        readOnly,
        required,
        ...(field.getText() !== undefined ? { value: field.getText() ?? "" } : {}),
        ...(field.getMaxLength() !== undefined ? { maxLength: field.getMaxLength() } : {}),
      };
      return out;
    }
    if (field instanceof PDFCheckBox) {
      return { name, type: "checkbox", value: field.isChecked(), readOnly, required };
    }
    if (field instanceof PDFRadioGroup) {
      return {
        name,
        type: "radio",
        value: field.getSelected() ?? "",
        options: field.getOptions(),
        readOnly,
        required,
      };
    }
    if (field instanceof PDFDropdown) {
      return {
        name,
        type: "dropdown",
        value: field.getSelected(),
        options: field.getOptions(),
        readOnly,
        required,
      };
    }
    if (field instanceof PDFOptionList) {
      return {
        name,
        type: "option-list",
        value: field.getSelected(),
        options: field.getOptions(),
        readOnly,
        required,
      };
    }
    if (field instanceof PDFSignature) {
      return { name, type: "signature", readOnly, required };
    }
    return { name, type: "unknown", readOnly, required };
  });
};
