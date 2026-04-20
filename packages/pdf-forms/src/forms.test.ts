import { describe, expect, it } from "vitest";
import { PDFDocument, StandardFonts } from "pdf-lib";
import { listFormFields } from "./list.js";
import { fillForm } from "./fill.js";
import { flattenForm } from "./flatten.js";
import { resetForm } from "./reset.js";

const buildAcroFormPdf = async (): Promise<Uint8Array> => {
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const page = pdf.addPage([612, 792]);
  page.drawText("Form fixture", { x: 50, y: 740, size: 18, font });

  const form = pdf.getForm();
  const text = form.createTextField("first.name");
  text.setText("");
  text.addToPage(page, { x: 50, y: 700, width: 200, height: 24 });

  const cb = form.createCheckBox("agree");
  cb.addToPage(page, { x: 50, y: 660, width: 16, height: 16 });

  const dd = form.createDropdown("color");
  dd.setOptions(["red", "green", "blue"]);
  dd.addToPage(page, { x: 50, y: 620, width: 120, height: 24 });

  return pdf.save();
};

describe("pdf-forms", () => {
  it("lists fields with type and metadata", async () => {
    const buf = await buildAcroFormPdf();
    const fields = await listFormFields(buf);
    const byName = Object.fromEntries(fields.map((f) => [f.name, f]));
    expect(byName["first.name"].type).toBe("text");
    expect(byName["agree"].type).toBe("checkbox");
    expect(byName["color"].type).toBe("dropdown");
    expect(byName["color"].options).toEqual(["red", "green", "blue"]);
  });

  it("fills a text + checkbox + dropdown and round-trips", async () => {
    const buf = await buildAcroFormPdf();
    const filled = await fillForm(buf, {
      values: { "first.name": "Ada", agree: true, color: "green" },
    });
    const fields = await listFormFields(filled);
    const byName = Object.fromEntries(fields.map((f) => [f.name, f]));
    expect(byName["first.name"].value).toBe("Ada");
    expect(byName["agree"].value).toBe(true);
    expect(byName["color"].value).toBe("green");
  });

  it("flattens the form so widgets become non-fillable", async () => {
    const buf = await buildAcroFormPdf();
    const flattened = await flattenForm(buf);
    expect((await listFormFields(flattened)).length).toBe(0);
  });

  it("resets fields to defaults", async () => {
    const buf = await buildAcroFormPdf();
    const filled = await fillForm(buf, { values: { "first.name": "x", agree: true } });
    const reset = await resetForm(filled);
    const fields = await listFormFields(reset);
    const byName = Object.fromEntries(fields.map((f) => [f.name, f]));
    expect(byName["first.name"].value).toBe("");
    expect(byName["agree"].value).toBe(false);
  });

  it("rejects values for unknown fields", async () => {
    const buf = await buildAcroFormPdf();
    await expect(fillForm(buf, { values: { "no.such.field": "x" } })).rejects.toThrow(/not found/);
  });
});
