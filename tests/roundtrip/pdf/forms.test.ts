import { describe, expect, it } from "vitest";
import { PdfAgent } from "@officeai/pdf";
import { fillForm, flattenForm, listFormFields, resetForm } from "@officeai/pdf-forms";
import { isPdfBytes, loadFixture } from "./helpers.js";

/**
 * Form fixtures: blank + pre-filled AcroForm pairs. Tests verify the
 * widget-kind enumeration, end-to-end fill via pdf-forms, flatten
 * effect, reset effect, and that the signed-then-modified fixture
 * surfaces a non-zero `signatureCount` even though no PKCS#7 blob is
 * present (it only has a /Sig field + SigFlags).
 */

describe("PDF roundtrip — forms", () => {
  it("acroform-basic.pdf enumerates every widget, all empty", async () => {
    const bytes = await loadFixture("acroform-basic.pdf");
    const agent = await PdfAgent.fromBuffer(bytes);
    const fields = agent.getSnapshot().root.formFields;
    // Snapshot projection is per-widget. The fixture defines:
    //   text(first.name) + checkbox(agree) + dropdown(country) +
    //   2 radio widgets (plan: free, pro). PDF.js can't tell radio
    //   widgets apart from checkbox widgets at the annotation level
    //   (both surface as Btn), so we assert by name + presence.
    expect(fields).toHaveLength(5);
    const names = fields.map((f) => f.name).sort();
    expect(names).toEqual(["agree", "country", "first.name", "plan", "plan"]);
    expect(fields.find((f) => f.name === "first.name")?.type).toBe("text");
    expect(fields.find((f) => f.name === "country")?.type).toBe("choice");
    expect(agent.getSnapshot().root.signatureCount).toBe(0);

    const listed = await listFormFields(bytes);
    const byName = new Map(listed.map((f) => [f.name, f] as const));
    expect(byName.get("first.name")?.type).toBe("text");
    expect(byName.get("first.name")?.value ?? "").toBe("");
    expect(byName.get("agree")?.type).toBe("checkbox");
    expect(byName.get("agree")?.value).toBe(false);
    expect(byName.get("country")?.type).toBe("dropdown");
    expect(byName.get("country")?.value).toEqual([]);
    expect(byName.get("plan")?.type).toBe("radio");
    expect(byName.get("plan")?.value).toBe("");
  });

  it("acroform-prefilled.pdf carries the seeded values", async () => {
    const bytes = await loadFixture("acroform-prefilled.pdf");
    const listed = await listFormFields(bytes);
    const byName = new Map(listed.map((f) => [f.name, f] as const));
    expect(byName.get("first.name")?.value).toBe("Ada");
    expect(byName.get("agree")?.value).toBe(true);
    // Dropdown values come back as string[] from pdf-lib's getSelected().
    expect(byName.get("country")?.value).toEqual(["DE"]);
    expect(byName.get("plan")?.value).toBe("pro");
  });

  it("fillForm + re-parse persists every field type", async () => {
    const bytes = await loadFixture("acroform-basic.pdf");
    const out = await fillForm(bytes, {
      values: {
        "first.name": "Linus",
        agree: true,
        country: "FR",
        plan: "free",
      },
    });
    expect(isPdfBytes(out)).toBe(true);
    const listed = await listFormFields(out);
    const byName = new Map(listed.map((f) => [f.name, f] as const));
    expect(byName.get("first.name")?.value).toBe("Linus");
    expect(byName.get("agree")?.value).toBe(true);
    expect(byName.get("country")?.value).toEqual(["FR"]);
    expect(byName.get("plan")?.value).toBe("free");
  });

  it("flattenForm strips every fillable widget", async () => {
    const bytes = await loadFixture("acroform-prefilled.pdf");
    const flat = await flattenForm(bytes);
    expect(isPdfBytes(flat)).toBe(true);
    const listed = await listFormFields(flat);
    expect(listed).toHaveLength(0);
  });

  it("resetForm clears values back to defaults", async () => {
    const bytes = await loadFixture("acroform-prefilled.pdf");
    const reset = await resetForm(bytes);
    const listed = await listFormFields(reset);
    const byName = new Map(listed.map((f) => [f.name, f] as const));
    expect(byName.get("first.name")?.value ?? "").toBe("");
    expect(byName.get("agree")?.value).toBe(false);
    expect(byName.get("plan")?.value).toBe("");
  });

  it("signed-then-modified.pdf reports signatureCount > 0", async () => {
    const agent = await PdfAgent.fromBuffer(await loadFixture("signed-then-modified.pdf"));
    const root = agent.getSnapshot().root;
    expect(root.signatureCount).toBeGreaterThanOrEqual(1);
    // The signature widget also surfaces as a form field of type "signature".
    const sig = root.formFields.find((f) => f.type === "signature");
    expect(sig).toBeDefined();
  });
});
