/**
 * @officeai/pdf-forms — AcroForm read / fill / flatten for PDFs.
 *
 * Spec: /spec/pdf/form-engine.md
 *
 * Backed by pdf-lib's PDFForm. JavaScript actions in field calc-order
 * are intentionally NOT executed (sandboxed; no JS interpreter is
 * shipped). Signature widgets are detected and reported but not signed.
 */
export { listFormFields, type ListedFormField } from "./list.js";
export { fillForm, type FillFormOptions } from "./fill.js";
export { flattenForm } from "./flatten.js";
export { resetForm } from "./reset.js";
