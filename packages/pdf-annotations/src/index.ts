/**
 * @officeai/pdf-annotations — typed annotation model + writers + I/O.
 *
 * Spec: /spec/pdf/annotation-model.md
 *
 * The model layer is fully implemented; the AP-stream writer ships with
 * the highest-value subset (highlight + sticky-note + free-text + link)
 * and a clean extension point for the rest. FDF/XFDF I/O ships in JSON
 * lossless form; the binary FDF format is a follow-up.
 */
export type {
  AnnotationInput,
  HighlightInput,
  StickyNoteInput,
  FreeTextInput,
  LinkInput,
  RectInput,
  ColorInput,
  AnnotationKind,
} from "./types.js";

export { addAnnotations, type AddAnnotationsOptions } from "./writer.js";
export { exportXfdf, importXfdf, type XfdfDocument } from "./xfdf.js";
