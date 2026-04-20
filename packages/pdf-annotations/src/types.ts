export type RectInput = readonly [number, number, number, number];
export interface ColorInput {
  readonly r: number;
  readonly g: number;
  readonly b: number;
  readonly a?: number;
}

export type AnnotationKind = "highlight" | "sticky-note" | "free-text" | "link";

interface BaseAnnotation {
  readonly pageNumber: number;
  readonly rect: RectInput;
  readonly author?: string;
  readonly contents?: string;
}

export interface HighlightInput extends BaseAnnotation {
  readonly kind: "highlight";
  readonly color?: ColorInput;
  /** Optional text that the highlight is on (preserved in /Contents). */
  readonly text?: string;
}

export interface StickyNoteInput extends BaseAnnotation {
  readonly kind: "sticky-note";
  readonly contents: string;
  readonly color?: ColorInput;
}

export interface FreeTextInput extends BaseAnnotation {
  readonly kind: "free-text";
  readonly contents: string;
  readonly fontSize?: number;
  readonly color?: ColorInput;
}

export interface LinkInput extends BaseAnnotation {
  readonly kind: "link";
  readonly url?: string;
  readonly destPage?: number;
}

export type AnnotationInput = HighlightInput | StickyNoteInput | FreeTextInput | LinkInput;
