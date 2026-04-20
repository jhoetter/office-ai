/**
 * Hand-rolled type declarations for the bundled PPTX editor component.
 * See `./docx.d.ts` for the rationale.
 */
import type { ComponentType } from "react";
import type { EmbeddedEditorProps } from "../contract";

export interface PptxEditorProps extends EmbeddedEditorProps {
  readonly onBootstrapReady?: (ready: boolean) => void;
  readonly initialBlank?: boolean;
  readonly initialSource?: { url: string; name?: string };
}

export declare const PptxEditor: ComponentType<PptxEditorProps>;
export type { EmbeddedEditorProps };
