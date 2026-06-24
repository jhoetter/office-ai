/**
 * Hand-rolled type declarations for the bundled PDF editor component.
 * See `./docx.d.ts` for the rationale.
 *
 * Note: hosts must serve the `pdfjs-dist` worker assets at `/pdfjs/` —
 * see `docs/embedding.md` in the office-ai repo and the embedding host
 * integration guide for the static-asset wiring.
 */
import type { ComponentType } from "react";
import type { EmbeddedEditorProps } from "../contract";

export interface PdfEditorProps extends EmbeddedEditorProps {
  readonly onBootstrapReady?: (ready: boolean) => void;
  readonly initialBlank?: boolean;
  readonly initialSource?: { url: string; name?: string };
}

export declare const PdfEditor: ComponentType<PdfEditorProps>;
export type { EmbeddedEditorProps };
