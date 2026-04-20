/**
 * Hand-rolled type declarations for the bundled DOCX editor component.
 *
 * The implementation lives in `apps/web/app/editor/DocxEditor.tsx` and
 * is bundled at build time via `build.mjs`. Hosts should rely on the
 * `EmbeddedEditorProps` shape (re-exported below) rather than the full
 * component-specific props to stay decoupled from the apps/web internal
 * API as it evolves.
 */
import type { ComponentType } from "react";
import type { EmbeddedEditorProps } from "../contract";

export interface DocxEditorProps extends EmbeddedEditorProps {
  readonly onBootstrapReady?: (ready: boolean) => void;
  readonly initialBlank?: boolean;
  readonly initialSource?: { url: string; name?: string };
}

export declare const DocxEditor: ComponentType<DocxEditorProps>;
export type { EmbeddedEditorProps };
