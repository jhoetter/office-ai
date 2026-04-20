/**
 * Type-only contract that the four embeddable editor components in
 * `@officeai/react-editors/{docx,xlsx,pptx,pdf}` will satisfy.
 *
 * This file is intentionally `.ts` (not `.tsx`) so it carries zero
 * runtime cost — hosts can import the types to build their own
 * wrapper components without pulling in the editor implementation.
 *
 * The Phase 0 in-repo refactor wired these props into
 * `apps/web/app/{editor,xlsx-editor,pptx-editor,pdf-viewer}/*Editor.tsx`
 * directly (see the office-ai source tree). Phase 1.5 extracts those
 * components into this package; until then hosts that want to embed
 * an editor either point an iframe at the apps/web URL or wait for
 * the next release that ships the components here.
 */

export type Locale = "en" | "de";
export type Theme = "light" | "dark";

/**
 * Host-supplied save handler. Receives the freshly-exported document
 * bytes, the canonical MIME (one of the constants in
 * `@officeai/react-editors/mime`), and the working filename.
 *
 * The host is responsible for persistence (presigned S3 PUT, an HTTP
 * PUT to its own backend, FileSystemAccess, etc). The editor will
 * await the returned promise and surface a toast based on the outcome
 * (success when the promise resolves, error when it rejects).
 */
export type EmbeddedEditorOnSave = (bytes: Uint8Array, mime: string, filename: string) => Promise<void>;

/**
 * Common props every embeddable editor accepts. Each editor adds its
 * own `onBootstrapReady` (and any format-specific extras), so hosts
 * may want to extend this rather than reach for the per-format types
 * directly.
 */
export interface EmbeddedEditorProps {
  /**
   * Pre-loaded document bytes. When set, takes priority over any
   * URL-based or blank bootstrap path so embedding hosts can stream
   * a `Uint8Array` straight into the editor without first stashing
   * it under a URL. The companion `initialFilename` controls the
   * working filename used for Save / Export.
   */
  readonly initialBytes?: Uint8Array;
  /**
   * Working filename to display in the editor header and use as the
   * default Save / Export name. Required when only `initialBytes` is
   * set; otherwise inferred from the URL-based source.
   */
  readonly initialFilename?: string;
  /**
   * Host save handler — see `EmbeddedEditorOnSave`. When provided,
   * the editor's Save action invokes this with the exported bytes
   * instead of falling through to its own File-System-Access /
   * download fallback.
   */
  readonly onSave?: EmbeddedEditorOnSave;
  /**
   * Host close handler. When provided, the editor surfaces a "Back"
   * affordance that calls this — the host route is responsible for
   * the actual navigation back to its own UI (e.g. asset list).
   */
  readonly onClose?: () => void;
  /**
   * Override the editor's i18n locale. When set, the editor mounts a
   * self-contained `<I18nProvider initialLocale={locale}>` so a host
   * whose own provider is in a different locale (or absent entirely)
   * still gets the correct UI language for the editor surface.
   */
  readonly locale?: Locale;
  /**
   * Theme override forwarded to the editor's theme provider. When
   * omitted, the editor falls back to its own default (currently the
   * `prefers-color-scheme` system pref / the host's own next-themes
   * `.dark` class on `<html>`).
   */
  readonly theme?: Theme;
}
