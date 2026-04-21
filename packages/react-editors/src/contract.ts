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
 * Host-supplied presence identity. Overrides the editor's default
 * "Anonymous Quokka" identity so multi-user cursors / avatars / comment
 * authorship show the *real* logged-in user (e.g. "Johannes Hötter")
 * instead of a generated handle.
 *
 * Embedding hosts that already have an authentication layer
 * (hof-os, internal portals, etc.) should always pass this — the
 * anonymous fallback is intended for the standalone office-ai web app
 * where users may not be signed in.
 *
 *   - `id`     stable per human (NOT per tab). Used to dedupe
 *              multi-tab presence and to attribute commands. We
 *              recommend the host's user UUID.
 *   - `name`   shown in the avatar tooltip and on tracked-changes
 *              authorship.
 *   - `color`  optional hex paint for the cursor / outline. When
 *              omitted the editor derives a stable colour from `id`
 *              via the same FNV-1a palette used for anonymous peers.
 */
export interface PresenceUser {
  readonly id: string;
  readonly name: string;
  readonly color?: string;
}

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
  /**
   * Identity to publish in the realtime awareness payload. When set,
   * remote cursors render the host-provided `name` (e.g. the
   * authenticated user's first + last name) instead of a generated
   * anonymous handle. Omit to keep the existing anonymous-peer
   * behavior. See `PresenceUser` for the per-field contract.
   */
  readonly presenceUser?: PresenceUser;
  /**
   * Explicit realtime room id. When set, the editor joins this exact
   * room instead of deriving one from the source URL or per-tab
   * fallback. Use a stable host-side hash of the document identity
   * (e.g. the S3 object key) so two browsers opening the same file
   * land in the same room without coordinating URLs.
   *
   * Pass `null` (not undefined) to explicitly disable realtime in this
   * mount — useful for previews and read-only embeds. Undefined keeps
   * the editor's built-in default room resolution.
   */
  readonly room?: string | null;
  /**
   * Hide affordances that load a *local* file into the editor (the
   * 📁 Open toolbar button, the matching Cmd+O / drag-drop
   * shortcuts). Embedded hosts manage their own document corpus
   * (S3 / DB / etc.) so a local-file open is a foot-gun: the user
   * loads bytes from their disk into the editor, edits, hits Save,
   * and the host's `onSave` writes those bytes back to the wrong
   * S3 key (or, before this flag, sometimes back to the local file
   * via the File System Access fallback).
   *
   * Defaults to `false` so the standalone office-ai web app keeps
   * its current behavior. Hosts with their own asset browser
   * (hof-os' `/edit-asset`, customer portals, …) should set this to
   * `true`.
   */
  readonly hideLocalFileOpen?: boolean;
}
