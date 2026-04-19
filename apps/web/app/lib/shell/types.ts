/**
 * Shared types for the cross-product editor shell.
 *
 * The shell (top bar, right rail, status bar, command palette,
 * find/replace) is product-agnostic. Each product (DOCX/XLSX/PPTX)
 * exposes a typed adapter that the shell consumes; everything inside
 * this file is the contract between the two layers.
 */

import type { ReactNode } from "react";

/** Identity of the active editor product. Drives the per-product
 * shortcuts catalogue, the export menu, and the right-rail tabs
 * (Outline is DOCX-only). */
export type ProductKind = "docx" | "xlsx" | "pptx";

/** Save / dirty state surfaced as a small pill in the top bar. The
 * product reports its current state and the shell decides how to
 * render it. `unknown` is the no-document-open fallback. */
export type SaveState = "saved" | "modified" | "saving" | "error" | "unknown";

/** Whether selecting a format runs immediately ("instant") or first
 * opens the rich Export dialog so the user can pick options
 * ("dialog"). The shell branches on this from the dropdown. */
export type ExportFormatKind = "instant" | "dialog";

/** Visual grouping in the Export dropdown / dialog left rail. The
 * shell renders the groups in the order they appear here:
 *  - `current` — exports of the current page / slide / sheet
 *    (always rendered first because it's almost always what the user
 *    wants when they reach for "Export this");
 *  - `native` — the product's own format;
 *  - `pdf-web` — PDF / HTML (server-side conversion);
 *  - `data`   — CSV/TSV/JSON/MD;
 *  - `images` — PNG/SVG/JPEG bundles. */
export type ExportFormatGroup = "current" | "native" | "pdf-web" | "data" | "images";

/** Lucide icon family used for the format row. The shell maps these
 * to concrete icons (kept here so adapters don't import lucide). */
export type ExportFormatIcon = "doc" | "sheet" | "slides" | "pdf" | "image" | "code" | "text";

/** Declarative description of a single option control rendered in
 * the rich Export dialog. The shell renders these without knowing
 * what the option does — the product reads the values back from the
 * `ExportOptionValues` map in its `onExport` handler. */
export interface ExportFormatOptionField {
  readonly id: string;
  readonly label: string;
  /** Optional helper text shown beneath the control. */
  readonly hint?: string;
  readonly control:
    | {
        readonly type: "select";
        readonly options: ReadonlyArray<{ readonly id: string; readonly label: string }>;
        readonly defaultId: string;
      }
    | { readonly type: "toggle"; readonly defaultValue: boolean }
    | { readonly type: "text"; readonly placeholder?: string; readonly defaultValue?: string }
    | {
        readonly type: "multiSelect";
        readonly options: ReadonlyArray<{ readonly id: string; readonly label: string }>;
        readonly defaultIds: ReadonlyArray<string>;
      };
}

/** Values collected from the dialog and forwarded to `onExport`. */
export type ExportOptionValue = string | boolean | ReadonlyArray<string>;
export type ExportOptionValues = Readonly<Record<string, ExportOptionValue>>;

/** A single export choice surfaced in the Export dropdown / dialog. */
export interface ExportFormat {
  readonly id: string;
  /** Human label, e.g. "Word document (.docx)". */
  readonly label: string;
  /** File extension without leading dot — used for the download
   * filename when the product doesn't override it. */
  readonly extension: string;
  /** MIME type for the Blob. */
  readonly mime: string;
  /** "instant" runs `onExport` straight from the dropdown; "dialog"
   * routes through the rich Export dialog so the user can configure
   * options first. Optional for legacy adapters; defaults to
   * "instant". */
  readonly kind?: ExportFormatKind;
  /** Visual grouping in the dialog left rail. Defaults to "native"
   * for the primary format and "pdf-web" otherwise. */
  readonly group?: ExportFormatGroup;
  /** Icon hint for the dropdown / dialog. */
  readonly icon?: ExportFormatIcon;
  /** Optional one-line description shown in the dialog under the
   * format label. */
  readonly description?: string;
  /** Declarative form fields surfaced in the dialog options pane.
   * Omit for formats that have nothing to configure. */
  readonly optionFields?: ReadonlyArray<ExportFormatOptionField>;
}

/** A single command exposed in the command palette. The `id` is used
 * for "recent commands first" persistence; the `run` callback is what
 * fires on Enter. Keep run() side-effect-free with respect to the
 * shell — the product is responsible for any toast or modal it wants
 * to surface as part of the action. */
export interface PaletteCommand {
  readonly id: string;
  readonly label: string;
  readonly hint?: string;
  /** Optional keyboard shortcut to display (display-only — the actual
   * binding lives in the shortcuts catalogue). */
  readonly shortcut?: string;
  /** Optional grouping label so the palette can section commands. */
  readonly section?: string;
  readonly run: () => void | Promise<void>;
  /** Hidden when false; allows the product to gate commands by
   * selection / cursor position without re-listing the palette. */
  readonly enabled?: boolean;
}

/** A single comment thread surfaced in the right-rail Comments tab.
 * The shell does not own comment rendering — it delegates to the
 * existing per-product `CommentsSidebar` provider — but it still needs
 * to know the count to badge the tab. */
export interface CommentsBadge {
  readonly openCount: number;
  readonly resolvedCount: number;
}

/** A heading entry for the Outline tab (DOCX-only today). */
export interface OutlineEntry {
  readonly id: string;
  readonly level: number; // 1..9
  readonly text: string;
  readonly onActivate: () => void;
  /**
   * True when the caret is currently inside (or after, but before the
   * next heading of equal-or-shallower level) this entry. The shell
   * paints the active row and scrolls it into view so the user can
   * always see "where they are" in the document at a glance.
   */
  readonly active?: boolean;
}

/** Status-bar selection summary. The product computes a short string
 * (or a structured summary for XLSX aggregates) and the shell renders
 * it dead-centre. */
export interface SelectionSummary {
  /** A single short string. Use for DOCX ("Paragraph 4 · 12 words")
   * and PPTX ("3 shapes selected"). */
  readonly text?: string;
  /** Optional structured aggregates for XLSX — Sum / Avg / Count /
   * Min / Max. The shell renders them as `Sum: x   Avg: y   …`. */
  readonly aggregates?: ReadonlyArray<{ readonly label: string; readonly value: string }>;
}

/** A product-side find adapter. Consumed by the shared
 * `FindReplacePanel`. */
export interface FindAdapter {
  readonly findAll: (query: string, options: FindOptions) => readonly FindMatch[];
  readonly gotoMatch: (match: FindMatch) => void;
  readonly replaceMatch: (match: FindMatch, replacement: string) => Promise<void> | void;
  readonly replaceAll: (query: string, replacement: string, options: FindOptions) => Promise<number> | number;
}

export interface FindOptions {
  readonly caseSensitive: boolean;
  readonly wholeWord: boolean;
  readonly regex: boolean;
}

/** Opaque match descriptor — the product owns the shape. */
export interface FindMatch {
  readonly id: string;
  /** A short preview of the match context, surfaced in the panel. */
  readonly preview: string;
}

/** The full product adapter that the shell consumes. Every editor
 * (DOCX/XLSX/PPTX) builds one of these and passes it as a prop to
 * `EditorShell`. */
export interface ProductAdapter {
  readonly product: ProductKind;
  /** Filename shown in the top bar. */
  readonly filename: string;
  /** Save state pill (saved / modified / saving / error). */
  readonly saveState: SaveState;
  /** Right-rail comments badge. Undefined hides the badge. */
  readonly comments?: CommentsBadge;
  /** DOCX-only outline. Undefined hides the Outline tab. */
  readonly outline?: ReadonlyArray<OutlineEntry>;
  /** Status-bar selection summary. */
  readonly selectionSummary?: SelectionSummary;

  // ── File ops ──
  readonly canOpen: boolean;
  readonly canSave: boolean;
  readonly canExport: boolean;
  readonly exportFormats: ReadonlyArray<ExportFormat>;
  readonly onOpenFile: () => void;
  readonly onSave: () => Promise<void> | void;
  readonly onExport: (format: ExportFormat, options?: ExportOptionValues) => Promise<void> | void;

  // ── History ──
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  readonly onUndo: () => void;
  readonly onRedo: () => void;

  // ── Shortcut surface ──
  readonly onOpenShortcuts: () => void;

  // ── Command palette ──
  readonly paletteCommands: ReadonlyArray<PaletteCommand>;

  // ── Find / Replace ──
  readonly findAdapter?: FindAdapter;

  // ── Right rail ──
  /** Renders the comments panel. The shell calls this when the rail
   * is open and the Comments tab is active. */
  readonly renderCommentsPanel?: () => ReactNode;
  /**
   * D11 — Optional Animations panel renderer (PPTX only). When
   * provided, the shell exposes an "Animations" tab in the right rail
   * alongside Comments. We delegate the entire panel UI to the
   * adapter so each product can shape it without leaking command
   * specifics into the shared shell.
   */
  readonly renderAnimationsPanel?: () => ReactNode;
  /** Optional: produce a programmatic "open the comment composer"
   * trigger that the shell can fire from the top-bar comments icon. */
  readonly onAddComment?: () => void;
}
