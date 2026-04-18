import { Schema, type MarkSpec, type NodeSpec } from "prosemirror-model";

/**
 * ProseMirror schema for the DOCX renderer. Mirrors `DocxSnapshot` 1:1
 * so that PM <-> model conversion is mechanical. See spec/docx/renderer.md.
 */

/**
 * Map a typed `ParagraphProperties.styleId` to the HTML tag the renderer
 * should emit. Word's heading styles map to `h1`-`h6`; everything else
 * stays as `<p>` and relies on the `data-style` attr + CSS rules in
 * `apps/web/app/globals.css` for visual differentiation.
 *
 * Both English (`Heading1`, `Title`, `Subtitle`) and German Word style
 * ids (`berschrift1`, `Titel`, `Untertitel`) are recognised — the
 * masterthesis fixture uses the German variants because Word strips the
 * leading `Ü` from `Überschrift` when emitting style ids (ASCII-only).
 */
export function paragraphHtmlTag(styleId: string): string {
  switch (styleId) {
    case "Title":
    case "Titel":
      return "h1";
    case "Subtitle":
    case "Untertitel":
      return "h2";
    case "Heading1":
    case "berschrift1":
      return "h1";
    case "Heading2":
    case "berschrift2":
      return "h2";
    case "Heading3":
    case "berschrift3":
      return "h3";
    case "Heading4":
    case "berschrift4":
      return "h4";
    case "Heading5":
    case "berschrift5":
      return "h5";
    case "Heading6":
    case "berschrift6":
    case "Heading7":
    case "berschrift7":
    case "Heading8":
    case "berschrift8":
    case "Heading9":
    case "berschrift9":
      return "h6";
    default:
      return "p";
  }
}

const nodes: Record<string, NodeSpec> = {
  doc: { content: "block+" },
  paragraph: {
    group: "block",
    content: "inline*",
    attrs: {
      paragraphId: { default: null },
      styleId: { default: null },
      alignment: { default: null },
      propsJson: { default: null },
    },
    parseDOM: [
      { tag: "p" },
      { tag: "h1" },
      { tag: "h2" },
      { tag: "h3" },
      { tag: "h4" },
      { tag: "h5" },
      { tag: "h6" },
    ],
    toDOM(node) {
      const styleId = typeof node.attrs.styleId === "string" ? (node.attrs.styleId as string) : "";
      const tag = paragraphHtmlTag(styleId);
      const cls = styleId ? `pm-p style-${styleId}` : "pm-p";
      const attrs: Record<string, string> = {
        class: cls,
        "data-style": styleId,
      };
      const alignment = typeof node.attrs.alignment === "string" ? (node.attrs.alignment as string) : "";
      if (alignment) {
        attrs.style = `text-align:${alignment}`;
        attrs["data-align"] = alignment;
      }
      return [tag, attrs, 0];
    },
  },
  table: {
    group: "block",
    atom: true,
    attrs: {
      tableId: { default: null },
      rawJson: { default: null },
      /**
       * Structured projection of the typed `Table` model:
       *   { rows: Array<{ header: boolean, cells: Array<{
       *       gridSpan: number, vMerge: "restart"|"continue"|null,
       *       blocks: Array<{ kind: "paragraph", text: string,
       *                       styleId?: string, alignment?: string }>
       *   }>}> }
       * The renderer materialises this into a `<table>` DOM subtree so the
       * user sees the actual cell content instead of a `[table]` chip.
       * Atom-ness is intentional — cells are read-only in this iteration to
       * keep top-level paragraph indexing in `transactionToCommands` stable.
       */
      tableJson: { default: null },
    },
    toDOM(node) {
      const data = parseTableJson(node.attrs.tableJson);
      if (!data || data.rows.length === 0) {
        return ["div", { class: "pm-table-placeholder", contenteditable: "false" }, "[table]"];
      }
      const tbody: unknown[] = ["tbody"];
      for (const row of data.rows) {
        const tr: unknown[] = [
          "tr",
          { class: row.header ? "pm-table-row pm-table-header-row" : "pm-table-row" },
        ];
        for (const cell of row.cells) {
          if (cell.vMerge === "continue") continue;
          const cellAttrs: Record<string, string> = { class: "pm-table-cell" };
          if (cell.gridSpan > 1) cellAttrs.colspan = String(cell.gridSpan);
          const cellChildren: unknown[] = [];
          for (const block of cell.blocks) {
            const pAttrs: Record<string, string> = { class: "pm-table-cell-p" };
            if (block.alignment) pAttrs.style = `text-align:${block.alignment}`;
            cellChildren.push(["p", pAttrs, block.text]);
          }
          tr.push([row.header ? "th" : "td", cellAttrs, ...cellChildren]);
        }
        tbody.push(tr);
      }
      return ["table", { class: "pm-table", contenteditable: "false" }, tbody];
    },
  },
  opaque_block: {
    group: "block",
    atom: true,
    attrs: {
      blockId: { default: null },
      rawJson: { default: null },
      tag: { default: null },
      previewText: { default: null },
    },
    toDOM(node) {
      const preview = typeof node.attrs.previewText === "string" ? (node.attrs.previewText as string) : "";
      if (preview.length > 0) {
        return [
          "div",
          {
            class: "pm-opaque-block pm-opaque-block-preview",
            contenteditable: "false",
            "data-tag": node.attrs.tag ?? "",
          },
          preview,
        ];
      }
      return [
        "div",
        { class: "pm-opaque-block", contenteditable: "false", "data-tag": node.attrs.tag ?? "" },
        `[${node.attrs.tag ?? "opaque"}]`,
      ];
    },
  },
  section_break: {
    group: "block",
    atom: true,
    attrs: { blockId: { default: null }, rawJson: { default: null } },
    toDOM() {
      return ["hr", { class: "pm-section-break", contenteditable: "false" }];
    },
  },
  text: { group: "inline" },
  hard_break: {
    group: "inline",
    inline: true,
    selectable: false,
    attrs: { breakType: { default: null } },
    toDOM() {
      return ["br"];
    },
  },
  tab: {
    group: "inline",
    inline: true,
    selectable: false,
    toDOM() {
      return ["span", { class: "pm-tab" }, "\t"];
    },
  },
  image: {
    group: "inline",
    inline: true,
    atom: true,
    attrs: { runId: { default: null }, drawingJson: { default: null } },
    toDOM() {
      return ["span", { class: "pm-image-placeholder" }, "[image]"];
    },
  },
  opaque_inline: {
    group: "inline",
    inline: true,
    atom: true,
    attrs: {
      inlineId: { default: null },
      rawJson: { default: null },
      tag: { default: null },
      previewText: { default: null },
    },
    toDOM(node) {
      const preview = typeof node.attrs.previewText === "string" ? (node.attrs.previewText as string) : "";
      if (preview.length > 0) {
        return [
          "span",
          {
            class: "pm-opaque-inline pm-opaque-inline-preview",
            "data-tag": node.attrs.tag ?? "",
          },
          preview,
        ];
      }
      return ["span", { class: "pm-opaque-inline", "data-tag": node.attrs.tag ?? "" }, "[opaque]"];
    },
  },
};

const marks: Record<string, MarkSpec> = {
  bold: { toDOM: () => ["strong", 0], parseDOM: [{ tag: "strong" }, { tag: "b" }] },
  italic: { toDOM: () => ["em", 0], parseDOM: [{ tag: "em" }, { tag: "i" }] },
  underline: {
    attrs: { value: { default: true } },
    toDOM: () => ["u", 0],
    parseDOM: [{ tag: "u" }],
  },
  strikethrough: { toDOM: () => ["s", 0], parseDOM: [{ tag: "s" }, { tag: "strike" }] },
  font_family: {
    attrs: { family: { default: "" } },
    toDOM: (mark) => [
      "span",
      { class: "pm-font-family", style: `font-family: ${String(mark.attrs.family)}` },
      0,
    ],
  },
  font_size: {
    attrs: { halfPoints: { default: 22 } },
    toDOM: (mark) => [
      "span",
      { class: "pm-font-size", style: `font-size: ${Number(mark.attrs.halfPoints) / 2}pt` },
      0,
    ],
  },
  color: {
    attrs: { rgb: { default: "000000" } },
    toDOM: (mark) => ["span", { class: "pm-color", style: `color: #${String(mark.attrs.rgb)}` }, 0],
  },
  highlight: {
    attrs: { name: { default: "yellow" } },
    toDOM: (mark) => ["span", { class: "pm-highlight", "data-highlight": String(mark.attrs.name) }, 0],
  },
  hyperlink: {
    attrs: { relationshipId: { default: null }, anchor: { default: null }, hyperlinkId: { default: null } },
    toDOM: (mark) => [
      "a",
      {
        class: "pm-hyperlink",
        "data-rid": String(mark.attrs.relationshipId ?? ""),
        "data-anchor": String(mark.attrs.anchor ?? ""),
      },
      0,
    ],
  },
  comment_mark: {
    attrs: { commentId: { default: "" } },
    toDOM: (mark) => [
      "span",
      { class: "pm-comment-mark", "data-comment-id": String(mark.attrs.commentId) },
      0,
    ],
  },
  revision_mark: {
    attrs: {
      revisionType: { default: "ins" },
      author: { default: "" },
      date: { default: "" },
      revisionId: { default: "" },
    },
    toDOM: (mark) => [
      "span",
      {
        class: `pm-revision pm-revision-${String(mark.attrs.revisionType)}`,
        "data-revision-id": String(mark.attrs.revisionId),
      },
      0,
    ],
  },
};

export interface RenderableTableBlock {
  readonly kind: "paragraph";
  readonly text: string;
  readonly styleId?: string;
  readonly alignment?: string;
}

export interface RenderableTableCell {
  readonly gridSpan: number;
  readonly vMerge: "restart" | "continue" | null;
  readonly blocks: ReadonlyArray<RenderableTableBlock>;
}

export interface RenderableTableRow {
  readonly header: boolean;
  readonly cells: ReadonlyArray<RenderableTableCell>;
}

export interface RenderableTable {
  readonly rows: ReadonlyArray<RenderableTableRow>;
}

function parseTableJson(value: unknown): RenderableTable | null {
  if (typeof value !== "string" || value.length === 0) return null;
  try {
    const parsed = JSON.parse(value) as unknown;
    if (!parsed || typeof parsed !== "object") return null;
    const rows = (parsed as { rows?: unknown }).rows;
    if (!Array.isArray(rows)) return null;
    return parsed as RenderableTable;
  } catch {
    return null;
  }
}

export const docxSchema = new Schema({ nodes, marks });
