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
      return renderTableDom(data) as import("prosemirror-model").DOMOutputSpec;
    },
  },
  opaque_block_wrapper: {
    group: "block",
    atom: true,
    attrs: {
      blockId: { default: null },
      rawJson: { default: null },
      tag: { default: null },
      /**
       * Structured projection of the children of an unwrapped content-wrapper
       * carrier (SDT / fldSimple / mc:AlternateContent / smartTag /
       * customXml). Same shape as the `table` node's `tableJson`: a list of
       * lightweight `{ kind, text, styleId?, alignment? }` records, so the
       * renderer can surface the wrapped paragraphs as real `<h1>` / `<p>`
       * nodes instead of collapsing the whole subtree into a single italic
       * preview chip.
       *
       * Read-only on purpose for this iteration — editing inside an opaque
       * carrier still requires the dirty-flag plumbing landed in P2.3 to be
       * exercised by mutating commands, which is deferred.
       */
      contentJson: { default: null },
    },
    toDOM(node) {
      const data = parseWrapperContent(node.attrs.contentJson);
      const tag = typeof node.attrs.tag === "string" ? node.attrs.tag : "";
      const children: unknown[] = [];
      if (data && data.blocks.length > 0) {
        for (const b of data.blocks) {
          const html = paragraphHtmlTag(b.styleId ?? "");
          const attrs: Record<string, string> = {
            class: "pm-opaque-wrapper-p",
            "data-style": b.styleId ?? "",
          };
          if (b.alignment) attrs.style = `text-align:${b.alignment}`;
          children.push([html, attrs, b.text]);
        }
      } else {
        children.push(["div", { class: "pm-opaque-wrapper-empty" }, `[${tag}]`]);
      }
      return [
        "div",
        {
          class: "pm-opaque-wrapper-block",
          "data-tag": tag,
          contenteditable: "false",
        },
        ...children,
      ] as unknown as import("prosemirror-model").DOMOutputSpec;
    },
  },
  opaque_inline_wrapper: {
    group: "inline",
    inline: true,
    atom: true,
    attrs: {
      inlineId: { default: null },
      rawJson: { default: null },
      tag: { default: null },
      /** `{ text: string }` — flattened text of the wrapped run children. */
      contentJson: { default: null },
    },
    toDOM(node) {
      const tag = typeof node.attrs.tag === "string" ? node.attrs.tag : "";
      const text = parseInlineWrapperContent(node.attrs.contentJson);
      return [
        "span",
        { class: "pm-opaque-wrapper-inline", "data-tag": tag },
        text.length > 0 ? text : `[${tag}]`,
      ];
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
  /**
   * Invisible PM marker that mirrors a `WrapperMarker` body block (the
   * begin / end brackets the parser inserts around a *lifted*
   * content-wrapper carrier — `<w:sdt>`, `mc:AlternateContent`, …).
   *
   * - `atom: true` so PM does not let the user place a caret inside it
   *   (the marker has no inner content; it is purely structural).
   * - Renders as a zero-size `<span>` with `display:none` so it does
   *   not contribute height to its enclosing page block — the page
   *   chunker treats it the same way.
   * - Carries `wrapperId` + `side` (`begin` | `end`) attrs so a future
   *   command pipeline (typed editing inside SDT carriers) can locate
   *   the bracketing markers without re-reading the snapshot.
   */
  wrapper_marker: {
    group: "block",
    atom: true,
    attrs: {
      blockId: { default: null },
      wrapperId: { default: null },
      side: { default: "begin" },
      tag: { default: null },
    },
    toDOM(node) {
      return [
        "span",
        {
          class: "pm-wrapper-marker",
          contenteditable: "false",
          "data-wrapper-id": String(node.attrs.wrapperId ?? ""),
          "data-wrapper-side": String(node.attrs.side ?? ""),
          "data-tag": String(node.attrs.tag ?? ""),
          style: "display:none",
        },
      ];
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
    attrs: {
      runId: { default: null },
      drawingJson: { default: null },
      /**
       * `data:` URL for the image bytes. Built at render time by
       * `doc-to-pm.ts` from the snapshot's `media` + `relationships` so
       * the editor can show real pixels instead of an `[image]` chip.
       * Falls back to the placeholder when the relationship cannot be
       * resolved (e.g. broken doc, unknown rId, missing media part).
       */
      dataUrl: { default: null },
      /** Display width in CSS pixels (intrinsic 1× size). */
      width: { default: null },
      /** Display height in CSS pixels (intrinsic 1× size). */
      height: { default: null },
      /** Alt text for screen readers / accessibility. */
      alt: { default: "" },
    },
    toDOM(node) {
      const dataUrl = typeof node.attrs.dataUrl === "string" ? node.attrs.dataUrl : "";
      if (dataUrl.length > 0) {
        const attrs: Record<string, string> = {
          class: "pm-image",
          src: dataUrl,
          alt: typeof node.attrs.alt === "string" ? node.attrs.alt : "",
          draggable: "false",
        };
        const w = typeof node.attrs.width === "number" ? node.attrs.width : 0;
        const h = typeof node.attrs.height === "number" ? node.attrs.height : 0;
        if (w > 0) attrs.width = String(w);
        if (h > 0) attrs.height = String(h);
        return ["img", attrs];
      }
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

/**
 * Lightweight projection of a `<w:r>` for renderable table cells. Phase 2
 * keeps cells read-only but preserves typed runs (text + the marks the
 * editor already understands) so cell paragraphs render with bold,
 * italic, color, highlight, font, etc. instead of being flattened to
 * plain text.
 *
 * The shape is deliberately structural — only fields the renderer needs
 * to paint Word-flavoured tables. The full typed `Run` lives on the
 * snapshot for round-trip; the renderer doesn't need it.
 */
export interface RenderableTableRun {
  readonly text: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean | string;
  readonly strike?: boolean;
  readonly fontFamily?: string;
  /** Half-points (Word units). 22 = 11pt. */
  readonly fontSize?: number;
  /** Hex RGB without leading `#`. */
  readonly color?: string;
  /** OOXML highlight enum value (yellow, green, …). */
  readonly highlight?: string;
}

export interface RenderableTableBlock {
  readonly kind: "paragraph";
  readonly runs: ReadonlyArray<RenderableTableRun>;
  readonly styleId?: string;
  readonly alignment?: string;
}

export interface RenderableTableBorderSide {
  readonly style: string;
  /** OOXML `w:sz` units (eighths of a point). */
  readonly size?: number;
  readonly color?: string;
}

export interface RenderableTableBorders {
  readonly top?: RenderableTableBorderSide;
  readonly left?: RenderableTableBorderSide;
  readonly bottom?: RenderableTableBorderSide;
  readonly right?: RenderableTableBorderSide;
  readonly insideH?: RenderableTableBorderSide;
  readonly insideV?: RenderableTableBorderSide;
}

export interface RenderableTableCellProps {
  /** Hex RGB cell shading fill (no leading `#`). `auto` → undefined. */
  readonly shadingFill?: string;
  readonly borders?: RenderableTableBorders;
  readonly vAlign?: "top" | "center" | "bottom";
  /** Twips. */
  readonly widthTw?: number;
  /** Per-side cell padding overrides, twips. */
  readonly padTop?: number;
  readonly padRight?: number;
  readonly padBottom?: number;
  readonly padLeft?: number;
}

export interface RenderableTableCell {
  readonly gridSpan: number;
  readonly vMerge: "restart" | "continue" | null;
  readonly blocks: ReadonlyArray<RenderableTableBlock>;
  readonly props?: RenderableTableCellProps;
}

export interface RenderableTableRow {
  readonly header: boolean;
  readonly cells: ReadonlyArray<RenderableTableCell>;
}

export interface RenderableTableProps {
  readonly borders?: RenderableTableBorders;
  /** Default cell padding in twips (`<w:tblCellMar>`). */
  readonly padTop?: number;
  readonly padRight?: number;
  readonly padBottom?: number;
  readonly padLeft?: number;
  readonly layout?: "auto" | "fixed";
  /** Table width in twips when `widthType` is `"dxa"`. */
  readonly widthTw?: number;
  /** Table width as 50ths of a percent when `widthType` is `"pct"`. */
  readonly widthPct?: number;
  readonly widthType?: "auto" | "dxa" | "pct" | "nil";
  readonly jc?: "left" | "center" | "right" | "start" | "end";
  /** Table indent in twips (positive = inset from left margin). */
  readonly indentTw?: number;
}

export interface RenderableTable {
  readonly rows: ReadonlyArray<RenderableTableRow>;
  /** Per-column widths in twips. Drives `<colgroup>`. */
  readonly gridCols?: ReadonlyArray<number>;
  readonly props?: RenderableTableProps;
}

export interface WrapperContentBlock {
  readonly kind: "paragraph";
  readonly text: string;
  readonly styleId?: string;
  readonly alignment?: string;
}

export interface WrapperContent {
  readonly blocks: ReadonlyArray<WrapperContentBlock>;
}

function parseWrapperContent(value: unknown): WrapperContent | null {
  if (typeof value !== "string" || value.length === 0) return null;
  try {
    const parsed = JSON.parse(value) as unknown;
    if (!parsed || typeof parsed !== "object") return null;
    const blocks = (parsed as { blocks?: unknown }).blocks;
    if (!Array.isArray(blocks)) return null;
    return parsed as WrapperContent;
  } catch {
    return null;
  }
}

function parseInlineWrapperContent(value: unknown): string {
  if (typeof value !== "string" || value.length === 0) return "";
  try {
    const parsed = JSON.parse(value) as unknown;
    if (!parsed || typeof parsed !== "object") return "";
    const text = (parsed as { text?: unknown }).text;
    return typeof text === "string" ? text : "";
  } catch {
    return "";
  }
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

/* ── Phase 2 table rendering helpers ─────────────────────────────────────── */

const TWIPS_PER_INCH = 1440;

function twipsToInches(tw: number): number {
  return tw / TWIPS_PER_INCH;
}

function eighthPtToPx(sz: number): number {
  // OOXML `w:sz` on a border is eighths of a point. Convert to CSS px
  // assuming the standard 96 DPI / 72 pt-per-inch ratio so the rendered
  // border is visually close to Word.
  return Math.max(0.5, (sz / 8) * (96 / 72));
}

function isMeaningfulBorderStyle(style: string | undefined): boolean {
  if (!style) return false;
  const v = style.toLowerCase();
  return v !== "nil" && v !== "none";
}

function borderStyleToCss(style: string): string {
  // OOXML border styles are richer than CSS; map the common ones.
  switch (style.toLowerCase()) {
    case "single":
      return "solid";
    case "double":
      return "double";
    case "dashed":
    case "dashsmallgap":
    case "dashedheavy":
      return "dashed";
    case "dotted":
    case "dottedheavy":
      return "dotted";
    case "thick":
    case "thickthinsmallgap":
      return "solid";
    default:
      return "solid";
  }
}

function borderSideCss(side: RenderableTableBorderSide | undefined): string | null {
  if (!side || !isMeaningfulBorderStyle(side.style)) return null;
  const widthPx = side.size !== undefined ? eighthPtToPx(side.size) : 1;
  const color = side.color && side.color !== "auto" ? `#${side.color}` : "currentColor";
  return `${widthPx.toFixed(2)}px ${borderStyleToCss(side.style)} ${color}`;
}

function isHexColor(value: string | undefined): value is string {
  if (!value) return false;
  if (value === "auto") return false;
  return /^[0-9A-Fa-f]{6}$/.test(value);
}

function tableCellPaddingStyle(
  cell: RenderableTableCell,
  table: RenderableTable
): { padding: string } | null {
  const tableProps = table.props;
  const cellProps = cell.props;
  const top = cellProps?.padTop ?? tableProps?.padTop;
  const right = cellProps?.padRight ?? tableProps?.padRight;
  const bottom = cellProps?.padBottom ?? tableProps?.padBottom;
  const left = cellProps?.padLeft ?? tableProps?.padLeft;
  if (top === undefined && right === undefined && bottom === undefined && left === undefined) {
    return null;
  }
  const t = top !== undefined ? `${twipsToInches(top).toFixed(3)}in` : "0";
  const r = right !== undefined ? `${twipsToInches(right).toFixed(3)}in` : "0";
  const b = bottom !== undefined ? `${twipsToInches(bottom).toFixed(3)}in` : "0";
  const l = left !== undefined ? `${twipsToInches(left).toFixed(3)}in` : "0";
  return { padding: `${t} ${r} ${b} ${l}` };
}

function tableStyleString(table: RenderableTable): string {
  const parts: string[] = [];
  const props = table.props;
  if (props?.widthType === "dxa" && props.widthTw !== undefined && props.widthTw > 0) {
    parts.push(`width: ${twipsToInches(props.widthTw).toFixed(3)}in`);
  } else if (props?.widthType === "pct" && props.widthPct !== undefined) {
    // OOXML pct unit is 50ths of a percent.
    const pct = props.widthPct / 50;
    parts.push(`width: ${pct}%`);
  } else if (props?.widthType === "auto") {
    parts.push("width: auto");
  }
  if (props?.layout === "fixed") parts.push("table-layout: fixed");
  if (props?.indentTw !== undefined && props.indentTw > 0) {
    parts.push(`margin-left: ${twipsToInches(props.indentTw).toFixed(3)}in`);
  }
  if (props?.jc === "center") parts.push("margin-left: auto; margin-right: auto");
  else if (props?.jc === "right" || props?.jc === "end") parts.push("margin-left: auto; margin-right: 0");

  const borders = props?.borders;
  if (borders) {
    const top = borderSideCss(borders.top);
    const right = borderSideCss(borders.right);
    const bottom = borderSideCss(borders.bottom);
    const left = borderSideCss(borders.left);
    if (top) parts.push(`border-top: ${top}`);
    if (right) parts.push(`border-right: ${right}`);
    if (bottom) parts.push(`border-bottom: ${bottom}`);
    if (left) parts.push(`border-left: ${left}`);
  }
  return parts.join("; ");
}

function cellStyleString(
  cell: RenderableTableCell,
  table: RenderableTable,
  rowIndex: number,
  rowCount: number,
  cellIndex: number,
  visibleCellCount: number,
  colSpan: number
): string {
  const parts: string[] = [];
  const props = cell.props;
  const tableProps = table.props;

  if (props?.vAlign) parts.push(`vertical-align: ${props.vAlign}`);
  if (isHexColor(props?.shadingFill)) parts.push(`background-color: #${props.shadingFill}`);

  // Cell border resolution: per-cell `tcBorders` wins; otherwise fall
  // back to the table's `tblBorders` (top/left/bottom/right for the
  // table edges, insideH/insideV for the interior). This mirrors
  // Word's resolution order.
  const isFirstRow = rowIndex === 0;
  const isLastRow = rowIndex === rowCount - 1;
  const isFirstCol = cellIndex === 0;
  const isLastCol = cellIndex === visibleCellCount - 1;

  const cellBorders = props?.borders;
  const tblBorders = tableProps?.borders;

  const pickSide = (
    cellSide: RenderableTableBorderSide | undefined,
    edgeSide: RenderableTableBorderSide | undefined,
    insideSide: RenderableTableBorderSide | undefined,
    isEdge: boolean
  ): string | null => {
    if (cellSide) return borderSideCss(cellSide);
    return borderSideCss(isEdge ? edgeSide : insideSide);
  };

  const top = pickSide(cellBorders?.top, tblBorders?.top, tblBorders?.insideH, isFirstRow);
  const right = pickSide(cellBorders?.right, tblBorders?.right, tblBorders?.insideV, isLastCol);
  const bottom = pickSide(cellBorders?.bottom, tblBorders?.bottom, tblBorders?.insideH, isLastRow);
  const left = pickSide(cellBorders?.left, tblBorders?.left, tblBorders?.insideV, isFirstCol);

  if (top) parts.push(`border-top: ${top}`);
  if (right) parts.push(`border-right: ${right}`);
  if (bottom) parts.push(`border-bottom: ${bottom}`);
  if (left) parts.push(`border-left: ${left}`);

  const padding = tableCellPaddingStyle(cell, table);
  if (padding) parts.push(`padding: ${padding.padding}`);

  // Colspan-aware width hint when grid columns are known and cell width
  // is implicit. For the explicit `<w:tcW>` case we let the column
  // group drive the layout instead so we don't double-constrain.
  void colSpan;

  return parts.join("; ");
}

function runToDom(run: RenderableTableRun): unknown {
  const styleParts: string[] = [];
  if (run.fontFamily) styleParts.push(`font-family: ${run.fontFamily}`);
  if (run.fontSize !== undefined) styleParts.push(`font-size: ${run.fontSize / 2}pt`);
  if (run.color && run.color !== "auto") styleParts.push(`color: #${run.color}`);

  let node: unknown = run.text;
  // Wrap inside-out so the outermost element is the most specific
  // (matches the order used by `pm-to-doc` reverse mapping for runs).
  if (run.bold) node = ["strong", {}, node];
  if (run.italic) node = ["em", {}, node];
  if (run.underline) node = ["u", {}, node];
  if (run.strike) node = ["s", {}, node];

  const attrs: Record<string, string> = { class: "pm-table-run" };
  if (styleParts.length > 0) attrs.style = styleParts.join("; ");
  if (run.highlight) attrs["data-highlight"] = run.highlight;
  return ["span", attrs, node];
}

function paragraphToDom(block: RenderableTableBlock): unknown {
  const tag = paragraphHtmlTag(block.styleId ?? "");
  const attrs: Record<string, string> = { class: "pm-table-cell-p" };
  if (block.alignment) attrs.style = `text-align: ${block.alignment}`;
  if (block.styleId) attrs["data-style"] = block.styleId;
  // Render each run as its own inline span so marks survive; if the
  // block has no runs at all, emit an empty paragraph (still valid
  // HTML and Word does the same for empty cells).
  if (block.runs.length === 0) return [tag, attrs];
  const children: unknown[] = block.runs.map(runToDom);
  return [tag, attrs, ...children];
}

function renderTableDom(data: RenderableTable): unknown {
  const tableAttrs: Record<string, string> = {
    class: "pm-table",
    contenteditable: "false",
  };
  const tableStyle = tableStyleString(data);
  if (tableStyle.length > 0) tableAttrs.style = tableStyle;

  const children: unknown[] = [];

  if (data.gridCols && data.gridCols.length > 0) {
    const colgroup: unknown[] = ["colgroup"];
    for (const w of data.gridCols) {
      if (w > 0) {
        colgroup.push(["col", { style: `width: ${twipsToInches(w).toFixed(3)}in` }]);
      } else {
        colgroup.push(["col", {}]);
      }
    }
    children.push(colgroup);
  }

  const tbody: unknown[] = ["tbody"];
  const rowCount = data.rows.length;
  for (let r = 0; r < rowCount; r++) {
    const row = data.rows[r];
    const tr: unknown[] = ["tr", { class: row.header ? "pm-table-row pm-table-header-row" : "pm-table-row" }];
    const visibleCells = row.cells.filter((c) => c.vMerge !== "continue");
    let visibleIndex = 0;
    for (let c = 0; c < row.cells.length; c++) {
      const cell = row.cells[c];
      if (cell.vMerge === "continue") continue;
      const colSpan = cell.gridSpan;
      const cellAttrs: Record<string, string> = { class: "pm-table-cell" };
      if (colSpan > 1) cellAttrs.colspan = String(colSpan);
      const style = cellStyleString(cell, data, r, rowCount, visibleIndex, visibleCells.length, colSpan);
      if (style.length > 0) cellAttrs.style = style;
      const cellChildren: unknown[] = cell.blocks.map(paragraphToDom);
      tr.push([row.header ? "th" : "td", cellAttrs, ...cellChildren]);
      visibleIndex++;
    }
    tbody.push(tr);
  }
  children.push(tbody);

  return ["table", tableAttrs, ...children];
}

export const docxSchema = new Schema({ nodes, marks });
