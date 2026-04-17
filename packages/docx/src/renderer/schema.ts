import { Schema, type MarkSpec, type NodeSpec } from "prosemirror-model";

/**
 * ProseMirror schema for the DOCX renderer. Mirrors `DocxSnapshot` 1:1
 * so that PM <-> model conversion is mechanical. See spec/docx/renderer.md.
 */

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
    parseDOM: [{ tag: "p" }],
    toDOM(node) {
      const cls = node.attrs.styleId ? `pm-p style-${String(node.attrs.styleId)}` : "pm-p";
      return ["p", { class: cls, "data-style": node.attrs.styleId ?? "" }, 0];
    },
  },
  table: {
    group: "block",
    atom: true,
    attrs: { tableId: { default: null }, rawJson: { default: null } },
    toDOM() {
      return ["div", { class: "pm-table-placeholder", contenteditable: "false" }, "[table]"];
    },
  },
  opaque_block: {
    group: "block",
    atom: true,
    attrs: { blockId: { default: null }, rawJson: { default: null }, tag: { default: null } },
    toDOM(node) {
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
    attrs: { inlineId: { default: null }, rawJson: { default: null } },
    toDOM() {
      return ["span", { class: "pm-opaque-inline" }, "[opaque]"];
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

export const docxSchema = new Schema({ nodes, marks });
