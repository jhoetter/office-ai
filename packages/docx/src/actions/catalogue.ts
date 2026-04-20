/**
 * Action catalogue for DOCX. The single source of truth tying
 * together the bus command type, the CLI subcommand, and the Cmd+K
 * palette entry for every public action.
 *
 * Seeded from the existing `office-agent docx …` subcommands in
 * `packages/agent/src/cli.ts` and the `paletteCommands` array in
 * `apps/web/app/editor/DocxEditor.tsx`. Each entry is a mechanical
 * translation — no behavioural change vs the hand-rolled wiring it
 * replaces.
 *
 * Add new public actions HERE first; the parity check
 * (`scripts/check-action-parity.mjs`) refuses to land a registered
 * command-bus handler that has no catalogue entry.
 */

import type { ActionDescriptor } from "@officeai/core";

/**
 * Free-form section taxonomy for DOCX. Kept stable so palette /
 * help-screen grouping doesn't churn between releases.
 */
export type DocxActionSection =
  | "File"
  | "Read"
  | "Edit"
  | "Insert"
  | "Format"
  | "Layout"
  | "Tables"
  | "Lists"
  | "Hyperlinks"
  | "Headers"
  | "Comments"
  | "Review"
  | "View"
  | "Mode"
  | "Collaboration";

export const docxActions: ReadonlyArray<ActionDescriptor> = [
  // ── File ──────────────────────────────────────────────────────────
  {
    id: "docx.create",
    commandType: null,
    label: "Create blank document",
    description: "Create a brand-new blank .docx file at --out (one empty paragraph, no styles part).",
    section: "File",
    surfaces: ["cli"],
    args: [{ name: "out", flag: "--out <path>", kind: "filepath", required: true, description: "Path to write the new .docx file" }],
  },
  {
    id: "docx.inspect",
    commandType: null,
    label: "Inspect document",
    description: "Print a structural summary (paragraphs, tables, comments, parts) as JSON.",
    section: "Read",
    surfaces: ["cli"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "withRuns", flag: "--with-runs", kind: "boolean", description: "Include per-paragraph run breakdown (offset, length, text) so callers can target precise text ranges" },
    ],
  },
  {
    id: "docx.read",
    commandType: null,
    label: "Read document",
    description: "Read a DOCX file as Markdown, structured JSON, or plain text.",
    section: "Read",
    surfaces: ["cli"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "format", flag: "--format <fmt>", kind: "enum", choices: ["markdown", "json", "text"], default: "markdown", description: "Output format" },
      { name: "range", flag: "--range <selector>", kind: "selector", description: "Selector e.g. paragraph:0..paragraph:5" },
      { name: "withTables", flag: "--with-tables", kind: "boolean", description: "JSON: include a top-level `tables` array with table id, dimensions, and per-cell paragraph ids/text" },
    ],
  },
  {
    id: "docx.search",
    commandType: null,
    label: "Search document",
    description: "Search a DOCX file for text and print matches as JSON.",
    section: "Read",
    surfaces: ["cli"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "query", flag: "-q, --query <text>", kind: "string", required: true, description: "Search query" },
      { name: "caseSensitive", flag: "--case-sensitive", kind: "boolean", description: "Case-sensitive search" },
      { name: "regex", flag: "--regex", kind: "boolean", description: "Treat the query as a regular expression" },
    ],
  },
  {
    id: "docx.diff",
    commandType: null,
    label: "Diff documents",
    description: "Compute a structural diff between two DOCX files.",
    section: "Read",
    surfaces: ["cli"],
    args: [
      { name: "before", flag: "--before <path>", kind: "filepath", required: true, description: "Path to the baseline .docx file" },
      { name: "after", flag: "--after <path>", kind: "filepath", required: true, description: "Path to the modified .docx file" },
    ],
  },

  // ── Text mutations ────────────────────────────────────────────────
  {
    id: "docx.insert-text",
    commandType: "docx:insert-text",
    label: "Insert text",
    description: "Insert text at a position selector and write the result.",
    section: "Edit",
    surfaces: ["cli"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "at", flag: "--at <selector>", kind: "selector", required: true, description: "Position selector e.g. section:0/paragraph:0/run:0/text:5" },
      { name: "text", flag: "--text <text>", kind: "string", required: true, description: "Text to insert" },
    ],
  },
  {
    id: "docx.delete-range",
    commandType: "docx:delete-range",
    label: "Delete range",
    description: "Delete a range of text. Range may span runs/paragraphs (e.g. paragraph:0/text:0..20 or paragraph:0..2).",
    section: "Edit",
    surfaces: ["cli"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "range", flag: "--range <selector>", kind: "selector", required: true, description: "Range selector e.g. paragraph:0/text:0..5 or paragraph:0..2" },
    ],
  },
  {
    id: "docx.replace-text",
    commandType: null,
    label: "Replace paragraph text",
    description: "Replace the entire text content of a paragraph (delete-range + insert-text in one mutation batch).",
    section: "Edit",
    surfaces: ["cli"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "paragraph", flag: "--paragraph <n>", kind: "number", required: true, description: "0-based paragraph index" },
      { name: "text", flag: "--text <text>", kind: "string", required: true, description: "New text content for the paragraph (may be empty)" },
    ],
  },
  {
    id: "docx.format-range",
    commandType: "docx:format-range",
    label: "Format text range",
    description: "Apply text formatting to a range (bold/italic/underline/strikethrough/font/size/color/highlight).",
    section: "Format",
    surfaces: ["cli"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "range", flag: "--range <selector>", kind: "selector", required: true, description: "Range selector e.g. paragraph:0/text:0..5 or paragraph:0..2" },
      { name: "bold", flag: "--bold <bool>", kind: "string", description: "true|false to set, omit to leave unchanged" },
      { name: "italic", flag: "--italic <bool>", kind: "string", description: "true|false" },
      { name: "underline", flag: "--underline <bool>", kind: "string", description: "true|false" },
      { name: "strike", flag: "--strike <bool>", kind: "string", description: "true|false" },
      { name: "fontFamily", flag: "--font-family <name>", kind: "string", description: "Run font family" },
      { name: "fontSize", flag: "--font-size <halfPoints>", kind: "number", description: "Run font size in half-points (e.g. 24 = 12pt)" },
      { name: "color", flag: "--color <RRGGBB>", kind: "string", description: "Hex color without leading #" },
      { name: "highlight", flag: "--highlight <name>", kind: "string", description: "Highlight color name (yellow|green|cyan|red|...)" },
    ],
  },
  {
    id: "docx.style",
    commandType: "docx:set-paragraph-style",
    label: "Set paragraph style",
    description: "Set the paragraph style at a position selector and write the result.",
    section: "Format",
    surfaces: ["cli"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "at", flag: "--at <selector>", kind: "selector", required: true, description: "Position selector targeting a paragraph" },
      { name: "style", flag: "--style <styleId>", kind: "string", required: true, description: "Style id, e.g. Heading1" },
    ],
  },
  {
    id: "docx.insert-paragraph",
    commandType: "docx:insert-paragraph",
    label: "Insert paragraph",
    description: "Insert a new paragraph at the given position selector. Optional --style applies a paragraph style id.",
    section: "Edit",
    surfaces: ["cli"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "at", flag: "--at <selector>", kind: "selector", required: true, description: "Position selector targeting a paragraph (e.g. paragraph:0)" },
      { name: "style", flag: "--style <styleId>", kind: "string", description: "Paragraph style id to apply (e.g. Heading1, Heading2, ListParagraph)" },
    ],
  },

  // ── Comments ──────────────────────────────────────────────────────
  {
    id: "docx.add-comment",
    commandType: "docx:add-comment",
    label: "Add comment",
    description: "Add a comment to a range selector and write the result.",
    section: "Comments",
    surfaces: ["cli", "palette", "toolbar"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "range", flag: "--range <selector>", kind: "selector", required: true, description: "Range selector e.g. paragraph:0/text:0..5" },
      { name: "text", flag: "--text <text>", kind: "string", required: true, description: "Comment text" },
      { name: "author", flag: "--author <name>", kind: "string", default: "office-agent", description: "Comment author" },
      { name: "initials", flag: "--initials <initials>", kind: "string", default: "OA", description: "Comment author initials" },
    ],
  },
  {
    id: "docx.resolve-comment",
    commandType: "docx:resolve-comment",
    label: "Resolve comment",
    description: "Mark a comment as resolved (or re-open with --reopen) and write the result.",
    section: "Comments",
    surfaces: ["cli"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "id", flag: "--id <commentId>", kind: "string", required: true, description: "Target comment id (as exposed by `docx inspect`)" },
      { name: "reopen", flag: "--reopen", kind: "boolean", description: "Re-open a previously resolved comment instead of resolving it" },
    ],
  },
  {
    id: "docx.reply-comment",
    commandType: "docx:reply-comment",
    label: "Reply to comment",
    description: "Append a reply to an existing comment and write the result.",
    section: "Comments",
    surfaces: ["cli"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "parent", flag: "--parent <commentId>", kind: "string", required: true, description: "Parent comment id" },
      { name: "text", flag: "--text <text>", kind: "string", required: true, description: "Reply text" },
      { name: "author", flag: "--author <name>", kind: "string", default: "office-agent", description: "Reply author" },
      { name: "initials", flag: "--initials <initials>", kind: "string", description: "Reply author initials" },
    ],
  },
  {
    id: "docx.delete-comment",
    commandType: "docx:delete-comment",
    label: "Delete comment",
    description: "Delete a comment (and its reply thread) and write the result.",
    section: "Comments",
    surfaces: ["cli"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "id", flag: "--id <commentId>", kind: "string", required: true, description: "Target comment id" },
    ],
  },

  // ── Tracked changes ───────────────────────────────────────────────
  {
    id: "docx.accept-change",
    commandType: "docx:accept-change",
    label: "Accept tracked change",
    description: "Accept a tracked change (insertion folds in, deletion lands).",
    section: "Review",
    surfaces: ["cli"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "id", flag: "--id <revisionId>", kind: "string", required: true, description: "Tracked-change revision id (the w:id on <w:ins>/<w:del>)" },
    ],
  },
  {
    id: "docx.reject-change",
    commandType: "docx:reject-change",
    label: "Reject tracked change",
    description: "Reject a tracked change (insertion is dropped, deletion is undone).",
    section: "Review",
    surfaces: ["cli"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "id", flag: "--id <revisionId>", kind: "string", required: true, description: "Tracked-change revision id" },
    ],
  },

  // ── Insert ────────────────────────────────────────────────────────
  {
    id: "docx.insert-image",
    commandType: "docx:insert-image",
    label: "Insert image",
    description: "Insert an inline image at a position selector. Reads image bytes from --image.",
    section: "Insert",
    surfaces: ["cli", "toolbar", "palette"],
    icon: "Image",
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "at", flag: "--at <selector>", kind: "selector", required: true, description: "Position selector targeting a paragraph" },
      { name: "image", flag: "--image <path>", kind: "filepath", required: true, description: "Path to image file (PNG/JPEG/GIF/BMP/TIFF/WebP/SVG)" },
      { name: "width", flag: "--width <px>", kind: "number", default: 200, description: "Display width in pixels @ 96 DPI (default: 200)" },
      { name: "height", flag: "--height <px>", kind: "number", default: 200, description: "Display height in pixels @ 96 DPI (default: 200)" },
      { name: "mime", flag: "--mime <type>", kind: "string", description: "Override mime type (auto-detected from extension by default)" },
      { name: "alt", flag: "--alt <text>", kind: "string", description: "Alt text (populates wp:docPr@descr)" },
      { name: "name", flag: "--name <name>", kind: "string", description: "Display name (wp:docPr@name)" },
    ],
  },
  {
    id: "docx.insert-table",
    commandType: "docx:insert-table",
    label: "Insert table",
    description: "Insert an empty rows × cols table at a position selector.",
    section: "Tables",
    surfaces: ["cli", "palette"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "at", flag: "--at <selector>", kind: "selector", required: true, description: "Position selector targeting a paragraph" },
      { name: "rows", flag: "--rows <n>", kind: "number", required: true, description: "Row count" },
      { name: "cols", flag: "--cols <n>", kind: "number", required: true, description: "Column count" },
      { name: "colWidths", flag: "--col-widths <csv>", kind: "string", description: "Optional comma-separated column widths in twips" },
    ],
  },
  {
    id: "docx.insert-row",
    commandType: "docx:insert-row",
    label: "Insert table row",
    description: "Insert a row into a table.",
    section: "Tables",
    surfaces: ["cli"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "tableId", flag: "--table-id <id>", kind: "string", required: true, description: "Target table id (use `docx inspect` to find ids)" },
      { name: "at", flag: "--at <n>", kind: "number", required: true, description: "0-based row index; equal to row count to append" },
    ],
  },
  {
    id: "docx.insert-column",
    commandType: "docx:insert-column",
    label: "Insert table column",
    description: "Insert a column into a table.",
    section: "Tables",
    surfaces: ["cli"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "tableId", flag: "--table-id <id>", kind: "string", required: true, description: "Target table id" },
      { name: "at", flag: "--at <n>", kind: "number", required: true, description: "0-based column index" },
      { name: "width", flag: "--width <twips>", kind: "number", description: "Column width in twips (defaults to equal split)" },
    ],
  },
  {
    id: "docx.set-cell-text",
    commandType: "docx:set-cell-content",
    label: "Set table cell text",
    description: "Replace one table cell's content with a single plain-text paragraph. Discover --table-id values via `docx read --format json --with-tables`.",
    section: "Tables",
    surfaces: ["cli"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "tableId", flag: "--table-id <id>", kind: "string", required: true, description: "Target table id (see `docx read --format json --with-tables`)" },
      { name: "row", flag: "--row <n>", kind: "number", required: true, description: "0-based row index" },
      { name: "col", flag: "--col <n>", kind: "number", required: true, description: "0-based column index" },
      { name: "text", flag: "--text <text>", kind: "string", required: true, description: "Text to place in the cell" },
    ],
  },

  // ── Charts ────────────────────────────────────────────────────────
  {
    id: "docx.insert-chart",
    commandType: "docx:insert-chart",
    label: "Insert chart",
    description: "Insert an inline chart at a position selector. Embeds a backing .xlsx workbook so 'Edit Data' works in Word.",
    section: "Insert",
    surfaces: ["cli", "palette"],
    icon: "BarChart3",
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "at", flag: "--at <selector>", kind: "selector", required: true, description: "Position selector targeting a paragraph" },
      { name: "chartType", flag: "--chart-type <type>", kind: "enum", choices: ["bar", "line", "pie", "area"], default: "bar", description: "Chart type" },
      { name: "title", flag: "--title <text>", kind: "string", description: "Chart title" },
      { name: "categories", flag: "--categories <csv>", kind: "string", required: true, description: "Comma-separated category labels (x-axis)" },
      { name: "series", flag: "--series <json>", kind: "string", required: true, description: "JSON array: [{name?, values:number[]}, ...]. Each series.values length must equal --categories length" },
      { name: "width", flag: "--width <px>", kind: "number", default: 480, description: "Display width in pixels @ 96 DPI" },
      { name: "height", flag: "--height <px>", kind: "number", default: 320, description: "Display height in pixels @ 96 DPI" },
    ],
  },
  {
    id: "docx.set-chart-data",
    commandType: "docx:set-chart-data",
    label: "Set chart data",
    description: "Replace categories + series for an existing chart and refresh the embedded workbook.",
    section: "Insert",
    surfaces: ["cli"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "chartPartPath", flag: "--chart-part-path <path>", kind: "string", required: true, description: "Chart part path, e.g. word/charts/chart1.xml" },
      { name: "categories", flag: "--categories <csv>", kind: "string", required: true, description: "Comma-separated category labels" },
      { name: "series", flag: "--series <json>", kind: "string", required: true, description: "JSON array: [{name?, values:number[]}, ...]" },
    ],
  },
  {
    id: "docx.set-chart-title",
    commandType: "docx:set-chart-title",
    label: "Set chart title",
    description: "Set / clear the title of an existing chart.",
    section: "Insert",
    surfaces: ["cli"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "chartPartPath", flag: "--chart-part-path <path>", kind: "string", required: true, description: "Chart part path" },
      { name: "title", flag: "--title <text>", kind: "string", description: "New title text; omit (or pass --clear) to remove" },
      { name: "clear", flag: "--clear", kind: "boolean", description: "Clear the existing title" },
    ],
  },
  {
    id: "docx.set-chart-type",
    commandType: "docx:set-chart-type",
    label: "Set chart type",
    description: "Switch the active plot type of an existing chart (bar/line/pie/area).",
    section: "Insert",
    surfaces: ["cli"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "chartPartPath", flag: "--chart-part-path <path>", kind: "string", required: true, description: "Chart part path" },
      { name: "chartType", flag: "--chart-type <type>", kind: "enum", choices: ["bar", "line", "pie", "area"], required: true, description: "Chart type" },
    ],
  },
  {
    id: "docx.insert-spreadsheet",
    commandType: "docx:insert-spreadsheet",
    label: "Insert spreadsheet",
    description: "Insert a live OLE-embedded Excel workbook at a position selector. Double-clicking the embed in Word opens the underlying .xlsx in Excel.",
    section: "Insert",
    surfaces: ["cli", "palette"],
    icon: "Table2",
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "at", flag: "--at <selector>", kind: "selector", required: true, description: "Position selector targeting a paragraph" },
      { name: "data", flag: "--data <json>", kind: "string", required: true, description: "JSON 2D array of cell values: [[\"A1\", \"B1\"], [1, 2], …]" },
      { name: "sheetName", flag: "--sheet-name <name>", kind: "string", default: "Sheet1", description: "Worksheet name" },
    ],
  },
  // ── "Insert from xlsx" palette wrappers ──────────────────────────
  // These three actions don't dispatch a bus command directly; they
  // open the XlsxRangePickerDialog with a preselected mode and the
  // dialog's submit handler dispatches the right `docx:insert-*`
  // command via the shared `applyXlsxEmbed` dispatcher. Catalogued
  // here purely so they appear in Cmd+K with proper i18n + icons;
  // CLI users keep using the lower-level `docx:insert-spreadsheet`
  // / `docx:insert-chart` / `docx:insert-table` commands directly.
  {
    id: "docx.insert-table-from-xlsx",
    commandType: null,
    label: "Insert table from xlsx",
    description: "Pick a range from an .xlsx file and insert it as a native Word table.",
    section: "Insert",
    surfaces: ["palette"],
    icon: "Table",
  },
  {
    id: "docx.insert-spreadsheet-from-xlsx",
    commandType: null,
    label: "Insert spreadsheet (live)",
    description: "Pick a range from an .xlsx file and embed the workbook as a live OLE object.",
    section: "Insert",
    surfaces: ["palette"],
    icon: "Table2",
  },
  {
    id: "docx.insert-chart-from-xlsx",
    commandType: null,
    label: "Insert chart from xlsx",
    description: "Pick a range from an .xlsx file and project it into a chart (first row = series, first column = categories).",
    section: "Insert",
    surfaces: ["palette"],
    icon: "BarChart3",
  },
  {
    id: "docx.update-spreadsheet",
    commandType: "docx:update-spreadsheet",
    label: "Update spreadsheet",
    description: "Replace the bytes of an existing OLE-embedded Excel workbook (used by the editor's double-click → edit → save loop).",
    section: "Insert",
    surfaces: ["cli"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "embeddingPartPath", flag: "--embedding-part-path <path>", kind: "string", required: true, description: "Part path of the embedded workbook, e.g. word/embeddings/oleObject1.xlsx" },
      { name: "xlsx", flag: "--xlsx <path>", kind: "filepath", required: true, description: "Path to the new .xlsx file whose bytes will replace the embed" },
    ],
  },

  // ── Hyperlinks ────────────────────────────────────────────────────
  {
    id: "docx.insert-hyperlink",
    commandType: "docx:insert-hyperlink",
    label: "Insert hyperlink",
    description: "Wrap a flat-text range in a paragraph with a hyperlink.",
    section: "Hyperlinks",
    surfaces: ["cli", "palette"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "paragraphId", flag: "--paragraph-id <id>", kind: "string", required: true, description: "Target paragraph id (use `docx read --format json` to find ids)" },
      { name: "start", flag: "--start <n>", kind: "number", required: true, description: "Inclusive start offset" },
      { name: "end", flag: "--end <n>", kind: "number", required: true, description: "Exclusive end offset" },
      { name: "target", flag: "--target <url>", kind: "string", description: "External URL target (mutually exclusive with --anchor)" },
      { name: "anchor", flag: "--anchor <name>", kind: "string", description: "Internal bookmark anchor (mutually exclusive with --target)" },
      { name: "tooltip", flag: "--tooltip <text>", kind: "string", description: "Tooltip text" },
    ],
  },
  {
    id: "docx.remove-hyperlink",
    commandType: "docx:remove-hyperlink",
    label: "Remove hyperlink",
    description: "Unwrap a hyperlink from a paragraph (optionally reaping its relationship).",
    section: "Hyperlinks",
    surfaces: ["cli"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "paragraphId", flag: "--paragraph-id <id>", kind: "string", required: true, description: "Target paragraph id" },
      { name: "hyperlinkId", flag: "--hyperlink-id <id>", kind: "string", required: true, description: "Hyperlink node id" },
    ],
  },

  // ── Lists ─────────────────────────────────────────────────────────
  {
    id: "docx.set-list",
    commandType: "docx:set-paragraph-list",
    label: "Set paragraph list",
    description: "Set or replace numbering (list) on a paragraph.",
    section: "Lists",
    surfaces: ["cli"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "paragraphId", flag: "--paragraph-id <id>", kind: "string", required: true, description: "Target paragraph id" },
      { name: "numId", flag: "--num-id <n>", kind: "number", required: true, description: "Numbering instance id (matches w:num/@w:numId)" },
      { name: "ilvl", flag: "--ilvl <n>", kind: "number", default: 0, description: "0-based level within the abstract numbering definition" },
    ],
  },
  {
    id: "docx.remove-list",
    commandType: "docx:remove-paragraph-list",
    label: "Remove paragraph list",
    description: "Remove numbering (list) from a paragraph.",
    section: "Lists",
    surfaces: ["cli"],
    args: [
      { name: "paragraphId", flag: "--paragraph-id <id>", kind: "string", required: true, description: "Target paragraph id" },
    ],
    buildPayload: (parsed) => ({ paragraphId: parsed.paragraphId as string }),
  },

  // ── Layout ────────────────────────────────────────────────────────
  {
    id: "docx.align",
    commandType: "docx:set-paragraph-alignment",
    label: "Set paragraph alignment",
    description: "Set (or clear with --clear) a paragraph's <w:jc> alignment.",
    section: "Layout",
    surfaces: ["cli"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "paragraphId", flag: "--paragraph-id <id>", kind: "string", required: true, description: "Target paragraph id" },
      { name: "alignment", flag: "--alignment <value>", kind: "string", description: "left | center | right | justify" },
      { name: "clear", flag: "--clear", kind: "boolean", description: "Clear any existing alignment" },
    ],
  },
  {
    id: "docx.indent",
    commandType: "docx:set-paragraph-indent",
    label: "Set paragraph indent",
    description: "Step a paragraph's left indent by --delta twips (negative outdents).",
    section: "Layout",
    surfaces: ["cli"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "paragraphId", flag: "--paragraph-id <id>", kind: "string", required: true, description: "Target paragraph id" },
      { name: "delta", flag: "--delta <twips>", kind: "number", required: true, description: "Signed delta in twips applied to indentation.left" },
    ],
  },

  // ── Headers / Footers ─────────────────────────────────────────────
  {
    id: "docx.header",
    commandType: "docx:set-header-text",
    label: "Set header text",
    description: "Replace one header paragraph's text content.",
    section: "Headers",
    surfaces: ["cli"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "part", flag: "--part <path>", kind: "string", required: true, description: "Header part path, e.g. word/header1.xml" },
      { name: "paragraphIndex", flag: "--paragraph-index <n>", kind: "number", required: true, description: "0-based paragraph index inside the header" },
      { name: "text", flag: "--text <text>", kind: "string", required: true, description: "New plain-text content" },
    ],
  },
  {
    id: "docx.footer",
    commandType: "docx:set-footer-text",
    label: "Set footer text",
    description: "Replace one footer paragraph's text content.",
    section: "Headers",
    surfaces: ["cli"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "part", flag: "--part <path>", kind: "string", required: true, description: "Footer part path, e.g. word/footer1.xml" },
      { name: "paragraphIndex", flag: "--paragraph-index <n>", kind: "number", required: true, description: "0-based paragraph index inside the footer" },
      { name: "text", flag: "--text <text>", kind: "string", required: true, description: "New plain-text content" },
    ],
  },

  // ── Generic apply / pending ───────────────────────────────────────
  {
    id: "docx.apply",
    commandType: null,
    label: "Apply JSON commands",
    description: "Apply a JSON command file (single command or { commands: [...] }) and write the result. Pass -c/--commands to read from disk or --from-stdin to read JSON piped on stdin (mutually exclusive).",
    section: "Edit",
    surfaces: ["cli"],
    args: [
      { name: "file", flag: "--file <path>", kind: "filepath", required: true, description: "Path to a .docx file" },
      { name: "commands", flag: "-c, --commands <path>", kind: "filepath", description: "Path to a JSON file containing one or more commands" },
      { name: "fromStdin", flag: "--from-stdin", kind: "boolean", description: "Read the JSON command body from stdin instead of -c <path>" },
    ],
  },

  // ── Palette-only / non-mutating actions (UI surfaces) ─────────────
  // These wrap toolbar interactions so every toolbar button is also
  // discoverable via Cmd+K. They piggy-back on existing handlers
  // (apply-list-format, set-paragraph-alignment, set-paragraph-indent,
  // insert-section-break) that the CLI exposes via richer subcommands;
  // the palette versions take no arguments and operate on the current
  // selection.
  {
    id: "docx.bullet-list",
    commandType: null,
    label: "Bullet list",
    description: "Toggle a bullet list on the current paragraph.",
    section: "Lists",
    surfaces: ["palette"],
    icon: "List",
  },
  {
    id: "docx.ordered-list",
    commandType: null,
    label: "Numbered list",
    description: "Toggle a numbered list on the current paragraph.",
    section: "Lists",
    surfaces: ["palette"],
    icon: "ListOrdered",
  },
  {
    id: "docx.align-left",
    commandType: null,
    label: "Align left",
    description: "Left-align the current paragraph.",
    section: "Layout",
    surfaces: ["palette"],
    icon: "AlignLeft",
  },
  {
    id: "docx.align-center",
    commandType: null,
    label: "Align center",
    description: "Center-align the current paragraph.",
    section: "Layout",
    surfaces: ["palette"],
    icon: "AlignCenter",
  },
  {
    id: "docx.align-right",
    commandType: null,
    label: "Align right",
    description: "Right-align the current paragraph.",
    section: "Layout",
    surfaces: ["palette"],
    icon: "AlignRight",
  },
  {
    id: "docx.align-justify",
    commandType: null,
    label: "Justify",
    description: "Justify the current paragraph.",
    section: "Layout",
    surfaces: ["palette"],
    icon: "AlignJustify",
  },
  {
    id: "docx.indent-increase",
    commandType: null,
    label: "Increase indent",
    description: "Indent the current paragraph by one step (¼ inch).",
    section: "Layout",
    surfaces: ["palette"],
    icon: "Indent",
  },
  {
    id: "docx.indent-decrease",
    commandType: null,
    label: "Decrease indent",
    description: "Outdent the current paragraph by one step (¼ inch).",
    section: "Layout",
    surfaces: ["palette"],
    icon: "Outdent",
  },
  {
    id: "docx.section-break-next-page",
    commandType: null,
    label: "Insert section break (next page)",
    description: "Insert a next-page section break after the current paragraph.",
    section: "Layout",
    surfaces: ["palette"],
  },
  {
    id: "docx.section-break-continuous",
    commandType: null,
    label: "Insert section break (continuous)",
    description: "Insert a continuous section break after the current paragraph.",
    section: "Layout",
    surfaces: ["palette"],
  },
  {
    id: "docx.insert-table-3x3",
    commandType: null,
    label: "Insert table (3 × 3)",
    description: "Insert a 3 × 3 table at the current selection.",
    section: "Insert",
    surfaces: ["palette"],
  },
  {
    id: "docx.insert-table-2x2",
    commandType: null,
    label: "Insert table (2 × 2)",
    description: "Insert a 2 × 2 table at the current selection.",
    section: "Insert",
    surfaces: ["palette"],
  },
  {
    id: "docx.page-setup",
    commandType: null,
    label: "Page setup…",
    description: "Open the page setup dialog to configure margins, paper size, and orientation.",
    section: "Layout",
    surfaces: ["palette"],
  },
  {
    id: "docx.toggle-marks",
    commandType: null,
    label: "Toggle formatting marks",
    description: "Show or hide non-printing characters (paragraph marks, tabs, spaces).",
    section: "View",
    surfaces: ["palette"],
  },
  {
    id: "docx.set-mode-edit",
    commandType: null,
    label: "Switch to Editing mode",
    description: "Direct edits commit to the document.",
    section: "Mode",
    surfaces: ["palette"],
  },
  {
    id: "docx.set-mode-suggest",
    commandType: null,
    label: "Switch to Suggesting mode",
    description: "Edits land as tracked changes for review.",
    section: "Mode",
    surfaces: ["palette"],
  },
  {
    id: "docx.set-mode-view",
    commandType: null,
    label: "Switch to Viewing mode",
    description: "Read-only mode — block all mutations.",
    section: "Mode",
    surfaces: ["palette"],
  },

  // ── Handlers without a public surface (still tracked) ─────────────
  // The bus exposes these handlers but the CLI / palette currently do
  // not — they are reached only via the `docx apply` JSON escape hatch
  // or via the in-editor dialogs that wrap them. The catalogue still
  // lists them so the parity check accepts the handler registration;
  // each entry carries a `hidden.reason` documenting why.
  {
    id: "docx.apply-list-format",
    commandType: "docx:apply-list-format",
    label: "Apply list format",
    description: "Toggle bullet/ordered list on a range (used by the toolbar list menu).",
    section: "Lists",
    surfaces: [],
    hidden: { reason: "Toolbar-only — invoked via the bullet/ordered-list menu in DocxEditor; CLI users compose set-list / remove-list directly." },
  },
  {
    id: "docx.accept-all-changes",
    commandType: "docx:accept-all-changes",
    label: "Accept all changes",
    description: "Accept every pending tracked change in one mutation.",
    section: "Review",
    surfaces: [],
    hidden: { reason: "Reachable today via the tracked-changes UI; CLI exposure is deferred." },
  },
  {
    id: "docx.reject-all-changes",
    commandType: "docx:reject-all-changes",
    label: "Reject all changes",
    description: "Reject every pending tracked change in one mutation.",
    section: "Review",
    surfaces: [],
    hidden: { reason: "Reachable today via the tracked-changes UI; CLI exposure is deferred." },
  },
  {
    id: "docx.insert-text-tracked",
    commandType: "docx:insert-text-tracked",
    label: "Insert text (tracked)",
    description: "Insert text as a tracked-change insertion (suggesting mode).",
    section: "Edit",
    surfaces: [],
    hidden: { reason: "Used implicitly when the editor is in suggesting mode; CLI users pick docx:insert-text directly." },
  },
  {
    id: "docx.delete-range-tracked",
    commandType: "docx:delete-range-tracked",
    label: "Delete range (tracked)",
    description: "Delete a range as a tracked-change deletion (suggesting mode).",
    section: "Edit",
    surfaces: [],
    hidden: { reason: "Used implicitly when the editor is in suggesting mode; CLI users pick docx:delete-range directly." },
  },
  {
    id: "docx.set-paragraph-spacing",
    commandType: "docx:set-paragraph-spacing",
    label: "Set paragraph spacing",
    description: "Set line spacing / before / after on a paragraph.",
    section: "Layout",
    surfaces: [],
    hidden: { reason: "Wired to the toolbar spacing dropdown; CLI exposure is deferred until the flag matrix stabilises." },
  },
  {
    id: "docx.insert-page-number",
    commandType: "docx:insert-page-number",
    label: "Insert page number",
    description: "Insert a page-number field at a position selector.",
    section: "Insert",
    surfaces: [],
    hidden: { reason: "Reached via the header/footer toolbar; CLI exposure is deferred." },
  },
  {
    id: "docx.set-section-different-first",
    commandType: "docx:set-section-different-first",
    label: "Set 'different first page'",
    description: "Toggle the first-page header/footer split on a section.",
    section: "Layout",
    surfaces: [],
    hidden: { reason: "Reached via the page-setup dialog; CLI exposure is deferred." },
  },
  {
    id: "docx.set-page-setup",
    commandType: "docx:set-page-setup",
    label: "Set page setup",
    description: "Update page size / margins / orientation on a section.",
    section: "Layout",
    surfaces: [],
    hidden: { reason: "Reached via the page-setup dialog; CLI exposure is deferred." },
  },
  {
    id: "docx.set-image-properties",
    commandType: "docx:set-image-properties",
    label: "Set image properties",
    description: "Update inline image alt text, name, or size.",
    section: "Insert",
    surfaces: [],
    hidden: { reason: "Reached via the image context toolbar; CLI exposure is deferred." },
  },
  {
    id: "docx.delete-image",
    commandType: "docx:delete-image",
    label: "Delete image",
    description: "Remove an inline image from a paragraph.",
    section: "Insert",
    surfaces: [],
    hidden: { reason: "Reached via the image context toolbar; CLI exposure is deferred." },
  },
  {
    id: "docx.insert-section-break",
    commandType: "docx:insert-section-break",
    label: "Insert section break",
    description: "Insert a section break (next page / continuous / even / odd).",
    section: "Layout",
    surfaces: [],
    hidden: { reason: "Reached via the toolbar Insert ▸ Breaks menu; CLI exposure is deferred." },
  },
  {
    id: "docx.insert-page-break",
    commandType: "docx:insert-page-break",
    label: "Insert page break",
    description: "Insert a hard page break.",
    section: "Layout",
    surfaces: [],
    hidden: { reason: "Reached via the toolbar Insert ▸ Breaks menu; CLI exposure is deferred." },
  },
];
