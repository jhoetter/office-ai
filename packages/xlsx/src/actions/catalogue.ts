/**
 * Action catalogue for XLSX. Single source of truth tying together
 * each command-bus handler, its CLI subcommand, and its Cmd+K palette
 * entry.
 *
 * Seeded mechanically from `packages/agent/src/cli-xlsx.ts` and
 * `apps/web/app/xlsx-editor/XlsxEditor.tsx`. Add new public actions
 * HERE first; the parity check refuses to land a handler with no
 * catalogue entry.
 *
 * Notes on `cliKind`:
 *   - "dispatch" (default): The CLI adapter generates a subcommand
 *     that loads the workbook, dispatches `commandType` with the
 *     payload built by `buildPayload`, and writes the result. Common
 *     write flags (`--source`, `--agent-id`, `--no-approve`, `--out`,
 *     `--pretty`) are added automatically.
 *   - "custom": Hand-rolled commander block in `cli-xlsx.ts` retains
 *     ownership (custom batch logic, stdin handling, multi-step
 *     pipelines). The catalogue still owns the label/description and
 *     the parity check verifies a subcommand by that name exists.
 */

import type { ActionDescriptor } from "@officeai/core";

export const xlsxActions: ReadonlyArray<ActionDescriptor> = [
  // ── File / Read (no commandType — they don't mutate) ──────────────
  {
    id: "xlsx.create",
    commandType: null,
    label: "Create blank workbook",
    description: "Create a brand-new blank .xlsx file at --out (one empty worksheet named Sheet1).",
    section: "File",
    surfaces: ["cli"],
  },
  {
    id: "xlsx.inspect",
    commandType: null,
    label: "Inspect workbook",
    description: "Print a structural summary (sheets, cells, parts, comments, merges) as JSON.",
    section: "Read",
    surfaces: ["cli"],
  },
  {
    id: "xlsx.read",
    commandType: null,
    label: "Read sheet / range",
    description: "Render a sheet (or range) as Markdown or JSON.",
    section: "Read",
    surfaces: ["cli"],
  },
  {
    id: "xlsx.search",
    commandType: null,
    label: "Search workbook",
    description: "Search workbook cells for text and print matches as JSON.",
    section: "Read",
    surfaces: ["cli"],
  },
  {
    id: "xlsx.diff",
    commandType: null,
    label: "Diff workbooks",
    description: "Compute a structural diff between two on-disk .xlsx files.",
    section: "Read",
    surfaces: ["cli"],
  },
  {
    id: "xlsx.apply-file",
    commandType: null,
    label: "Apply JSON commands",
    description: "Apply a JSON command file (single command or { commands: [...] }) and write the result.",
    section: "Edit",
    surfaces: ["cli"],
  },
  {
    id: "xlsx.apply",
    commandType: null,
    label: "Apply raw command",
    description:
      "Generic command escape hatch — pass --type and --payload (JSON), or --type with --payload-stdin.",
    section: "Edit",
    surfaces: ["cli"],
  },
  {
    id: "xlsx.fill-formula",
    commandType: null,
    label: "Fill formula across range",
    description:
      "Fill a formula across an A1 range, expanding {row}/{col}/{rowN}/{colN} placeholders per cell.",
    section: "Edit",
    surfaces: ["cli"],
  },
  {
    id: "xlsx.clear-range",
    commandType: null,
    label: "Clear range",
    description: "Wipe every cell in an A1 range to null in a single round-trip.",
    section: "Edit",
    surfaces: ["cli"],
  },

  // ── Cell value / formula / format ─────────────────────────────────
  {
    id: "xlsx.set-cell",
    commandType: "xlsx:set-cell-value",
    label: "Set cell value",
    description: "Set a single cell's literal value (string/number/boolean/null/error).",
    section: "Edit",
    surfaces: ["cli", "palette"],
  },
  {
    id: "xlsx.set-formula",
    commandType: "xlsx:set-cell-formula",
    label: "Set formula",
    description: "Set a single cell's formula (with or without leading '=').",
    section: "Edit",
    surfaces: ["cli", "palette"],
  },
  {
    id: "xlsx.set-range",
    commandType: "xlsx:set-range-values",
    label: "Set range values",
    description: "Set a row-major 2-D matrix of cell values across a range.",
    section: "Edit",
    surfaces: ["cli"],
  },
  {
    id: "xlsx.set-format",
    commandType: "xlsx:set-cell-format",
    label: "Apply cell format",
    description: "Apply a format patch to a range (font/fill/border/alignment/numberFormat). Pass --format as JSON.",
    section: "Format",
    surfaces: ["cli", "palette"],
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Target sheet name" },
      { name: "range", flag: "--range <a1>", kind: "string", required: true, description: "A1 cell or range, e.g. A1 or A1:E5" },
      { name: "format", flag: "--format <json>", kind: "string", required: true, description: "JSON CellFormatPatch (font/fill/border/alignment/numberFormat)" },
    ],
    buildPayload: ({ sheet, range, format }) => ({
      sheet: String(sheet),
      range: String(range),
      format: typeof format === "string" ? JSON.parse(format) : format,
    }),
  },
  {
    id: "xlsx.set-wrap-text",
    commandType: "xlsx:set-cell-format",
    label: "Toggle wrap text",
    description: "Toggle wrap-text on a range. Wrapped cells display long content on multiple lines instead of overflowing.",
    section: "Format",
    surfaces: ["cli", "palette", "toolbar"],
    icon: "WrapText",
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Target sheet name" },
      { name: "range", flag: "--range <a1>", kind: "string", required: true, description: "A1 cell or range, e.g. A1 or A1:E5" },
      { name: "wrap", flag: "--wrap <bool>", kind: "boolean", default: true, description: "true to enable, false to disable wrapping" },
    ],
    buildPayload: ({ sheet, range, wrap }) => ({
      sheet: String(sheet),
      range: String(range),
      format: { alignment: { wrapText: wrap === false || wrap === "false" ? false : true } },
    }),
  },

  // ── Sheet management ──────────────────────────────────────────────
  {
    id: "xlsx.add-sheet",
    commandType: "xlsx:add-sheet",
    label: "Add sheet",
    description: "Append (or insert at --at) a new worksheet.",
    section: "Sheet",
    surfaces: ["cli", "palette"],
  },
  {
    id: "xlsx.rename-sheet",
    commandType: "xlsx:rename-sheet",
    label: "Rename sheet",
    description: "Rename a sheet by current name.",
    section: "Sheet",
    surfaces: ["cli", "palette"],
  },
  {
    id: "xlsx.delete-sheet",
    commandType: "xlsx:delete-sheet",
    label: "Delete sheet",
    description: "Delete a worksheet by name (the workbook must keep ≥1 visible sheet).",
    section: "Sheet",
    surfaces: ["cli", "palette"],
  },
  {
    id: "xlsx.move-sheet",
    commandType: "xlsx:move-sheet",
    label: "Move sheet",
    description: "Reorder a sheet within the workbook.",
    section: "Sheet",
    surfaces: ["palette"],
    hidden: { reason: "Reachable via the sheet tabs context menu; CLI exposure is deferred." },
  },
  {
    id: "xlsx.set-sheet-state",
    commandType: "xlsx:set-sheet-state",
    label: "Set sheet visibility",
    description: "Hide / show / very-hide a worksheet.",
    section: "Sheet",
    surfaces: ["palette"],
    hidden: { reason: "Reachable via the sheet tabs context menu; CLI exposure is deferred." },
  },
  {
    id: "xlsx.set-sheet-tab-color",
    commandType: "xlsx:set-sheet-tab-color",
    label: "Set sheet tab color",
    description: "Paint the sheet-tab strip with a custom color (or clear it).",
    section: "Sheet",
    surfaces: ["cli", "contextMenu"],
    icon: "Palette",
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Target sheet name" },
      {
        name: "color",
        flag: "--color <argb>",
        kind: "string",
        required: false,
        description: "ARGB hex (e.g. FFCC0000); omit or pass empty string to clear",
      },
    ],
    buildPayload: ({ sheet, color }) => ({
      name: String(sheet),
      color: color === undefined || color === null || color === "" ? null : String(color),
    }),
  },

  // ── Rows / columns / merges ───────────────────────────────────────
  {
    id: "xlsx.insert-row",
    commandType: "xlsx:insert-row",
    label: "Insert rows",
    description: "Insert N blank rows BEFORE a 1-based row index.",
    section: "Edit",
    surfaces: ["cli", "palette", "toolbar", "contextMenu"],
    icon: "Plus",
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Target sheet name" },
      { name: "at", flag: "--at <row>", kind: "number", required: true, description: "1-based row index — insertion happens BEFORE this row" },
      { name: "count", flag: "--count <n>", kind: "number", default: 1, description: "Number of blank rows to insert" },
    ],
    buildPayload: ({ sheet, at, count }) => ({
      sheet: String(sheet),
      at: Number(at),
      count: Number(count ?? 1),
    }),
  },
  {
    id: "xlsx.insert-column",
    commandType: "xlsx:insert-column",
    label: "Insert columns",
    description: "Insert N blank columns BEFORE a 1-based column index (A=1).",
    section: "Edit",
    surfaces: ["cli", "palette", "toolbar", "contextMenu"],
    icon: "Plus",
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Target sheet name" },
      { name: "at", flag: "--at <col>", kind: "number", required: true, description: "1-based column index (A=1)" },
      { name: "count", flag: "--count <n>", kind: "number", default: 1, description: "Number of blank columns to insert" },
    ],
    buildPayload: ({ sheet, at, count }) => ({
      sheet: String(sheet),
      at: Number(at),
      count: Number(count ?? 1),
    }),
  },
  {
    id: "xlsx.delete-row",
    commandType: "xlsx:delete-row",
    label: "Delete rows",
    description: "Delete N rows starting at a 1-based row index.",
    section: "Edit",
    surfaces: ["cli", "palette", "toolbar", "contextMenu"],
    icon: "Minus",
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Target sheet name" },
      { name: "at", flag: "--at <row>", kind: "number", required: true, description: "1-based row index of the first row to drop" },
      { name: "count", flag: "--count <n>", kind: "number", default: 1, description: "Number of rows to drop" },
    ],
    buildPayload: ({ sheet, at, count }) => ({
      sheet: String(sheet),
      at: Number(at),
      count: Number(count ?? 1),
    }),
  },
  {
    id: "xlsx.delete-column",
    commandType: "xlsx:delete-column",
    label: "Delete columns",
    description: "Delete N columns starting at a 1-based column index (A=1).",
    section: "Edit",
    surfaces: ["cli", "palette", "toolbar", "contextMenu"],
    icon: "Minus",
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Target sheet name" },
      { name: "at", flag: "--at <col>", kind: "number", required: true, description: "1-based column index (A=1)" },
      { name: "count", flag: "--count <n>", kind: "number", default: 1, description: "Number of columns to drop" },
    ],
    buildPayload: ({ sheet, at, count }) => ({
      sheet: String(sheet),
      at: Number(at),
      count: Number(count ?? 1),
    }),
  },
  {
    id: "xlsx.merge-cells",
    commandType: "xlsx:merge-cells",
    label: "Merge cells",
    description: "Merge an A1 range covering ≥2 cells.",
    section: "Format",
    surfaces: ["cli", "palette", "toolbar"],
    icon: "Combine",
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Target sheet name" },
      { name: "range", flag: "--range <a1>", kind: "string", required: true, description: "A1 range, e.g. A1:B2" },
    ],
    buildPayload: ({ sheet, range }) => ({ sheet: String(sheet), range: String(range) }),
  },
  {
    id: "xlsx.unmerge-cells",
    commandType: "xlsx:unmerge-cells",
    label: "Unmerge cells",
    description: "Unmerge an existing merged range (must match exactly).",
    section: "Format",
    surfaces: ["cli", "palette", "toolbar"],
    icon: "Split",
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Target sheet name" },
      { name: "range", flag: "--range <a1>", kind: "string", required: true, description: "A1 range matching an existing merge exactly" },
    ],
    buildPayload: ({ sheet, range }) => ({ sheet: String(sheet), range: String(range) }),
  },
  {
    id: "xlsx.set-column-width",
    commandType: "xlsx:set-column-width",
    label: "Set column width",
    description: "Override a column's width in CSS pixels (pass --reset to clear).",
    section: "Format",
    surfaces: ["cli", "palette", "toolbar"],
    icon: "MoveHorizontal",
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Target sheet name" },
      { name: "column", flag: "--column <n>", kind: "number", required: true, description: "1-based column index (A=1)" },
      { name: "width", flag: "--width <px>", kind: "number", description: "Width in CSS pixels" },
      { name: "reset", flag: "--reset", kind: "boolean", description: "Reset to default width" },
    ],
    buildPayload: ({ sheet, column, width, reset }) => ({
      sheet: String(sheet),
      column: Number(column),
      width: reset ? null : Number(width),
    }),
  },
  {
    id: "xlsx.set-row-height",
    commandType: "xlsx:set-row-height",
    label: "Set row height",
    description: "Override a row's height in CSS pixels (pass --reset to clear).",
    section: "Format",
    surfaces: ["cli", "palette", "toolbar"],
    icon: "MoveVertical",
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Target sheet name" },
      { name: "row", flag: "--row <n>", kind: "number", required: true, description: "1-based row index" },
      { name: "height", flag: "--height <px>", kind: "number", description: "Height in CSS pixels" },
      { name: "reset", flag: "--reset", kind: "boolean", description: "Reset to default height" },
    ],
    buildPayload: ({ sheet, row, height, reset }) => ({
      sheet: String(sheet),
      row: Number(row),
      height: reset ? null : Number(height),
    }),
  },
  {
    id: "xlsx.set-row-visibility",
    commandType: "xlsx:set-row-visibility",
    label: "Hide / unhide row",
    description: "Show or hide a single row (row-header context-menu parity).",
    section: "Format",
    surfaces: ["cli", "contextMenu"],
    icon: "EyeOff",
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Target sheet name" },
      { name: "row", flag: "--row <n>", kind: "number", required: true, description: "1-based row index" },
      { name: "hidden", flag: "--hidden <bool>", kind: "boolean", required: true, description: "true = hide, false = show" },
    ],
    buildPayload: ({ sheet, row, hidden }) => ({
      sheet: String(sheet),
      row: Number(row),
      hidden: Boolean(hidden),
    }),
  },
  {
    id: "xlsx.set-column-visibility",
    commandType: "xlsx:set-column-visibility",
    label: "Hide / unhide column",
    description: "Show or hide a single column (column-header context-menu parity).",
    section: "Format",
    surfaces: ["cli", "contextMenu"],
    icon: "EyeOff",
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Target sheet name" },
      { name: "column", flag: "--column <n>", kind: "number", required: true, description: "1-based column index (A=1)" },
      { name: "hidden", flag: "--hidden <bool>", kind: "boolean", required: true, description: "true = hide, false = show" },
    ],
    buildPayload: ({ sheet, column, hidden }) => ({
      sheet: String(sheet),
      column: Number(column),
      hidden: Boolean(hidden),
    }),
  },

  // ── Comments / notes ──────────────────────────────────────────────
  {
    id: "xlsx.add-comment",
    commandType: "xlsx:add-comment",
    label: "Add comment",
    description: "Attach a classic note (single-cell anchor).",
    section: "Comments",
    surfaces: ["cli", "palette"],
  },
  {
    id: "xlsx.reply-comment",
    commandType: "xlsx:reply-comment",
    label: "Reply to comment",
    description: "Append a reply to an existing comment thread.",
    section: "Comments",
    surfaces: ["palette"],
    hidden: { reason: "Reached via the comment thread UI; CLI exposure is deferred." },
  },
  {
    id: "xlsx.resolve-comment",
    commandType: "xlsx:resolve-comment",
    label: "Resolve comment",
    description: "Mark a comment as resolved.",
    section: "Comments",
    surfaces: ["palette"],
    hidden: { reason: "Reached via the comment thread UI; CLI exposure is deferred." },
  },
  {
    id: "xlsx.edit-comment",
    commandType: "xlsx:edit-comment",
    label: "Edit comment",
    description: "Edit a comment's text.",
    section: "Comments",
    surfaces: ["palette"],
    hidden: { reason: "Reached via the comment thread UI; CLI exposure is deferred." },
  },
  {
    id: "xlsx.delete-comment",
    commandType: "xlsx:delete-comment",
    label: "Delete comment",
    description: "Delete a comment thread.",
    section: "Comments",
    surfaces: ["palette"],
    hidden: { reason: "Reached via the comment thread UI; CLI exposure is deferred." },
  },

  // ── Conditional formatting / validation / defined names ───────────
  {
    id: "xlsx.add-conditional-format",
    commandType: "xlsx:add-conditional-format",
    label: "Add conditional format",
    description: "Add a conditional formatting rule to a range. Pass --rule as JSON matching the ConditionalFormat schema.",
    section: "Format",
    surfaces: ["cli", "palette"],
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Target sheet name" },
      { name: "rule", flag: "--rule <json>", kind: "string", required: true, description: "JSON ConditionalFormat (id, range, type, …) — see model/types.ts" },
    ],
    buildPayload: ({ sheet, rule }) => ({
      sheet: String(sheet),
      rule: typeof rule === "string" ? JSON.parse(rule) : rule,
    }),
  },
  {
    id: "xlsx.remove-conditional-format",
    commandType: "xlsx:remove-conditional-format",
    label: "Remove conditional format",
    description: "Remove a conditional formatting rule by id.",
    section: "Format",
    surfaces: ["cli", "palette"],
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Target sheet name" },
      { name: "id", flag: "--id <ruleId>", kind: "string", required: true, description: "Rule id (returned by add-conditional-format)" },
    ],
    buildPayload: ({ sheet, id }) => ({ sheet: String(sheet), id: String(id) }),
  },
  {
    id: "xlsx.clear-conditional-formats",
    commandType: "xlsx:clear-conditional-formats",
    label: "Clear conditional formats",
    description: "Drop every conditional formatting rule from a sheet.",
    section: "Format",
    surfaces: ["cli", "palette"],
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Target sheet name" },
    ],
    buildPayload: ({ sheet }) => ({ sheet: String(sheet) }),
  },
  {
    id: "xlsx.add-data-validation",
    commandType: "xlsx:add-data-validation",
    label: "Add data validation",
    description: "Add a data validation rule to a range. Pass --rule as JSON matching the DataValidation schema.",
    section: "Data",
    surfaces: ["cli", "palette"],
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Target sheet name" },
      { name: "rule", flag: "--rule <json>", kind: "string", required: true, description: "JSON DataValidation (id, range, type, …)" },
    ],
    buildPayload: ({ sheet, rule }) => ({
      sheet: String(sheet),
      rule: typeof rule === "string" ? JSON.parse(rule) : rule,
    }),
  },
  {
    id: "xlsx.remove-data-validation",
    commandType: "xlsx:remove-data-validation",
    label: "Remove data validation",
    description: "Remove a data validation rule by id.",
    section: "Data",
    surfaces: ["cli", "palette"],
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Target sheet name" },
      { name: "id", flag: "--id <ruleId>", kind: "string", required: true, description: "Rule id" },
    ],
    buildPayload: ({ sheet, id }) => ({ sheet: String(sheet), id: String(id) }),
  },
  {
    id: "xlsx.clear-data-validations",
    commandType: "xlsx:clear-data-validations",
    label: "Clear data validations",
    description: "Drop every data validation rule on a sheet.",
    section: "Data",
    surfaces: ["cli", "palette"],
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Target sheet name" },
    ],
    buildPayload: ({ sheet }) => ({ sheet: String(sheet) }),
  },
  {
    id: "xlsx.add-defined-name",
    commandType: "xlsx:add-defined-name",
    label: "Add defined name",
    description: "Define a workbook- or sheet-scoped name (named range).",
    section: "Data",
    surfaces: ["cli", "palette"],
    icon: "Bookmark",
    args: [
      { name: "name", flag: "--name <name>", kind: "string", required: true, description: "Display name (no spaces, ≤ 255 chars, not a cell ref)" },
      { name: "refersTo", flag: "--refers-to <expr>", kind: "string", required: true, description: 'OOXML refersTo expression without leading = (e.g. "Sheet1!$A$1:$B$5")' },
      { name: "scope", flag: "--scope <sheet>", kind: "string", description: "Sheet-name scope (omit for workbook-scoped)" },
      { name: "comment", flag: "--comment <text>", kind: "string", description: "Optional comment" },
    ],
    buildPayload: ({ name, refersTo, scope, comment }) => ({
      name: String(name),
      refersTo: String(refersTo),
      ...(scope !== undefined ? { scope: String(scope) } : {}),
      ...(comment !== undefined ? { comment: String(comment) } : {}),
    }),
  },
  {
    id: "xlsx.update-defined-name",
    commandType: "xlsx:update-defined-name",
    label: "Update defined name",
    description: "Update an existing defined name's reference, scope, or rename it.",
    section: "Data",
    surfaces: ["cli", "palette"],
    args: [
      { name: "name", flag: "--name <name>", kind: "string", required: true, description: "Existing name to update" },
      { name: "scope", flag: "--scope <sheet>", kind: "string", description: "Existing scope (omit for workbook-scoped)" },
      { name: "nextName", flag: "--next-name <name>", kind: "string", description: "New name (rename)" },
      { name: "refersTo", flag: "--refers-to <expr>", kind: "string", description: "New refersTo expression" },
      { name: "comment", flag: "--comment <text>", kind: "string", description: "New comment" },
    ],
    buildPayload: ({ name, scope, nextName, refersTo, comment }) => ({
      name: String(name),
      ...(scope !== undefined ? { scope: String(scope) } : {}),
      ...(nextName !== undefined ? { nextName: String(nextName) } : {}),
      ...(refersTo !== undefined ? { refersTo: String(refersTo) } : {}),
      ...(comment !== undefined ? { comment: String(comment) } : {}),
    }),
  },
  {
    id: "xlsx.remove-defined-name",
    commandType: "xlsx:remove-defined-name",
    label: "Remove defined name",
    description: "Delete a defined name.",
    section: "Data",
    surfaces: ["cli", "palette"],
    args: [
      { name: "name", flag: "--name <name>", kind: "string", required: true, description: "Name to delete" },
      { name: "scope", flag: "--scope <sheet>", kind: "string", description: "Scope (omit for workbook-scoped)" },
    ],
    buildPayload: ({ name, scope }) => ({
      name: String(name),
      ...(scope !== undefined ? { scope: String(scope) } : {}),
    }),
  },

  // ── Tables / charts / images ──────────────────────────────────────
  {
    id: "xlsx.add-table",
    commandType: "xlsx:add-table",
    label: "Add table",
    description: "Convert a range into a structured table (with auto-filter and styles).",
    section: "Insert",
    surfaces: ["cli", "palette", "toolbar"],
    icon: "Table",
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Target sheet name" },
      { name: "range", flag: "--range <a1>", kind: "string", required: true, description: 'A1 range covering header + body, e.g. "A1:E25"' },
      { name: "hasHeaders", flag: "--has-headers <bool>", kind: "boolean", default: true, description: "Whether the first row is the header (default true)" },
      { name: "name", flag: "--name <name>", kind: "string", description: "Optional explicit table name" },
    ],
    buildPayload: ({ sheet, range, hasHeaders, name }) => ({
      sheet: String(sheet),
      range: String(range),
      ...(hasHeaders !== undefined ? { hasHeaders: hasHeaders === false || hasHeaders === "false" ? false : true } : {}),
      ...(name !== undefined ? { name: String(name) } : {}),
    }),
  },
  {
    id: "xlsx.remove-table",
    commandType: "xlsx:remove-table",
    label: "Remove table",
    description: "Convert a structured table back to a plain range.",
    section: "Insert",
    surfaces: ["cli", "palette"],
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Target sheet name" },
      { name: "tableId", flag: "--table-id <id>", kind: "string", required: true, description: "Table id" },
    ],
    buildPayload: ({ sheet, tableId }) => ({ sheet: String(sheet), tableId: String(tableId) }),
  },
  {
    id: "xlsx.add-chart",
    commandType: "xlsx:add-chart",
    label: "Insert chart",
    description: "Insert a chart anchored to a sheet position.",
    section: "Insert",
    surfaces: ["palette"],
    icon: "BarChart3",
    hidden: { reason: "Reached via the Insert ▸ Chart menu; CLI exposure is deferred." },
  },
  {
    id: "xlsx.remove-chart",
    commandType: "xlsx:remove-chart",
    label: "Remove chart",
    description: "Remove an inserted chart by id.",
    section: "Insert",
    surfaces: ["cli", "palette"],
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Target sheet name" },
      { name: "chartId", flag: "--chart-id <id>", kind: "string", required: true, description: "Chart id" },
    ],
    buildPayload: ({ sheet, chartId }) => ({ sheet: String(sheet), chartId: String(chartId) }),
  },
  {
    id: "xlsx.move-chart",
    commandType: "xlsx:move-chart",
    label: "Move chart",
    description: "Reposition a chart on the sheet.",
    section: "Insert",
    surfaces: ["cli", "palette"],
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Target sheet name" },
      { name: "chartId", flag: "--chart-id <id>", kind: "string", required: true, description: "Chart id" },
      { name: "fromRow", flag: "--from-row <n>", kind: "number", required: true, description: "Anchor cell row (0-based)" },
      { name: "fromCol", flag: "--from-col <n>", kind: "number", required: true, description: "Anchor cell column (0-based)" },
      { name: "fromOffsetXPx", flag: "--from-offset-x <px>", kind: "number", default: 0, description: "X offset within the anchor cell, in CSS pixels" },
      { name: "fromOffsetYPx", flag: "--from-offset-y <px>", kind: "number", default: 0, description: "Y offset within the anchor cell, in CSS pixels" },
    ],
    buildPayload: ({ sheet, chartId, fromRow, fromCol, fromOffsetXPx, fromOffsetYPx }) => ({
      sheet: String(sheet),
      chartId: String(chartId),
      fromRow: Number(fromRow),
      fromCol: Number(fromCol),
      fromOffsetXPx: Number(fromOffsetXPx ?? 0),
      fromOffsetYPx: Number(fromOffsetYPx ?? 0),
    }),
  },
  {
    id: "xlsx.resize-chart",
    commandType: "xlsx:resize-chart",
    label: "Resize chart",
    description: "Resize a chart on the sheet (in CSS pixels).",
    section: "Insert",
    surfaces: ["cli", "palette"],
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Target sheet name" },
      { name: "chartId", flag: "--chart-id <id>", kind: "string", required: true, description: "Chart id" },
      { name: "widthPx", flag: "--width <px>", kind: "number", required: true, description: "Width in CSS pixels" },
      { name: "heightPx", flag: "--height <px>", kind: "number", required: true, description: "Height in CSS pixels" },
    ],
    buildPayload: ({ sheet, chartId, widthPx, heightPx }) => ({
      sheet: String(sheet),
      chartId: String(chartId),
      widthPx: Number(widthPx),
      heightPx: Number(heightPx),
    }),
  },
  {
    id: "xlsx.update-chart",
    commandType: "xlsx:update-chart",
    label: "Update chart",
    description:
      "Patch a chart's type, data range, header/category toggles, title, palette, legend/data-label/gridline toggles, or axis titles.",
    section: "Insert",
    surfaces: ["cli", "palette"],
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Target sheet name" },
      { name: "chartId", flag: "--chart-id <id>", kind: "string", required: true, description: "Chart id" },
      { name: "kind", flag: "--kind <type>", kind: "enum", choices: ["bar", "column", "line", "pie", "area"], description: "Chart kind" },
      { name: "dataRange", flag: "--data-range <a1>", kind: "string", description: "A1 data range (≥ 2 cells)" },
      { name: "title", flag: "--title <text|null>", kind: "string", description: 'Chart title; pass "null" to clear' },
      { name: "hasHeaderRow", flag: "--has-header-row <bool>", kind: "boolean", description: "Whether the first row is the header" },
      { name: "hasCategoryColumn", flag: "--has-category-column <bool>", kind: "boolean", description: "Whether the first column is the category" },
    ],
    buildPayload: ({ sheet, chartId, kind, dataRange, title, hasHeaderRow, hasCategoryColumn }) => ({
      sheet: String(sheet),
      chartId: String(chartId),
      ...(kind !== undefined ? { kind: String(kind) } : {}),
      ...(dataRange !== undefined ? { dataRange: String(dataRange) } : {}),
      ...(title !== undefined ? { title: title === "null" ? null : String(title) } : {}),
      ...(hasHeaderRow !== undefined ? { hasHeaderRow: hasHeaderRow === false || hasHeaderRow === "false" ? false : true } : {}),
      ...(hasCategoryColumn !== undefined ? { hasCategoryColumn: hasCategoryColumn === false || hasCategoryColumn === "false" ? false : true } : {}),
    }),
  },
  {
    id: "xlsx.add-image",
    commandType: "xlsx:add-image",
    label: "Insert image",
    description: "Insert an image anchored to a sheet position.",
    section: "Insert",
    surfaces: ["palette"],
    icon: "Image",
    hidden: { reason: "Reached via the Insert ▸ Image menu; CLI exposure is deferred." },
  },
  {
    id: "xlsx.move-image",
    commandType: "xlsx:move-image",
    label: "Move image",
    description: "Reposition an image on the sheet.",
    section: "Insert",
    surfaces: [],
    hidden: { reason: "Reached via image drag handles." },
  },
  {
    id: "xlsx.resize-image",
    commandType: "xlsx:resize-image",
    label: "Resize image",
    description: "Resize an image on the sheet.",
    section: "Insert",
    surfaces: [],
    hidden: { reason: "Reached via image resize handles." },
  },
  {
    id: "xlsx.remove-image",
    commandType: "xlsx:remove-image",
    label: "Remove image",
    description: "Remove an image from the sheet.",
    section: "Insert",
    surfaces: [],
    hidden: { reason: "Reached via image context menu." },
  },

  // ── Auto-filter / sort / paste / fill / freeze / split ────────────
  {
    id: "xlsx.set-auto-filter",
    commandType: "xlsx:set-auto-filter",
    label: "Set auto-filter",
    description: "Apply an auto-filter to a range. Pass --range null to remove.",
    section: "Data",
    surfaces: ["cli", "palette"],
    icon: "Filter",
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Target sheet name" },
      { name: "range", flag: "--range <a1|null>", kind: "string", description: 'A1 range covering header + body, e.g. "A1:E100", or "null" to remove' },
    ],
    buildPayload: ({ sheet, range }) => ({
      sheet: String(sheet),
      range: range === "null" || range === null || range === undefined ? null : String(range),
    }),
  },
  {
    id: "xlsx.set-filter-column",
    commandType: "xlsx:set-filter-column",
    label: "Filter column",
    description: "Configure a column filter inside an existing auto-filter. Pass --criterion as JSON.",
    section: "Data",
    surfaces: ["cli", "palette"],
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Target sheet name" },
      { name: "colId", flag: "--col-id <n>", kind: "number", required: true, description: "0-based offset from autoFilter.range.c1" },
      { name: "criterion", flag: "--criterion <json>", kind: "string", required: true, description: "JSON FilterColumn (e.g. {kind:'values',values:['A','B']})" },
    ],
    buildPayload: ({ sheet, colId, criterion }) => ({
      sheet: String(sheet),
      colId: Number(colId),
      criterion: typeof criterion === "string" ? JSON.parse(criterion) : criterion,
    }),
  },
  {
    id: "xlsx.clear-filter-column",
    commandType: "xlsx:clear-filter-column",
    label: "Clear column filter",
    description: "Clear the filter on a single column.",
    section: "Data",
    surfaces: ["cli", "palette"],
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Target sheet name" },
      { name: "colId", flag: "--col-id <n>", kind: "number", required: true, description: "0-based offset from autoFilter.range.c1" },
    ],
    buildPayload: ({ sheet, colId }) => ({ sheet: String(sheet), colId: Number(colId) }),
  },
  {
    id: "xlsx.sort-range",
    commandType: "xlsx:sort-range",
    label: "Sort range",
    description: "Sort the rows inside a range (excluding the header) by a single column.",
    section: "Data",
    surfaces: ["cli", "palette"],
    icon: "ArrowUpDown",
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Target sheet name" },
      { name: "range", flag: "--range <a1>", kind: "string", required: true, description: 'A1 range; the first row is treated as the header (e.g. "A1:E25")' },
      { name: "colId", flag: "--col-id <n>", kind: "number", required: true, description: "0-based column offset from the range's first column" },
      { name: "order", flag: "--order <asc|desc>", kind: "enum", choices: ["asc", "desc"], default: "asc", description: "Sort direction" },
    ],
    buildPayload: ({ sheet, range, colId, order }) => ({
      sheet: String(sheet),
      range: String(range),
      sortBy: {
        colId: Number(colId),
        order: order === "desc" ? "desc" : "asc",
      },
    }),
  },
  {
    id: "xlsx.paste-range",
    commandType: "xlsx:paste-range",
    label: "Paste range",
    description: "Paste a 2-D matrix of values into a sheet.",
    section: "Edit",
    surfaces: [],
    hidden: { reason: "Reached via clipboard paste; not surfaced standalone." },
  },
  {
    id: "xlsx.text-to-columns",
    commandType: "xlsx:text-to-columns",
    label: "Text to columns",
    description: "Split a column's text into multiple columns by delimiter.",
    section: "Data",
    surfaces: ["palette"],
    hidden: { reason: "Reached via the Data ▸ Text to columns dialog; CLI exposure is deferred." },
  },
  {
    id: "xlsx.fill-range",
    commandType: "xlsx:fill-range",
    label: "Fill range",
    description: "Fill a range with a series (linear/growth/date/etc).",
    section: "Edit",
    surfaces: [],
    hidden: { reason: "Reached via the autofill drag handle." },
  },
  {
    id: "xlsx.freeze-panes",
    commandType: "xlsx:freeze-panes",
    label: "Freeze panes",
    description: "Freeze rows / columns above and to the left of a cell.",
    section: "View",
    surfaces: ["palette"],
    hidden: { reason: "Reached via the View ▸ Freeze menu; CLI exposure is deferred." },
  },
  {
    id: "xlsx.unfreeze-panes",
    commandType: "xlsx:unfreeze-panes",
    label: "Unfreeze panes",
    description: "Remove all frozen-pane splits on a sheet.",
    section: "View",
    surfaces: ["palette"],
    hidden: { reason: "Reached via the View ▸ Freeze menu; CLI exposure is deferred." },
  },

  // ── Palette-only convenience entries (UI wrappers / dialogs) ──────
  // Each entry is reachable from Cmd+K via the editor's runner map.
  // commandType is null when the action is sugar (toggles, dialogs)
  // around a richer catalogue entry; the CLI exposes the underlying
  // verb separately.
  {
    id: "xlsx.toggle-filter",
    commandType: null,
    label: "Toggle auto-filter",
    description: "Apply or remove an auto-filter on the active sheet.",
    section: "Data",
    surfaces: ["palette"],
  },
  {
    id: "xlsx.format-cells",
    commandType: null,
    label: "Format cells…",
    description: "Open the Format Cells dialog (Number tab).",
    section: "Format",
    surfaces: ["palette"],
  },
  {
    id: "xlsx.format-cells-alignment",
    commandType: null,
    label: "Format cells: Alignment",
    description: "Open the Format Cells dialog on the Alignment tab.",
    section: "Format",
    surfaces: ["palette"],
  },
  {
    id: "xlsx.format-cells-font",
    commandType: null,
    label: "Format cells: Font",
    description: "Open the Format Cells dialog on the Font tab.",
    section: "Format",
    surfaces: ["palette"],
  },
  {
    id: "xlsx.format-cells-border",
    commandType: null,
    label: "Format cells: Border",
    description: "Open the Format Cells dialog on the Border tab.",
    section: "Format",
    surfaces: ["palette"],
  },
  {
    id: "xlsx.format-cells-fill",
    commandType: null,
    label: "Format cells: Fill",
    description: "Open the Format Cells dialog on the Fill tab.",
    section: "Format",
    surfaces: ["palette"],
  },
  {
    id: "xlsx.format-cells-protection",
    commandType: null,
    label: "Format cells: Protection",
    description: "Open the Format Cells dialog on the Protection tab.",
    section: "Format",
    surfaces: ["palette"],
  },
  {
    id: "xlsx.borders-all",
    commandType: null,
    label: "Borders: All",
    description: "Apply borders to every cell in the selection.",
    section: "Format",
    surfaces: ["palette"],
  },
  {
    id: "xlsx.borders-outside",
    commandType: null,
    label: "Borders: Outside",
    description: "Apply borders only to the perimeter of the selection.",
    section: "Format",
    surfaces: ["palette"],
  },
  {
    id: "xlsx.borders-thick-outside",
    commandType: null,
    label: "Borders: Thick outside",
    description: "Apply a thick border to the perimeter of the selection.",
    section: "Format",
    surfaces: ["palette"],
  },
  {
    id: "xlsx.borders-none",
    commandType: null,
    label: "Borders: Clear",
    description: "Remove all borders from the selection.",
    section: "Format",
    surfaces: ["palette"],
  },
  {
    id: "xlsx.paste-special",
    commandType: null,
    label: "Paste Special…",
    description: "Open the Paste Special dialog.",
    section: "Edit",
    surfaces: ["palette"],
  },
  {
    id: "xlsx.format-painter",
    commandType: null,
    label: "Format Painter",
    description: "Activate the Format Painter to copy cell formatting.",
    section: "Format",
    surfaces: ["palette"],
  },
  {
    id: "xlsx.hide-sheet",
    commandType: null,
    label: "Hide sheet",
    description: "Hide the active sheet (preserves data; toggle via the sheet tab menu).",
    section: "Sheet",
    surfaces: ["palette"],
  },
  {
    id: "xlsx.conditional-format",
    commandType: null,
    label: "Conditional Formatting…",
    description: "Open the Conditional Formatting manager dialog.",
    section: "Format",
    surfaces: ["palette"],
  },
  {
    id: "xlsx.conditional-format-clear",
    commandType: null,
    label: "Clear conditional formatting from this sheet",
    description: "Drop every conditional-formatting rule on the active sheet.",
    section: "Format",
    surfaces: ["palette"],
  },
  {
    id: "xlsx.data-validation",
    commandType: null,
    label: "Data Validation…",
    description: "Open the Data Validation dialog for the active selection.",
    section: "Data",
    surfaces: ["palette"],
  },
  {
    id: "xlsx.data-validation-clear",
    commandType: null,
    label: "Clear data validation from this sheet",
    description: "Drop every data-validation rule on the active sheet.",
    section: "Data",
    surfaces: ["palette"],
  },
  {
    id: "xlsx.name-manager",
    commandType: null,
    label: "Name Manager…",
    description: "Open the Name Manager dialog to inspect and edit defined names.",
    section: "Data",
    surfaces: ["palette"],
    shortcut: "F3",
  },
  {
    id: "xlsx.define-name",
    commandType: null,
    label: "Define name from selection…",
    description: "Open the Name Manager scoped to the current selection.",
    section: "Data",
    surfaces: ["palette"],
  },
  {
    id: "xlsx.merge",
    commandType: null,
    label: "Merge cells",
    description: "Merge the current selection into one cell.",
    section: "Format",
    surfaces: ["palette"],
  },
  {
    id: "xlsx.unmerge",
    commandType: null,
    label: "Unmerge cells",
    description: "Split previously merged cells in the current selection.",
    section: "Format",
    surfaces: ["palette"],
  },
  {
    id: "xlsx.insert-image",
    commandType: null,
    label: "Insert image",
    description: "Pick and insert an image anchored to the active selection.",
    section: "Insert",
    surfaces: ["palette"],
  },
  {
    id: "xlsx.format-as-table",
    commandType: null,
    label: "Format as Table…",
    description: "Convert the current selection into a structured table.",
    section: "Insert",
    surfaces: ["palette"],
    shortcut: "Mod+T",
  },
  {
    id: "xlsx.insert-chart",
    commandType: null,
    label: "Insert chart…",
    description: "Open the Insert Chart dialog for the current selection.",
    section: "Insert",
    surfaces: ["palette"],
  },
  {
    id: "xlsx.insert-pivot-table",
    commandType: null,
    label: "Insert pivot table…",
    description:
      "Summarise the current selection by grouping a column and aggregating another. Result writes as plain values (full OOXML pivot round-trip lands in a follow-up).",
    section: "Insert",
    surfaces: ["palette"],
  },
  {
    id: "xlsx.edit-chart",
    commandType: null,
    label: "Edit chart…",
    description: "Edit the currently selected chart's data range, type, and title.",
    section: "Insert",
    surfaces: ["palette"],
  },

  // ── Phase 6 — Protection ─────────────────────────────────────────
  {
    id: "xlsx.set-sheet-protection",
    commandType: "xlsx:set-sheet-protection",
    label: "Protect sheet",
    description:
      "Toggle Excel's Protect Sheet element on a worksheet. Pass --enabled false to clear; precomputed password hashes are accepted via --password-hash + --algorithm-name + --salt-value + --spin-count.",
    section: "Review",
    surfaces: ["cli", "palette", "toolbar"],
    icon: "Lock",
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Sheet name." },
      {
        name: "enabled",
        flag: "--enabled <bool>",
        kind: "boolean",
        default: true,
        description: "true to protect, false to clear protection.",
      },
      { name: "passwordHash", flag: "--password-hash <hex>", kind: "string", description: "Precomputed password hash." },
      { name: "algorithmName", flag: "--algorithm-name <name>", kind: "string", description: "Hash algorithm name." },
      { name: "hashValue", flag: "--hash-value <hex>", kind: "string", description: "Hash value (algorithm output)." },
      { name: "saltValue", flag: "--salt-value <b64>", kind: "string", description: "Salt for the hash." },
      { name: "spinCount", flag: "--spin-count <n>", kind: "number", description: "Iteration count." },
      { name: "objects", flag: "--objects <bool>", kind: "boolean", description: "Allow editing objects." },
      { name: "scenarios", flag: "--scenarios <bool>", kind: "boolean", description: "Allow editing scenarios." },
      { name: "formatCells", flag: "--format-cells <bool>", kind: "boolean", description: "Allow formatting cells." },
      { name: "formatColumns", flag: "--format-columns <bool>", kind: "boolean", description: "Allow formatting columns." },
      { name: "formatRows", flag: "--format-rows <bool>", kind: "boolean", description: "Allow formatting rows." },
      { name: "insertColumns", flag: "--insert-columns <bool>", kind: "boolean", description: "Allow inserting columns." },
      { name: "insertRows", flag: "--insert-rows <bool>", kind: "boolean", description: "Allow inserting rows." },
      { name: "insertHyperlinks", flag: "--insert-hyperlinks <bool>", kind: "boolean", description: "Allow inserting hyperlinks." },
      { name: "deleteColumns", flag: "--delete-columns <bool>", kind: "boolean", description: "Allow deleting columns." },
      { name: "deleteRows", flag: "--delete-rows <bool>", kind: "boolean", description: "Allow deleting rows." },
      { name: "selectLockedCells", flag: "--select-locked-cells <bool>", kind: "boolean", description: "Allow selecting locked cells." },
      { name: "sort", flag: "--sort <bool>", kind: "boolean", description: "Allow sorting." },
      { name: "autoFilter", flag: "--auto-filter <bool>", kind: "boolean", description: "Allow using AutoFilter." },
      { name: "pivotTables", flag: "--pivot-tables <bool>", kind: "boolean", description: "Allow editing pivot tables." },
      { name: "selectUnlockedCells", flag: "--select-unlocked-cells <bool>", kind: "boolean", description: "Allow selecting unlocked cells." },
    ],
    buildPayload: (parsed) => {
      const out: Record<string, unknown> = {
        sheet: String(parsed.sheet),
        enabled: parsed.enabled === false || parsed.enabled === "false" ? false : true,
      };
      const passthrough = [
        "passwordHash",
        "algorithmName",
        "hashValue",
        "saltValue",
        "objects",
        "scenarios",
        "formatCells",
        "formatColumns",
        "formatRows",
        "insertColumns",
        "insertRows",
        "insertHyperlinks",
        "deleteColumns",
        "deleteRows",
        "selectLockedCells",
        "sort",
        "autoFilter",
        "pivotTables",
        "selectUnlockedCells",
      ] as const;
      for (const k of passthrough) {
        const v = (parsed as Record<string, unknown>)[k];
        if (v !== undefined) out[k] = v;
      }
      if (parsed.spinCount !== undefined) out.spinCount = Number(parsed.spinCount);
      return out;
    },
  },
  {
    id: "xlsx.set-workbook-protection",
    commandType: "xlsx:set-workbook-protection",
    label: "Protect workbook",
    description:
      "Toggle Excel's Protect Workbook element. Defaults to lockStructure=1 when no flags are supplied. Pass --enabled false to clear.",
    section: "Review",
    surfaces: ["cli", "palette", "toolbar"],
    icon: "Lock",
    args: [
      {
        name: "enabled",
        flag: "--enabled <bool>",
        kind: "boolean",
        default: true,
        description: "true to protect, false to clear protection.",
      },
      { name: "workbookPasswordHash", flag: "--password-hash <hex>", kind: "string", description: "Precomputed password hash." },
      { name: "algorithmName", flag: "--algorithm-name <name>", kind: "string", description: "Hash algorithm name." },
      { name: "hashValue", flag: "--hash-value <hex>", kind: "string", description: "Hash value." },
      { name: "saltValue", flag: "--salt-value <b64>", kind: "string", description: "Salt for the hash." },
      { name: "spinCount", flag: "--spin-count <n>", kind: "number", description: "Iteration count." },
      { name: "lockStructure", flag: "--lock-structure <bool>", kind: "boolean", description: "Lock workbook structure." },
      { name: "lockWindows", flag: "--lock-windows <bool>", kind: "boolean", description: "Lock workbook windows." },
      { name: "lockRevision", flag: "--lock-revision <bool>", kind: "boolean", description: "Lock revision tracking." },
    ],
    buildPayload: (parsed) => {
      const out: Record<string, unknown> = {
        enabled: parsed.enabled === false || parsed.enabled === "false" ? false : true,
      };
      const passthrough = [
        "workbookPasswordHash",
        "algorithmName",
        "hashValue",
        "saltValue",
        "lockStructure",
        "lockWindows",
        "lockRevision",
      ] as const;
      for (const k of passthrough) {
        const v = (parsed as Record<string, unknown>)[k];
        if (v !== undefined) out[k] = v;
      }
      if (parsed.spinCount !== undefined) out.spinCount = Number(parsed.spinCount);
      return out;
    },
  },

  // ── Phase 4c — Formulas tab ──────────────────────────────────────
  {
    id: "xlsx.set-calc-mode",
    commandType: "xlsx:set-calc-mode",
    label: "Calculation options",
    description:
      "Set Excel's Formulas → Calculation Options (auto / autoNoTable / manual) and the related calcOnSave / iterate knobs.",
    section: "Formulas",
    surfaces: ["cli", "palette", "toolbar"],
    icon: "Calculator",
    args: [
      {
        name: "calcMode",
        flag: "--mode <mode>",
        kind: "enum",
        choices: ["auto", "autoNoTable", "manual"],
        description: "Calculation mode.",
      },
      { name: "calcOnSave", flag: "--calc-on-save <bool>", kind: "boolean", description: "Recalculate before saving." },
      { name: "iterate", flag: "--iterate <bool>", kind: "boolean", description: "Enable iterative calculation." },
      { name: "iterateCount", flag: "--iterate-count <n>", kind: "number", description: "Maximum iterations." },
      { name: "iterateDelta", flag: "--iterate-delta <n>", kind: "number", description: "Convergence delta." },
    ],
    buildPayload: (parsed) => {
      const out: Record<string, unknown> = {};
      if (parsed.calcMode) out.calcMode = String(parsed.calcMode);
      if (parsed.calcOnSave !== undefined) out.calcOnSave = parsed.calcOnSave !== false && parsed.calcOnSave !== "false";
      if (parsed.iterate !== undefined) out.iterate = parsed.iterate !== false && parsed.iterate !== "false";
      if (parsed.iterateCount !== undefined) out.iterateCount = Number(parsed.iterateCount);
      if (parsed.iterateDelta !== undefined) out.iterateDelta = Number(parsed.iterateDelta);
      return out;
    },
  },
  {
    id: "xlsx.set-show-formulas",
    commandType: "xlsx:set-show-formulas",
    label: "Show formulas",
    description:
      "Toggle the Formulas → Formula Auditing → Show Formulas view mode on a worksheet (writes showFormulas=\"1\" to the sheetView).",
    section: "Formulas",
    surfaces: ["cli", "palette", "toolbar"],
    icon: "Eye",
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Sheet name." },
      { name: "show", flag: "--show <bool>", kind: "boolean", default: true, description: "true to show formulas, false to show values." },
    ],
    buildPayload: ({ sheet, show }) => ({
      sheet: String(sheet),
      show: show === false || show === "false" ? false : true,
    }),
  },

  // ── Phase 4b: Page Layout ──────────────────────────────────────────
  {
    id: "xlsx.set-page-setup",
    commandType: "xlsx:set-page-setup",
    label: "Set page setup",
    description:
      "Configure the worksheet's <pageSetup> element (orientation, paper size, scaling, fit-to-pages, page numbering). Pass --clear to reset.",
    section: "Page Layout",
    surfaces: ["cli", "palette", "toolbar"],
    icon: "FileText",
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Sheet name." },
      {
        name: "orientation",
        flag: "--orientation <kind>",
        kind: "enum",
        choices: ["default", "portrait", "landscape"],
        description: "Page orientation.",
      },
      { name: "paperSize", flag: "--paper-size <n>", kind: "number", description: "OOXML paperSize enum (1=letter, 9=A4)." },
      { name: "scale", flag: "--scale <n>", kind: "number", description: "Print scaling percentage (10–400)." },
      { name: "fitToWidth", flag: "--fit-to-width <n>", kind: "number", description: "Pages wide; 0 = use scale." },
      { name: "fitToHeight", flag: "--fit-to-height <n>", kind: "number", description: "Pages tall; 0 = use scale." },
      { name: "firstPageNumber", flag: "--first-page-number <n>", kind: "number", description: "First page number." },
      { name: "useFirstPageNumber", flag: "--use-first-page-number <bool>", kind: "boolean", description: "Honour firstPageNumber." },
      { name: "blackAndWhite", flag: "--black-and-white <bool>", kind: "boolean", description: "Print in B&W." },
      { name: "draft", flag: "--draft <bool>", kind: "boolean", description: "Draft quality (skip graphics)." },
      { name: "clear", flag: "--clear", kind: "boolean", description: "Drop the <pageSetup> element entirely." },
    ],
    buildPayload: (a) => {
      const out: Record<string, unknown> = { sheet: String(a.sheet) };
      if (a.orientation) out.orientation = a.orientation;
      if (a.paperSize !== undefined) out.paperSize = Number(a.paperSize);
      if (a.scale !== undefined) out.scale = Number(a.scale);
      if (a.fitToWidth !== undefined) out.fitToWidth = Number(a.fitToWidth);
      if (a.fitToHeight !== undefined) out.fitToHeight = Number(a.fitToHeight);
      if (a.firstPageNumber !== undefined) out.firstPageNumber = Number(a.firstPageNumber);
      if (typeof a.useFirstPageNumber === "boolean") out.useFirstPageNumber = a.useFirstPageNumber;
      if (typeof a.blackAndWhite === "boolean") out.blackAndWhite = a.blackAndWhite;
      if (typeof a.draft === "boolean") out.draft = a.draft;
      if (a.clear === true) out.clear = true;
      return out;
    },
  },
  {
    id: "xlsx.set-page-margins",
    commandType: "xlsx:set-page-margins",
    label: "Set page margins",
    description:
      "Set worksheet page margins. Pass --preset (normal | wide | narrow), individual *-in overrides, or --clear.",
    section: "Page Layout",
    surfaces: ["cli", "palette", "toolbar"],
    icon: "Frame",
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Sheet name." },
      {
        name: "preset",
        flag: "--preset <kind>",
        kind: "enum",
        choices: ["normal", "wide", "narrow"],
        description: "Built-in margin preset.",
      },
      { name: "leftIn", flag: "--left-in <n>", kind: "number", description: "Left margin (inches)." },
      { name: "rightIn", flag: "--right-in <n>", kind: "number", description: "Right margin (inches)." },
      { name: "topIn", flag: "--top-in <n>", kind: "number", description: "Top margin (inches)." },
      { name: "bottomIn", flag: "--bottom-in <n>", kind: "number", description: "Bottom margin (inches)." },
      { name: "headerIn", flag: "--header-in <n>", kind: "number", description: "Header margin (inches)." },
      { name: "footerIn", flag: "--footer-in <n>", kind: "number", description: "Footer margin (inches)." },
      { name: "clear", flag: "--clear", kind: "boolean", description: "Drop the <pageMargins> element entirely." },
    ],
    buildPayload: (a) => {
      const out: Record<string, unknown> = { sheet: String(a.sheet) };
      if (a.preset) out.preset = a.preset;
      if (a.leftIn !== undefined) out.leftIn = Number(a.leftIn);
      if (a.rightIn !== undefined) out.rightIn = Number(a.rightIn);
      if (a.topIn !== undefined) out.topIn = Number(a.topIn);
      if (a.bottomIn !== undefined) out.bottomIn = Number(a.bottomIn);
      if (a.headerIn !== undefined) out.headerIn = Number(a.headerIn);
      if (a.footerIn !== undefined) out.footerIn = Number(a.footerIn);
      if (a.clear === true) out.clear = true;
      return out;
    },
  },
  {
    id: "xlsx.set-print-options",
    commandType: "xlsx:set-print-options",
    label: "Set print options",
    description:
      "Toggle Page Layout → Sheet Options (gridlines, headings, centering on page). Pass --clear to remove the <printOptions> element.",
    section: "Page Layout",
    surfaces: ["cli", "palette", "toolbar"],
    icon: "Printer",
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Sheet name." },
      { name: "horizontalCentered", flag: "--horizontal-centered <bool>", kind: "boolean", description: "Centre horizontally on page." },
      { name: "verticalCentered", flag: "--vertical-centered <bool>", kind: "boolean", description: "Centre vertically on page." },
      { name: "headings", flag: "--headings <bool>", kind: "boolean", description: "Print row/column headings." },
      { name: "gridLines", flag: "--grid-lines <bool>", kind: "boolean", description: "Print gridlines." },
      { name: "clear", flag: "--clear", kind: "boolean", description: "Drop the <printOptions> element entirely." },
    ],
    buildPayload: (a) => {
      const out: Record<string, unknown> = { sheet: String(a.sheet) };
      if (typeof a.horizontalCentered === "boolean") out.horizontalCentered = a.horizontalCentered;
      if (typeof a.verticalCentered === "boolean") out.verticalCentered = a.verticalCentered;
      if (typeof a.headings === "boolean") out.headings = a.headings;
      if (typeof a.gridLines === "boolean") out.gridLines = a.gridLines;
      if (a.clear === true) out.clear = true;
      return out;
    },
  },
  {
    id: "xlsx.set-print-area",
    commandType: "xlsx:set-print-area",
    label: "Set print area",
    description: "Define the worksheet's _xlnm.Print_Area defined name to a range, or --clear to remove it.",
    section: "Page Layout",
    surfaces: ["cli", "palette", "toolbar"],
    icon: "PrinterCheck",
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Sheet name." },
      { name: "range", flag: "--range <a1>", kind: "string", description: "A1 range, e.g. A1:D20." },
      { name: "clear", flag: "--clear", kind: "boolean", description: "Remove the existing print area." },
    ],
    buildPayload: ({ sheet, range, clear }) => ({
      sheet: String(sheet),
      ...(range ? { range: String(range) } : {}),
      ...(clear === true ? { clear: true } : {}),
    }),
  },
  {
    id: "xlsx.set-print-titles",
    commandType: "xlsx:set-print-titles",
    label: "Set print titles",
    description:
      "Define _xlnm.Print_Titles (rows/columns repeated on each printed page). Pass --rows (e.g. '1:1'), --cols (e.g. 'A:B'), or --clear.",
    section: "Page Layout",
    surfaces: ["cli", "palette", "toolbar"],
    icon: "Repeat",
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Sheet name." },
      { name: "rows", flag: "--rows <range>", kind: "string", description: "Row range to repeat (e.g. '1:1')." },
      { name: "cols", flag: "--cols <range>", kind: "string", description: "Column range to repeat (e.g. 'A:B')." },
      { name: "clear", flag: "--clear", kind: "boolean", description: "Remove the existing print titles." },
    ],
    buildPayload: ({ sheet, rows, cols, clear }) => ({
      sheet: String(sheet),
      ...(rows ? { rows: String(rows) } : {}),
      ...(cols ? { cols: String(cols) } : {}),
      ...(clear === true ? { clear: true } : {}),
    }),
  },

  // ── Phase 7: View tab ─────────────────────────────────────────────
  {
    // Palette-only opener for the Zoom dialog. Kept separate from
    // `xlsx.set-sheet-view` so the palette entry that says "Set sheet
    // view" actually configures the sheet view (mode + toggles) and
    // the entry that says "Zoom" actually opens Zoom — instead of one
    // running the other (the original wiring made Cmd+K → "Set sheet
    // view" pop the Zoom dialog, which was a lying button).
    id: "xlsx.open-zoom-dialog",
    commandType: null,
    label: "Zoom",
    description: "Open the Excel-style zoom dialog (View → Zoom).",
    section: "View",
    surfaces: ["palette"],
    icon: "ZoomIn",
  },
  {
    id: "xlsx.set-sheet-view",
    commandType: "xlsx:set-sheet-view",
    label: "Set sheet view",
    description:
      "Patch the worksheet's <sheetView> element. Mirrors Excel's View tab toggles (view mode, gridlines, headings, ruler, zoom, RTL).",
    section: "View",
    surfaces: ["cli", "palette", "toolbar"],
    icon: "Eye",
    args: [
      { name: "sheet", flag: "--sheet <name>", kind: "string", required: true, description: "Sheet name." },
      {
        name: "view",
        flag: "--view <mode>",
        kind: "enum",
        choices: ["normal", "pageBreakPreview", "pageLayout"],
        description: "View mode.",
      },
      { name: "showGridLines", flag: "--show-grid-lines <bool>", kind: "boolean", description: "Show worksheet gridlines." },
      { name: "showRowColHeaders", flag: "--show-row-col-headers <bool>", kind: "boolean", description: "Show row & column headings." },
      { name: "showZeros", flag: "--show-zeros <bool>", kind: "boolean", description: "Show zero values." },
      { name: "showRuler", flag: "--show-ruler <bool>", kind: "boolean", description: "Show the page ruler (Page Layout view)." },
      { name: "showOutlineSymbols", flag: "--show-outline-symbols <bool>", kind: "boolean", description: "Show outline-level symbols." },
      { name: "rightToLeft", flag: "--right-to-left <bool>", kind: "boolean", description: "Render the sheet right-to-left." },
      { name: "zoomScale", flag: "--zoom-scale <n>", kind: "number", description: "Zoom scale (10–400)." },
      { name: "zoomScaleNormal", flag: "--zoom-scale-normal <n>", kind: "number", description: "Zoom scale for the normal view (10–400)." },
    ],
    buildPayload: (a) => {
      const out: Record<string, unknown> = { sheet: String(a.sheet) };
      if (a.view) out.view = a.view;
      if (typeof a.showGridLines === "boolean") out.showGridLines = a.showGridLines;
      if (typeof a.showRowColHeaders === "boolean") out.showRowColHeaders = a.showRowColHeaders;
      if (typeof a.showZeros === "boolean") out.showZeros = a.showZeros;
      if (typeof a.showRuler === "boolean") out.showRuler = a.showRuler;
      if (typeof a.showOutlineSymbols === "boolean") out.showOutlineSymbols = a.showOutlineSymbols;
      if (typeof a.rightToLeft === "boolean") out.rightToLeft = a.rightToLeft;
      if (a.zoomScale !== undefined) out.zoomScale = Number(a.zoomScale);
      if (a.zoomScaleNormal !== undefined) out.zoomScaleNormal = Number(a.zoomScaleNormal);
      return out;
    },
  },
];
