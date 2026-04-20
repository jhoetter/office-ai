/**
 * Action catalogue for PDF. Single source of truth tying together
 * each command-bus handler, its CLI subcommand, and its Cmd+K palette
 * entry.
 *
 * Note: the PDF surface is split. Some `office-agent pdf …`
 * subcommands ride the bus (rotate-pages, delete-pages, …); others
 * (merge, split, fill-form, add-watermark) live in `pdf-edit` /
 * `pdf-forms` / `pdf-ocr` and run directly on the input bytes
 * because they aren't yet modelled as bus commands. The catalogue
 * surfaces both so the parity check can validate every public
 * subcommand has a known home, while marking the non-bus ones with
 * `commandType: null`.
 */

import type { ActionDescriptor } from "@officeai/core";

export const pdfActions: ReadonlyArray<ActionDescriptor> = [
  // ── Read (no commandType) ─────────────────────────────────────────
  {
    id: "pdf.read-metadata",
    commandType: null,
    label: "Read metadata",
    description: "Print the PDF's document-info dictionary, page count, and engine details.",
    section: "Read",
    surfaces: ["cli"],
  },
  {
    id: "pdf.read-page",
    commandType: null,
    label: "Read page",
    description: "Read a single page's text content.",
    section: "Read",
    surfaces: ["cli"],
  },
  {
    id: "pdf.read-outline",
    commandType: null,
    label: "Read outline",
    description: "Print the document outline (bookmarks) as JSON.",
    section: "Read",
    surfaces: ["cli"],
  },
  {
    id: "pdf.read-annotations",
    commandType: null,
    label: "Read annotations",
    description: "List every annotation across the document as JSON.",
    section: "Read",
    surfaces: ["cli"],
  },
  {
    id: "pdf.list-form-fields",
    commandType: null,
    label: "List form fields",
    description: "List every interactive form field with its current value.",
    section: "Read",
    surfaces: ["cli"],
  },
  {
    id: "pdf.search-text",
    commandType: null,
    label: "Search text",
    description: "Search the PDF's text layer; emits matches as JSON.",
    section: "Read",
    surfaces: ["cli"],
  },
  {
    id: "pdf.export-markdown",
    commandType: null,
    label: "Export Markdown",
    description: "Project the PDF's text content as Markdown.",
    section: "Read",
    surfaces: ["cli"],
  },

  // ── Page-level mutations (bus-backed) ─────────────────────────────
  {
    id: "pdf.rotate-pages",
    commandType: "pdf:rotate-pages",
    label: "Rotate pages",
    description: "Rotate one or more pages by 90/180/270°.",
    section: "Pages",
    surfaces: ["cli", "palette"],
    icon: "RotateCw",
  },
  {
    id: "pdf.set-page-rotation",
    commandType: "pdf:set-page-rotation",
    label: "Set page rotation",
    description: "Set absolute rotation on a single page.",
    section: "Pages",
    surfaces: ["palette"],
    hidden: { reason: "Wrapped by pdf rotate-pages on the CLI; surfaced separately in the palette." },
  },
  {
    id: "pdf.reorder-pages",
    commandType: "pdf:reorder-pages",
    label: "Reorder pages",
    description: "Reorder the document's pages.",
    section: "Pages",
    surfaces: ["cli", "palette"],
  },
  {
    id: "pdf.delete-pages",
    commandType: "pdf:delete-pages",
    label: "Delete pages",
    description: "Delete one or more pages from the document.",
    section: "Pages",
    surfaces: ["cli", "palette"],
  },

  // ── Page-level mutations (pdf-edit; not yet bus-backed) ───────────
  {
    id: "pdf.merge",
    commandType: null,
    label: "Merge PDFs",
    description: "Merge two or more PDFs in order; writes a new PDF to --out.",
    section: "Pages",
    surfaces: ["cli", "palette"],
  },
  {
    id: "pdf.split",
    commandType: null,
    label: "Split PDF",
    description: "Split a PDF by page range or every-N-pages; writes one PDF per chunk.",
    section: "Pages",
    surfaces: ["cli", "palette"],
  },
  {
    id: "pdf.extract-pages",
    commandType: null,
    label: "Extract pages",
    description: "Extract a page range into a new PDF.",
    section: "Pages",
    surfaces: ["cli", "palette"],
  },
  {
    id: "pdf.crop-pages",
    commandType: null,
    label: "Crop pages",
    description: "Crop a page range to a rectangle.",
    section: "Pages",
    surfaces: ["cli"],
  },
  {
    id: "pdf.add-watermark",
    commandType: null,
    label: "Add watermark",
    description: "Stamp a text or image watermark across a page range.",
    section: "Pages",
    surfaces: ["cli", "palette"],
  },
  {
    id: "pdf.add-page-numbers",
    commandType: null,
    label: "Add page numbers",
    description: "Stamp page numbers across a page range.",
    section: "Pages",
    surfaces: ["cli", "palette"],
  },

  // ── Metadata ──────────────────────────────────────────────────────
  {
    id: "pdf.set-metadata",
    commandType: "pdf:set-metadata",
    label: "Set metadata",
    description: "Set / clear metadata fields (title, author, subject, keywords).",
    section: "Document",
    surfaces: ["cli", "palette"],
  },

  // ── Bookmarks ─────────────────────────────────────────────────────
  {
    id: "pdf.add-bookmark",
    commandType: "pdf:add-bookmark",
    label: "Add bookmark",
    description: "Add a bookmark / outline entry to the document.",
    section: "Document",
    surfaces: ["palette"],
    hidden: { reason: "Reached via the bookmarks pane; CLI exposure is deferred." },
  },

  // ── Annotations ───────────────────────────────────────────────────
  {
    id: "pdf.add-annotation",
    commandType: "pdf:add-annotation",
    label: "Add annotation",
    description: "Add an annotation (highlight / underline / text / shape) to a page.",
    section: "Annotations",
    surfaces: ["palette"],
    hidden: { reason: "Reached via the annotation toolbar; CLI exposure is deferred." },
  },
  {
    id: "pdf.update-annotation",
    commandType: "pdf:update-annotation",
    label: "Update annotation",
    description: "Update an annotation's color / text / geometry.",
    section: "Annotations",
    surfaces: [],
    hidden: { reason: "Reached via the annotation property pane." },
  },
  {
    id: "pdf.remove-annotation",
    commandType: "pdf:remove-annotation",
    label: "Remove annotation",
    description: "Remove an annotation.",
    section: "Annotations",
    surfaces: ["palette"],
    hidden: { reason: "Reached via the annotation context menu." },
  },

  // ── Comments (bus-backed, share annotation surface) ───────────────
  {
    id: "pdf.add-comment",
    commandType: "pdf:add-comment",
    label: "Add comment",
    description: "Anchor a sticky-note comment to a page.",
    section: "Comments",
    surfaces: ["palette"],
    hidden: { reason: "Reached via the comment toolbar; CLI exposure is deferred." },
  },
  {
    id: "pdf.reply-comment",
    commandType: "pdf:reply-comment",
    label: "Reply to comment",
    description: "Append a reply to an existing comment thread.",
    section: "Comments",
    surfaces: [],
    hidden: { reason: "Reached via the comment thread UI." },
  },
  {
    id: "pdf.edit-comment",
    commandType: "pdf:edit-comment",
    label: "Edit comment",
    description: "Edit a comment's text.",
    section: "Comments",
    surfaces: [],
    hidden: { reason: "Reached via the comment thread UI." },
  },
  {
    id: "pdf.resolve-comment",
    commandType: "pdf:resolve-comment",
    label: "Resolve comment",
    description: "Mark a comment thread as resolved.",
    section: "Comments",
    surfaces: [],
    hidden: { reason: "Reached via the comment thread UI." },
  },
  {
    id: "pdf.delete-comment",
    commandType: "pdf:delete-comment",
    label: "Delete comment",
    description: "Delete a comment thread.",
    section: "Comments",
    surfaces: [],
    hidden: { reason: "Reached via the comment thread UI." },
  },

  // ── Forms ─────────────────────────────────────────────────────────
  {
    id: "pdf.fill-form",
    commandType: null,
    label: "Fill form",
    description: "Fill an interactive form from a JSON map of field → value.",
    section: "Forms",
    surfaces: ["cli", "palette"],
  },
  {
    id: "pdf.flatten-form",
    commandType: null,
    label: "Flatten form",
    description: "Flatten an interactive form into static PDF content.",
    section: "Forms",
    surfaces: ["cli", "palette"],
  },
  {
    id: "pdf.reset-form",
    commandType: null,
    label: "Reset form",
    description: "Reset every interactive form field to its default value.",
    section: "Forms",
    surfaces: ["cli", "palette"],
  },

  // ── OCR ───────────────────────────────────────────────────────────
  {
    id: "pdf.add-text-layer",
    commandType: null,
    label: "Add OCR text layer",
    description: "Run OCR over each page and embed an invisible selectable text layer.",
    section: "OCR",
    surfaces: ["cli", "palette"],
  },

  // ── Palette-only viewer sugar (navigation + zoom + display) ───────
  // These are local viewer state changes; no command-bus mutation is
  // involved. They live in the catalogue so Cmd+K stays the single
  // surface for everything the user can reach from a UI button.
  {
    id: "pdf.next-page",
    commandType: null,
    label: "Next page",
    description: "Navigate to the next page in the viewer.",
    section: "Navigate",
    surfaces: ["palette"],
  },
  {
    id: "pdf.prev-page",
    commandType: null,
    label: "Previous page",
    description: "Navigate to the previous page in the viewer.",
    section: "Navigate",
    surfaces: ["palette"],
  },
  {
    id: "pdf.first-page",
    commandType: null,
    label: "First page",
    description: "Jump to the first page.",
    section: "Navigate",
    surfaces: ["palette"],
  },
  {
    id: "pdf.last-page",
    commandType: null,
    label: "Last page",
    description: "Jump to the last page.",
    section: "Navigate",
    surfaces: ["palette"],
  },
  {
    id: "pdf.zoom-in",
    commandType: null,
    label: "Zoom in",
    description: "Increase the viewer zoom level.",
    section: "View",
    surfaces: ["palette"],
  },
  {
    id: "pdf.zoom-out",
    commandType: null,
    label: "Zoom out",
    description: "Decrease the viewer zoom level.",
    section: "View",
    surfaces: ["palette"],
  },
  {
    id: "pdf.fit-width",
    commandType: null,
    label: "Fit width",
    description: "Fit the page width to the viewer.",
    section: "View",
    surfaces: ["palette"],
  },
  {
    id: "pdf.fit-page",
    commandType: null,
    label: "Fit page",
    description: "Fit a whole page into the viewer.",
    section: "View",
    surfaces: ["palette"],
  },
  {
    id: "pdf.actual-size",
    commandType: null,
    label: "Actual size (100 %)",
    description: "Reset the viewer zoom to 100%.",
    section: "View",
    surfaces: ["palette"],
  },
  {
    id: "pdf.rotate-cw",
    commandType: null,
    label: "Rotate view clockwise",
    description: "Rotate the on-screen view clockwise (does not modify the PDF).",
    section: "View",
    surfaces: ["palette"],
  },
  {
    id: "pdf.rotate-ccw",
    commandType: null,
    label: "Rotate view counter-clockwise",
    description: "Rotate the on-screen view counter-clockwise (does not modify the PDF).",
    section: "View",
    surfaces: ["palette"],
  },
  {
    id: "pdf.toggle-dark",
    commandType: null,
    label: "Toggle dark mode",
    description: "Cycle the viewer through light, sepia, and dark display modes.",
    section: "View",
    surfaces: ["palette"],
  },
  {
    id: "pdf.toggle-reflow",
    commandType: null,
    label: "Toggle reflow",
    description: "Toggle reflowed reading mode (single-column text reflow).",
    section: "View",
    surfaces: ["palette"],
  },
  {
    id: "pdf.rotate-page",
    commandType: null,
    label: "Rotate current page 90°",
    description: "Rotate the current page by 90° (writes to the document; see also pdf.rotate-pages).",
    section: "Pages",
    surfaces: ["palette"],
  },
  {
    id: "pdf.delete-page",
    commandType: null,
    label: "Delete current page",
    description: "Delete the current page from the document (see also pdf.delete-pages).",
    section: "Pages",
    surfaces: ["palette"],
  },
  {
    id: "pdf.print",
    commandType: null,
    label: "Print…",
    description: "Open the browser's native print dialog for the PDF.",
    section: "File",
    surfaces: ["palette"],
  },
];
