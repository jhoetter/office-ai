"use client";

import { useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import {
  Image as ImageIcon,
  MessageSquarePlus,
  AlignLeft,
  AlignCenter,
  AlignRight,
  AlignJustify,
  Indent,
  Outdent,
  List,
  ListOrdered,
  ChevronDown,
  Pencil,
  PenLine,
  Eye,
  Pilcrow,
  Check,
  X,
  ScrollText,
  SeparatorHorizontal,
  StickyNote,
  Table2,
  Hash,
  ListOrdered as ListOrderedIcon,
  Lock,
  AlignVerticalJustifyStart,
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  PaintBucket,
  Combine,
  Type,
  Trash2,
  Minus,
  ArrowUpToLine,
  ArrowDownToLine,
  ArrowLeftToLine,
  ArrowRightToLine,
  Bookmark,
  Brush,
  FileText,
  Square,
  ListTree,
  Quote,
  Link2,
  Palette,
} from "@officeai/ui/sonaloop-icons";
import { TextFormatBar, cn } from "@officeai/ui";
import { InsertTableMenu } from "./InsertTableMenu";
import type { ActiveTextFormat, TextFormatProvider } from "@officeai/text-formatting";
import { Ribbon, ToolbarMenu, useAction, type RibbonCatalogue } from "../lib/shell";
import { docxActions } from "@officeai/docx";
import type { SelectedImageInfo } from "./ImageContextToolbar";
import type { PageZoneFocusDetail } from "@/lib/page-decorations";

export interface ToolbarStyleOption {
  value: string;
  label: string;
}

export type AlignmentValue = "left" | "center" | "right" | "justify";

/**
 * Effective spacing values surfaced to the toolbar for display. All
 * fields are optional because the cascade may not resolve them (e.g. a
 * paragraph in a synthetic doc with no styles part may have no `line`).
 */
export interface ResolvedSpacingDisplay {
  readonly line?: number;
  readonly lineRule?: "auto" | "exact" | "atLeast";
  readonly before?: number;
  readonly after?: number;
}

export interface ToolbarProps {
  agentReady: boolean;
  docInfo: { paragraphs: number; revision: number; commentThreads: number } | null;
  activeStyle: string;
  /**
   * Shared text-formatting provider. Owns selection-aware reads
   * (B/I/U/S, font family/size/color/highlight) and dispatches the
   * corresponding `docx:format-range` patches via the agent. Wired
   * up in `DocxEditor.tsx` per render.
   */
  textFormatProvider: TextFormatProvider;
  textFormatActive: ActiveTextFormat;
  /** Alignment of the paragraph containing the caret, or `null`. */
  activeAlignment: AlignmentValue | null;
  /**
   * Effective `<w:spacing>` for the paragraph containing the caret, after
   * the style cascade has been resolved. Drives the line-spacing dropdown
   * and the before/after numeric readouts. `null` when no paragraph is
   * focused.
   */
  activeSpacing: ResolvedSpacingDisplay | null;
  /**
   * Effective left indent in twips for the paragraph containing the
   * caret. `null` when no paragraph is focused. Pure display — the
   * existing ±360 buttons mutate via `onAdjustIndent`.
   */
  activeIndentLeft: number | null;
  /** Style picker contents derived from the loaded document. */
  styleOptions: ReadonlyArray<ToolbarStyleOption>;
  onInsertImage: () => void;
  /**
   * Insert a footnote reference at the current caret position via
   * `docx:insert-footnote`. Disabled when the caret is not in a body /
   * header / footer paragraph (the command's own rejection handler
   * surfaces the friendlier toast either way).
   */
  onInsertFootnote: () => void;
  onInsertTable: (rows: number, cols: number) => void;
  /**
   * Open the XlsxRangePicker dialog so the user can pick a .xlsx
   * file + sheet + range and decide whether to insert it as a
   * native table, a live-OLE embedded workbook, or a chart. Mirrors
   * the three "Insert from xlsx" entries in the action catalogue.
   */
  onInsertFromXlsx: () => void;
  onSetParagraphStyle: (style: string) => void;
  onSetAlignment: (alignment: AlignmentValue) => void;
  onAdjustIndent: (deltaTwips: number) => void;
  /**
   * Apply a `docx:set-paragraph-spacing` to the paragraph the caret is in.
   * Each field follows the command's `null = clear` / `undefined = leave`
   * semantics. The toolbar always sends `lineRule` alongside `line` to
   * keep the OOXML pair consistent.
   */
  onSetParagraphSpacing: (patch: {
    line?: number | null;
    lineRule?: "auto" | "exact" | "atLeast" | null;
    before?: number | null;
    after?: number | null;
  }) => void;
  onToggleList: (kind: "bullet" | "ordered") => void;
  /**
   * 9b — Word-parity list-level mini-menu. Set to a number 0..8 when
   * the caret sits in a list paragraph; null otherwise. The mini-menu
   * surfaces nine indents (matching Word's bullets dropdown) and
   * dispatches `docx:set-paragraph-list` preserving the existing
   * numId.
   */
  currentListLevel: number | null;
  onSetListLevel: (ilvl: number) => void;
  onAddComment: () => void;
  /**
   * B11 — Section break menu. The four OOXML section types map to
   * Word's Insert › Breaks submenu: Next page, Continuous, Even
   * page, Odd page. The toolbar dispatches this with `paragraphIndex`
   * resolved at the editor layer.
   */
  onInsertSectionBreak: (type: "nextPage" | "continuous" | "evenPage" | "oddPage") => void;
  /**
   * Surface "not yet supported" toasts for buttons whose backing command
   * does not yet exist.
   */
  onUnsupported: (label: string) => void;
  /**
   * Editor interaction mode (Word "Track Changes" surface):
   *   - `"edit"` — direct edits;
   *   - `"suggest"` — every insert/delete becomes a tracked
   *     `<w:ins>` / `<w:del>` revision (the Suggesting / redlining
   *     mode);
   *   - `"view"` — read-only.
   * The picker writes through to PM's `setEditMode` without
   * remounting so the cursor and selection survive mode changes.
   */
  editMode: EditModeValue;
  onSetEditMode: (mode: EditModeValue) => void;
  /** Word's pilcrow toggle — show/hide nonprinting characters. */
  formattingMarksOn: boolean;
  onToggleFormattingMarks: () => void;
  /**
   * B8 — Review tab. Number of unresolved tracked-change wrappers in
   * the document; when 0 the menu disables its Accept/Reject items.
   */
  trackedChangesCount: number;
  onAcceptAllChanges: () => void;
  onRejectAllChanges: () => void;
  /**
   * Open the Word-style Restrict Editing dialog (Review → Protect).
   * Mounted by `DocxEditor` so the toolbar stays presentational.
   */
  onOpenProtectDocument: () => void;
  /**
   * Whether the document currently has a `<w:documentProtection>`
   * element that's actively enforcing edits. Used to render the
   * Protect button's pressed state so users see at a glance that
   * the document is locked down.
   */
  documentProtectionActive: boolean;
  /**
   * Header/footer focus state. When set, the contextual "Kopf- und
   * Fußzeile" tab auto-activates and exposes its actions (insert
   * page number / page count / image, toggle different-first-page,
   * close H/F).
   */
  hfFocus: PageZoneFocusDetail | null;
  onCloseHeaderFooter: () => void;
  onToggleSectionDifferentFirst: (checked: boolean) => void;
  /**
   * Reflects the current section's `<w:titlePg/>` flag so the "Erste
   * Seite anders" checkbox shows whether the active section already
   * suppresses its first-page header/footer (Word's "Different First
   * Page" gesture). Computed by the editor from the snapshot at the
   * caret's section.
   */
  currentSectionTitlePg: boolean;
  onInsertHFField: (kind: "PAGE" | "NUMPAGES") => void;
  onInsertHFImage: () => void;
  /**
   * The currently NodeSelection-selected inline image, or `null`. Drives the
   * contextual "Bildtools" tab.
   */
  selectedImage: SelectedImageInfo | null;
  onEditImageAlt: (info: SelectedImageInfo) => void;
  onDeleteImage: (id: string) => void;
  /**
   * Tabledn id of the currently NodeSelection-selected table, or `null`.
   * Drives the contextual "Tabellentools" tab (row/column insertion).
   */
  selectedTableId: string | null;
  onInsertTableRow: (tableId: string, where: "top" | "bottom") => void;
  onInsertTableColumn: (tableId: string, where: "start" | "end") => void;
  onDeleteTableRow: (tableId: string, row: number) => void;
  onDeleteTableColumn: (tableId: string, column: number) => void;
  onDeleteTable: (tableId: string) => void;
  /**
   * Currently-targeted cell within the selected table. The Tabellentools
   * tab uses this to drive row-height / column-width / shading /
   * alignment commands without forcing the user to drag-select cell
   * boundaries (DOCX tables render as PM atoms today, so cell-level
   * caret editing isn't available — Word's own Table Properties
   * dialog falls back to the same row/column inputs).
   */
  activeTableCell: { row: number; column: number };
  onSetActiveTableCell: (cell: { row: number; column: number }) => void;
  /** Apply `<w:shd>` to the targeted cell. Pass `null` to clear. */
  onSetCellShading: (tableId: string, row: number, column: number, fill: string | null) => void;
  /** Apply `<w:vAlign>` to the targeted cell. Pass `null` to clear. */
  onSetCellAlignment: (
    tableId: string,
    row: number,
    column: number,
    vAlign: "top" | "center" | "bottom" | null
  ) => void;
  /**
   * Open the Word-style Bookmark dialog (References → Bookmark). The
   * dialog itself owns its modal state in {@link DocxEditor}; the
   * toolbar callback only flips it open. This pattern matches the
   * Restrict Editing dialog so the toolbar stays presentational.
   */
  onOpenBookmarkDialog: () => void;
  /**
   * Synthesize / refresh a Table of Contents (References → TOC).
   * Composed via the existing `docx:insert-paragraph` /
   * `docx:set-paragraph-style` commands — see
   * `handleInsertOrUpdateToc` in {@link DocxEditor}. `mode === "update"`
   * regenerates an existing TOC block in place; `"insert"` adds a
   * fresh one above the caret.
   */
  onInsertOrUpdateToc: (mode: "insert" | "update") => void;
  /**
   * Insert a Word-style "Caption" paragraph (e.g. "Figure 1: Diagram").
   * Composed via `docx:insert-paragraph` + `docx:set-paragraph-style`
   * with the `Caption` style id; the editor opens a small inline
   * prompt for the caption label so the user picks "Figure" /
   * "Table" / "Equation" + free text.
   */
  onInsertCaption: () => void;
  /**
   * Insert a cross-reference to an existing bookmark. The dialog
   * surfaces every bookmark currently in the document; "Insert"
   * splices a `→ {name}` reference text at the caret (an MVP that
   * resolves to a real `<w:fldSimple instr=" REF … "/>` in a future
   * pass — for now the text is plain so the link is at least
   * visually obvious).
   */
  onInsertCrossReference: () => void;
  /** Open the page-color picker (Design tab). */
  onOpenPageColorPicker: () => void;
  /** Open the page-borders dialog (Design tab). */
  onOpenPageBordersDialog: () => void;
  /** Open the watermark composer (Design tab). */
  onOpenWatermarkDialog: () => void;
  /** Open the document-theme picker (Design tab). */
  onOpenThemePicker: () => void;
  /** Apply `<w:trHeight>` to the targeted row. Pass `null` to clear. */
  onSetRowHeight: (
    tableId: string,
    row: number,
    heightTwips: number | null,
    rule?: "auto" | "exact" | "atLeast"
  ) => void;
  /** Apply `<w:tcW>` to the targeted column. */
  onSetColumnWidth: (tableId: string, column: number, widthTwips: number) => void;
  /** Merge `[fromColumn..toColumn]` in the targeted row. */
  onMergeCellsHorizontal: (tableId: string, row: number, fromColumn: number, toColumn: number) => void;
}

export type EditModeValue = "edit" | "suggest" | "view";

/**
 * Editor toolbar. Layout is intentionally close to Word's: file/style on
 * the left, inline marks + color in the middle, alignment / indent /
 * list on the right, and the export action pinned to the far right
 * with the doc metadata strip.
 *
 * Selection-binding contract (P2.2): every dropdown / pressed-state
 * button derives its value from `activeXxx` props that are recomputed
 * on every selection change in `DocxEditor.tsx`. `MIXED` selections
 * render as "—" in the FontSize/FontFamily pickers; `undefined` means
 * "no mark on this run" and renders blank.
 */
export function Toolbar(props: ToolbarProps): ReactNode {
  // Toolbar buttons that double as Cmd+K palette actions read their
  // label/shortcut from the central docx action catalogue. A typo'd
  // id throws at first render (loud failure beats a silently
  // mislabelled button) so adding/renaming an action only happens in
  // one place — packages/docx/src/actions/catalogue.ts.
  const insertImageAction = useAction(docxActions, "docx.insert-image");
  const insertFootnoteAction = useAction(docxActions, "docx.insert-footnote");
  const addCommentAction = useAction(docxActions, "docx.add-comment");

  // Catalogue is rebuilt every render — that's cheap (it's plain
  // descriptor objects with closures over current `props`), and it
  // ensures every render fn captures the latest callbacks. The
  // Ribbon does its own visibility/auto-activate work over this
  // structure.
  const catalogue = useMemo<RibbonCatalogue<DocxRibbonCtx>>(
    () =>
      buildDocxRibbonCatalogue({
        agentReady: props.agentReady,
        insertImageLabel: insertImageAction.label,
        insertFootnoteLabel: insertFootnoteAction.label,
        addCommentLabel: addCommentAction.label,
      }),
    [props.agentReady, insertImageAction.label, insertFootnoteAction.label, addCommentAction.label]
  );

  const ctx: DocxRibbonCtx = props;

  return (
    <Ribbon
      ariaLabel="Document toolbar"
      testId="docx-toolbar"
      catalogue={catalogue}
      ctx={ctx}
      trailing={
        <div className="flex items-center gap-3 text-xs text-secondary">
          {props.docInfo && (
            <span className="hidden whitespace-nowrap md:inline">
              {props.docInfo.paragraphs} paragraphs · rev {props.docInfo.revision} ·{" "}
              {props.docInfo.commentThreads} comment{props.docInfo.commentThreads === 1 ? "" : "s"}
            </span>
          )}
          <ReviewMenu
            count={props.trackedChangesCount}
            onAcceptAll={props.onAcceptAllChanges}
            onRejectAll={props.onRejectAllChanges}
          />
          <EditModePicker value={props.editMode} onChange={props.onSetEditMode} />
        </div>
      }
    />
  );
}

/**
 * Ribbon ctx for DOCX. We just pass `ToolbarProps` through; the
 * group `render` fns destructure what they need. Keeping the ctx
 * shape identical to the props avoids having to thread two parallel
 * shapes through `DocxEditor.tsx`.
 */
type DocxRibbonCtx = ToolbarProps;

interface DocxRibbonOptions {
  readonly agentReady: boolean;
  readonly insertImageLabel: string;
  readonly insertFootnoteLabel: string;
  readonly addCommentLabel: string;
}

/**
 * Office subset-pragmatic catalogue. Tab names mirror Word DE
 * (Start / Einfügen / Layout / Überprüfen / Ansicht). Contextual
 * tabs follow the same naming (Bildtools, Tabellentools, Kopf- und
 * Fußzeile). Groups inside a tab are bottom-labelled by the Ribbon
 * primitive so the surface reads like a real Word ribbon.
 */
function buildDocxRibbonCatalogue(opts: DocxRibbonOptions): RibbonCatalogue<DocxRibbonCtx> {
  const { agentReady, insertImageLabel, insertFootnoteLabel, addCommentLabel } = opts;
  return {
    defaultTabId: "start",
    tabs: [
      {
        id: "start",
        label: "Start",
        groups: [
          {
            id: "styles",
            label: "Formatvorlagen",
            render: (ctx) => (
              <ParagraphStylePicker
                value={ctx.activeStyle}
                options={ctx.styleOptions}
                onChange={ctx.onSetParagraphStyle}
                disabled={!agentReady}
              />
            ),
          },
          {
            id: "font",
            label: "Schriftart",
            render: (ctx) => (
              <TextFormatBar
                provider={ctx.textFormatProvider}
                active={ctx.textFormatActive}
                disabled={!agentReady}
              />
            ),
          },
          {
            id: "paragraph",
            label: "Absatz",
            render: (ctx) => (
              <>
                <ToolbarBtn label="Bullet list" onClick={() => ctx.onToggleList("bullet")}>
                  <List size={14} />
                </ToolbarBtn>
                <ToolbarBtn label="Numbered list" onClick={() => ctx.onToggleList("ordered")}>
                  <ListOrdered size={14} />
                </ToolbarBtn>
                <ListLevelMenu
                  disabled={!agentReady || ctx.currentListLevel === null}
                  level={ctx.currentListLevel}
                  onPick={ctx.onSetListLevel}
                />
                <ToolbarBtn label="Decrease indent" onClick={() => ctx.onAdjustIndent(-360)}>
                  <Outdent size={14} />
                </ToolbarBtn>
                <ToolbarBtn label="Increase indent" onClick={() => ctx.onAdjustIndent(360)}>
                  <Indent size={14} />
                </ToolbarBtn>
                <ToolbarBtn
                  label="Align left"
                  active={ctx.activeAlignment === "left"}
                  onClick={() => ctx.onSetAlignment("left")}
                >
                  <AlignLeft size={14} />
                </ToolbarBtn>
                <ToolbarBtn
                  label="Align center"
                  active={ctx.activeAlignment === "center"}
                  onClick={() => ctx.onSetAlignment("center")}
                >
                  <AlignCenter size={14} />
                </ToolbarBtn>
                <ToolbarBtn
                  label="Align right"
                  active={ctx.activeAlignment === "right"}
                  onClick={() => ctx.onSetAlignment("right")}
                >
                  <AlignRight size={14} />
                </ToolbarBtn>
                <ToolbarBtn
                  label="Align justify"
                  active={ctx.activeAlignment === "justify"}
                  onClick={() => ctx.onSetAlignment("justify")}
                >
                  <AlignJustify size={14} />
                </ToolbarBtn>
                <SpacingMenu
                  spacing={ctx.activeSpacing}
                  onApply={ctx.onSetParagraphSpacing}
                  disabled={!agentReady || ctx.activeSpacing === null}
                />
                <ToolbarBtn
                  label="Show formatting marks"
                  active={ctx.formattingMarksOn}
                  onClick={ctx.onToggleFormattingMarks}
                >
                  <Pilcrow size={14} />
                </ToolbarBtn>
              </>
            ),
          },
        ],
      },
      {
        id: "insert",
        label: "Einfügen",
        groups: [
          {
            id: "tables",
            label: "Tabellen",
            render: (ctx) => (
              <>
                <InsertTableMenu disabled={!agentReady} onInsert={ctx.onInsertTable} />
                <ToolbarBtn label="Insert from xlsx" onClick={ctx.onInsertFromXlsx}>
                  <Table2 size={14} />
                </ToolbarBtn>
              </>
            ),
          },
          {
            id: "illustrations",
            label: "Illustrationen",
            render: (ctx) => (
              <ToolbarBtn label={insertImageLabel} onClick={ctx.onInsertImage}>
                <ImageIcon size={14} />
              </ToolbarBtn>
            ),
          },
          {
            id: "comments",
            label: "Kommentare",
            render: (ctx) => (
              <ToolbarBtn label={addCommentLabel} onClick={ctx.onAddComment}>
                <MessageSquarePlus size={14} />
              </ToolbarBtn>
            ),
          },
          {
            id: "footnotes",
            label: "Fußnoten",
            render: (ctx) => (
              <ToolbarBtn label={insertFootnoteLabel} onClick={ctx.onInsertFootnote}>
                <StickyNote size={14} />
              </ToolbarBtn>
            ),
          },
        ],
      },
      {
        id: "layout",
        label: "Layout",
        groups: [
          {
            id: "page-setup",
            label: "Seite einrichten",
            render: (ctx) => <SectionBreakMenu disabled={!agentReady} onInsert={ctx.onInsertSectionBreak} />,
          },
          {
            id: "paragraph-spacing",
            label: "Absatz",
            render: (ctx) => (
              <SpacingMenu
                spacing={ctx.activeSpacing}
                onApply={ctx.onSetParagraphSpacing}
                disabled={!agentReady || ctx.activeSpacing === null}
              />
            ),
          },
          {
            id: "indent",
            label: "Einzug",
            render: (ctx) => (
              <>
                <ToolbarBtn label="Decrease indent" onClick={() => ctx.onAdjustIndent(-360)}>
                  <Outdent size={14} />
                </ToolbarBtn>
                <ToolbarBtn label="Increase indent" onClick={() => ctx.onAdjustIndent(360)}>
                  <Indent size={14} />
                </ToolbarBtn>
                <span
                  className="inline-block min-w-[3.25rem] px-1 text-[11px] tabular-nums text-secondary"
                  title="Left indent"
                  aria-hidden={ctx.activeIndentLeft === null || ctx.activeIndentLeft <= 0}
                >
                  {ctx.activeIndentLeft !== null && ctx.activeIndentLeft > 0
                    ? `${twipsToInches(ctx.activeIndentLeft)}"`
                    : ""}
                </span>
              </>
            ),
          },
        ],
      },
      {
        // Phase 9c §3b — Design / Entwurf tab.
        //
        // The four backend commands listed in the plan (`docx:set-page-color`,
        // `docx:set-page-borders`, `docx:set-page-watermark`,
        // `docx:set-document-theme`) require new typed model fields on
        // `DocxDocument` (background block, sectPr.pgBorders writer, header
        // part for watermarks, theme1.xml writer) plus parser/serializer
        // support that we don't have yet. Per the same pattern used for
        // the 9b PPTX shape-outline placeholders, we render disabled
        // ribbon buttons here so the planned ribbon shape is visible
        // and a follow-up plan only needs to flip `disabled` and wire
        // `onClick`. CLI exposure stays deferred until backends land.
        id: "design",
        label: "Entwurf",
        groups: [
          {
            id: "themes",
            label: "Designs",
            render: (ctx) => (
              <ToolbarBtn
                label="Document theme"
                onClick={ctx.onOpenThemePicker}
                testId="docx-set-document-theme"
              >
                <Palette size={14} />
                <span className="ml-1 text-xs">Designs</span>
              </ToolbarBtn>
            ),
          },
          {
            id: "page-background",
            label: "Seitenhintergrund",
            render: (ctx) => (
              <>
                <ToolbarBtn
                  label="Page color"
                  onClick={ctx.onOpenPageColorPicker}
                  testId="docx-set-page-color"
                >
                  <PaintBucket size={14} />
                  <span className="ml-1 text-xs">Seitenfarbe</span>
                </ToolbarBtn>
                <ToolbarBtn
                  label="Page borders"
                  onClick={ctx.onOpenPageBordersDialog}
                  testId="docx-set-page-borders"
                >
                  <Square size={14} />
                  <span className="ml-1 text-xs">Seitenränder</span>
                </ToolbarBtn>
                <ToolbarBtn
                  label="Watermark"
                  onClick={ctx.onOpenWatermarkDialog}
                  testId="docx-set-page-watermark"
                >
                  <Brush size={14} />
                  <span className="ml-1 text-xs">Wasserzeichen</span>
                </ToolbarBtn>
              </>
            ),
          },
        ],
      },
      {
        // Phase 9c §3c — References / Verweise tab.
        //
        // Identical Coming-soon strategy: bookmarks need a typed
        // `<w:bookmarkStart/End>` model + serializer pair, TOC needs
        // `<w:fldSimple>` + cached field result, captions/cross-refs
        // need new run kinds. Each will get its own dedicated plan.
        // Until then, surfacing the planned ribbon group here is the
        // honest progress: CLI users see no entries, UI users see the
        // intended group with disabled triggers.
        id: "references",
        label: "Verweise",
        groups: [
          {
            id: "bookmarks",
            label: "Lesezeichen",
            render: (ctx) => (
              <ToolbarBtn
                label="Insert bookmark"
                onClick={ctx.onOpenBookmarkDialog}
                testId="docx-insert-bookmark"
              >
                <Bookmark size={14} />
                <span className="ml-1 text-xs">Lesezeichen</span>
              </ToolbarBtn>
            ),
          },
          {
            id: "toc",
            label: "Inhaltsverzeichnis",
            render: (ctx) => (
              <>
                <ToolbarBtn
                  label="Insert table of contents"
                  onClick={() => ctx.onInsertOrUpdateToc("insert")}
                  testId="docx-insert-toc"
                >
                  <ListTree size={14} />
                  <span className="ml-1 text-xs">TOC</span>
                </ToolbarBtn>
                <ToolbarBtn
                  label="Update existing TOC"
                  onClick={() => ctx.onInsertOrUpdateToc("update")}
                  testId="docx-update-toc"
                >
                  <FileText size={14} />
                  <span className="ml-1 text-xs">Aktualisieren</span>
                </ToolbarBtn>
              </>
            ),
          },
          {
            id: "captions",
            label: "Beschriftungen",
            render: (ctx) => (
              <>
                <ToolbarBtn label="Insert caption" onClick={ctx.onInsertCaption} testId="docx-insert-caption">
                  <Quote size={14} />
                  <span className="ml-1 text-xs">Beschriftung</span>
                </ToolbarBtn>
                <ToolbarBtn
                  label="Insert cross-reference"
                  onClick={ctx.onInsertCrossReference}
                  testId="docx-insert-cross-reference"
                >
                  <Link2 size={14} />
                  <span className="ml-1 text-xs">Querverweis</span>
                </ToolbarBtn>
              </>
            ),
          },
        ],
      },
      {
        id: "review",
        label: "Überprüfen",
        groups: [
          {
            id: "comments",
            label: "Kommentare",
            render: (ctx) => (
              <ToolbarBtn label={addCommentLabel} onClick={ctx.onAddComment}>
                <MessageSquarePlus size={14} />
              </ToolbarBtn>
            ),
          },
          {
            id: "tracked",
            label: "Nachverfolgung",
            render: (ctx) => (
              <ReviewMenu
                count={ctx.trackedChangesCount}
                onAcceptAll={ctx.onAcceptAllChanges}
                onRejectAll={ctx.onRejectAllChanges}
              />
            ),
          },
          {
            id: "protect",
            label: "Schützen",
            render: (ctx) => (
              <ToolbarBtn
                label={
                  ctx.documentProtectionActive
                    ? "Bearbeitung einschränken (aktiv)"
                    : "Bearbeitung einschränken"
                }
                onClick={ctx.onOpenProtectDocument}
                active={ctx.documentProtectionActive}
                testId="docx-protect-document"
              >
                <Lock size={14} />
                <span className="ml-1 text-xs">Schützen</span>
              </ToolbarBtn>
            ),
          },
        ],
      },
      {
        id: "view",
        label: "Ansicht",
        groups: [
          {
            id: "show",
            label: "Anzeigen",
            render: (ctx) => (
              <ToolbarBtn
                label="Show formatting marks"
                active={ctx.formattingMarksOn}
                onClick={ctx.onToggleFormattingMarks}
              >
                <Pilcrow size={14} />
              </ToolbarBtn>
            ),
          },
        ],
      },
      {
        id: "image-tools",
        label: "Bildtools",
        contextual: { accent: "image" },
        visible: (ctx) => ctx.selectedImage !== null,
        autoActivateWhen: (ctx) => ctx.selectedImage !== null,
        groups: [
          {
            id: "alt",
            label: "Bildbeschreibung",
            render: (ctx) =>
              ctx.selectedImage ? (
                <ToolbarBtn
                  label="Edit alt text"
                  onClick={() => {
                    const sel = ctx.selectedImage;
                    if (sel) ctx.onEditImageAlt(sel);
                  }}
                >
                  <Type size={14} />
                  <span className="ml-1 text-xs">Alt-Text</span>
                </ToolbarBtn>
              ) : null,
          },
          {
            id: "size",
            label: "Größe",
            render: (ctx) =>
              ctx.selectedImage ? (
                <span className="px-2 text-xs text-secondary tabular-nums">
                  {ctx.selectedImage.widthPx} × {ctx.selectedImage.heightPx} px
                </span>
              ) : null,
          },
          {
            id: "delete",
            label: "Aktionen",
            render: (ctx) =>
              ctx.selectedImage ? (
                <ToolbarBtn
                  label="Delete image"
                  onClick={() => {
                    const sel = ctx.selectedImage;
                    if (sel) ctx.onDeleteImage(sel.imageId);
                  }}
                >
                  <Trash2 size={14} />
                </ToolbarBtn>
              ) : null,
          },
        ],
      },
      {
        id: "table-tools",
        label: "Tabellentools",
        contextual: { accent: "table" },
        visible: (ctx) => ctx.selectedTableId !== null,
        autoActivateWhen: (ctx) => ctx.selectedTableId !== null,
        groups: [
          {
            id: "rows",
            label: "Zeilen",
            render: (ctx) =>
              ctx.selectedTableId ? (
                <>
                  <ToolbarBtn
                    label="Insert row above"
                    onClick={() => {
                      const id = ctx.selectedTableId;
                      if (id) ctx.onInsertTableRow(id, "top");
                    }}
                  >
                    <ArrowUpToLine size={14} />
                  </ToolbarBtn>
                  <ToolbarBtn
                    label="Insert row below"
                    onClick={() => {
                      const id = ctx.selectedTableId;
                      if (id) ctx.onInsertTableRow(id, "bottom");
                    }}
                  >
                    <ArrowDownToLine size={14} />
                  </ToolbarBtn>
                </>
              ) : null,
          },
          {
            id: "columns",
            label: "Spalten",
            render: (ctx) =>
              ctx.selectedTableId ? (
                <>
                  <ToolbarBtn
                    label="Insert column at start"
                    onClick={() => {
                      const id = ctx.selectedTableId;
                      if (id) ctx.onInsertTableColumn(id, "start");
                    }}
                  >
                    <ArrowLeftToLine size={14} />
                  </ToolbarBtn>
                  <ToolbarBtn
                    label="Insert column at end"
                    onClick={() => {
                      const id = ctx.selectedTableId;
                      if (id) ctx.onInsertTableColumn(id, "end");
                    }}
                  >
                    <ArrowRightToLine size={14} />
                  </ToolbarBtn>
                </>
              ) : null,
          },
          {
            id: "delete",
            label: "Löschen",
            render: (ctx) =>
              ctx.selectedTableId ? (
                <>
                  <ToolbarBtn
                    label="Zeile löschen"
                    testId="docx-delete-row"
                    onClick={() => {
                      const id = ctx.selectedTableId;
                      if (id) ctx.onDeleteTableRow(id, ctx.activeTableCell.row);
                    }}
                  >
                    <Minus size={14} />
                    <span className="ml-1 text-xs">Zeile</span>
                  </ToolbarBtn>
                  <ToolbarBtn
                    label="Spalte löschen"
                    testId="docx-delete-column"
                    onClick={() => {
                      const id = ctx.selectedTableId;
                      if (id) ctx.onDeleteTableColumn(id, ctx.activeTableCell.column);
                    }}
                  >
                    <Minus size={14} />
                    <span className="ml-1 text-xs">Spalte</span>
                  </ToolbarBtn>
                  <ToolbarBtn
                    label="Tabelle löschen"
                    testId="docx-delete-table"
                    onClick={() => {
                      const id = ctx.selectedTableId;
                      if (id) ctx.onDeleteTable(id);
                    }}
                  >
                    <Trash2 size={14} />
                    <span className="ml-1 text-xs">Tabelle</span>
                  </ToolbarBtn>
                </>
              ) : null,
          },
          {
            id: "active-cell",
            label: "Zielzelle",
            render: (ctx) => (
              <ActiveCellPicker
                row={ctx.activeTableCell.row}
                column={ctx.activeTableCell.column}
                onChange={ctx.onSetActiveTableCell}
              />
            ),
          },
          {
            id: "cell-size",
            label: "Größe",
            render: (ctx) =>
              ctx.selectedTableId ? (
                <CellSizeControls
                  tableId={ctx.selectedTableId}
                  cell={ctx.activeTableCell}
                  onSetRowHeight={ctx.onSetRowHeight}
                  onSetColumnWidth={ctx.onSetColumnWidth}
                />
              ) : null,
          },
          {
            id: "cell-align",
            label: "Ausrichtung",
            render: (ctx) =>
              ctx.selectedTableId ? (
                <>
                  <ToolbarBtn
                    label="Vertikal oben"
                    testId="docx-cell-align-top"
                    onClick={() => {
                      const id = ctx.selectedTableId;
                      if (id)
                        ctx.onSetCellAlignment(
                          id,
                          ctx.activeTableCell.row,
                          ctx.activeTableCell.column,
                          "top"
                        );
                    }}
                  >
                    <AlignVerticalJustifyStart size={14} />
                  </ToolbarBtn>
                  <ToolbarBtn
                    label="Vertikal mittig"
                    testId="docx-cell-align-center"
                    onClick={() => {
                      const id = ctx.selectedTableId;
                      if (id)
                        ctx.onSetCellAlignment(
                          id,
                          ctx.activeTableCell.row,
                          ctx.activeTableCell.column,
                          "center"
                        );
                    }}
                  >
                    <AlignVerticalJustifyCenter size={14} />
                  </ToolbarBtn>
                  <ToolbarBtn
                    label="Vertikal unten"
                    testId="docx-cell-align-bottom"
                    onClick={() => {
                      const id = ctx.selectedTableId;
                      if (id)
                        ctx.onSetCellAlignment(
                          id,
                          ctx.activeTableCell.row,
                          ctx.activeTableCell.column,
                          "bottom"
                        );
                    }}
                  >
                    <AlignVerticalJustifyEnd size={14} />
                  </ToolbarBtn>
                </>
              ) : null,
          },
          {
            id: "cell-design",
            label: "Entwurf",
            render: (ctx) =>
              ctx.selectedTableId ? (
                <CellShadingPicker
                  tableId={ctx.selectedTableId}
                  cell={ctx.activeTableCell}
                  onSetCellShading={ctx.onSetCellShading}
                />
              ) : null,
          },
          {
            id: "cell-merge",
            label: "Verbinden",
            render: (ctx) =>
              ctx.selectedTableId ? (
                <MergeCellsControl
                  tableId={ctx.selectedTableId}
                  row={ctx.activeTableCell.row}
                  fromColumn={ctx.activeTableCell.column}
                  onMerge={ctx.onMergeCellsHorizontal}
                />
              ) : null,
          },
        ],
      },
      {
        id: "hf-tools",
        label: "Kopf- und Fußzeile",
        contextual: { accent: "hf" },
        visible: (ctx) => ctx.hfFocus !== null,
        autoActivateWhen: (ctx) => ctx.hfFocus !== null,
        groups: [
          {
            id: "hf-status",
            label: "Position",
            render: (ctx) =>
              ctx.hfFocus ? (
                <span className="inline-flex items-center gap-1.5 px-2 text-xs font-medium text-foreground">
                  <span className="inline-block h-1.5 w-1.5 rounded-full bg-[var(--accent)]" aria-hidden />
                  {ctx.hfFocus.slot === "header" ? "Kopfzeile" : "Fußzeile"}
                  {ctx.hfFocus.target && ctx.hfFocus.target !== "default" ? (
                    <span className="rounded bg-divider/60 px-1.5 py-0.5 text-[10px] uppercase tracking-wide text-secondary">
                      {ctx.hfFocus.target}
                    </span>
                  ) : null}
                  {ctx.hfFocus.pageNumber ? (
                    <span className="font-normal text-secondary">· Seite {ctx.hfFocus.pageNumber}</span>
                  ) : null}
                </span>
              ) : null,
          },
          {
            id: "hf-options",
            label: "Optionen",
            render: (ctx) => (
              <div className="flex flex-col gap-0.5">
                <label className="inline-flex items-center gap-1.5 px-2 text-xs text-foreground hover:text-foreground">
                  <input
                    type="checkbox"
                    onMouseDown={(e) => e.preventDefault()}
                    checked={ctx.currentSectionTitlePg}
                    onChange={(e) => ctx.onToggleSectionDifferentFirst(e.target.checked)}
                    data-testid="docx-hf-toggle-different-first"
                    className="h-3 w-3 accent-[var(--accent)]"
                  />
                  Erste Seite anders
                </label>
                {/* "Different odd & even" is a real Word feature (`<w:evenAndOddHeaders/>`
                 * on `settings.xml`) but the backend command does not exist yet.
                 * Surfacing the toggle anyway and routing it through `onUnsupported`
                 * means users get a "not yet supported" toast instead of a silently
                 * dead checkbox — and exercises the prop wired through Toolbar/Catalogue. */}
                <label className="inline-flex items-center gap-1.5 px-2 text-xs text-secondary hover:text-foreground">
                  <input
                    type="checkbox"
                    onMouseDown={(e) => e.preventDefault()}
                    checked={false}
                    onChange={() => ctx.onUnsupported("Verschiedene gerade/ungerade Seiten")}
                    data-testid="docx-hf-toggle-different-odd-even"
                    className="h-3 w-3 accent-[var(--accent)]"
                  />
                  Verschiedene gerade/ungerade
                </label>
              </div>
            ),
          },
          {
            id: "hf-insert",
            label: "Einfügen",
            render: (ctx) => (
              <>
                <ToolbarBtn
                  label="Insert page #"
                  testId="docx-hf-insert-page-number"
                  onClick={() => ctx.onInsertHFField("PAGE")}
                >
                  <Hash size={14} />
                  <span className="ml-1 text-xs">Seitenzahl</span>
                </ToolbarBtn>
                <ToolbarBtn
                  label="Insert page count"
                  testId="docx-hf-insert-page-count"
                  onClick={() => ctx.onInsertHFField("NUMPAGES")}
                >
                  <ListOrderedIcon size={14} />
                  <span className="ml-1 text-xs">Seitenanzahl</span>
                </ToolbarBtn>
                <ToolbarBtn label="Insert image" testId="docx-hf-insert-image" onClick={ctx.onInsertHFImage}>
                  <ImageIcon size={14} />
                  <span className="ml-1 text-xs">Bild</span>
                </ToolbarBtn>
              </>
            ),
          },
          {
            id: "hf-close",
            label: "Schließen",
            render: (ctx) => (
              <ToolbarBtn
                label="Close header & footer"
                testId="docx-hf-close"
                onClick={ctx.onCloseHeaderFooter}
              >
                <X size={14} />
                <span className="ml-1 text-xs">Schließen</span>
              </ToolbarBtn>
            ),
          },
        ],
      },
    ],
  };
}

// Re-export for E2E tests that key off the H/F insert buttons by their
// stable testIds (docx-hf-insert-page-number / -count / -image / close /
// toggle-different-first). The Ribbon catalogue render fns above carry
// these testIds inline via the underlying ToolbarBtn `data-testid`
// attribute; we keep the constants here so future renames are caught
// statically by the type system.
const HF_TESTIDS = {
  pageNumber: "docx-hf-insert-page-number",
  pageCount: "docx-hf-insert-page-count",
  image: "docx-hf-insert-image",
  close: "docx-hf-close",
  toggleDifferentFirst: "docx-hf-toggle-different-first",
} as const;
void HF_TESTIDS;

/**
 * Edit-mode picker (Editing / Suggesting / Viewing — the Word and
 * Google Docs surface). Visualises the current mode with a colored
 * pill so the user always knows whether their next keystroke goes
 * to plain text, becomes a tracked suggestion, or is rejected
 * outright by the read-only surface.
 *
 * Click the pill to open a simple menu; the previously-selected
 * option is highlighted. Mode changes are dispatched immediately
 * (no confirm step) — the user can always revert by re-opening the
 * picker.
 */
function EditModePicker(props: { value: EditModeValue; onChange: (v: EditModeValue) => void }): ReactNode {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    if (!open) return;
    const onDoc = (e: MouseEvent) => {
      if (!ref.current) return;
      if (!ref.current.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, [open]);

  const meta = EDIT_MODE_META[props.value];
  return (
    <div ref={ref} className="relative">
      <button
        type="button"
        title={`Edit mode: ${meta.label}`}
        aria-label={`Edit mode: ${meta.label}`}
        aria-haspopup="menu"
        aria-expanded={open}
        onClick={() => setOpen((v) => !v)}
        className={cn(
          "inline-flex items-center gap-1.5 rounded-md border px-2 py-1 text-xs font-medium transition-colors",
          meta.pillClass
        )}
        data-testid="edit-mode-picker"
        data-edit-mode={props.value}
      >
        <meta.Icon size={12} />
        {meta.label}
        <ChevronDown size={10} />
      </button>
      {open && (
        <div
          role="menu"
          className="absolute right-0 z-30 mt-1 w-56 rounded-md border border-divider bg-surface p-1 text-xs shadow-md"
        >
          {(Object.keys(EDIT_MODE_META) as EditModeValue[]).map((key) => {
            const m = EDIT_MODE_META[key];
            const active = key === props.value;
            return (
              <button
                key={key}
                type="button"
                role="menuitemradio"
                aria-checked={active}
                onClick={() => {
                  props.onChange(key);
                  setOpen(false);
                }}
                data-testid={`edit-mode-option-${key}`}
                className={cn(
                  "flex w-full items-start gap-2 rounded px-2 py-1.5 text-left hover:bg-hover",
                  active && "bg-hover"
                )}
              >
                <m.Icon size={12} className={cn("mt-0.5 shrink-0", m.iconColorClass)} />
                <span className="min-w-0 flex-1">
                  <span className="block font-medium text-foreground">{m.label}</span>
                  <span className="block text-[11px] text-secondary">{m.description}</span>
                </span>
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}

/**
 * 9b — List-level mini-menu next to the bullet/numbered buttons.
 *
 * Surfaces the nine indent levels Word exposes in its bullets ribbon
 * dropdown. Disabled (and visually muted) whenever the caret isn't in
 * a list paragraph — that's the only time `set-paragraph-list` would
 * have a meaningful target. The menu also supports the keyboard chord
 * Tab / Shift+Tab via `wordShortcutsKeymapPlugin`; this dropdown is
 * the discoverable / point-and-click affordance for the same action.
 *
 * The active level is highlighted so the user can see at a glance how
 * deep they are without counting bullets on screen.
 */
function ListLevelMenu(props: {
  disabled: boolean;
  level: number | null;
  onPick: (ilvl: number) => void;
}): ReactNode {
  const [open, setOpen] = useState(false);
  const triggerRef = useRef<HTMLButtonElement | null>(null);

  const choose = (ilvl: number) => {
    props.onPick(ilvl);
    setOpen(false);
  };

  return (
    <span className="inline-flex">
      <button
        ref={triggerRef}
        type="button"
        title={
          props.level === null
            ? "List level (place caret in a list)"
            : `List level (currently ${props.level + 1})`
        }
        aria-label="Change list level"
        aria-haspopup="menu"
        aria-expanded={open}
        disabled={props.disabled}
        onClick={() => setOpen((v) => !v)}
        className={cn(
          "inline-flex h-7 items-center gap-1 rounded-md border border-transparent px-1.5 text-xs text-foreground transition-colors hover:bg-hover",
          props.disabled && "cursor-not-allowed opacity-50 hover:bg-transparent"
        )}
        data-testid="list-level-menu-trigger"
      >
        <Indent size={12} />
        <span className="tabular-nums">{props.level === null ? "—" : props.level + 1}</span>
        <ChevronDown size={10} />
      </button>
      <ToolbarMenu
        open={open}
        onClose={() => setOpen(false)}
        triggerRef={triggerRef}
        role="menu"
        className="w-48 rounded-md border border-divider bg-surface p-1 text-xs shadow-md"
      >
        {Array.from({ length: 9 }, (_, i) => (
          <button
            key={i}
            type="button"
            role="menuitem"
            onClick={() => choose(i)}
            data-testid={`list-level-${i}`}
            className={cn(
              "flex w-full items-center gap-2 rounded px-2 py-1 text-left text-xs hover:bg-hover",
              props.level === i && "bg-hover font-semibold"
            )}
          >
            <span className="tabular-nums text-secondary">{i + 1}.</span>
            <span style={{ paddingLeft: i * 8 }}>•</span>
            <span className="ml-auto text-tertiary">ilvl {i}</span>
          </button>
        ))}
      </ToolbarMenu>
    </span>
  );
}

/**
 * B11 — Section break menu. Mirrors Word's Insert › Breaks submenu
 * with the four legal OOXML section types. The shortcut for the
 * common "Next page" entry (Mod+Shift+Enter) is surfaced inline so
 * keyboard-first users can discover it.
 */
function SectionBreakMenu(props: {
  disabled: boolean;
  onInsert: (type: "nextPage" | "continuous" | "evenPage" | "oddPage") => void;
}): ReactNode {
  const [open, setOpen] = useState(false);
  const triggerRef = useRef<HTMLButtonElement | null>(null);

  const choose = (type: "nextPage" | "continuous" | "evenPage" | "oddPage") => {
    props.onInsert(type);
    setOpen(false);
  };

  return (
    <span className="inline-flex">
      <button
        ref={triggerRef}
        type="button"
        title="Insert section break"
        aria-label="Insert section break"
        aria-haspopup="menu"
        aria-expanded={open}
        disabled={props.disabled}
        onClick={() => setOpen((v) => !v)}
        className={cn(
          "inline-flex h-7 items-center gap-1 rounded-md border border-transparent px-1.5 text-xs text-foreground transition-colors hover:bg-hover",
          props.disabled && "cursor-not-allowed opacity-50 hover:bg-transparent"
        )}
        data-testid="section-break-menu-button"
      >
        <SeparatorHorizontal size={14} />
        <ChevronDown size={10} />
      </button>
      <ToolbarMenu
        open={open}
        onClose={() => setOpen(false)}
        triggerRef={triggerRef}
        role="menu"
        className="w-64 rounded-md border border-divider bg-surface p-1 text-xs shadow-md"
      >
        <SectionBreakMenuItem
          label="Next page"
          description="Start the next section on a new page."
          shortcut="Mod+Shift+Enter"
          onClick={() => choose("nextPage")}
          testId="section-break-next-page"
        />
        <SectionBreakMenuItem
          label="Continuous"
          description="Begin a new section without a page break."
          onClick={() => choose("continuous")}
          testId="section-break-continuous"
        />
        <SectionBreakMenuItem
          label="Even page"
          description="Start the next section on the next even-numbered page."
          onClick={() => choose("evenPage")}
          testId="section-break-even"
        />
        <SectionBreakMenuItem
          label="Odd page"
          description="Start the next section on the next odd-numbered page."
          onClick={() => choose("oddPage")}
          testId="section-break-odd"
        />
      </ToolbarMenu>
    </span>
  );
}

function SectionBreakMenuItem(props: {
  label: string;
  description: string;
  shortcut?: string;
  onClick: () => void;
  testId: string;
}): ReactNode {
  return (
    <button
      type="button"
      role="menuitem"
      onClick={props.onClick}
      data-testid={props.testId}
      className="flex w-full items-start gap-2 rounded px-2 py-1.5 text-left hover:bg-hover"
    >
      <span className="min-w-0 flex-1">
        <span className="block font-medium text-foreground">{props.label}</span>
        <span className="block text-[11px] text-secondary">{props.description}</span>
      </span>
      {props.shortcut ? (
        <kbd className="ml-2 shrink-0 rounded bg-hover px-1 py-0.5 font-mono text-[10px] text-secondary">
          {props.shortcut}
        </kbd>
      ) : null}
    </button>
  );
}

/**
 * B8 — Review menu.
 *
 * Word's "Review" tab condensed into a single menu next to the edit
 * mode picker so it stays visible at every viewport width. Carries
 * a small badge with the live count of unresolved revisions; the
 * Accept-all / Reject-all entries are disabled when the count is
 * zero so the menu becomes a stable affordance instead of jumping
 * in and out of the layout.
 */
function ReviewMenu(props: { count: number; onAcceptAll: () => void; onRejectAll: () => void }): ReactNode {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    if (!open) return;
    const onDoc = (e: MouseEvent) => {
      if (!ref.current) return;
      if (!ref.current.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, [open]);

  const empty = props.count === 0;
  return (
    <div ref={ref} className="relative">
      <button
        type="button"
        title="Review tracked changes"
        aria-label="Review tracked changes"
        aria-haspopup="menu"
        aria-expanded={open}
        onClick={() => setOpen((v) => !v)}
        className={cn(
          "inline-flex items-center gap-1.5 rounded-md border border-divider bg-surface px-2 py-1 text-xs font-medium text-foreground transition-colors hover:bg-hover"
        )}
        data-testid="review-menu-button"
      >
        <ScrollText size={12} />
        Review
        {!empty && (
          <span
            className="rounded-full bg-accent px-1.5 py-0.5 text-[10px] font-semibold leading-none text-white"
            aria-label={`${props.count} unresolved tracked change${props.count === 1 ? "" : "s"}`}
          >
            {props.count}
          </span>
        )}
        <ChevronDown size={10} />
      </button>
      {open && (
        <div
          role="menu"
          className="absolute right-0 z-30 mt-1 w-64 rounded-md border border-divider bg-surface p-1 text-xs shadow-md"
        >
          <ReviewMenuItem
            disabled={empty}
            icon={<Check size={12} className="text-[var(--success)]" />}
            label="Accept all changes"
            description="Fold every insertion into the document and remove every deletion."
            onClick={() => {
              if (empty) return;
              props.onAcceptAll();
              setOpen(false);
            }}
            testId="review-accept-all"
          />
          <ReviewMenuItem
            disabled={empty}
            icon={<X size={12} className="text-[var(--error)]" />}
            label="Reject all changes"
            description="Drop every insertion and restore every deletion to the original text."
            onClick={() => {
              if (empty) return;
              props.onRejectAll();
              setOpen(false);
            }}
            testId="review-reject-all"
          />
        </div>
      )}
    </div>
  );
}

function ReviewMenuItem(props: {
  disabled: boolean;
  icon: ReactNode;
  label: string;
  description: string;
  onClick: () => void;
  testId: string;
}): ReactNode {
  return (
    <button
      type="button"
      role="menuitem"
      disabled={props.disabled}
      onClick={props.onClick}
      data-testid={props.testId}
      className={cn(
        "flex w-full items-start gap-2 rounded px-2 py-1.5 text-left",
        props.disabled ? "cursor-not-allowed opacity-50" : "hover:bg-hover"
      )}
    >
      <span className="mt-0.5 shrink-0">{props.icon}</span>
      <span className="min-w-0 flex-1">
        <span className="block font-medium text-foreground">{props.label}</span>
        <span className="block text-[11px] text-secondary">{props.description}</span>
      </span>
    </button>
  );
}

const EDIT_MODE_META: Record<
  EditModeValue,
  {
    label: string;
    description: string;
    Icon: typeof Pencil;
    pillClass: string;
    iconColorClass: string;
  }
> = {
  edit: {
    label: "Editing",
    description: "Type and delete directly. Changes apply immediately.",
    Icon: Pencil,
    pillClass: "border-divider bg-surface text-foreground hover:bg-hover",
    iconColorClass: "text-foreground",
  },
  suggest: {
    label: "Suggesting",
    description: "Every insert and delete is recorded as a tracked change you can accept or reject later.",
    Icon: PenLine,
    pillClass:
      "border-[var(--ai-violet)] bg-[var(--ai-violet-light)] text-[var(--ai-violet)] hover:brightness-95",
    iconColorClass: "text-[var(--ai-violet)]",
  },
  view: {
    label: "Viewing",
    description: "Read-only. Typing and edits are blocked.",
    Icon: Eye,
    pillClass: "border-divider bg-hover text-secondary hover:bg-divider",
    iconColorClass: "text-secondary",
  },
};

function ToolbarBtn(props: {
  label: string;
  onClick: () => void;
  active?: boolean;
  testId?: string;
  /**
   * Disabled trigger. Used both for state-conditional buttons (e.g.
   * "Merge cells" with no multi-cell selection) and for Phase 9c
   * Coming-soon placeholders that surface the planned ribbon shape
   * before the underlying backend lands. Renders the same control,
   * with `aria-disabled` and reduced opacity, so the layout doesn't
   * jump when the button gates on/off.
   */
  disabled?: boolean;
  children: ReactNode;
}): ReactNode {
  return (
    <button
      type="button"
      title={props.label}
      aria-label={props.label}
      aria-pressed={props.active ?? undefined}
      aria-disabled={props.disabled ?? undefined}
      disabled={props.disabled ?? undefined}
      data-testid={props.testId}
      onMouseDown={(e) => e.preventDefault()}
      onClick={props.onClick}
      className={cn(
        "inline-flex items-center rounded-md p-1.5 text-secondary hover:bg-hover hover:text-foreground",
        props.active && "bg-accent-light text-foreground",
        props.disabled && "cursor-not-allowed opacity-40 hover:bg-transparent hover:text-secondary"
      )}
    >
      {props.children}
    </button>
  );
}

function ParagraphStylePicker(props: {
  value: string;
  options: ReadonlyArray<ToolbarStyleOption>;
  onChange: (v: string) => void;
  disabled?: boolean;
}): ReactNode {
  // Ensure the active value is in the option list so the <select> doesn't
  // silently drop the displayed value to "Normal" when the doc carries a
  // style id we haven't surfaced yet.
  const hasActive = props.options.some((o) => o.value === props.value);
  const items = hasActive
    ? props.options
    : [{ value: props.value, label: props.value || "—" }, ...props.options];
  return (
    <label className="inline-flex items-center gap-1 text-xs text-secondary">
      <span className="sr-only">Paragraph style</span>
      <select
        title="Paragraph style"
        aria-label="Paragraph style"
        value={props.value}
        disabled={props.disabled}
        onChange={(e) => props.onChange(e.target.value)}
        className="h-7 max-w-40 rounded-md border border-divider bg-surface px-2 text-xs text-foreground hover:bg-hover focus:outline-none"
      >
        {items.map((s) => (
          <option key={s.value} value={s.value}>
            {s.label}
          </option>
        ))}
      </select>
    </label>
  );
}

/**
 * Spacing dropdown — line spacing (single / 1.15 / 1.5 / double) and
 * per-paragraph before/after spacing in points. Reads the resolved
 * effective spacing for display so a Heading style's inherited "1.5
 * line" surfaces here even when the paragraph has no direct
 * `<w:spacing>`.
 *
 * Word stores `<w:spacing w:line>` in twentieths of a line for
 * `lineRule="auto"` (so 240 = 1.0, 360 = 1.5) and in twips for
 * `exact` / `atLeast`. We mutate the `auto` family from the picker
 * because that covers the >95% case; the existing exact/atLeast
 * value is shown but the picker resets `lineRule` to `auto` on
 * change. Before/after stay as twips for OOXML parity (1pt = 20
 * twips); inputs accept points and convert.
 */
function SpacingMenu(props: {
  spacing: ResolvedSpacingDisplay | null;
  onApply: (patch: {
    line?: number | null;
    lineRule?: "auto" | "exact" | "atLeast" | null;
    before?: number | null;
    after?: number | null;
  }) => void;
  disabled?: boolean;
}): ReactNode {
  const [open, setOpen] = useState(false);
  const triggerRef = useRef<HTMLButtonElement | null>(null);

  const lineDisplay = formatLineSpacing(props.spacing);
  const beforePts = toPoints(props.spacing?.before);
  const afterPts = toPoints(props.spacing?.after);

  return (
    <span className="inline-flex">
      <button
        ref={triggerRef}
        type="button"
        title="Line and paragraph spacing"
        aria-label="Line and paragraph spacing"
        aria-haspopup="menu"
        aria-expanded={open}
        disabled={props.disabled}
        onClick={() => setOpen((v) => !v)}
        className="inline-flex items-center gap-1 rounded-md px-2 py-1.5 text-xs text-secondary hover:bg-hover hover:text-foreground disabled:opacity-50"
      >
        <span className="tabular-nums">{lineDisplay}</span>
        <ChevronDown size={10} />
      </button>
      <ToolbarMenu
        open={open}
        onClose={() => setOpen(false)}
        triggerRef={triggerRef}
        align="right"
        role="menu"
        className="w-56 rounded-md border border-divider bg-surface p-2 text-xs shadow-md"
      >
        <div className="mb-2 font-medium text-foreground">Line spacing</div>
        {LINE_SPACING_PRESETS.map((preset) => (
          <button
            key={preset.label}
            type="button"
            role="menuitem"
            onClick={() => {
              props.onApply({ line: preset.line, lineRule: "auto" });
              setOpen(false);
            }}
            className={cn(
              "flex w-full items-center justify-between rounded px-2 py-1 text-left hover:bg-hover",
              isActiveLine(props.spacing, preset.line) && "bg-accent-light text-foreground"
            )}
          >
            <span>{preset.label}</span>
            <span className="text-secondary tabular-nums">{(preset.line / 240).toFixed(2)}×</span>
          </button>
        ))}
        <div className="mt-3 border-t border-divider pt-2">
          <div className="mb-1 font-medium text-foreground">Paragraph spacing (pt)</div>
          <label className="mb-1 flex items-center justify-between gap-2">
            <span className="text-secondary">Before</span>
            <input
              type="number"
              min={0}
              step={1}
              defaultValue={beforePts ?? ""}
              onBlur={(e) => {
                const v = e.target.value === "" ? null : Number(e.target.value);
                if (v === null) props.onApply({ before: null });
                else if (Number.isFinite(v) && v >= 0) props.onApply({ before: Math.round(v * 20) });
              }}
              className="h-6 w-16 rounded border border-divider bg-surface px-1 text-right tabular-nums focus:outline-none"
            />
          </label>
          <label className="flex items-center justify-between gap-2">
            <span className="text-secondary">After</span>
            <input
              type="number"
              min={0}
              step={1}
              defaultValue={afterPts ?? ""}
              onBlur={(e) => {
                const v = e.target.value === "" ? null : Number(e.target.value);
                if (v === null) props.onApply({ after: null });
                else if (Number.isFinite(v) && v >= 0) props.onApply({ after: Math.round(v * 20) });
              }}
              className="h-6 w-16 rounded border border-divider bg-surface px-1 text-right tabular-nums focus:outline-none"
            />
          </label>
        </div>
      </ToolbarMenu>
    </span>
  );
}

const LINE_SPACING_PRESETS: ReadonlyArray<{ label: string; line: number }> = [
  { label: "Single", line: 240 },
  { label: "1.15", line: 276 },
  { label: "1.5", line: 360 },
  { label: "Double", line: 480 },
];

function formatLineSpacing(s: ResolvedSpacingDisplay | null): string {
  if (!s || s.line === undefined) return "Spacing";
  if (s.lineRule && s.lineRule !== "auto") return `${(s.line / 20).toFixed(1)}pt`;
  return `${(s.line / 240).toFixed(2)}×`;
}

function isActiveLine(s: ResolvedSpacingDisplay | null, line: number): boolean {
  if (!s) return false;
  if (s.lineRule && s.lineRule !== "auto") return false;
  return s.line === line;
}

function toPoints(twips: number | undefined): number | undefined {
  if (twips === undefined) return undefined;
  return Math.round(twips / 20);
}

function twipsToInches(twips: number): string {
  return (twips / 1440).toFixed(2);
}

function ActiveCellPicker(props: {
  row: number;
  column: number;
  onChange: (cell: { row: number; column: number }) => void;
}): ReactNode {
  return (
    <div className="flex items-center gap-1 px-1 text-xs text-secondary">
      <span>Zelle</span>
      <label className="flex items-center gap-1">
        <span>Z</span>
        <input
          type="number"
          min={0}
          value={props.row}
          onChange={(e) => {
            const v = Math.max(0, Number(e.target.value) || 0);
            props.onChange({ row: v, column: props.column });
          }}
          className="w-12 rounded border border-borderSubtle bg-background px-1 text-foreground"
          data-testid="docx-table-active-row"
        />
      </label>
      <label className="flex items-center gap-1">
        <span>S</span>
        <input
          type="number"
          min={0}
          value={props.column}
          onChange={(e) => {
            const v = Math.max(0, Number(e.target.value) || 0);
            props.onChange({ row: props.row, column: v });
          }}
          className="w-12 rounded border border-borderSubtle bg-background px-1 text-foreground"
          data-testid="docx-table-active-col"
        />
      </label>
    </div>
  );
}

function CellSizeControls(props: {
  tableId: string;
  cell: { row: number; column: number };
  onSetRowHeight: (
    tableId: string,
    row: number,
    heightTwips: number | null,
    rule?: "auto" | "exact" | "atLeast"
  ) => void;
  onSetColumnWidth: (tableId: string, column: number, widthTwips: number) => void;
}): ReactNode {
  const [rowCm, setRowCm] = useState<string>("");
  const [colCm, setColCm] = useState<string>("");
  const cmToTwips = (cm: number): number => Math.round((cm / 2.54) * 1440);
  return (
    <div className="flex items-center gap-1 px-1 text-xs text-secondary">
      <label className="flex items-center gap-1">
        <span>Höhe (cm)</span>
        <input
          type="number"
          step="0.1"
          min={0}
          value={rowCm}
          onChange={(e) => setRowCm(e.target.value)}
          className="w-16 rounded border border-borderSubtle bg-background px-1 text-foreground"
          data-testid="docx-row-height-input"
        />
      </label>
      <button
        type="button"
        onClick={() => {
          const n = Number(rowCm);
          if (!Number.isFinite(n) || n <= 0) {
            props.onSetRowHeight(props.tableId, props.cell.row, null);
          } else {
            props.onSetRowHeight(props.tableId, props.cell.row, cmToTwips(n), "atLeast");
          }
        }}
        className="rounded px-1 text-xs hover:bg-hover"
        data-testid="docx-row-height-apply"
      >
        Setzen
      </button>
      <label className="flex items-center gap-1">
        <span>Breite (cm)</span>
        <input
          type="number"
          step="0.1"
          min={0}
          value={colCm}
          onChange={(e) => setColCm(e.target.value)}
          className="w-16 rounded border border-borderSubtle bg-background px-1 text-foreground"
          data-testid="docx-col-width-input"
        />
      </label>
      <button
        type="button"
        onClick={() => {
          const n = Number(colCm);
          if (!Number.isFinite(n) || n <= 0) return;
          props.onSetColumnWidth(props.tableId, props.cell.column, cmToTwips(n));
        }}
        className="rounded px-1 text-xs hover:bg-hover"
        data-testid="docx-col-width-apply"
      >
        Setzen
      </button>
    </div>
  );
}

const CELL_SHADING_PRESETS: ReadonlyArray<{ label: string; fill: string | null }> = [
  { label: "Keine", fill: null },
  { label: "Hellgelb", fill: "FFF2CC" },
  { label: "Hellblau", fill: "DEEBF7" },
  { label: "Hellgrün", fill: "E2EFDA" },
  { label: "Hellrot", fill: "FCE4D6" },
  { label: "Hellgrau", fill: "EDEDED" },
  { label: "Mittelgrau", fill: "BFBFBF" },
];

function CellShadingPicker(props: {
  tableId: string;
  cell: { row: number; column: number };
  onSetCellShading: (tableId: string, row: number, column: number, fill: string | null) => void;
}): ReactNode {
  const [open, setOpen] = useState(false);
  const triggerRef = useRef<HTMLButtonElement | null>(null);
  return (
    <>
      <button
        ref={triggerRef}
        type="button"
        title="Zellenschattierung"
        aria-label="Zellenschattierung"
        aria-haspopup="dialog"
        aria-expanded={open}
        onClick={() => setOpen((v) => !v)}
        className={cn(
          "inline-flex items-center gap-0.5 rounded-md p-1.5 text-secondary hover:bg-hover hover:text-foreground",
          open && "bg-hover text-foreground"
        )}
        data-testid="docx-cell-shading-trigger"
      >
        <PaintBucket size={14} />
        <ChevronDown size={10} />
      </button>
      <ToolbarMenu
        open={open}
        onClose={() => setOpen(false)}
        triggerRef={triggerRef}
        role="dialog"
        className="w-max rounded-md border border-divider bg-surface p-1 shadow-md"
      >
        <div className="grid w-44 grid-cols-2 gap-1">
          {CELL_SHADING_PRESETS.map((p) => (
            <button
              key={p.label}
              type="button"
              onClick={() => {
                props.onSetCellShading(props.tableId, props.cell.row, props.cell.column, p.fill);
                setOpen(false);
              }}
              className="flex items-center gap-2 rounded px-2 py-1 text-left text-xs hover:bg-hover"
            >
              <span
                className="inline-block h-3 w-3 rounded border border-divider"
                style={{ background: p.fill ? `#${p.fill}` : "transparent" }}
              />
              <span>{p.label}</span>
            </button>
          ))}
        </div>
      </ToolbarMenu>
    </>
  );
}

function MergeCellsControl(props: {
  tableId: string;
  row: number;
  fromColumn: number;
  onMerge: (tableId: string, row: number, fromColumn: number, toColumn: number) => void;
}): ReactNode {
  const [toCol, setToCol] = useState<string>("");
  return (
    <div className="flex items-center gap-1 px-1 text-xs text-secondary">
      <label className="flex items-center gap-1">
        <span>bis Spalte</span>
        <input
          type="number"
          min={props.fromColumn + 1}
          value={toCol}
          onChange={(e) => setToCol(e.target.value)}
          className="w-12 rounded border border-borderSubtle bg-background px-1 text-foreground"
          data-testid="docx-merge-to-col"
        />
      </label>
      <button
        type="button"
        onClick={() => {
          const n = Number(toCol);
          if (!Number.isFinite(n) || n <= props.fromColumn) return;
          props.onMerge(props.tableId, props.row, props.fromColumn, n);
        }}
        className="inline-flex items-center gap-1 rounded px-1 text-xs hover:bg-hover"
        data-testid="docx-merge-horizontal-apply"
      >
        <Combine size={14} />
        <span>Verbinden</span>
      </button>
    </div>
  );
}
