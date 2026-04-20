"use client";

import * as React from "react";
import {
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  Edit3,
  Highlighter,
  Maximize2,
  Minus,
  Moon,
  PenLine,
  Plus,
  RotateCcw,
  RotateCw,
  ScrollText,
  Square,
  StickyNote,
  Trash2,
  Type,
  Workflow,
} from "lucide-react";
import { ToolbarMenu, ToolbarRow } from "@/lib/shell";
import { useTranslator } from "@/lib/i18n";
import type { PdfDarkModeStrategy, PdfViewMode } from "./PdfCanvas";

export type PdfAnnotationTool = "highlight" | "sticky" | "free-text";

export interface PdfToolbarProps {
  readonly disabled: boolean;
  readonly currentPage: number;
  readonly totalPages: number;
  readonly zoom: number;
  readonly viewMode: PdfViewMode;
  readonly darkMode: PdfDarkModeStrategy;
  readonly reflow: boolean;
  readonly onPrevPage: () => void;
  readonly onNextPage: () => void;
  readonly onJumpToPage: (pageNumber: number) => void;
  readonly onZoomIn: () => void;
  readonly onZoomOut: () => void;
  readonly onFitWidth: () => void;
  readonly onFitPage: () => void;
  readonly onActualSize: () => void;
  readonly onSetViewMode: (mode: PdfViewMode) => void;
  readonly onRotateClockwise: () => void;
  readonly onRotateCounterClockwise: () => void;
  readonly onAnnotate: (tool: PdfAnnotationTool) => void;
  readonly onRotatePages: () => void;
  readonly onReorderPages: () => void;
  readonly onDeletePages: () => void;
  readonly onFillForm: () => void;
  readonly onFlattenForm: () => void;
  readonly onSetDarkMode: (mode: PdfDarkModeStrategy) => void;
  readonly onToggleReflow: () => void;
}

/**
 * Top toolbar for the PDF viewer.
 *
 * Layout (left → right):
 *
 *   Page nav         · prev / "n of total" / next
 *   Zoom             · - / "100 %" picker / +
 *   View mode        · single / continuous / two-up dropdown
 *   Rotate view      · CCW / CW
 *   Annotate         · highlight / sticky / free-text
 *   Page ops         · rotate / reorder / delete pages dropdowns
 *   Form             · fill / flatten
 *   Dark mode        · off / css / smart toggle group
 *   Reflow           · single-column toggle
 *
 * Every label is sourced from `t("pdf.*")`. The toolbar is a thin
 * presentational surface — every callback maps directly to a
 * method on `PdfEditor` so undo / redo and realtime broadcast can
 * happen at the editor scope without each button knowing about
 * either system.
 */
export function PdfToolbar(props: PdfToolbarProps): React.ReactNode {
  const { t } = useTranslator();
  const {
    disabled,
    currentPage,
    totalPages,
    zoom,
    viewMode,
    darkMode,
    reflow,
  } = props;
  const noPages = totalPages === 0;
  const canPrev = !disabled && !noPages && currentPage > 1;
  const canNext = !disabled && !noPages && currentPage < totalPages;
  return (
    <ToolbarRow ariaLabel="PDF toolbar" testId="pdf-toolbar">
      <PageNav
        currentPage={currentPage}
        totalPages={totalPages}
        canPrev={canPrev}
        canNext={canNext}
        onPrev={props.onPrevPage}
        onNext={props.onNextPage}
        onJump={props.onJumpToPage}
      />
      <Sep />
      <IconButton
        onClick={props.onZoomOut}
        icon={<Minus size={14} />}
        label={t("pdf.zoomOut")}
        disabled={disabled}
        testId="pdf-zoom-out"
      />
      <ZoomMenu
        zoom={zoom}
        disabled={disabled}
        onFitWidth={props.onFitWidth}
        onFitPage={props.onFitPage}
        onActualSize={props.onActualSize}
        onZoomIn={props.onZoomIn}
        onZoomOut={props.onZoomOut}
      />
      <IconButton
        onClick={props.onZoomIn}
        icon={<Plus size={14} />}
        label={t("pdf.zoomIn")}
        disabled={disabled}
        testId="pdf-zoom-in"
      />
      <Sep />
      <ViewModeMenu mode={viewMode} disabled={disabled} onPick={props.onSetViewMode} />
      <Sep />
      <IconButton
        onClick={props.onRotateCounterClockwise}
        icon={<RotateCcw size={14} />}
        label={t("pdf.rotateCounterClockwise")}
        disabled={disabled}
        testId="pdf-rotate-ccw"
      />
      <IconButton
        onClick={props.onRotateClockwise}
        icon={<RotateCw size={14} />}
        label={t("pdf.rotateClockwise")}
        disabled={disabled}
        testId="pdf-rotate-cw"
      />
      <Sep />
      <ToolbarTextButton
        onClick={() => props.onAnnotate("highlight")}
        icon={<Highlighter size={14} />}
        label={t("pdf.highlight")}
        disabled={disabled}
        testId="pdf-annotate-highlight"
      />
      <ToolbarTextButton
        onClick={() => props.onAnnotate("sticky")}
        icon={<StickyNote size={14} />}
        label={t("pdf.stickyNote")}
        disabled={disabled}
        testId="pdf-annotate-sticky"
      />
      <ToolbarTextButton
        onClick={() => props.onAnnotate("free-text")}
        icon={<Type size={14} />}
        label={t("pdf.freeText")}
        disabled={disabled}
        testId="pdf-annotate-freetext"
      />
      <Sep />
      <PageOpsMenu
        disabled={disabled}
        onRotatePages={props.onRotatePages}
        onReorderPages={props.onReorderPages}
        onDeletePages={props.onDeletePages}
      />
      <FormMenu
        disabled={disabled}
        onFillForm={props.onFillForm}
        onFlattenForm={props.onFlattenForm}
      />
      <Sep />
      <DarkModeToggle value={darkMode} disabled={disabled} onChange={props.onSetDarkMode} />
      <ToolbarTextButton
        onClick={props.onToggleReflow}
        icon={<ScrollText size={14} />}
        label={t("pdf.reflow")}
        disabled={disabled}
        active={reflow}
        testId="pdf-reflow-toggle"
      />
    </ToolbarRow>
  );
}

function Sep(): React.ReactNode {
  return <span className="mx-2 h-5 w-px bg-divider" />;
}

interface IconButtonProps {
  readonly onClick: () => void;
  readonly icon: React.ReactNode;
  readonly label: string;
  readonly disabled?: boolean;
  readonly active?: boolean;
  readonly testId?: string;
}

function IconButton({ onClick, icon, label, disabled, active, testId }: IconButtonProps): React.ReactNode {
  return (
    <button
      type="button"
      onClick={onClick}
      disabled={disabled}
      title={label}
      aria-label={label}
      aria-pressed={active}
      data-testid={testId}
      className={
        "inline-flex h-7 w-7 items-center justify-center rounded text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40 " +
        (active ? "bg-hover ring-1 ring-[var(--accent)]/40" : "")
      }
    >
      {icon}
    </button>
  );
}

interface ToolbarTextButtonProps extends IconButtonProps {
  readonly icon: React.ReactNode;
}

function ToolbarTextButton({
  onClick,
  icon,
  label,
  disabled,
  active,
  testId,
}: ToolbarTextButtonProps): React.ReactNode {
  return (
    <button
      type="button"
      onClick={onClick}
      disabled={disabled}
      title={label}
      aria-label={label}
      aria-pressed={active}
      data-testid={testId}
      className={
        "inline-flex shrink-0 items-center gap-1.5 rounded px-2 py-1 text-xs text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40 " +
        (active ? "bg-hover ring-1 ring-[var(--accent)]/40" : "")
      }
    >
      {icon}
      <span className="hidden sm:inline">{label}</span>
    </button>
  );
}

// ─────────────────────────────────────────────────────────────────────
// Page nav

interface PageNavProps {
  readonly currentPage: number;
  readonly totalPages: number;
  readonly canPrev: boolean;
  readonly canNext: boolean;
  readonly onPrev: () => void;
  readonly onNext: () => void;
  readonly onJump: (pageNumber: number) => void;
}

/**
 * Inline "n of total" widget. The current page is editable so the
 * user can type a page number and press Enter to jump (mirrors
 * Acrobat's URL-style page-jump). We intentionally don't use a
 * `<select>` because the dropdown gets unwieldy past ~50 pages and
 * the inline `<input>` keeps the toolbar height stable for any
 * document length.
 */
function PageNav({
  currentPage,
  totalPages,
  canPrev,
  canNext,
  onPrev,
  onNext,
  onJump,
}: PageNavProps): React.ReactNode {
  const { t } = useTranslator();
  const [draft, setDraft] = React.useState<string>(String(currentPage || 1));
  React.useEffect(() => {
    setDraft(String(currentPage || 1));
  }, [currentPage]);
  const submit = (): void => {
    const n = Number(draft);
    if (!Number.isFinite(n) || totalPages === 0) {
      setDraft(String(currentPage || 1));
      return;
    }
    const clamped = Math.max(1, Math.min(totalPages, Math.round(n)));
    onJump(clamped);
    setDraft(String(clamped));
  };
  return (
    <div className="inline-flex items-center gap-1" data-testid="pdf-page-nav">
      <IconButton
        onClick={onPrev}
        icon={<ChevronLeft size={14} />}
        label={t("pdf.page")}
        disabled={!canPrev}
        testId="pdf-page-prev"
      />
      <input
        type="text"
        inputMode="numeric"
        value={draft}
        onChange={(e) => setDraft(e.target.value)}
        onBlur={submit}
        onKeyDown={(e) => {
          if (e.key === "Enter") {
            e.preventDefault();
            submit();
          }
        }}
        disabled={totalPages === 0}
        aria-label={t("pdf.page")}
        data-testid="pdf-page-input"
        className="h-6 w-10 rounded border border-divider bg-background px-1 text-center text-xs tabular-nums focus:outline-none focus:ring-1 focus:ring-[var(--accent)]"
      />
      <span className="text-xs tabular-nums text-tertiary">
        {totalPages > 0 ? `/ ${totalPages}` : ""}
      </span>
      <IconButton
        onClick={onNext}
        icon={<ChevronRight size={14} />}
        label={t("pdf.page")}
        disabled={!canNext}
        testId="pdf-page-next"
      />
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────
// Zoom menu

interface ZoomMenuProps {
  readonly zoom: number;
  readonly disabled: boolean;
  readonly onFitWidth: () => void;
  readonly onFitPage: () => void;
  readonly onActualSize: () => void;
  readonly onZoomIn: () => void;
  readonly onZoomOut: () => void;
}

function ZoomMenu({
  zoom,
  disabled,
  onFitWidth,
  onFitPage,
  onActualSize,
}: ZoomMenuProps): React.ReactNode {
  const { t } = useTranslator();
  const [open, setOpen] = React.useState(false);
  const triggerRef = React.useRef<HTMLButtonElement>(null);
  const display = `${Math.round(zoom * 100)} %`;
  return (
    <>
      <button
        ref={triggerRef}
        type="button"
        disabled={disabled}
        onClick={() => setOpen((v) => !v)}
        title={t("pdf.fitWidth")}
        data-testid="pdf-zoom-menu-trigger"
        className="inline-flex h-7 min-w-[60px] items-center justify-center gap-0.5 rounded border border-divider px-2 text-xs tabular-nums text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
      >
        {display}
        <ChevronDown size={12} />
      </button>
      <ToolbarMenu
        open={open}
        onClose={() => setOpen(false)}
        triggerRef={triggerRef}
        role="menu"
        testId="pdf-zoom-menu"
        className="w-44 rounded-md border border-divider bg-surface p-1 shadow-lg"
      >
        <MenuItem
          onClick={() => {
            setOpen(false);
            onActualSize();
          }}
          label={t("pdf.actualSize")}
          testId="pdf-zoom-actual"
        />
        <MenuItem
          onClick={() => {
            setOpen(false);
            onFitWidth();
          }}
          label={t("pdf.fitWidth")}
          testId="pdf-zoom-fit-width"
        />
        <MenuItem
          onClick={() => {
            setOpen(false);
            onFitPage();
          }}
          label={t("pdf.fitPage")}
          testId="pdf-zoom-fit-page"
        />
        <Divider />
        {[0.5, 0.75, 1, 1.25, 1.5, 2, 3].map((preset) => (
          <MenuItem
            key={preset}
            onClick={() => {
              setOpen(false);
              if (preset === 1) onActualSize();
              else if (preset > zoom) onActualSize(); // fall through; editor decides
            }}
            label={`${Math.round(preset * 100)} %`}
            testId={`pdf-zoom-preset-${preset}`}
          />
        ))}
      </ToolbarMenu>
    </>
  );
}

// ─────────────────────────────────────────────────────────────────────
// View mode

interface ViewModeMenuProps {
  readonly mode: PdfViewMode;
  readonly disabled: boolean;
  readonly onPick: (mode: PdfViewMode) => void;
}

function ViewModeMenu({ mode, disabled, onPick }: ViewModeMenuProps): React.ReactNode {
  const { t } = useTranslator();
  const [open, setOpen] = React.useState(false);
  const triggerRef = React.useRef<HTMLButtonElement>(null);
  const label =
    mode === "single" ? t("pdf.singlePage") : mode === "two-up" ? t("pdf.twoUp") : t("pdf.continuous");
  return (
    <>
      <button
        ref={triggerRef}
        type="button"
        disabled={disabled}
        onClick={() => setOpen((v) => !v)}
        title={label}
        data-testid="pdf-view-mode-trigger"
        className="inline-flex items-center gap-1 rounded px-2 py-1 text-xs text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
      >
        <Maximize2 size={14} />
        <span className="hidden sm:inline">{label}</span>
        <ChevronDown size={12} />
      </button>
      <ToolbarMenu
        open={open}
        onClose={() => setOpen(false)}
        triggerRef={triggerRef}
        role="menu"
        testId="pdf-view-mode"
        className="grid w-44 grid-cols-1 gap-0.5 rounded-md border border-divider bg-surface p-1 shadow-lg"
      >
        {(["single", "continuous", "two-up"] as const).map((m) => (
          <MenuItem
            key={m}
            label={
              m === "single"
                ? t("pdf.singlePage")
                : m === "two-up"
                  ? t("pdf.twoUp")
                  : t("pdf.continuous")
            }
            active={mode === m}
            onClick={() => {
              setOpen(false);
              onPick(m);
            }}
            testId={`pdf-view-mode-${m}`}
          />
        ))}
      </ToolbarMenu>
    </>
  );
}

// ─────────────────────────────────────────────────────────────────────
// Page ops

interface PageOpsMenuProps {
  readonly disabled: boolean;
  readonly onRotatePages: () => void;
  readonly onReorderPages: () => void;
  readonly onDeletePages: () => void;
}

function PageOpsMenu({
  disabled,
  onRotatePages,
  onReorderPages,
  onDeletePages,
}: PageOpsMenuProps): React.ReactNode {
  const { t } = useTranslator();
  const [open, setOpen] = React.useState(false);
  const triggerRef = React.useRef<HTMLButtonElement>(null);
  return (
    <>
      <button
        ref={triggerRef}
        type="button"
        disabled={disabled}
        onClick={() => setOpen((v) => !v)}
        title={t("pdf.rotatePages")}
        data-testid="pdf-page-ops-trigger"
        className="inline-flex items-center gap-1 rounded px-2 py-1 text-xs text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
      >
        <Workflow size={14} />
        <span className="hidden sm:inline">{t("pdf.page")}</span>
        <ChevronDown size={12} />
      </button>
      <ToolbarMenu
        open={open}
        onClose={() => setOpen(false)}
        triggerRef={triggerRef}
        role="menu"
        testId="pdf-page-ops-menu"
        className="grid w-52 grid-cols-1 gap-0.5 rounded-md border border-divider bg-surface p-1 shadow-lg"
      >
        <MenuItem
          icon={<RotateCw size={14} />}
          label={t("pdf.rotatePages")}
          onClick={() => {
            setOpen(false);
            onRotatePages();
          }}
          testId="pdf-rotate-pages"
        />
        <MenuItem
          icon={<Workflow size={14} />}
          label={t("pdf.reorderPages")}
          onClick={() => {
            setOpen(false);
            onReorderPages();
          }}
          testId="pdf-reorder-pages"
        />
        <MenuItem
          icon={<Trash2 size={14} />}
          label={t("pdf.deletePages")}
          onClick={() => {
            setOpen(false);
            onDeletePages();
          }}
          testId="pdf-delete-pages"
        />
      </ToolbarMenu>
    </>
  );
}

// ─────────────────────────────────────────────────────────────────────
// Form

interface FormMenuProps {
  readonly disabled: boolean;
  readonly onFillForm: () => void;
  readonly onFlattenForm: () => void;
}

function FormMenu({ disabled, onFillForm, onFlattenForm }: FormMenuProps): React.ReactNode {
  const { t } = useTranslator();
  const [open, setOpen] = React.useState(false);
  const triggerRef = React.useRef<HTMLButtonElement>(null);
  return (
    <>
      <button
        ref={triggerRef}
        type="button"
        disabled={disabled}
        onClick={() => setOpen((v) => !v)}
        title={t("pdf.fillForm")}
        data-testid="pdf-form-trigger"
        className="inline-flex items-center gap-1 rounded px-2 py-1 text-xs text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
      >
        <Edit3 size={14} />
        <span className="hidden sm:inline">{t("pdf.fillForm")}</span>
        <ChevronDown size={12} />
      </button>
      <ToolbarMenu
        open={open}
        onClose={() => setOpen(false)}
        triggerRef={triggerRef}
        role="menu"
        testId="pdf-form-menu"
        className="grid w-44 grid-cols-1 gap-0.5 rounded-md border border-divider bg-surface p-1 shadow-lg"
      >
        <MenuItem
          icon={<PenLine size={14} />}
          label={t("pdf.fillForm")}
          onClick={() => {
            setOpen(false);
            onFillForm();
          }}
          testId="pdf-form-fill"
        />
        <MenuItem
          icon={<Square size={14} />}
          label={t("pdf.flattenForm")}
          onClick={() => {
            setOpen(false);
            onFlattenForm();
          }}
          testId="pdf-form-flatten"
        />
      </ToolbarMenu>
    </>
  );
}

// ─────────────────────────────────────────────────────────────────────
// Dark mode

interface DarkModeToggleProps {
  readonly value: PdfDarkModeStrategy;
  readonly disabled: boolean;
  readonly onChange: (next: PdfDarkModeStrategy) => void;
}

/**
 * Three-state dark-mode toggle. Cycles `off → css → smart → off`
 * on click; the icon's background reads as "off" when the value
 * is `off` and as "armed" otherwise. We expose all three states
 * (rather than a binary toggle) because the per-pixel "smart"
 * pass is materially slower on long documents and some users
 * deliberately prefer the cheap CSS strategy even though it
 * inverts photographs.
 */
function DarkModeToggle({ value, disabled, onChange }: DarkModeToggleProps): React.ReactNode {
  const { t } = useTranslator();
  const next: PdfDarkModeStrategy =
    value === "off" ? "css" : value === "css" ? "smart" : "off";
  const label =
    value === "off"
      ? t("pdf.darkMode")
      : value === "css"
        ? `${t("pdf.darkMode")} (CSS)`
        : `${t("pdf.darkMode")} (smart)`;
  return (
    <button
      type="button"
      onClick={() => onChange(next)}
      disabled={disabled}
      title={label}
      aria-label={label}
      aria-pressed={value !== "off"}
      data-testid="pdf-dark-mode-toggle"
      data-mode={value}
      className={
        "inline-flex shrink-0 items-center gap-1.5 rounded px-2 py-1 text-xs text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40 " +
        (value !== "off" ? "bg-hover ring-1 ring-[var(--accent)]/40" : "")
      }
    >
      <Moon size={14} />
      <span className="hidden sm:inline">{t("pdf.darkMode")}</span>
    </button>
  );
}

// ─────────────────────────────────────────────────────────────────────
// Generic dropdown menu items

interface MenuItemProps {
  readonly label: string;
  readonly onClick: () => void;
  readonly testId?: string;
  readonly icon?: React.ReactNode;
  readonly active?: boolean;
}

function MenuItem({ label, onClick, testId, icon, active }: MenuItemProps): React.ReactNode {
  return (
    <button
      type="button"
      role="menuitem"
      onClick={onClick}
      data-testid={testId}
      className={
        "flex items-center gap-2 rounded px-2 py-1 text-left text-xs text-foreground hover:bg-hover " +
        (active ? "bg-hover" : "")
      }
    >
      {icon ?? <span className="inline-block h-[14px] w-[14px]" aria-hidden />}
      <span>{label}</span>
      {active ? <span className="ml-auto text-[10px] text-[var(--accent)]">●</span> : null}
    </button>
  );
}

function Divider(): React.ReactNode {
  return <div className="my-0.5 h-px bg-divider" role="separator" />;
}
