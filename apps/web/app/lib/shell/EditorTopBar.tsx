"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import Link from "next/link";
import {
  ArrowLeft,
  ChevronDown,
  Download,
  FileCode,
  FileImage,
  FileSpreadsheet,
  FileText,
  FileType2,
  FolderOpen,
  Keyboard,
  MessageSquare,
  Presentation,
  Redo2,
  Save,
  Search,
  Sliders,
  Undo2,
} from "lucide-react";
import { ThemeToggle, cn } from "@officeai/ui";
import { ExportDialog } from "./ExportDialog";
import { InlineSpinner } from "./InlineSpinner";
import type {
  ExportFormat,
  ExportFormatGroup,
  ExportFormatIcon,
  ProductAdapter,
  SaveState,
} from "./types";

export interface EditorTopBarProps {
  readonly adapter: ProductAdapter;
  readonly onOpenCommandPalette: () => void;
  readonly onOpenFindReplace: () => void;
  readonly onToggleRail: () => void;
  readonly railOpen: boolean;
  /** Editing the filename triggers this. */
  readonly onRenameFilename?: (next: string) => void;
}

/**
 * The single top bar used by all three editors.
 *
 * Layout:
 *   [Back]  ·  [Filename]  [save-state]
 *                                                   [Open] [Save] [Export ▾] | [Undo] [Redo] | [Search] [Pal] [Cmts] | [Theme] [Shortcuts]
 *
 * Visual rules:
 *  - 44 px tall, no second header row above or below.
 *  - Icon-only by default, tooltip surfaces the label.
 *  - The shell never renders the AI panel button (no AI in the
 *    product surface). The `Comments` icon is a quick toggle for the
 *    right rail.
 */
export function EditorTopBar({
  adapter,
  onOpenCommandPalette,
  onOpenFindReplace,
  onToggleRail,
  railOpen,
  onRenameFilename,
}: EditorTopBarProps): React.ReactNode {
  return (
    <header className="flex h-11 items-center gap-2 border-b border-divider bg-background px-3" role="banner">
      <Link
        href="/"
        className="inline-flex h-8 items-center gap-1 rounded-md px-1.5 text-sm text-secondary hover:bg-hover hover:text-foreground"
        aria-label="Back to home"
        title="Back to home"
      >
        <ArrowLeft size={14} />
      </Link>

      <span className="text-tertiary" aria-hidden>
        ·
      </span>

      <FilenameField value={adapter.filename} onCommit={onRenameFilename} />

      <SaveStatePill state={adapter.saveState} />

      <div className="flex-1" />

      {/* Primary file ops */}
      <ToolbarIcon
        label="Open"
        shortcut="Cmd+O"
        onClick={adapter.onOpenFile}
        disabled={!adapter.canOpen}
        testId="shell-open"
      >
        <FolderOpen size={15} />
      </ToolbarIcon>
      <ToolbarIcon
        label="Save"
        shortcut="Cmd+S"
        onClick={() => void adapter.onSave()}
        disabled={!adapter.canSave}
        testId="shell-save"
      >
        <Save size={15} />
      </ToolbarIcon>
      <ExportMenu adapter={adapter} />

      <Sep />

      <ToolbarIcon
        label="Undo"
        shortcut="Cmd+Z"
        onClick={adapter.onUndo}
        disabled={!adapter.canUndo}
        testId="shell-undo"
      >
        <Undo2 size={15} />
      </ToolbarIcon>
      <ToolbarIcon
        label="Redo"
        shortcut="Cmd+Shift+Z"
        onClick={adapter.onRedo}
        disabled={!adapter.canRedo}
        testId="shell-redo"
      >
        <Redo2 size={15} />
      </ToolbarIcon>

      <Sep />

      {adapter.findAdapter ? (
        <ToolbarIcon label="Find" shortcut="Cmd+F" onClick={onOpenFindReplace} testId="shell-find">
          <Search size={15} />
        </ToolbarIcon>
      ) : null}
      <ToolbarIcon
        label="Command palette"
        shortcut="Cmd+K"
        onClick={onOpenCommandPalette}
        testId="shell-palette"
      >
        <PaletteIcon />
      </ToolbarIcon>
      <ToolbarIcon
        label={railOpen ? "Hide comments" : "Show comments"}
        shortcut="Cmd+Alt+M"
        onClick={onToggleRail}
        testId="shell-comments-toggle"
        active={railOpen}
        badge={adapter.comments?.openCount}
      >
        <MessageSquare size={15} />
      </ToolbarIcon>

      <Sep />

      <ThemeToggle />
      <ToolbarIcon
        label="Keyboard shortcuts"
        shortcut="Cmd+/"
        onClick={adapter.onOpenShortcuts}
        testId="shell-shortcuts"
      >
        <Keyboard size={15} />
      </ToolbarIcon>
    </header>
  );
}

function FilenameField({
  value,
  onCommit,
}: {
  readonly value: string;
  readonly onCommit?: (next: string) => void;
}): React.ReactNode {
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(value);
  const inputRef = useRef<HTMLInputElement | null>(null);

  useEffect(() => {
    if (!editing) setDraft(value);
  }, [editing, value]);

  useEffect(() => {
    if (editing && inputRef.current) {
      inputRef.current.select();
    }
  }, [editing]);

  if (!editing || !onCommit) {
    return (
      <button
        type="button"
        className="max-w-[260px] truncate rounded-sm px-1 text-sm font-medium text-foreground hover:bg-hover"
        title={onCommit ? "Click to rename" : value}
        onClick={() => onCommit && setEditing(true)}
        data-testid="shell-filename"
      >
        {value || "Untitled"}
      </button>
    );
  }

  const commit = () => {
    const next = draft.trim();
    if (next.length === 0) {
      setEditing(false);
      setDraft(value);
      return;
    }
    onCommit(next);
    setEditing(false);
  };

  return (
    <input
      ref={inputRef}
      className="h-7 max-w-[260px] rounded-sm border border-divider bg-background px-1 text-sm font-medium text-foreground focus:border-[var(--accent)] focus:outline-none"
      value={draft}
      onChange={(e) => setDraft(e.target.value)}
      onBlur={commit}
      onKeyDown={(e) => {
        if (e.key === "Enter") {
          e.preventDefault();
          commit();
        } else if (e.key === "Escape") {
          e.preventDefault();
          setEditing(false);
          setDraft(value);
        }
      }}
      data-testid="shell-filename-input"
    />
  );
}

function SaveStatePill({ state }: { readonly state: SaveState }): React.ReactNode {
  if (state === "unknown") return null;
  const label =
    state === "saved"
      ? "Saved"
      : state === "modified"
        ? "Modified"
        : state === "saving"
          ? "Saving…"
          : "Save error";
  const tone =
    state === "error"
      ? "text-[color:var(--error)]"
      : state === "modified"
        ? "text-[color:var(--warning)]"
        : "text-tertiary";
  return (
    <span
      className={cn("ml-1 inline-flex items-center gap-1 text-xs tabular-nums", tone)}
      aria-live="polite"
      data-testid={`shell-save-state-${state}`}
    >
      {state === "saving" ? <InlineSpinner size={11} /> : null}
      {label}
    </span>
  );
}

const GROUP_ORDER: ReadonlyArray<ExportFormatGroup> = [
  "deck",
  "native",
  "pdf-web",
  "data",
  "images",
  "current",
];
const GROUP_LABEL: Record<ExportFormatGroup, string> = {
  deck: "Whole deck",
  native: "Native",
  "pdf-web": "PDF & web",
  data: "Data",
  images: "Images",
  current: "This slide",
};

function ExportMenu({ adapter }: { readonly adapter: ProductAdapter }): React.ReactNode {
  const [open, setOpen] = useState(false);
  const [dialogOpen, setDialogOpen] = useState(false);
  const [dialogInitialId, setDialogInitialId] = useState<string | undefined>(undefined);
  const ref = useRef<HTMLDivElement | null>(null);

  const groups = useMemo(
    () => groupFormats(adapter.exportFormats),
    [adapter.exportFormats]
  );

  useEffect(() => {
    if (!open) return;
    const onDoc = (e: MouseEvent) => {
      if (!ref.current?.contains(e.target as Node)) setOpen(false);
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") setOpen(false);
    };
    document.addEventListener("mousedown", onDoc);
    document.addEventListener("keydown", onKey);
    return () => {
      document.removeEventListener("mousedown", onDoc);
      document.removeEventListener("keydown", onKey);
    };
  }, [open]);

  if (!adapter.canExport || adapter.exportFormats.length === 0) return null;

  const baseFilename = stripExtension(adapter.filename);

  // Single-format products skip the dropdown — Export is still a
  // single icon-only button (no chevron). Preserves the original
  // top-bar density when only one format is on offer.
  if (adapter.exportFormats.length === 1) {
    const fmt = adapter.exportFormats[0]!;
    const needsDialog = fmt.kind === "dialog" || (fmt.optionFields && fmt.optionFields.length > 0);
    return (
      <>
        <ToolbarIcon
          label={`Export ${fmt.label}`}
          onClick={() => {
            if (needsDialog) {
              setDialogInitialId(fmt.id);
              setDialogOpen(true);
            } else {
              void adapter.onExport(fmt);
            }
          }}
          testId="shell-export"
        >
          <Download size={15} />
        </ToolbarIcon>
        <ExportDialog
          open={dialogOpen}
          onClose={() => setDialogOpen(false)}
          formats={adapter.exportFormats}
          initialFormatId={dialogInitialId}
          baseFilename={baseFilename}
          onExport={(format, options) => adapter.onExport(format, options)}
        />
      </>
    );
  }

  const openDialog = (id?: string) => {
    setOpen(false);
    setDialogInitialId(id);
    setDialogOpen(true);
  };

  return (
    <>
      <div className="relative" ref={ref}>
        <button
          type="button"
          className="inline-flex h-8 items-center gap-1 rounded-md px-2 text-secondary transition-colors hover:bg-hover hover:text-foreground focus:outline-none focus-visible:ring-2 focus-visible:ring-[var(--accent)]/40"
          onClick={() => setOpen((v) => !v)}
          aria-haspopup="menu"
          aria-expanded={open}
          title="Export"
          data-testid="shell-export"
        >
          <Download size={15} />
          <ChevronDown size={12} />
        </button>
        {open && (
          <div
            role="menu"
            className="absolute right-0 top-full z-30 mt-1 min-w-[260px] rounded-md border border-divider bg-surface p-1 shadow-md"
          >
            {groups.map(({ group, items }, gi) => (
              <div key={group}>
                {gi > 0 ? <div className="my-1 h-px bg-divider" aria-hidden /> : null}
                <div className="px-2 pb-0.5 pt-1 text-[10px] font-semibold uppercase tracking-wide text-tertiary">
                  {GROUP_LABEL[group]}
                </div>
                {items.map((fmt) => {
                  const dialogish =
                    fmt.kind === "dialog" || (fmt.optionFields && fmt.optionFields.length > 0);
                  return (
                    <button
                      key={fmt.id}
                      type="button"
                      role="menuitem"
                      className="flex w-full items-center gap-2 rounded-sm px-2 py-1.5 text-left text-sm text-foreground hover:bg-hover"
                      onClick={() => {
                        if (dialogish) {
                          openDialog(fmt.id);
                        } else {
                          setOpen(false);
                          void adapter.onExport(fmt);
                        }
                      }}
                      data-testid={`shell-export-${fmt.id}`}
                    >
                      <span className="text-secondary">
                        <FormatIcon icon={fmt.icon ?? guessIcon(fmt)} />
                      </span>
                      <span className="min-w-0 flex-1 truncate">{fmt.label}</span>
                      {dialogish ? (
                        <span
                          className="text-tertiary"
                          aria-label="Has options"
                          title="Has options"
                        >
                          <Sliders size={11} />
                        </span>
                      ) : null}
                    </button>
                  );
                })}
              </div>
            ))}
            <div className="my-1 h-px bg-divider" aria-hidden />
            <button
              type="button"
              role="menuitem"
              className="flex w-full items-center gap-2 rounded-sm px-2 py-1.5 text-left text-sm text-foreground hover:bg-hover"
              onClick={() => openDialog(undefined)}
              data-testid="shell-export-open-dialog"
            >
              <span className="text-secondary">
                <Sliders size={12} />
              </span>
              <span>More options…</span>
            </button>
          </div>
        )}
      </div>
      <ExportDialog
        open={dialogOpen}
        onClose={() => setDialogOpen(false)}
        formats={adapter.exportFormats}
        initialFormatId={dialogInitialId}
        baseFilename={baseFilename}
        onExport={(format, options) => adapter.onExport(format, options)}
      />
    </>
  );
}

function FormatIcon({ icon }: { readonly icon: ExportFormatIcon }): React.ReactNode {
  switch (icon) {
    case "doc":
      return <FileText size={13} />;
    case "sheet":
      return <FileSpreadsheet size={13} />;
    case "slides":
      return <Presentation size={13} />;
    case "pdf":
      return <FileType2 size={13} />;
    case "image":
      return <FileImage size={13} />;
    case "code":
      return <FileCode size={13} />;
    case "text":
      return <FileText size={13} />;
    default: {
      const _exhaustive: never = icon;
      void _exhaustive;
      return <FileText size={13} />;
    }
  }
}

function guessIcon(format: ExportFormat): ExportFormatIcon {
  switch (format.extension) {
    case "docx":
      return "doc";
    case "xlsx":
      return "sheet";
    case "pptx":
      return "slides";
    case "pdf":
      return "pdf";
    case "html":
    case "json":
      return "code";
    case "csv":
    case "tsv":
    case "txt":
    case "md":
      return "text";
    case "png":
    case "jpg":
    case "jpeg":
    case "svg":
    case "zip":
      return "image";
    default:
      return "doc";
  }
}

function defaultGroup(format: ExportFormat): ExportFormatGroup {
  switch (format.extension) {
    case "docx":
    case "xlsx":
    case "pptx":
      return "native";
    case "pdf":
    case "html":
      return "pdf-web";
    case "csv":
    case "tsv":
    case "json":
    case "txt":
    case "md":
      return "data";
    default:
      return "images";
  }
}

function groupFormats(
  formats: ReadonlyArray<ExportFormat>
): ReadonlyArray<{
  readonly group: ExportFormatGroup;
  readonly items: ReadonlyArray<ExportFormat>;
}> {
  const buckets = new Map<ExportFormatGroup, ExportFormat[]>();
  for (const fmt of formats) {
    const g = fmt.group ?? defaultGroup(fmt);
    const list = buckets.get(g) ?? [];
    list.push(fmt);
    buckets.set(g, list);
  }
  return GROUP_ORDER.filter((g) => (buckets.get(g)?.length ?? 0) > 0).map((g) => ({
    group: g,
    items: buckets.get(g) ?? [],
  }));
}

function stripExtension(name: string): string {
  const idx = name.lastIndexOf(".");
  if (idx <= 0) return name;
  return name.slice(0, idx);
}

function ToolbarIcon({
  children,
  label,
  shortcut,
  onClick,
  disabled,
  active,
  badge,
  testId,
}: {
  readonly children: React.ReactNode;
  readonly label: string;
  readonly shortcut?: string;
  readonly onClick: () => void;
  readonly disabled?: boolean;
  readonly active?: boolean;
  readonly badge?: number;
  readonly testId?: string;
}): React.ReactNode {
  const tip = shortcut ? `${label} (${shortcut})` : label;
  return (
    <button
      type="button"
      onClick={onClick}
      disabled={disabled}
      title={tip}
      aria-label={label}
      data-testid={testId}
      className={cn(
        "relative inline-flex h-8 w-8 items-center justify-center rounded-md text-secondary transition-colors",
        "hover:bg-hover hover:text-foreground",
        "focus:outline-none focus-visible:ring-2 focus-visible:ring-[var(--accent)]/40",
        active && "bg-hover text-foreground",
        disabled && "pointer-events-none opacity-40"
      )}
    >
      {children}
      {badge && badge > 0 ? (
        <span
          className="absolute -right-0.5 -top-0.5 inline-flex h-4 min-w-4 items-center justify-center rounded-full bg-[var(--accent)] px-1 text-[10px] font-semibold leading-none text-white"
          aria-hidden
        >
          {badge > 99 ? "99+" : badge}
        </span>
      ) : null}
    </button>
  );
}

function Sep(): React.ReactNode {
  return <span aria-hidden className="mx-1 h-5 w-px bg-divider" />;
}

function PaletteIcon(): React.ReactNode {
  return (
    <svg
      width="15"
      height="15"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
      strokeLinecap="round"
      strokeLinejoin="round"
    >
      <rect x="3" y="3" width="7" height="7" rx="1" />
      <rect x="14" y="3" width="7" height="7" rx="1" />
      <rect x="3" y="14" width="7" height="7" rx="1" />
      <path d="M14 17.5h7" />
      <path d="M17.5 14v7" />
    </svg>
  );
}
