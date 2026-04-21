"use client";

import { useEffect, useRef, useState, type DragEvent as ReactDragEvent, type ReactNode } from "react";
import { cn } from "@officeai/ui";
import { CommandPalette } from "./CommandPalette";
import { EditorStatusBar } from "./EditorStatusBar";
import { EditorTopBar } from "./EditorTopBar";
import { FindReplacePanel } from "./FindReplacePanel";
import { RightRail, useRightRailController, type RightRailTab } from "./RightRail";
import { Toaster, type ToastItem } from "./Toaster";
import { useTranslator } from "@/lib/i18n";
import type { ProductAdapter } from "./types";

export interface EditorShellProps {
  readonly adapter: ProductAdapter;
  /** The toolbar specific to the active product. Rendered immediately
   * below the top bar. */
  readonly toolbar?: ReactNode;
  /** The actual document body — ProseMirror surface, grid, slide
   * canvas. */
  readonly body: ReactNode;
  /** Optional product-specific status bar slots. */
  readonly statusBarLeft?: ReactNode;
  readonly statusBarRight?: ReactNode;
  /** Toasts to display. The shell renders them; the product owns the
   * queue. */
  readonly toasts?: ReadonlyArray<ToastItem>;
  readonly onDismissToast?: (id: string) => void;
  /** Drag-drop onto the shell. The shell handles dragover styling
   * and hands the file off to the product. */
  readonly onFileDrop?: (file: File) => void;
  /** Optional accept filter for drop (extension match). E.g. `.docx`. */
  readonly dropExtension?: string;
  /** Filename rename callback wired to the top bar. */
  readonly onRenameFilename?: (next: string) => void;
  /**
   * Optional slot rendered in the top bar (between the save-state
   * pill and the file ops). Currently used for the realtime
   * `PresenceStack`; product-agnostic so other "ambient" UI can
   * land here later (e.g. a sync status pill for cloud docs).
   */
  readonly topBarExtras?: ReactNode;
  /**
   * Imperative request from the product to open (or switch) the
   * right rail to a specific tab. Bump the `nonce` to trigger the
   * effect a second time when the same tab needs re-opening (e.g. a
   * palette command run twice in a row). The shell honours the
   * request without taking ownership of the rail's open/close
   * persistence — once the user closes the rail it stays closed
   * until the next nonce change.
   */
  readonly requestRailTab?: { readonly tab: RightRailTab; readonly nonce: number };
  /**
   * Optional override for the back button. When provided, the back
   * button calls this callback (instead of the default `<Link href="/">`).
   * Hosts that embed the editor (hof-os' `/edit-asset`, etc.) pass this
   * to navigate back to the *embedding* app rather than the standalone
   * office-ai home.
   */
  readonly onBack?: () => void;
}

/**
 * The single layout primitive used by all three editors.
 *
 * Layout (CSS grid):
 *
 *    [               EditorTopBar               ] (44px)
 *    [           Toolbar (per-product)          ] (auto)
 *    [    Body            ][   RightRail        ] (1fr)
 *    [               EditorStatusBar            ] (28px)
 *
 * Owns:
 *   - Drag-drop highlighting for the body
 *   - Cmd+K palette, Cmd+F find/replace toggling
 *   - Right-rail open/close + tab state (auto-opens on comments)
 *   - Toaster
 *   - Filename rename
 *
 * Stays out of:
 *   - The toolbar (each product owns its own)
 *   - The document/grid/canvas itself
 *   - The shortcut catalogue (lives in lib/shortcuts/)
 */
export function EditorShell({
  adapter,
  toolbar,
  body,
  statusBarLeft,
  statusBarRight,
  toasts,
  onDismissToast,
  onFileDrop,
  dropExtension,
  onRenameFilename,
  topBarExtras,
  requestRailTab,
  onBack,
}: EditorShellProps): ReactNode {
  const rail = useRightRailController(adapter);

  // Honour imperative open requests from the product (e.g. palette
  // commands like "Add shape animation" reveal the Animations tab).
  // We key the effect on the nonce so back-to-back requests for the
  // same tab still re-open it after the user closed it.
  useEffect(() => {
    if (!requestRailTab) return;
    rail.setTab(requestRailTab.tab);
    rail.setOpen(true);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [requestRailTab?.nonce]);
  const [paletteOpen, setPaletteOpen] = useState(false);
  const [findOpen, setFindOpen] = useState(false);
  const [findMode, setFindMode] = useState<"find" | "replace">("find");
  const [dragHover, setDragHover] = useState(false);
  const dragCounterRef = useRef(0);

  // Global shortcut bindings: Cmd+S, Cmd+K, Cmd+F, Cmd+Alt+F, Cmd+Alt+M.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const mod = e.metaKey || e.ctrlKey;
      if (!mod) return;
      const key = e.key.toLowerCase();
      // Cmd+S — save (calls product adapter's onSave; in embedded hosts
      // like hof-os this routes to the host's S3 PUT, not the browser's
      // built-in "Save Page" dialog which would otherwise capture this).
      if (key === "s" && !e.shiftKey && !e.altKey) {
        if (!adapter.canSave) return;
        e.preventDefault();
        void adapter.onSave();
        return;
      }
      // Cmd+K — palette
      if (key === "k" && !e.shiftKey) {
        e.preventDefault();
        setPaletteOpen((v) => !v);
        return;
      }
      // Cmd+F — find ; Cmd+Alt+F — replace
      if (key === "f" && !e.shiftKey) {
        if (!adapter.findAdapter) return;
        e.preventDefault();
        setFindMode(e.altKey ? "replace" : "find");
        setFindOpen(true);
        return;
      }
      // Cmd+Alt+M — toggle right rail (Comments)
      if (key === "m" && e.altKey) {
        e.preventDefault();
        rail.setTab("comments");
        rail.setOpen(!rail.open);
        return;
      }
      // Cmd+Alt+A — toggle right rail (Animations) when the active
      // product exposes a renderAnimationsPanel adapter.
      if (key === "a" && e.altKey) {
        if (!adapter.renderAnimationsPanel) return;
        e.preventDefault();
        const wasOpenOnAnim = rail.open && rail.tab === "animations";
        rail.setTab("animations");
        rail.setOpen(!wasOpenOnAnim);
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [adapter, adapter.findAdapter, adapter.renderAnimationsPanel, rail]);

  const handleDragOver = (e: ReactDragEvent<HTMLDivElement>) => {
    if (!onFileDrop) return;
    if (!hasFileItem(e.dataTransfer?.items)) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = "copy";
  };
  const handleDragEnter = (e: ReactDragEvent<HTMLDivElement>) => {
    if (!onFileDrop) return;
    if (!hasFileItem(e.dataTransfer?.items)) return;
    e.preventDefault();
    dragCounterRef.current += 1;
    setDragHover(true);
  };
  const handleDragLeave = (e: ReactDragEvent<HTMLDivElement>) => {
    if (!onFileDrop) return;
    e.preventDefault();
    dragCounterRef.current = Math.max(0, dragCounterRef.current - 1);
    if (dragCounterRef.current === 0) setDragHover(false);
  };
  const handleDrop = (e: ReactDragEvent<HTMLDivElement>) => {
    if (!onFileDrop) return;
    e.preventDefault();
    dragCounterRef.current = 0;
    setDragHover(false);
    const file = pickAcceptableFile(e.dataTransfer, dropExtension);
    if (file) onFileDrop(file);
  };

  return (
    <div
      className="flex h-full min-h-0 w-full flex-col overflow-hidden bg-background text-foreground"
      onDragOver={handleDragOver}
      onDragEnter={handleDragEnter}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
      data-testid="editor-shell"
      data-product={adapter.product}
    >
      <EditorTopBar
        adapter={adapter}
        onOpenCommandPalette={() => setPaletteOpen(true)}
        onOpenFindReplace={() => {
          if (!adapter.findAdapter) return;
          setFindMode("find");
          setFindOpen(true);
        }}
        onToggleRail={() => {
          rail.setTab("comments");
          rail.setOpen(!rail.open);
        }}
        railOpen={rail.open}
        onRenameFilename={onRenameFilename}
        extras={topBarExtras}
        onBack={onBack}
      />

      {/*
        Toolbar slot has a fixed 40 px height regardless of selection
        state. Each product's toolbar uses the shared `ToolbarRow`
        primitive which never wraps; on narrow viewports the leading
        slot scrolls horizontally instead of growing taller. This
        keeps the body (canvas / grid / page) anchored — selection
        changes never push it down by a row.
      */}
      {toolbar ? (
        <div className="h-10 min-h-10 max-h-10 shrink-0 border-b border-divider bg-background">{toolbar}</div>
      ) : null}

      <div className="relative flex min-h-0 min-w-0 flex-1">
        <main className="relative flex min-h-0 min-w-0 flex-1 flex-col overflow-hidden" role="main">
          {body}
          {dragHover ? <DropOverlay extension={dropExtension} /> : null}
          <FindReplacePanel
            adapter={adapter.findAdapter}
            open={findOpen}
            mode={findMode}
            onClose={() => setFindOpen(false)}
          />
        </main>
        <RightRail
          adapter={adapter}
          open={rail.open}
          tab={rail.tab}
          onTabChange={rail.setTab}
          onClose={() => rail.setOpen(false)}
        />
      </div>

      <EditorStatusBar adapter={adapter} leftSlot={statusBarLeft} rightSlot={statusBarRight} />

      <CommandPalette
        open={paletteOpen}
        onClose={() => setPaletteOpen(false)}
        commands={adapter.paletteCommands}
      />

      <Toaster toasts={toasts ?? []} onDismiss={onDismissToast ?? (() => {})} />
    </div>
  );
}

function hasFileItem(items: DataTransferItemList | undefined): boolean {
  if (!items) return false;
  for (let i = 0; i < items.length; i += 1) {
    if (items[i]?.kind === "file") return true;
  }
  return false;
}

function pickAcceptableFile(dt: DataTransfer, ext?: string): File | null {
  const files = dt?.files;
  if (!files || files.length === 0) return null;
  const file = files[0]!;
  if (!ext) return file;
  const lower = file.name.toLowerCase();
  if (lower.endsWith(ext.toLowerCase())) return file;
  return null;
}

function DropOverlay({ extension }: { readonly extension?: string }): ReactNode {
  const { t } = useTranslator();
  return (
    <div
      className={cn(
        "pointer-events-none absolute inset-0 z-10 flex items-center justify-center",
        "bg-[color:color-mix(in_srgb,var(--accent-light)_70%,transparent)] backdrop-blur-[1px]"
      )}
      aria-hidden
    >
      <div className="rounded-md border-2 border-dashed border-[color:var(--accent)] bg-background px-4 py-3 text-sm font-medium text-foreground shadow-md">
        {t("common.draftFile", { ext: extension ?? "" })}
      </div>
    </div>
  );
}
