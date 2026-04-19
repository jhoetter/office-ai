"use client";

import { useCallback, useEffect, useRef, useState } from "react";

/**
 * Tiny state container for the shared {@link KeyboardShortcutsDialog}.
 *
 * Installs a global `keydown` listener for `Mod+/` (Cmd+/ on macOS,
 * Ctrl+/ elsewhere) that toggles the dialog. The listener is skipped
 * when focus sits on a contenteditable surface or an `<input>` /
 * `<textarea>` whose own keymap would otherwise be hijacked — but
 * `Mod+/` is not bound by ProseMirror's `baseKeymap`, the XLSX
 * formula bar, or any of the per-product surfaces, so callers
 * generally don't need to worry about conflicts.
 */
export interface UseShortcutsDialog {
  readonly open: boolean;
  readonly setOpen: (open: boolean) => void;
  readonly toggle: () => void;
}

export function useShortcutsDialog(): UseShortcutsDialog {
  const [open, setOpen] = useState(false);
  // Keep `open` in a ref so the global keydown handler stays stable
  // across re-renders (no listener churn on every toggle).
  const openRef = useRef(false);
  useEffect(() => {
    openRef.current = open;
  }, [open]);

  const toggle = useCallback(() => {
    setOpen((v) => !v);
  }, []);

  useEffect(() => {
    function onKeyDown(event: KeyboardEvent) {
      const isMod = event.metaKey || event.ctrlKey;
      if (!isMod) return;
      // `/` is layout-dependent, so prefer `event.key` here — it's
      // already normalised to "/" on US/UK/DE layouts when Shift is
      // up, and we don't want to listen for the Shift variant.
      if (event.key !== "/" && event.key !== "?") return;
      // Mod+/ is intentionally global. The dialog itself does not
      // consume the key and no per-product surface (formula bar,
      // comment composer, ProseMirror) binds it, so allow it through
      // for INPUT/TEXTAREA too — the user's muscle memory expects
      // the help dialog to appear no matter where focus sits.
      event.preventDefault();
      setOpen((v) => !v);
    }
    window.addEventListener("keydown", onKeyDown, true);
    return () => window.removeEventListener("keydown", onKeyDown, true);
  }, []);

  // Esc closes when open. Capture-phase so it wins over PM/grid Esc
  // handlers — we only act when the dialog is actually open.
  useEffect(() => {
    function onKeyDown(event: KeyboardEvent) {
      if (!openRef.current) return;
      if (event.key !== "Escape") return;
      event.preventDefault();
      event.stopPropagation();
      setOpen(false);
    }
    window.addEventListener("keydown", onKeyDown, true);
    return () => window.removeEventListener("keydown", onKeyDown, true);
  }, []);

  return { open, setOpen, toggle };
}
