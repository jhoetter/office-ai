/**
 * Window-scoped Alt-key state tracker. Browsers don't expose
 * `altKey` on `ClipboardEvent`s (which are dispatched in response to
 * the system clipboard, not directly to a key event), so we keep a
 * tiny mirror of the modifier-key state by listening to bubbling
 * `keydown`/`keyup`/`blur`. Editors subscribe once on mount and
 * read the state from the snapshot when they need to make a paste
 * decision.
 *
 * Lives in the embed directory because the only consumer today is
 * the cross-format paste path (`applyXlsxEmbed`) — Alt routes the
 * paste to "live OLE" instead of "materialised table". If other
 * surfaces start needing Alt-on-paste this should move up to
 * `apps/web/app/lib/keys/`.
 */

let altPressed = false;
let installed = false;

function onKeyDown(e: KeyboardEvent): void {
  if (e.key === "Alt" || e.altKey) altPressed = true;
}

function onKeyUp(e: KeyboardEvent): void {
  if (e.key === "Alt") altPressed = false;
  // Some browsers fire keyup with altKey=false even though Alt was
  // released; mirror that explicitly so we don't latch.
  if (!e.altKey) altPressed = false;
}

function onBlur(): void {
  altPressed = false;
}

/**
 * Install the global listeners. Idempotent — repeated calls are a
 * no-op. Returns the uninstall function so individual mounts can
 * dispose during HMR / unmount, although in practice the listeners
 * are cheap enough to leave installed for the document's lifetime.
 */
export function installAltKeyTracker(): () => void {
  if (typeof window === "undefined") return () => {};
  if (installed) return () => {};
  installed = true;
  window.addEventListener("keydown", onKeyDown, true);
  window.addEventListener("keyup", onKeyUp, true);
  window.addEventListener("blur", onBlur);
  return () => {
    if (!installed) return;
    installed = false;
    window.removeEventListener("keydown", onKeyDown, true);
    window.removeEventListener("keyup", onKeyUp, true);
    window.removeEventListener("blur", onBlur);
  };
}

export function isAltKeyPressed(): boolean {
  return altPressed;
}
