/**
 * Shared undo / redo plumbing.
 *
 * Every editor in this repo (DOCX, XLSX, PPTX) treats the headless
 * `CommandBus` (wrapped by each agent) as the single source of truth
 * for undo history. The toolbar buttons and the keyboard shortcut
 * MUST go through the same path so the toolbar's enabled state and
 * what `Cmd-Z` actually does can never disagree.
 *
 * This file owns:
 *   - chord detection (so Cmd-Z / Cmd-Shift-Z / Cmd-Y mean the same
 *     thing across editors and platforms);
 *   - the "is the focus inside a real form field?" guard so we don't
 *     fight the browser's native undo on `<input>` / `<textarea>` /
 *     contenteditable surfaces that aren't the editor itself;
 *   - the actual dispatch onto the agent.
 *
 * Renderers that need their own keymap plugin (e.g. ProseMirror's
 * `keymap`) should call into `runUndo` / `runRedo` from inside the
 * plugin's command rather than re-implementing the chord detection.
 */

/**
 * Minimal shape every editor agent already exposes (DocxAgent,
 * XlsxAgent, PptxAgent). We accept the loosest interface we can so
 * the helper has zero coupling to any one product.
 */
export interface UndoableAgent {
  canUndo(): boolean;
  canRedo(): boolean;
  undo(): unknown;
  redo(): unknown;
}

/**
 * Structural shape of a keyboard event. We accept both the DOM
 * `KeyboardEvent` and React's synthetic `KeyboardEvent<T>` (which
 * are NOT assignable to each other but share these fields). This
 * lets every editor — XLSX uses React onKeyDown, PPTX uses a
 * window-level DOM listener, the DOCX page-keymap uses PM-style
 * events — call into the same helper without ceremony.
 */
export interface KeyChord {
  readonly key: string;
  readonly metaKey: boolean;
  readonly ctrlKey: boolean;
  readonly shiftKey: boolean;
  readonly altKey: boolean;
  readonly target?: EventTarget | null;
  preventDefault(): void;
}

/**
 * True iff `e` represents the canonical "undo" chord:
 *   Cmd-Z (mac) / Ctrl-Z (everyone else), without Shift.
 *
 * We deliberately accept either modifier on every platform — Chrome
 * on macOS reports `metaKey` for Cmd, but a user with a Linux
 * keyboard plugged into a Mac will hit Ctrl. Matching both keeps
 * the editor predictable across browser / OS / hardware combos.
 */
export function isUndoChord(e: KeyChord): boolean {
  if (!(e.metaKey || e.ctrlKey)) return false;
  if (e.altKey) return false;
  const k = e.key.toLowerCase();
  return k === "z" && !e.shiftKey;
}

/**
 * True iff `e` is one of the canonical "redo" chords. We accept
 * both Cmd-Shift-Z (mac convention) and Cmd-Y (Windows convention)
 * because users come from both worlds and arguing over which one
 * "should" work is a quick way to lose them.
 */
export function isRedoChord(e: KeyChord): boolean {
  if (!(e.metaKey || e.ctrlKey)) return false;
  if (e.altKey) return false;
  const k = e.key.toLowerCase();
  if (k === "z" && e.shiftKey) return true;
  if (k === "y" && !e.shiftKey) return true;
  return false;
}

/**
 * True iff the event target is a real form-input surface where the
 * browser's native undo should win. This is what keeps `Cmd-Z`
 * inside the chat composer, the filename input in the top bar, and
 * the comment composer behaving like a normal text field instead of
 * undoing the last document edit.
 *
 * The editor surfaces themselves (ProseMirror's contenteditable,
 * the XLSX grid, the PPTX canvas) are NOT covered by this guard
 * because their keymaps run inside the editor and never reach the
 * window-level handler — by the time we're checking `isFormField`
 * we've already established that the focus is some *other* element.
 *
 * Note: ProseMirror sets `contenteditable="true"` on its root, so a
 * naive isContentEditable check would block the editor itself. The
 * call sites that use this guard live on `window`-level keydown
 * listeners *outside* the editor's PM keymap, where focus on the
 * editor surface still reports `isContentEditable === true` for
 * the wrapping div. To handle that, we let callers pass an
 * `editorSurface` element to opt that subtree out of the guard.
 */
export function isFormField(
  target: EventTarget | null,
  editorSurface?: Element | null
): boolean {
  if (!(target instanceof HTMLElement)) return false;
  if (editorSurface && editorSurface.contains(target)) return false;
  const tag = target.tagName;
  if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT") return true;
  // `isContentEditable` is the spec-y way to check, but jsdom doesn't
  // reflect the attribute onto the property — which made tests lie
  // about real-browser behaviour. Fall back to the attribute lookup
  // (and the inherited `contenteditable: true | inherit` ladder) so
  // both real browsers and jsdom agree.
  if (target.isContentEditable) return true;
  for (let el: HTMLElement | null = target; el; el = el.parentElement) {
    const attr = el.getAttribute("contenteditable");
    if (attr === "false") return false;
    if (attr === "true" || attr === "" || attr === "plaintext-only") return true;
  }
  return false;
}

/** Run undo on the agent if it can; returns true iff it ran. */
export function runUndo(agent: UndoableAgent | null | undefined): boolean {
  if (!agent || !agent.canUndo()) return false;
  agent.undo();
  return true;
}

/** Run redo on the agent if it can; returns true iff it ran. */
export function runRedo(agent: UndoableAgent | null | undefined): boolean {
  if (!agent || !agent.canRedo()) return false;
  agent.redo();
  return true;
}

export interface HandleUndoRedoOptions {
  /**
   * Element wrapping the editor surface. If the keydown originates
   * inside this element the form-field guard is suppressed (the
   * editor IS contenteditable, but it's not "a form field" in the
   * sense we care about — it's the thing we're trying to undo).
   */
  editorSurface?: Element | null;
}

/**
 * One-stop shop for window-level keydown handlers. Returns `true`
 * iff the event was handled (in which case the caller should NOT
 * run any other behaviour for this keydown — but `preventDefault`
 * is already taken care of here).
 *
 * Usage:
 *
 * ```ts
 * useEffect(() => {
 *   const onKey = (e: KeyboardEvent) => {
 *     if (handleUndoRedo(e, agentRef.current)) return;
 *     // ... other shortcuts ...
 *   };
 *   window.addEventListener("keydown", onKey);
 *   return () => window.removeEventListener("keydown", onKey);
 * }, []);
 * ```
 */
export function handleUndoRedo(
  e: KeyChord,
  agent: UndoableAgent | null | undefined,
  opts: HandleUndoRedoOptions = {}
): boolean {
  const isUndo = isUndoChord(e);
  const isRedo = !isUndo && isRedoChord(e);
  if (!isUndo && !isRedo) return false;
  if (isFormField(e.target ?? null, opts.editorSurface)) return false;
  e.preventDefault();
  if (isUndo) return runUndo(agent);
  return runRedo(agent);
}
