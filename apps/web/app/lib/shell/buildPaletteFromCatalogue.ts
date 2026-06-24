/**
 * Adapter that turns a per-format `ActionDescriptor[]` (the central
 * catalogue, e.g. `docxActions` from `@officeai/docx`) into the
 * `PaletteCommand[]` the shell's Cmd+K palette consumes.
 *
 * The catalogue owns metadata (id, label, section, description,
 * shortcut) so the palette and action metadata never drift. The
 * product editor (DocxEditor, XlsxEditor, …) supplies a `runners` map
 * keyed by action id with the closure-bound side effect for each entry
 * whose `surfaces` includes "palette".
 *
 * The function:
 *   • silently skips catalogue entries not flagged with "palette"
 *   • silently skips palette entries that have no runner (so a catalogue
 *     entry can be added before the editor wires it up — a temporary
 *     gap surfaces as a missing palette item rather than a runtime error)
 *   • throws if a runner key references an action id absent from the
 *     catalogue (typo guard — the editor is the place that breaks)
 *   • respects per-runner `enabled` flags so the editor can gate by
 *     selection / cursor without re-listing the palette
 *
 * Intent: every palette item visible in Cmd+K is a deliberately
 * web-callable action. CLI and MCP exposure are separate descriptor
 * decisions (`cliCallable`, `agentCallable`) so a terminal convenience
 * command does not automatically become an agent tool.
 */

import type { ActionDescriptor } from "@officeai/core";
import type { PaletteCommand } from "./types";
import { translateAction, type TranslateFn } from "./translateAction";

export interface PaletteRunner {
  readonly run: () => void | Promise<void>;
  /** Optional: hide the palette entry without removing it from the
   * runners map (e.g. "switch to editing mode" while already editing). */
  readonly enabled?: boolean;
}

export type PaletteRunners = Record<string, PaletteRunner | undefined>;

/**
 * Build the palette command list from a per-format catalogue.
 *
 * Pass `t` from `useTranslator()` to localise labels, descriptions, and
 * section headers via the convention-based keys documented in
 * `translateAction.ts`. Calls without `t` (legacy / tests) keep the
 * catalogue's English strings.
 */
export function buildPaletteFromCatalogue(
  catalogue: ReadonlyArray<ActionDescriptor>,
  runners: PaletteRunners,
  t?: TranslateFn
): ReadonlyArray<PaletteCommand> {
  const byId = new Map<string, ActionDescriptor>();
  for (const a of catalogue) byId.set(a.id, a);

  for (const id of Object.keys(runners)) {
    if (!byId.has(id)) {
      throw new Error(
        `buildPaletteFromCatalogue: runner declared for unknown action id "${id}". Add it to the catalogue or fix the typo.`
      );
    }
  }

  const out: PaletteCommand[] = [];
  for (const action of catalogue) {
    if (action.hidden) continue;
    if (!action.surfaces.includes("palette")) continue;
    const runner = runners[action.id];
    if (!runner) continue;
    const strings = t
      ? translateAction(action, t)
      : { label: action.label, description: action.description, section: action.section };
    out.push({
      id: action.id,
      label: strings.label,
      section: strings.section,
      ...(action.shortcut ? { shortcut: action.shortcut } : {}),
      ...(strings.description ? { hint: strings.description } : {}),
      run: runner.run,
      ...(runner.enabled !== undefined ? { enabled: runner.enabled } : {}),
    });
  }
  return out;
}
