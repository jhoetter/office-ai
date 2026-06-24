/**
 * Action catalogue — the single source of truth that ties together
 * every user-facing operation in OfficeAI.
 *
 * Today, the same logical action ("insert image", "set paragraph
 * alignment", "apply autofilter", …) is declared in four uncoordinated
 * places:
 *
 *  1. a handler in `packages/{format}/src/commands/registry.ts`
 *  2. an agent / CLI binding in `packages/agent/src/actions-to-*.ts`
 *  3. a palette entry in `apps/web/app/{format}-editor/*.tsx`
 *  4. bespoke JSX in `Toolbar.tsx`, context menus, and dialogs
 *
 * Drift between these surfaces was invisible — the docx CLI had ~32
 * subcommands but the docx Cmd+K palette only listed 9 entries, so
 * most "AI-doable" actions were unreachable from the keyboard.
 *
 * The action catalogue collapses surfaces 2 and 3 into one declarative
 * list per format (`packages/{format}/src/actions/catalogue.ts`),
 * which the MCP adapter / CLI / palette / parity check all read from. Surface 4
 * (toolbar JSX) keeps its bespoke layout but pulls label/icon/shortcut
 * out of the same descriptor via `useAction(id)` so typos fail at
 * compile time.
 *
 * The contract here is intentionally framework-free (no zod, no
 * lucide-react, no React) so it can live in `@officeai/core` and be
 * imported by every format package, the CLI, the parity script, and
 * the web app without dragging a UI dep into headless code.
 */

/**
 * Where an action shows up inside the human web editor. Headless
 * surfaces (`agentCallable`, `cliCallable`) are explicit capability
 * fields on the descriptor instead of being inferred from this UI list.
 */
export type ActionSurface = "toolbar" | "palette" | "contextMenu";

/** File family this action targets. */
export type ActionFormat = "docx" | "xlsx" | "pptx" | "pdf";

/**
 * How the action's command payload is described for generated
 * surfaces. `catalogue-args` means the descriptor owns both `args`
 * and `buildPayload`; `custom` means a hand-written adapter remains
 * authoritative; `none` is for non-bus actions.
 */
export type ActionCommandSchema = "catalogue-args" | "custom" | "none";

/**
 * Argument kind for CLI / palette form generation. Stays narrow on
 * purpose — every kind here is something both commander and a tiny
 * inline form can render without needing a full schema validator.
 */
export type ActionArgKind = "string" | "number" | "boolean" | "filepath" | "selector" | "enum" | "stringList";

export interface ActionArg {
  /** camelCase property name in the parsed-args record. */
  readonly name: string;
  /**
   * CLI flag spelling, e.g. `--at`, `-q, --query`. Commander parses
   * this directly; the palette form derives its label from `name`.
   */
  readonly flag: string;
  readonly kind: ActionArgKind;
  /** Required arg → commander uses `requiredOption`; palette form marks the field. */
  readonly required?: boolean;
  /** Help text. Surfaces verbatim in `--help` and palette tooltip. */
  readonly description: string;
  /** Default value, surfaced in `--help`. */
  readonly default?: string | number | boolean;
  /** For `kind === "enum"` — accepted values. */
  readonly choices?: ReadonlyArray<string>;
  /** Optional placeholder for palette form input. */
  readonly placeholder?: string;
}

/**
 * One user-facing action. Agent, CLI and palette adapters read
 * directly from this; toolbar JSX reads `label` / `icon` / `shortcut`
 * via `useAction(id)`.
 */
export interface ActionDescriptor {
  /**
   * Stable, dot-namespaced id, e.g. `docx.insert-image`. Drives
   * palette recents, the parity check, and the typed `useAction`
   * hook.
   */
  readonly id: string;
  /**
   * Bus command type this action dispatches, e.g. `docx:insert-image`.
   * Set to `null` for non-mutating actions ("open", "switch to
   * viewing mode"). The parity check uses this field to verify every
   * registered handler appears in some catalogue entry.
   */
  readonly commandType: string | null;
  /** File family this action belongs to. */
  readonly format: ActionFormat;
  /** Human-facing label — palette + CLI help summary. */
  readonly label: string;
  /** One-line description — palette hint + CLI `--help` body. */
  readonly description: string;
  /**
   * Free-form section label used to group entries in the palette and
   * the CLI help screen. We deliberately keep this as a `string` so
   * each format can pick its own taxonomy without coordinating types.
   */
  readonly section: string;
  /**
   * Optional lucide-react icon name (as a string so this file does
   * not depend on `lucide-react`). Surfaces consult their own icon
   * map to resolve it.
   */
  readonly icon?: string;
  /** Display-only keyboard shortcut, e.g. `Cmd+K`. */
  readonly shortcut?: string;
  /** Surfaces this action wants to be on. */
  readonly surfaces: ReadonlyArray<ActionSurface>;
  /** Expose this action through MCP / agent-facing tool adapters. */
  readonly agentCallable: boolean;
  /** Expose this action through the web editor. */
  readonly webCallable: boolean;
  /** Expose this action through the terminal CLI. */
  readonly cliCallable: boolean;
  /** Mutation requires an explicit review / approval affordance. */
  readonly requiresReview: boolean;
  /** Generated surfaces can evaluate the action without writing. */
  readonly supportsDryRun: boolean;
  /** Generated surfaces can show a before/after diff after running. */
  readonly supportsDiff: boolean;
  /** Payload-schema ownership for generated adapters. */
  readonly commandSchema: ActionCommandSchema;
  /**
   * Declarative arg list. Used to:
   *   - register commander flags on the auto-generated subcommand
   *   - render an inline form in the palette when run from Cmd+K
   *   - sanity-check the parsed args before they reach `buildPayload`
   *
   * Omit entirely for parameter-less actions.
   */
  readonly args?: ReadonlyArray<ActionArg>;
  /**
   * Build the bus payload from parsed args. Required when
   * `commandType` is set; ignored otherwise. Throw to surface a
   * usage error before the bus dispatch.
   */
  readonly buildPayload?: (parsed: Record<string, unknown>) => unknown;
  /**
   * Explicit opt-out from one or more surfaces. Still counted as
   * "covered" by the parity check, but surface-specific checks are
   * skipped. Use sparingly and document the reason.
   */
  readonly hidden?: { readonly reason: string };
}

/**
 * Index a catalogue by id for O(1) lookups from React hooks and the
 * parity check.
 */
export function indexActionsById<T extends ActionDescriptor>(
  actions: ReadonlyArray<T>
): ReadonlyMap<string, T> {
  const map = new Map<string, T>();
  for (const a of actions) {
    if (map.has(a.id)) {
      throw new Error(`Duplicate action id "${a.id}" in catalogue`);
    }
    map.set(a.id, a);
  }
  return map;
}

/**
 * Index a catalogue by `commandType` (skipping entries with
 * `commandType: null`). Used by the parity check to verify every
 * registered handler has at least one catalogue entry.
 */
export function indexActionsByCommandType<T extends ActionDescriptor>(
  actions: ReadonlyArray<T>
): ReadonlyMap<string, ReadonlyArray<T>> {
  const map = new Map<string, T[]>();
  for (const a of actions) {
    if (a.commandType === null) continue;
    const existing = map.get(a.commandType);
    if (existing) {
      existing.push(a);
    } else {
      map.set(a.commandType, [a]);
    }
  }
  return map;
}
