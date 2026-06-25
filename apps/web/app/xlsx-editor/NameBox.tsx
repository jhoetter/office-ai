"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import type { ReactNode } from "react";
import { ChevronDown } from "@officeai/ui/sonaloop-icons";
import type { DefinedName } from "@officeai/xlsx";

/**
 * C12 — Name Box.
 *
 * Sits to the left of the formula bar and serves three purposes:
 *
 *   1. **Display**: shows a defined name covering the current
 *      selection (workbook scope wins over sheet scope), or the
 *      A1 ref / size-summary of the selection itself.
 *   2. **Navigation**: the user types a defined name, an A1 cell
 *      ref, or an A1 range and presses Enter — we resolve the
 *      target and call `onJump`. Names take precedence over refs
 *      (matches Excel).
 *   3. **Authoring**: the user types a brand-new identifier and
 *      presses Enter — we ask the parent to create a workbook-
 *      scoped defined name pointing at the current selection
 *      via `onCreateName`. The Name Manager is reachable via
 *      the chevron menu.
 *
 * The component never owns the selection itself — it only
 * receives the resolved selection ref and dispatches navigation
 * / creation intents.
 */
export interface NameBoxProps {
  /** Display A1 ref of the current selection, e.g. "A1" or "A1:C5". */
  readonly selectionRef: string;
  /** Workbook + sheet-scoped defined names to surface in the dropdown. */
  readonly definedNames: ReadonlyArray<DefinedName>;
  /** Active sheet name — used to filter sheet-scoped names. */
  readonly activeSheet: string | undefined;
  /** Resolve typed input to a navigable target; receives the raw text. */
  readonly onJump: (input: string) => boolean;
  /** Create a workbook-scoped name pointing at the current selection. */
  readonly onCreateName: (name: string) => void;
  /** Open the Name Manager dialog. */
  readonly onOpenManager: () => void;
  readonly disabled?: boolean;
}

export function NameBox(props: NameBoxProps): ReactNode {
  const { selectionRef, definedNames, activeSheet, onJump, onCreateName, onOpenManager, disabled } = props;

  const [draft, setDraft] = useState<string>("");
  const [focused, setFocused] = useState<boolean>(false);
  const [menuOpen, setMenuOpen] = useState<boolean>(false);
  const inputRef = useRef<HTMLInputElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);

  const visibleNames = useMemo(
    () => definedNames.filter((d) => !d.hidden && (d.scope === undefined || d.scope === activeSheet)),
    [definedNames, activeSheet]
  );

  // Display value — what the user sees when not editing. Prefer a
  // defined name covering the exact selection ref over the bare A1
  // ref so Excel's "named range named display" behaviour holds.
  const display = useMemo(() => {
    if (focused) return draft;
    const namedMatch = definedNames.find((d) => {
      if (d.hidden) return false;
      if (d.scope !== undefined && d.scope !== activeSheet) return false;
      return refersToMatchesSelection(d.refersTo, selectionRef, activeSheet);
    });
    if (namedMatch) return namedMatch.name;
    return selectionRef || "";
  }, [focused, draft, definedNames, selectionRef, activeSheet]);

  useEffect(() => {
    if (!menuOpen) return;
    const onDoc = (e: MouseEvent) => {
      const t = e.target as HTMLElement | null;
      if (t && containerRef.current?.contains(t)) return;
      setMenuOpen(false);
    };
    window.addEventListener("mousedown", onDoc);
    return () => window.removeEventListener("mousedown", onDoc);
  }, [menuOpen]);

  function commit(value: string): void {
    const trimmed = value.trim();
    if (!trimmed) return;
    // First try jump (existing name or ref). If it fails *and* the
    // input looks like a valid identifier, treat it as a request to
    // mint a new workbook-scoped name covering the current selection.
    if (onJump(trimmed)) return;
    if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(trimmed)) {
      onCreateName(trimmed);
    }
  }

  return (
    <div ref={containerRef} className="relative inline-flex items-center">
      <input
        ref={inputRef}
        data-testid="name-box"
        aria-label="Name box"
        title="Name box — type a name or cell reference"
        value={focused ? draft : display}
        disabled={disabled}
        onFocus={() => {
          setFocused(true);
          setDraft(display);
          requestAnimationFrame(() => inputRef.current?.select());
        }}
        onBlur={() => {
          setFocused(false);
          setDraft("");
        }}
        onChange={(e) => setDraft(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter") {
            e.preventDefault();
            const v = draft;
            inputRef.current?.blur();
            commit(v);
          } else if (e.key === "Escape") {
            e.preventDefault();
            setFocused(false);
            setDraft("");
            inputRef.current?.blur();
          }
        }}
        className="h-7 w-[140px] rounded border border-divider bg-background px-2 text-xs font-mono text-foreground focus:border-accent focus:outline-none disabled:opacity-50"
      />
      <button
        type="button"
        data-testid="name-box-menu"
        aria-label="Defined names"
        title="Defined names"
        disabled={disabled}
        onClick={() => setMenuOpen((v) => !v)}
        className="ml-0.5 inline-flex h-7 w-5 items-center justify-center rounded text-secondary hover:bg-hover disabled:opacity-50"
      >
        <ChevronDown size={12} />
      </button>
      {menuOpen ? (
        <div
          data-testid="name-box-menu-popover"
          role="listbox"
          className="absolute left-0 top-full z-30 mt-1 min-w-[200px] rounded border border-divider bg-background py-1 shadow-lg"
        >
          {visibleNames.length === 0 ? (
            <div className="px-3 py-1.5 text-xs text-secondary">No defined names yet</div>
          ) : (
            visibleNames.map((d) => (
              <button
                key={`${d.name}-${d.scope ?? "wb"}`}
                type="button"
                role="option"
                data-testid={`name-box-option-${d.name}`}
                className="block w-full px-3 py-1 text-left text-xs font-mono text-foreground hover:bg-hover"
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => {
                  setMenuOpen(false);
                  onJump(d.name);
                }}
              >
                <span className="text-foreground">{d.name}</span>
                <span className="ml-2 text-tertiary">{d.refersTo}</span>
              </button>
            ))
          )}
          <div className="my-1 border-t border-divider" />
          <button
            type="button"
            data-testid="name-box-open-manager"
            className="block w-full px-3 py-1 text-left text-xs text-foreground hover:bg-hover"
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => {
              setMenuOpen(false);
              onOpenManager();
            }}
          >
            Manage names…
          </button>
        </div>
      ) : null}
    </div>
  );
}

function refersToMatchesSelection(
  refersTo: string,
  selectionRef: string,
  activeSheet: string | undefined
): boolean {
  // Strip leading `=`, sheet qualifier, and absolute markers so we
  // compare the same canonical form. We keep this comparison
  // intentionally narrow — only exact matches against the active
  // sheet's selection should swap the box's display.
  if (!selectionRef) return false;
  const m = /^(?:'(?:[^']|'')+'|[A-Za-z_][\w. ]*)!(.+)$/.exec(refersTo.trim().replace(/^=/, ""));
  let sheet = activeSheet;
  let body = refersTo.trim().replace(/^=/, "");
  if (m) {
    const sm = /^(?:'((?:[^']|'')+)'|([A-Za-z_][\w. ]*))!/.exec(refersTo.trim().replace(/^=/, ""));
    sheet = (sm?.[1] ?? sm?.[2] ?? activeSheet)?.replace(/''/g, "'");
    body = m[1]!;
  }
  if (sheet !== activeSheet) return false;
  return body.replace(/\$/g, "") === selectionRef.replace(/\$/g, "");
}
