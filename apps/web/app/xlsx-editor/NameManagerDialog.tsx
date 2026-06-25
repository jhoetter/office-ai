"use client";

/**
 * C12 — Name Manager dialog.
 *
 * Compact CRUD over `XlsxWorkbook.definedNames`. Mirrors Excel's
 * Formulas → Name Manager. New names default to workbook scope and
 * the current selection; edits stay in-place inside each row so
 * the keyboard journey from list ↔ form is short.
 */

import { useEffect, useMemo, useRef, useState } from "react";
import type { ReactNode } from "react";
import { Pencil, Plus, Trash2, X } from "@officeai/ui/sonaloop-icons";
import { useFocusTrap } from "@officeai/ui";
import type { DefinedName } from "@officeai/xlsx";

export interface NameManagerDialogProps {
  readonly open: boolean;
  readonly onClose: () => void;
  readonly definedNames: ReadonlyArray<DefinedName>;
  /** All sheet names — used to populate the scope picker. */
  readonly sheetNames: ReadonlyArray<string>;
  /** Default A1 ref to seed the "new name" form's refersTo field. */
  readonly defaultRefersTo: string;
  readonly onAdd: (entry: { name: string; refersTo: string; scope?: string; comment?: string }) => void;
  readonly onUpdate: (entry: {
    name: string;
    scope?: string;
    nextName?: string;
    refersTo?: string;
    comment?: string;
  }) => void;
  readonly onRemove: (entry: { name: string; scope?: string }) => void;
}

interface Draft {
  readonly name: string;
  readonly refersTo: string;
  readonly scope: string; // "" = workbook
  readonly comment: string;
}

export function NameManagerDialog(props: NameManagerDialogProps): ReactNode {
  const { open, onClose, definedNames, sheetNames, defaultRefersTo, onAdd, onUpdate, onRemove } = props;

  const [draft, setDraft] = useState<Draft>({
    name: "",
    refersTo: defaultRefersTo,
    scope: "",
    comment: "",
  });
  const panelRef = useRef<HTMLDivElement>(null);
  useFocusTrap(panelRef, { enabled: open, onEscape: onClose });
  const [editing, setEditing] = useState<{
    name: string;
    scope: string | undefined;
    draft: Draft;
  } | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!open) {
      setEditing(null);
      setError(null);
      return;
    }
    setDraft((d) => ({ ...d, refersTo: defaultRefersTo }));
  }, [open, defaultRefersTo]);

  useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        e.preventDefault();
        if (editing) setEditing(null);
        else onClose();
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [open, onClose, editing]);

  const sortedNames = useMemo(() => {
    return [...definedNames].sort((a, b) => {
      const sa = a.scope ?? "";
      const sb = b.scope ?? "";
      if (sa !== sb) return sa.localeCompare(sb);
      return a.name.localeCompare(b.name);
    });
  }, [definedNames]);

  if (!open) return null;

  const submitNew = () => {
    setError(null);
    const name = draft.name.trim();
    const refersTo = draft.refersTo.trim();
    if (!name) {
      setError("Name is required.");
      return;
    }
    if (!refersTo) {
      setError("Refers to is required.");
      return;
    }
    try {
      onAdd({
        name,
        refersTo,
        ...(draft.scope ? { scope: draft.scope } : {}),
        ...(draft.comment.trim() ? { comment: draft.comment.trim() } : {}),
      });
      setDraft({ name: "", refersTo: defaultRefersTo, scope: draft.scope, comment: "" });
    } catch (e) {
      setError(String(e));
    }
  };

  const submitEdit = () => {
    if (!editing) return;
    setError(null);
    try {
      onUpdate({
        name: editing.name,
        scope: editing.scope,
        nextName: editing.draft.name.trim() || undefined,
        refersTo: editing.draft.refersTo.trim(),
        comment: editing.draft.comment.trim(),
      });
      setEditing(null);
    } catch (e) {
      setError(String(e));
    }
  };

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/40"
      role="dialog"
      aria-modal="true"
      aria-label="Name Manager"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div
        ref={panelRef}
        tabIndex={-1}
        className="relative w-[640px] rounded-md border border-divider bg-background p-4 text-sm text-foreground shadow-xl outline-none"
      >
        <div className="mb-3 flex items-center justify-between">
          <h2 className="text-base font-semibold">Name Manager</h2>
          <button
            type="button"
            aria-label="Close"
            className="inline-flex h-6 w-6 items-center justify-center rounded text-secondary hover:bg-hover"
            onClick={onClose}
          >
            <X size={14} />
          </button>
        </div>

        <div className="mb-4 max-h-[280px] overflow-auto rounded border border-divider">
          <table className="w-full text-xs">
            <thead className="bg-surface text-secondary">
              <tr>
                <th className="px-2 py-1 text-left font-medium">Name</th>
                <th className="px-2 py-1 text-left font-medium">Scope</th>
                <th className="px-2 py-1 text-left font-medium">Refers to</th>
                <th className="px-2 py-1" />
              </tr>
            </thead>
            <tbody>
              {sortedNames.length === 0 ? (
                <tr>
                  <td colSpan={4} className="px-2 py-3 text-center text-secondary">
                    No defined names yet.
                  </td>
                </tr>
              ) : (
                sortedNames.map((d) => {
                  const isEditing =
                    editing &&
                    editing.name === d.name &&
                    (editing.scope ?? undefined) === (d.scope ?? undefined);
                  if (isEditing) {
                    return (
                      <tr
                        key={`${d.name}-${d.scope ?? "wb"}`}
                        className="border-t border-divider bg-accent-soft"
                      >
                        <td className="px-2 py-1">
                          <input
                            data-testid={`nm-edit-name-${d.name}`}
                            value={editing!.draft.name}
                            onChange={(e) =>
                              setEditing((cur) =>
                                cur ? { ...cur, draft: { ...cur.draft, name: e.target.value } } : cur
                              )
                            }
                            className="h-6 w-full rounded border border-divider bg-background px-1 text-xs font-mono"
                          />
                        </td>
                        <td className="px-2 py-1 text-secondary">{d.scope ?? "Workbook"}</td>
                        <td className="px-2 py-1">
                          <input
                            data-testid={`nm-edit-refersTo-${d.name}`}
                            value={editing!.draft.refersTo}
                            onChange={(e) =>
                              setEditing((cur) =>
                                cur
                                  ? {
                                      ...cur,
                                      draft: { ...cur.draft, refersTo: e.target.value },
                                    }
                                  : cur
                              )
                            }
                            className="h-6 w-full rounded border border-divider bg-background px-1 text-xs font-mono"
                          />
                        </td>
                        <td className="px-2 py-1 text-right">
                          <button
                            type="button"
                            data-testid={`nm-save-${d.name}`}
                            className="rounded bg-accent px-2 py-0.5 text-xs text-white hover:bg-accent/90"
                            onClick={submitEdit}
                          >
                            Save
                          </button>
                          <button
                            type="button"
                            className="ml-1 rounded border border-divider px-2 py-0.5 text-xs hover:bg-hover"
                            onClick={() => setEditing(null)}
                          >
                            Cancel
                          </button>
                        </td>
                      </tr>
                    );
                  }
                  return (
                    <tr
                      key={`${d.name}-${d.scope ?? "wb"}`}
                      className="border-t border-divider"
                      data-testid={`nm-row-${d.name}`}
                    >
                      <td className="px-2 py-1 font-mono">{d.name}</td>
                      <td className="px-2 py-1 text-secondary">{d.scope ?? "Workbook"}</td>
                      <td className="px-2 py-1 font-mono text-secondary">{d.refersTo}</td>
                      <td className="px-2 py-1 text-right">
                        <button
                          type="button"
                          aria-label="Edit"
                          title="Edit"
                          data-testid={`nm-edit-${d.name}`}
                          className="inline-flex h-6 w-6 items-center justify-center rounded text-secondary hover:bg-hover"
                          onClick={() =>
                            setEditing({
                              name: d.name,
                              scope: d.scope,
                              draft: {
                                name: d.name,
                                refersTo: d.refersTo,
                                scope: d.scope ?? "",
                                comment: d.comment ?? "",
                              },
                            })
                          }
                        >
                          <Pencil size={12} />
                        </button>
                        <button
                          type="button"
                          aria-label="Delete"
                          title="Delete"
                          data-testid={`nm-delete-${d.name}`}
                          className="inline-flex h-6 w-6 items-center justify-center rounded text-secondary hover:bg-hover"
                          onClick={() => onRemove({ name: d.name, scope: d.scope })}
                        >
                          <Trash2 size={12} />
                        </button>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>

        <div className="mb-3 grid grid-cols-[1fr,140px,1fr] gap-2">
          <label className="flex flex-col gap-0.5 text-xs">
            <span className="text-secondary">Name</span>
            <input
              data-testid="nm-new-name"
              value={draft.name}
              onChange={(e) => setDraft((d) => ({ ...d, name: e.target.value }))}
              placeholder="e.g. Revenue"
              className="h-7 rounded border border-divider bg-background px-2 text-xs font-mono"
            />
          </label>
          <label className="flex flex-col gap-0.5 text-xs">
            <span className="text-secondary">Scope</span>
            <select
              data-testid="nm-new-scope"
              value={draft.scope}
              onChange={(e) => setDraft((d) => ({ ...d, scope: e.target.value }))}
              className="h-7 rounded border border-divider bg-background px-1 text-xs"
            >
              <option value="">Workbook</option>
              {sheetNames.map((s) => (
                <option key={s} value={s}>
                  {s}
                </option>
              ))}
            </select>
          </label>
          <label className="flex flex-col gap-0.5 text-xs">
            <span className="text-secondary">Refers to</span>
            <input
              data-testid="nm-new-refersTo"
              value={draft.refersTo}
              onChange={(e) => setDraft((d) => ({ ...d, refersTo: e.target.value }))}
              placeholder="Sheet1!$A$1:$A$10"
              className="h-7 rounded border border-divider bg-background px-2 text-xs font-mono"
            />
          </label>
        </div>

        {error ? (
          <p className="mb-2 rounded border border-destructive/40 bg-destructive/10 px-2 py-1 text-xs text-destructive">
            {error}
          </p>
        ) : null}

        <div className="flex justify-end gap-2 border-t border-divider pt-3">
          <button
            type="button"
            onClick={onClose}
            className="h-7 rounded border border-divider bg-background px-3 text-xs hover:bg-hover"
          >
            Close
          </button>
          <button
            type="button"
            data-testid="nm-add"
            onClick={submitNew}
            className="inline-flex h-7 items-center gap-1 rounded bg-accent px-3 text-xs font-medium text-white hover:bg-accent/90"
          >
            <Plus size={12} />
            Add name
          </button>
        </div>
      </div>
    </div>
  );
}
