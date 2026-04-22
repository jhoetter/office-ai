"use client";

/**
 * Word's References → Bookmark dialog.
 *
 * Word's bookmark dialog has three responsibilities:
 *   1. add a new bookmark at the current selection (or a zero-length
 *      anchor at the caret if nothing is selected),
 *   2. list existing bookmarks so the user can jump to or delete them,
 *   3. validate names against Word's identifier rules
 *      (`[A-Za-z_][\w]*`).
 *
 * We dispatch `docx:insert-bookmark` / `docx:delete-bookmark` for (1)
 * and (3); "Go to" delegates to a callback so the editor can move the
 * caret to the anchor and scroll it into view (the bookmark commands
 * themselves stay pure functional mutations on the snapshot).
 */

import { useEffect, useRef, useState } from "react";
import type { ReactNode } from "react";
import { X } from "lucide-react";
import { useFocusTrap } from "@officeai/ui";

export interface BookmarkRow {
  readonly name: string;
  readonly paragraphId: string;
  readonly startOffset: number;
  readonly endOffset: number;
}

export interface BookmarkDialogProps {
  readonly open: boolean;
  readonly bookmarks: ReadonlyArray<BookmarkRow>;
  /** Whether a non-empty selection is currently active; controls the Add button copy. */
  readonly hasSelection: boolean;
  readonly onClose: () => void;
  readonly onAdd: (name: string) => void;
  readonly onDelete: (name: string) => void;
  readonly onGoTo: (b: BookmarkRow) => void;
}

const VALID_NAME = /^[A-Za-z_][\w]*$/;

export function BookmarkDialog(props: BookmarkDialogProps): ReactNode {
  const { open, bookmarks, hasSelection, onClose, onAdd, onDelete, onGoTo } = props;
  const [name, setName] = useState("");
  const [error, setError] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const panelRef = useRef<HTMLDivElement>(null);
  useFocusTrap(panelRef, { enabled: open, onEscape: onClose });

  useEffect(() => {
    if (open) {
      setName("");
      setError(null);
      const t = window.setTimeout(() => inputRef.current?.focus(), 30);
      return () => window.clearTimeout(t);
    }
    return undefined;
  }, [open]);

  if (!open) return null;

  const handleAdd = (): void => {
    const trimmed = name.trim();
    if (!VALID_NAME.test(trimmed)) {
      setError("Name must start with a letter or _ and contain only letters, digits and underscores.");
      return;
    }
    setError(null);
    onAdd(trimmed);
    setName("");
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40">
      <div
        ref={panelRef}
        role="dialog"
        aria-modal="true"
        aria-labelledby="docx-bookmark-dialog-title"
        className="w-[28rem] rounded-lg border border-divider bg-card text-card-foreground shadow-xl"
        data-testid="docx-bookmark-dialog"
      >
        <div className="flex items-center justify-between border-b border-divider px-4 py-3">
          <h2 id="docx-bookmark-dialog-title" className="text-sm font-semibold">
            Bookmark
          </h2>
          <button
            type="button"
            onClick={onClose}
            aria-label="Close"
            className="rounded p-1 text-secondary hover:bg-divider/40"
          >
            <X size={14} />
          </button>
        </div>
        <div className="space-y-3 px-4 py-3">
          <div>
            <label htmlFor="docx-bookmark-name" className="block text-xs font-medium text-secondary">
              Bookmark name
            </label>
            <input
              ref={inputRef}
              id="docx-bookmark-name"
              type="text"
              value={name}
              onChange={(e) => setName(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter") {
                  e.preventDefault();
                  handleAdd();
                }
              }}
              placeholder="MyAnchor"
              data-testid="docx-bookmark-name-input"
              className="mt-1 w-full rounded border border-divider bg-background px-2 py-1 text-sm"
            />
            {error && <p className="mt-1 text-xs text-destructive">{error}</p>}
            <p className="mt-1 text-[11px] text-secondary">
              {hasSelection
                ? "The selected range will be wrapped with the new bookmark anchors."
                : "A zero-length anchor will be inserted at the caret position."}
            </p>
          </div>
          <div>
            <h3 className="mb-1 text-xs font-medium text-secondary">Existing bookmarks</h3>
            {bookmarks.length === 0 ? (
              <p className="text-xs text-secondary">No bookmarks in this document.</p>
            ) : (
              <ul
                className="max-h-40 divide-y divide-divider overflow-y-auto rounded border border-divider"
                data-testid="docx-bookmark-list"
              >
                {bookmarks.map((b) => (
                  <li key={b.name} className="flex items-center justify-between px-2 py-1 text-xs">
                    <span className="truncate" title={b.name}>
                      {b.name}
                    </span>
                    <span className="flex items-center gap-1">
                      <button
                        type="button"
                        onClick={() => onGoTo(b)}
                        className="rounded px-1 py-0.5 text-[11px] text-link hover:bg-divider/40"
                      >
                        Go to
                      </button>
                      <button
                        type="button"
                        onClick={() => onDelete(b.name)}
                        className="rounded px-1 py-0.5 text-[11px] text-destructive hover:bg-divider/40"
                      >
                        Delete
                      </button>
                    </span>
                  </li>
                ))}
              </ul>
            )}
          </div>
        </div>
        <div className="flex items-center justify-end gap-2 border-t border-divider px-4 py-3">
          <button
            type="button"
            onClick={onClose}
            className="rounded border border-divider px-3 py-1 text-xs hover:bg-divider/40"
          >
            Close
          </button>
          <button
            type="button"
            onClick={handleAdd}
            disabled={name.trim().length === 0}
            data-testid="docx-bookmark-add"
            className="rounded bg-primary px-3 py-1 text-xs text-primary-foreground disabled:opacity-50"
          >
            Add
          </button>
        </div>
      </div>
    </div>
  );
}
