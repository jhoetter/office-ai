"use client";

import { useEffect, useState, type ReactNode } from "react";
import { ArrowDownToLine, ArrowUpToLine, ArrowLeftToLine, ArrowRightToLine } from "lucide-react";
import { cn } from "@officeai/ui";
import type { EditorView } from "prosemirror-view";

/**
 * B4 — contextual table toolbar.
 *
 * Tables are rendered as PM atom nodes (cells aren't directly
 * editable yet) — selecting a table puts a `NodeSelection` on its
 * `table` node which we observe by polling on every PM transaction.
 * When a table is selected the toolbar floats above it and exposes
 * the row/column insertion commands available today; once cell-level
 * editing lands the same surface will gain Delete row/col and Merge
 * cells.
 *
 * The bar is rendered in a fixed-position layer over the editor
 * scroll container so it always appears within the viewport, even
 * for tables anchored near the bottom of the page card.
 */

export interface TableContextToolbarProps {
  readonly view: EditorView | null;
  readonly host: HTMLElement | null;
  readonly onInsertRow: (tableId: string, where: "top" | "bottom") => void;
  readonly onInsertColumn: (tableId: string, where: "start" | "end") => void;
}

interface SelectedTable {
  readonly tableId: string;
  readonly rect: { left: number; top: number; width: number };
}

export function TableContextToolbar(props: TableContextToolbarProps): ReactNode {
  const { view, host, onInsertRow, onInsertColumn } = props;
  const [selected, setSelected] = useState<SelectedTable | null>(null);

  useEffect(() => {
    if (!view || !host) {
      setSelected(null);
      return;
    }
    const compute = () => {
      const sel = view.state.selection;
      const node = (sel as { node?: { type: { name: string }; attrs: Record<string, unknown> } }).node;
      if (!node || node.type.name !== "table") {
        setSelected(null);
        return;
      }
      const tableId = typeof node.attrs.tableId === "string" ? node.attrs.tableId : null;
      if (!tableId) {
        setSelected(null);
        return;
      }
      // Look for the rendered table inside the editor host so we can
      // anchor the toolbar over it. We tag tables with the `pm-table`
      // class in the DOM serializer; index by NodeSelection's
      // position via `view.nodeDOM`.
      const dom = view.nodeDOM(sel.from) as HTMLElement | null;
      if (!dom) {
        setSelected(null);
        return;
      }
      const tableEl = dom.tagName === "TABLE" ? dom : dom.querySelector?.("table.pm-table");
      const target = (tableEl as HTMLElement | null) ?? dom;
      const rect = target.getBoundingClientRect();
      setSelected({
        tableId,
        rect: { left: rect.left, top: rect.top, width: rect.width },
      });
    };
    compute();
    // PM emits selection changes via dispatchTransaction; the host
    // also fires DOM `selectionchange` when atom nodes get focused.
    const onSel = () => compute();
    document.addEventListener("selectionchange", onSel);
    const onScroll = () => compute();
    host.addEventListener("scroll", onScroll, { capture: true });
    window.addEventListener("resize", onSel);
    return () => {
      document.removeEventListener("selectionchange", onSel);
      host.removeEventListener("scroll", onScroll, { capture: true });
      window.removeEventListener("resize", onSel);
    };
  }, [view, host]);

  if (!selected) return null;

  const top = Math.max(8, selected.rect.top - 36);
  const left = selected.rect.left + selected.rect.width / 2;

  return (
    <div
      role="toolbar"
      aria-label="Table actions"
      className={cn(
        "fixed z-40 flex items-center gap-0.5 rounded-md border border-divider bg-surface p-1 shadow-md"
      )}
      style={{ top, left, transform: "translateX(-50%)" }}
      data-testid="table-context-toolbar"
    >
      <Btn label="Insert row above" onClick={() => onInsertRow(selected.tableId, "top")}>
        <ArrowUpToLine size={14} />
      </Btn>
      <Btn label="Insert row below" onClick={() => onInsertRow(selected.tableId, "bottom")}>
        <ArrowDownToLine size={14} />
      </Btn>
      <span className="mx-1 h-4 w-px bg-divider" aria-hidden />
      <Btn label="Insert column at start" onClick={() => onInsertColumn(selected.tableId, "start")}>
        <ArrowLeftToLine size={14} />
      </Btn>
      <Btn label="Insert column at end" onClick={() => onInsertColumn(selected.tableId, "end")}>
        <ArrowRightToLine size={14} />
      </Btn>
    </div>
  );
}

function Btn(props: { label: string; onClick: () => void; children: ReactNode }) {
  return (
    <button
      type="button"
      title={props.label}
      aria-label={props.label}
      onClick={props.onClick}
      className="rounded p-1 text-secondary hover:bg-hover hover:text-foreground"
    >
      {props.children}
    </button>
  );
}
