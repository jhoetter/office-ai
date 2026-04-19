"use client";

import { useRef, useState, type ReactNode } from "react";
import { ChevronDown, Table as TableIcon } from "lucide-react";
import { cn } from "@officeai/ui";
import { ToolbarMenu } from "../lib/shell";

/**
 * B4 — Word-style "Insert Table" grid picker.
 *
 * Renders a `MAX_ROWS × MAX_COLS` grid that highlights the
 * `(rows × cols)` rectangle the user is currently hovering. Click
 * commits an insert via {@link onInsert}. The picker stays purely
 * presentational — the surrounding editor owns the `docx:insert-table`
 * dispatch so the command can resolve its insertion point against the
 * current PM selection.
 *
 * The grid mirrors Word's affordance closely: 10x10 cells, 18px each,
 * with an "N × M Table" caption that updates on hover. The caption
 * also includes a manual entry fallback for very large tables which
 * we bind to "Insert table…" — opening a tiny inline dialog rather
 * than a separate modal so the flow stays one-handed.
 */

const MAX_ROWS = 10;
const MAX_COLS = 10;

export interface InsertTableMenuProps {
  readonly disabled?: boolean;
  readonly onInsert: (rows: number, cols: number) => void;
}

export function InsertTableMenu(props: InsertTableMenuProps): ReactNode {
  const [open, setOpen] = useState(false);
  const [hover, setHover] = useState<{ rows: number; cols: number } | null>(null);
  const [customRows, setCustomRows] = useState(3);
  const [customCols, setCustomCols] = useState(3);
  const triggerRef = useRef<HTMLButtonElement | null>(null);

  const commit = (rows: number, cols: number) => {
    props.onInsert(rows, cols);
    setOpen(false);
    setHover(null);
  };

  return (
    <>
      <button
        ref={triggerRef}
        type="button"
        title="Insert table"
        aria-label="Insert table"
        aria-haspopup="dialog"
        aria-expanded={open}
        disabled={props.disabled}
        onClick={() => setOpen((v) => !v)}
        className={cn(
          "inline-flex items-center gap-0.5 rounded-md p-1.5 text-secondary hover:bg-hover hover:text-foreground disabled:opacity-50",
          open && "bg-hover text-foreground"
        )}
        data-testid="insert-table-trigger"
      >
        <TableIcon size={14} />
        <ChevronDown size={10} />
      </button>
      <ToolbarMenu
        open={open}
        onClose={() => {
          setOpen(false);
          setHover(null);
        }}
        triggerRef={triggerRef}
        role="dialog"
        className="w-max rounded-md border border-divider bg-surface p-2 shadow-md"
      >
        <div
          className="grid gap-[2px]"
          style={{
            gridTemplateColumns: `repeat(${MAX_COLS}, 18px)`,
            gridTemplateRows: `repeat(${MAX_ROWS}, 18px)`,
          }}
          onMouseLeave={() => setHover(null)}
          data-testid="insert-table-grid"
        >
          {Array.from({ length: MAX_ROWS * MAX_COLS }, (_, i) => {
            const r = Math.floor(i / MAX_COLS) + 1;
            const c = (i % MAX_COLS) + 1;
            const inside = hover !== null && r <= hover.rows && c <= hover.cols;
            return (
              <button
                key={i}
                type="button"
                aria-label={`${r} × ${c}`}
                onMouseEnter={() => setHover({ rows: r, cols: c })}
                onClick={() => commit(r, c)}
                className={cn(
                  "h-[18px] w-[18px] rounded-sm border transition-colors",
                  inside ? "border-accent bg-accent-light" : "border-divider bg-background hover:bg-hover"
                )}
              />
            );
          })}
        </div>
        <div className="mt-2 text-center text-xs tabular-nums text-secondary">
          {hover ? `${hover.rows} × ${hover.cols} table` : "Drag to size"}
        </div>
        <div className="mt-2 flex items-center justify-center gap-2 border-t border-divider pt-2 text-xs">
          <span className="text-secondary">Custom</span>
          <NumberField value={customRows} min={1} max={50} onChange={setCustomRows} label="rows" />
          <span className="text-secondary">×</span>
          <NumberField value={customCols} min={1} max={20} onChange={setCustomCols} label="columns" />
          <button
            type="button"
            onClick={() => commit(customRows, customCols)}
            className="ml-1 rounded border border-divider bg-background px-2 py-0.5 hover:bg-hover"
            data-testid="insert-table-custom-apply"
          >
            Insert
          </button>
        </div>
      </ToolbarMenu>
    </>
  );
}

function NumberField(props: {
  value: number;
  min: number;
  max: number;
  label: string;
  onChange: (v: number) => void;
}) {
  return (
    <input
      type="number"
      aria-label={props.label}
      value={props.value}
      min={props.min}
      max={props.max}
      onChange={(e) => {
        const v = Number(e.currentTarget.value);
        if (!Number.isFinite(v)) return;
        props.onChange(Math.max(props.min, Math.min(props.max, Math.round(v))));
      }}
      className="h-6 w-12 rounded border border-divider bg-background px-1 text-right tabular-nums"
    />
  );
}
