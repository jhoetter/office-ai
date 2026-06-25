"use client";

/**
 * Generic "Selection pane" primitive — renders a vertical list of
 * named items (typically z-ordered, like PPTX shapes or XLSX chart
 * elements) with show/hide toggles, drag-reorder hooks, and click
 * to select.
 *
 * Each editor binds it to a different model (PPTX uses slide
 * shapes; XLSX chart elements; DOCX could use floating images /
 * shapes). The pane stays presentational.
 */

import type { ReactNode } from "react";
import { Eye, EyeOff } from "@officeai/ui/sonaloop-icons";
import { cn } from "@officeai/ui";

export interface SelectionPaneItem {
  readonly id: string;
  readonly label: string;
  readonly hidden?: boolean;
  readonly indent?: number;
}

export interface SelectionPaneProps {
  readonly items: ReadonlyArray<SelectionPaneItem>;
  readonly selectedId?: string | null;
  readonly onSelect: (id: string) => void;
  readonly onToggleHidden?: (id: string, hidden: boolean) => void;
  readonly onMoveUp?: (id: string) => void;
  readonly onMoveDown?: (id: string) => void;
  readonly testIdPrefix: string;
  readonly emptyMessage?: string;
}

export function SelectionPane(props: SelectionPaneProps): ReactNode {
  const {
    items,
    selectedId,
    onSelect,
    onToggleHidden,
    onMoveUp,
    onMoveDown,
    testIdPrefix,
    emptyMessage = "Nothing to show",
  } = props;

  if (items.length === 0) {
    return (
      <div className="px-3 py-2 text-xs text-tertiary" data-testid={`${testIdPrefix}-empty`}>
        {emptyMessage}
      </div>
    );
  }

  return (
    <ul className="flex flex-col gap-0.5" data-testid={testIdPrefix}>
      {items.map((item) => {
        const selected = selectedId === item.id;
        return (
          <li key={item.id}>
            <div
              className={cn(
                "flex items-center gap-1 rounded px-1 py-0.5 text-xs hover:bg-hover",
                selected && "bg-accent-soft text-accent"
              )}
              style={{ paddingLeft: `${(item.indent ?? 0) * 12 + 4}px` }}
              data-testid={`${testIdPrefix}-item-${item.id}`}
            >
              <button
                type="button"
                onClick={() => onSelect(item.id)}
                className="flex-1 truncate text-left"
                data-testid={`${testIdPrefix}-select-${item.id}`}
              >
                {item.label}
              </button>
              {onMoveUp ? (
                <button
                  type="button"
                  onClick={() => onMoveUp(item.id)}
                  title="Bring forward"
                  aria-label="Bring forward"
                  className="rounded px-1 text-secondary hover:bg-hover hover:text-foreground"
                  data-testid={`${testIdPrefix}-move-up-${item.id}`}
                >
                  ↑
                </button>
              ) : null}
              {onMoveDown ? (
                <button
                  type="button"
                  onClick={() => onMoveDown(item.id)}
                  title="Send backward"
                  aria-label="Send backward"
                  className="rounded px-1 text-secondary hover:bg-hover hover:text-foreground"
                  data-testid={`${testIdPrefix}-move-down-${item.id}`}
                >
                  ↓
                </button>
              ) : null}
              {onToggleHidden ? (
                <button
                  type="button"
                  onClick={() => onToggleHidden(item.id, !item.hidden)}
                  title={item.hidden ? "Show" : "Hide"}
                  aria-label={item.hidden ? "Show" : "Hide"}
                  aria-pressed={!item.hidden}
                  className="rounded px-1 text-secondary hover:bg-hover hover:text-foreground"
                  data-testid={`${testIdPrefix}-toggle-${item.id}`}
                >
                  {item.hidden ? <EyeOff size={11} /> : <Eye size={11} />}
                </button>
              ) : null}
            </div>
          </li>
        );
      })}
    </ul>
  );
}
