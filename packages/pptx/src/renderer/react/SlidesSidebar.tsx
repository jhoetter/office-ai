import * as React from "react";
import type { ChartPart, Slide, SlideSize, ThemeColorScheme } from "../../model/types.js";
import { SlideThumbnail } from "./SlideThumbnail.js";

export type SlideContextAction =
  | "insert-before"
  | "insert-after"
  | "duplicate"
  | "delete"
  | "move-up"
  | "move-down"
  | "move-to-start"
  | "move-to-end";

export interface SlidesSidebarProps {
  readonly slides: ReadonlyArray<Slide>;
  readonly slideSize: SlideSize;
  readonly mediaUrls?: ReadonlyMap<string, string>;
  readonly theme?: ThemeColorScheme;
  readonly charts?: ReadonlyMap<string, ChartPart>;
  readonly activeIndex: number;
  readonly onSelect: (index: number) => void;
  readonly thumbnailWidth?: number;
  /**
   * Native HTML5 drag-and-drop reorder. When supplied, slide rows
   * become draggable; releasing on another row dispatches
   * `(fromIndex, toIndex)`. The host is expected to apply the
   * reorder via `pptx:move-slide`.
   */
  readonly onReorder?: (fromIndex: number, toIndex: number) => void;
  /**
   * Right-click / context-menu handler. When supplied, the sidebar
   * intercepts `contextmenu` on each thumbnail and renders an inline
   * popover menu. Each entry calls back with the slide index and
   * action; the host is responsible for the actual command dispatch.
   */
  readonly onContextAction?: (slideIndex: number, action: SlideContextAction) => void;
}

interface MenuState {
  readonly slideIndex: number;
  readonly x: number;
  readonly y: number;
}

export function SlidesSidebar(props: SlidesSidebarProps): React.ReactElement {
  const {
    slides,
    slideSize,
    mediaUrls,
    theme,
    charts,
    activeIndex,
    onSelect,
    thumbnailWidth = 180,
    onReorder,
    onContextAction,
  } = props;

  const [dragIndex, setDragIndex] = React.useState<number | null>(null);
  const [dropTarget, setDropTarget] = React.useState<{ index: number; before: boolean } | null>(null);
  const [menu, setMenu] = React.useState<MenuState | null>(null);

  // Dismiss the context menu on any outside click / scroll / Esc so it
  // never gets stuck open.
  React.useEffect(() => {
    if (!menu) return;
    const onDoc = (e: MouseEvent | KeyboardEvent) => {
      if (e instanceof KeyboardEvent && e.key !== "Escape") return;
      setMenu(null);
    };
    document.addEventListener("mousedown", onDoc);
    document.addEventListener("keydown", onDoc);
    return () => {
      document.removeEventListener("mousedown", onDoc);
      document.removeEventListener("keydown", onDoc);
    };
  }, [menu]);

  const draggable = onReorder !== undefined;

  const handleDrop = (toIndex: number, before: boolean) => {
    if (!onReorder || dragIndex === null) return;
    let target = before ? toIndex : toIndex + 1;
    // Splice-style adjustment: dropping after a row earlier than the
    // dragged one shifts the destination by one (the source slot is
    // removed first, then re-inserted).
    if (dragIndex < target) target -= 1;
    if (target !== dragIndex && target >= 0 && target <= slides.length - 1) {
      onReorder(dragIndex, target);
    }
    setDragIndex(null);
    setDropTarget(null);
  };

  return (
    <ul
      className="officeai-pptx-sidebar"
      style={{
        listStyle: "none",
        margin: 0,
        padding: 8,
        display: "flex",
        flexDirection: "column",
        gap: 8,
        overflowY: "auto",
        height: "100%",
        position: "relative",
      }}
      onDragLeave={() => setDropTarget(null)}
    >
      {slides.map((s, i) => {
        const isActive = i === activeIndex;
        const isDragging = dragIndex === i;
        const showInsertBefore = dropTarget?.index === i && dropTarget.before;
        const showInsertAfter = dropTarget?.index === i && !dropTarget.before;
        return (
          <li
            key={s.id}
            data-testid={`pptx-slide-row-${i}`}
            data-active={isActive ? "true" : undefined}
            draggable={draggable}
            onDragStart={(e) => {
              if (!draggable) return;
              setDragIndex(i);
              e.dataTransfer.effectAllowed = "move";
              try {
                e.dataTransfer.setData("text/plain", String(i));
              } catch {
                /* some browsers throw on synthetic events */
              }
            }}
            onDragEnd={() => {
              setDragIndex(null);
              setDropTarget(null);
            }}
            onDragOver={(e) => {
              if (!draggable || dragIndex === null) return;
              e.preventDefault();
              e.dataTransfer.dropEffect = "move";
              const rect = (e.currentTarget as HTMLElement).getBoundingClientRect();
              const before = e.clientY < rect.top + rect.height / 2;
              if (dropTarget?.index !== i || dropTarget.before !== before) {
                setDropTarget({ index: i, before });
              }
            }}
            onDrop={(e) => {
              if (!draggable) return;
              e.preventDefault();
              const before = dropTarget?.before ?? false;
              handleDrop(i, before);
            }}
            onContextMenu={(e) => {
              if (!onContextAction) return;
              e.preventDefault();
              setMenu({ slideIndex: i, x: e.clientX, y: e.clientY });
            }}
            style={{
              display: "flex",
              alignItems: "center",
              gap: 8,
              opacity: isDragging ? 0.4 : 1,
              cursor: draggable ? "grab" : undefined,
              position: "relative",
              borderTop: showInsertBefore
                ? "2px solid var(--officeai-accent, #2563eb)"
                : "2px solid transparent",
              borderBottom: showInsertAfter
                ? "2px solid var(--officeai-accent, #2563eb)"
                : "2px solid transparent",
              transition: "border-color 80ms ease",
            }}
          >
            <span style={{ fontSize: 12, color: "#71717a", width: 18, textAlign: "right" }}>{i + 1}</span>
            <SlideThumbnail
              slide={s}
              slideSize={slideSize}
              mediaUrls={mediaUrls}
              theme={theme}
              charts={charts}
              width={thumbnailWidth}
              active={isActive}
              onClick={() => onSelect(i)}
              label={`Slide ${i + 1}`}
            />
          </li>
        );
      })}
      {menu && onContextAction ? (
        <SlideContextMenu
          state={menu}
          slideCount={slides.length}
          onPick={(action) => {
            const target = menu.slideIndex;
            setMenu(null);
            onContextAction(target, action);
          }}
        />
      ) : null}
    </ul>
  );
}

interface SlideContextMenuProps {
  readonly state: MenuState;
  readonly slideCount: number;
  readonly onPick: (action: SlideContextAction) => void;
}

function SlideContextMenu({ state, slideCount, onPick }: SlideContextMenuProps): React.ReactElement {
  const items: ReadonlyArray<{
    readonly action: SlideContextAction;
    readonly label: string;
    readonly enabled: boolean;
    readonly group?: number;
  }> = [
    { action: "insert-before", label: "New slide before", enabled: true, group: 0 },
    { action: "insert-after", label: "New slide after", enabled: true, group: 0 },
    { action: "duplicate", label: "Duplicate slide", enabled: true, group: 1 },
    {
      action: "move-up",
      label: "Move up",
      enabled: state.slideIndex > 0,
      group: 2,
    },
    {
      action: "move-down",
      label: "Move down",
      enabled: state.slideIndex < slideCount - 1,
      group: 2,
    },
    {
      action: "move-to-start",
      label: "Move to start",
      enabled: state.slideIndex > 0,
      group: 2,
    },
    {
      action: "move-to-end",
      label: "Move to end",
      enabled: state.slideIndex < slideCount - 1,
      group: 2,
    },
    {
      action: "delete",
      label: "Delete slide",
      enabled: slideCount > 1,
      group: 3,
    },
  ];
  return (
    <div
      role="menu"
      data-testid="pptx-slide-context-menu"
      onMouseDown={(e) => e.stopPropagation()}
      style={{
        position: "fixed",
        top: state.y,
        left: state.x,
        zIndex: 60,
        minWidth: 200,
        padding: 4,
        borderRadius: 6,
        border: "1px solid var(--officeai-divider, #e5e7eb)",
        background: "var(--officeai-surface, #ffffff)",
        boxShadow: "0 8px 24px rgba(0,0,0,0.12)",
        display: "flex",
        flexDirection: "column",
        gap: 2,
      }}
    >
      {items.map((item, idx) => {
        const prev = idx > 0 ? items[idx - 1] : undefined;
        const showSep = prev && prev.group !== item.group;
        return (
          <React.Fragment key={item.action}>
            {showSep ? (
              <div
                role="separator"
                style={{ height: 1, background: "var(--officeai-divider, #e5e7eb)", margin: "2px 4px" }}
              />
            ) : null}
            <button
              type="button"
              role="menuitem"
              data-testid={`pptx-slide-menu-${item.action}`}
              disabled={!item.enabled}
              onClick={() => onPick(item.action)}
              style={{
                appearance: "none",
                border: "none",
                background: "transparent",
                textAlign: "left",
                padding: "6px 8px",
                fontSize: 12,
                borderRadius: 4,
                cursor: item.enabled ? "pointer" : "not-allowed",
                opacity: item.enabled ? 1 : 0.5,
                color: "var(--officeai-foreground, #111827)",
              }}
              onMouseEnter={(e) => {
                if (item.enabled) {
                  (e.currentTarget as HTMLButtonElement).style.background = "var(--officeai-hover, #f3f4f6)";
                }
              }}
              onMouseLeave={(e) => {
                (e.currentTarget as HTMLButtonElement).style.background = "transparent";
              }}
            >
              {item.label}
            </button>
          </React.Fragment>
        );
      })}
    </div>
  );
}
