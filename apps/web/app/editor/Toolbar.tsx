"use client";

import { useEffect, useRef, useState, type ReactNode } from "react";
import {
  Bold,
  Italic,
  Underline,
  Strikethrough,
  Download,
  FileUp,
  MessageSquarePlus,
  AlignLeft,
  AlignCenter,
  AlignRight,
  AlignJustify,
  Indent,
  Outdent,
  List,
  ListOrdered,
  ChevronDown,
  Palette,
  Highlighter,
} from "lucide-react";
import { Button, cn } from "@officeai/ui";
import { COLOR_PALETTE, FONT_SIZES, HIGHLIGHT_PALETTE, PARAGRAPH_STYLES } from "@/lib/format-helpers";
import type { TextFormat } from "@officeai/docx";

export interface ToolbarProps {
  agentReady: boolean;
  docInfo: { paragraphs: number; revision: number; commentThreads: number } | null;
  activeStyle: string;
  activeMarks: ReadonlySet<string>;
  onOpenFile: () => void;
  onExport: () => void;
  onSetParagraphStyle: (style: string) => void;
  onApplyFormat: (format: TextFormat) => void;
  onToggleMark: (mark: "bold" | "italic" | "underline" | "strike") => void;
  onAddComment: () => void;
  /**
   * Surface "not yet supported" toasts for buttons whose backing command
   * does not yet exist (alignment, indent, lists). The brief mandates we
   * never silently no-op these.
   */
  onUnsupported: (label: string) => void;
}

/**
 * Editor toolbar. Layout is intentionally close to Word's: file/style on
 * the left, inline marks + color in the middle, alignment / indent / list
 * on the right, and the export action pinned to the far right with the
 * doc metadata strip (kept verbatim so the existing P1.1 e2e specs that
 * read "{N} blocks · rev {R} · {C} comments" continue to pass).
 *
 * The toolbar wraps with `flex-wrap` so it stays usable down to 360px;
 * the metadata strip hides under `md` and reappears alongside the export
 * button on larger viewports.
 */
export function Toolbar(props: ToolbarProps): ReactNode {
  return (
    <div
      role="toolbar"
      aria-label="Document toolbar"
      className="editor-toolbar flex flex-wrap items-center gap-1 border-b border-divider pb-3"
    >
      {/* File group */}
      <button
        type="button"
        onClick={props.onOpenFile}
        className="inline-flex items-center gap-1.5 rounded-md border border-divider bg-surface px-2.5 py-1 text-xs text-foreground hover:bg-hover"
      >
        <FileUp size={14} />
        Open .docx
      </button>

      <Divider />

      {/* Paragraph style */}
      <ParagraphStylePicker
        value={props.activeStyle}
        onChange={props.onSetParagraphStyle}
        disabled={!props.agentReady}
      />

      {/* Font size */}
      <FontSizePicker
        onChange={(halfPoints) => props.onApplyFormat({ fontSize: halfPoints })}
        disabled={!props.agentReady}
      />

      <Divider />

      {/* Inline marks */}
      <ToolbarBtn
        label="Bold"
        active={props.activeMarks.has("bold")}
        onClick={() => props.onToggleMark("bold")}
      >
        <Bold size={14} />
      </ToolbarBtn>
      <ToolbarBtn
        label="Italic"
        active={props.activeMarks.has("italic")}
        onClick={() => props.onToggleMark("italic")}
      >
        <Italic size={14} />
      </ToolbarBtn>
      <ToolbarBtn
        label="Underline"
        active={props.activeMarks.has("underline")}
        onClick={() => props.onToggleMark("underline")}
      >
        <Underline size={14} />
      </ToolbarBtn>
      <ToolbarBtn
        label="Strike"
        active={props.activeMarks.has("strikethrough")}
        onClick={() => props.onToggleMark("strike")}
      >
        <Strikethrough size={14} />
      </ToolbarBtn>

      <Divider />

      {/* Color + highlight */}
      <ColorPicker
        label="Font color"
        icon={<Palette size={14} />}
        items={COLOR_PALETTE.map((c) => ({ value: c.hex, label: c.name, swatch: `#${c.hex}` }))}
        onPick={(hex) => props.onApplyFormat({ color: hex })}
        disabled={!props.agentReady}
      />
      <ColorPicker
        label="Highlight"
        icon={<Highlighter size={14} />}
        items={HIGHLIGHT_PALETTE.map((h) => ({ value: h.name, label: h.label, swatch: h.swatch }))}
        onPick={(name) => props.onApplyFormat({ highlight: name })}
        disabled={!props.agentReady}
      />

      <Divider />

      {/* Alignment — backing command not in P1.2; surface graceful toast */}
      <ToolbarBtn label="Align left" onClick={() => props.onUnsupported("alignment")}>
        <AlignLeft size={14} />
      </ToolbarBtn>
      <ToolbarBtn label="Align center" onClick={() => props.onUnsupported("alignment")}>
        <AlignCenter size={14} />
      </ToolbarBtn>
      <ToolbarBtn label="Align right" onClick={() => props.onUnsupported("alignment")}>
        <AlignRight size={14} />
      </ToolbarBtn>
      <ToolbarBtn label="Align justify" onClick={() => props.onUnsupported("alignment")}>
        <AlignJustify size={14} />
      </ToolbarBtn>

      <Divider />

      {/* Indentation — backing command not in P1.2 */}
      <ToolbarBtn label="Decrease indent" onClick={() => props.onUnsupported("indentation")}>
        <Outdent size={14} />
      </ToolbarBtn>
      <ToolbarBtn label="Increase indent" onClick={() => props.onUnsupported("indentation")}>
        <Indent size={14} />
      </ToolbarBtn>

      <Divider />

      {/* Lists — backing command not in P1.2 */}
      <ToolbarBtn label="Bullet list" onClick={() => props.onUnsupported("bullet list")}>
        <List size={14} />
      </ToolbarBtn>
      <ToolbarBtn label="Numbered list" onClick={() => props.onUnsupported("numbered list")}>
        <ListOrdered size={14} />
      </ToolbarBtn>

      <Divider />

      {/* Comment */}
      <ToolbarBtn label="Add comment" onClick={props.onAddComment}>
        <MessageSquarePlus size={14} />
      </ToolbarBtn>

      <div className="ml-auto flex items-center gap-3 text-xs text-secondary">
        {props.docInfo && (
          <span className="hidden whitespace-nowrap md:inline">
            {props.docInfo.paragraphs} paragraphs · rev {props.docInfo.revision} ·{" "}
            {props.docInfo.commentThreads} comment{props.docInfo.commentThreads === 1 ? "" : "s"}
          </span>
        )}
        <Button variant="accent" size="sm" onClick={props.onExport}>
          <Download size={14} />
          Export
        </Button>
      </div>
    </div>
  );
}

function Divider(): ReactNode {
  return <div className="mx-1 h-4 w-px bg-divider" aria-hidden />;
}

function ToolbarBtn(props: {
  label: string;
  onClick: () => void;
  active?: boolean;
  children: ReactNode;
}): ReactNode {
  return (
    <button
      type="button"
      title={props.label}
      aria-label={props.label}
      aria-pressed={props.active ?? undefined}
      onClick={props.onClick}
      className={cn(
        "rounded-md p-1.5 text-secondary hover:bg-hover hover:text-foreground",
        props.active && "bg-accent-light text-foreground"
      )}
    >
      {props.children}
    </button>
  );
}

function ParagraphStylePicker(props: {
  value: string;
  onChange: (v: string) => void;
  disabled?: boolean;
}): ReactNode {
  return (
    <label className="inline-flex items-center gap-1 text-xs text-secondary">
      <span className="sr-only">Paragraph style</span>
      <select
        title="Paragraph style"
        aria-label="Paragraph style"
        value={props.value}
        disabled={props.disabled}
        onChange={(e) => props.onChange(e.target.value)}
        className="h-7 rounded-md border border-divider bg-surface px-2 text-xs text-foreground hover:bg-hover focus:outline-none"
      >
        {PARAGRAPH_STYLES.map((s) => (
          <option key={s.value} value={s.value}>
            {s.label}
          </option>
        ))}
      </select>
    </label>
  );
}

function FontSizePicker(props: { onChange: (halfPoints: number) => void; disabled?: boolean }): ReactNode {
  return (
    <label className="inline-flex items-center gap-1 text-xs text-secondary">
      <span className="sr-only">Font size</span>
      <select
        title="Font size"
        aria-label="Font size"
        defaultValue=""
        disabled={props.disabled}
        onChange={(e) => {
          const n = Number(e.target.value);
          if (!Number.isFinite(n) || n <= 0) return;
          props.onChange(n);
          e.target.value = "";
        }}
        className="h-7 w-16 rounded-md border border-divider bg-surface px-2 text-xs text-foreground hover:bg-hover focus:outline-none"
      >
        <option value="" disabled>
          Size
        </option>
        {FONT_SIZES.map((halfPts) => (
          <option key={halfPts} value={halfPts}>
            {halfPts / 2}
          </option>
        ))}
      </select>
    </label>
  );
}

interface ColorItem {
  value: string;
  label: string;
  swatch: string;
}

function ColorPicker(props: {
  label: string;
  icon: ReactNode;
  items: ReadonlyArray<ColorItem>;
  onPick: (value: string) => void;
  disabled?: boolean;
}): ReactNode {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    if (!open) return;
    const onDoc = (e: MouseEvent) => {
      if (!ref.current) return;
      if (!ref.current.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, [open]);

  return (
    <div ref={ref} className="relative inline-flex">
      <button
        type="button"
        title={props.label}
        aria-label={props.label}
        aria-haspopup="menu"
        aria-expanded={open}
        disabled={props.disabled}
        onClick={() => setOpen((v) => !v)}
        className="flex items-center gap-0.5 rounded-md p-1.5 text-secondary hover:bg-hover hover:text-foreground"
      >
        {props.icon}
        <ChevronDown size={10} />
      </button>
      {open && (
        <div
          role="menu"
          className="absolute left-0 top-full z-30 mt-1 grid w-44 grid-cols-4 gap-1 rounded-md border border-divider bg-surface p-2 shadow-md"
        >
          {props.items.map((item) => (
            <button
              key={item.value}
              type="button"
              role="menuitem"
              title={item.label}
              aria-label={`${props.label}: ${item.label}`}
              onClick={() => {
                props.onPick(item.value);
                setOpen(false);
              }}
              className="flex h-7 w-7 items-center justify-center rounded border border-divider hover:scale-110"
              style={{ background: item.swatch }}
            />
          ))}
        </div>
      )}
    </div>
  );
}
