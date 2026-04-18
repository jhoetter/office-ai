"use client";

import type { ReactNode } from "react";
import { useCallback, useMemo } from "react";
import {
  Bold,
  Italic,
  Underline,
  Strikethrough,
  AlignLeft,
  AlignCenter,
  AlignRight,
  Paintbrush,
  Type as TypeIcon,
} from "lucide-react";
import { cn } from "@officeai/ui";
import type { CellFormatPatch, EffectiveStyle, StyleTable } from "@officeai/xlsx";
import { flattenCellXf } from "@officeai/xlsx";
import { NUMBER_FORMAT_PRESETS, type NumberFormatPresetId } from "./styles";
import type { Selection } from "./selection";

export interface ToolbarProps {
  readonly disabled: boolean;
  /**
   * The active anchor cell's effective style, used to flip toggle
   * states (Bold pressed, etc.) so the toolbar reflects what's
   * actually on the cell.
   */
  readonly anchorStyleId: number | undefined;
  readonly styles: StyleTable;
  readonly selection: Selection | null;
  /**
   * Dispatch a `xlsx:set-cell-format` over the current selection. The
   * parent owns the agent + sheet name so this stays a pure
   * presentational component.
   */
  readonly onApply: (patch: CellFormatPatch) => void;
}

/**
 * Excel-style formatting toolbar — Bold / Italic / Underline /
 * Strike, Font colour, Fill colour, Horizontal alignment, and a
 * Number-format dropdown. Every action funnels through the parent's
 * `onApply` (which dispatches `xlsx:set-cell-format` on the active
 * range) so the agent path and the human path share the same
 * command surface.
 *
 * Toggle state for the bold/italic/etc. buttons is derived from the
 * *anchor* cell's effective style (matches Excel — the active cell
 * decides what the ribbon looks like).
 */
export function Toolbar(props: ToolbarProps): ReactNode {
  const { disabled, anchorStyleId, styles, selection, onApply } = props;

  const effective: EffectiveStyle = useMemo(
    () => flattenCellXf(styles, anchorStyleId),
    [styles, anchorStyleId]
  );

  const apply = useCallback(
    (patch: CellFormatPatch) => {
      if (!selection) return;
      onApply(patch);
    },
    [onApply, selection]
  );

  const onToggle = useCallback(
    (key: "bold" | "italic" | "underline" | "strike") => {
      const current = !!effective.font[key];
      apply({ font: { [key]: !current } });
    },
    [apply, effective]
  );

  const onAlign = useCallback(
    (h: "left" | "center" | "right") => {
      apply({ alignment: { horizontal: h } });
    },
    [apply]
  );

  const onFontColor = useCallback(
    (rrggbb: string) => {
      apply({ font: { color: rrggbb } });
    },
    [apply]
  );

  const onFillColor = useCallback(
    (rrggbb: string) => {
      apply({ fill: { color: rrggbb, pattern: "solid" } });
    },
    [apply]
  );

  const onNumberFormat = useCallback(
    (id: NumberFormatPresetId) => {
      const code = NUMBER_FORMAT_PRESETS.find((p) => p.id === id)?.code;
      if (!code) return;
      apply({ numberFormat: code });
    },
    [apply]
  );

  return (
    <div
      data-testid="xlsx-toolbar"
      className="flex flex-wrap items-center gap-1 rounded-md border border-divider bg-surface px-2 py-1.5"
    >
      <ToggleBtn
        icon={<Bold size={14} />}
        label="Bold"
        testId="format-bold"
        active={!!effective.font.bold}
        disabled={disabled}
        onClick={() => onToggle("bold")}
      />
      <ToggleBtn
        icon={<Italic size={14} />}
        label="Italic"
        testId="format-italic"
        active={!!effective.font.italic}
        disabled={disabled}
        onClick={() => onToggle("italic")}
      />
      <ToggleBtn
        icon={<Underline size={14} />}
        label="Underline"
        testId="format-underline"
        active={!!effective.font.underline}
        disabled={disabled}
        onClick={() => onToggle("underline")}
      />
      <ToggleBtn
        icon={<Strikethrough size={14} />}
        label="Strikethrough"
        testId="format-strike"
        active={!!effective.font.strike}
        disabled={disabled}
        onClick={() => onToggle("strike")}
      />

      <div className="mx-1 h-5 w-px bg-divider" aria-hidden />

      <ColorPicker
        icon={<TypeIcon size={14} />}
        title="Font color"
        testId="format-font-color"
        disabled={disabled}
        onPick={onFontColor}
      />
      <ColorPicker
        icon={<Paintbrush size={14} />}
        title="Fill color"
        testId="format-fill-color"
        disabled={disabled}
        onPick={onFillColor}
      />

      <div className="mx-1 h-5 w-px bg-divider" aria-hidden />

      <ToggleBtn
        icon={<AlignLeft size={14} />}
        label="Align left"
        testId="format-align-left"
        active={effective.alignment?.horizontal === "left"}
        disabled={disabled}
        onClick={() => onAlign("left")}
      />
      <ToggleBtn
        icon={<AlignCenter size={14} />}
        label="Align center"
        testId="format-align-center"
        active={effective.alignment?.horizontal === "center"}
        disabled={disabled}
        onClick={() => onAlign("center")}
      />
      <ToggleBtn
        icon={<AlignRight size={14} />}
        label="Align right"
        testId="format-align-right"
        active={effective.alignment?.horizontal === "right"}
        disabled={disabled}
        onClick={() => onAlign("right")}
      />

      <div className="mx-1 h-5 w-px bg-divider" aria-hidden />

      <select
        data-testid="format-number"
        aria-label="Number format"
        disabled={disabled}
        defaultValue=""
        onChange={(e) => {
          const v = e.target.value;
          if (v) onNumberFormat(v as NumberFormatPresetId);
          e.currentTarget.value = "";
        }}
        className="h-7 rounded border border-divider bg-background px-1 text-xs text-foreground disabled:opacity-50"
      >
        <option value="" disabled>
          Format…
        </option>
        {NUMBER_FORMAT_PRESETS.map((p) => (
          <option key={p.id} value={p.id}>
            {p.label}
          </option>
        ))}
      </select>
    </div>
  );
}

interface ToggleBtnProps {
  readonly icon: ReactNode;
  readonly label: string;
  readonly testId: string;
  readonly active: boolean;
  readonly disabled: boolean;
  readonly onClick: () => void;
}

function ToggleBtn(props: ToggleBtnProps): ReactNode {
  const { icon, label, testId, active, disabled, onClick } = props;
  return (
    <button
      type="button"
      data-testid={testId}
      title={label}
      aria-label={label}
      aria-pressed={active}
      disabled={disabled}
      onMouseDown={(e) => {
        // Don't steal focus from the surface — toolbar clicks should
        // leave the active cell selection (and therefore the next
        // type-to-edit target) intact.
        e.preventDefault();
      }}
      onClick={onClick}
      className={cn(
        "inline-flex h-7 w-7 items-center justify-center rounded text-foreground hover:bg-hover disabled:opacity-50",
        active && "bg-[var(--ai-violet-light)] text-[var(--ai-violet)]"
      )}
    >
      {icon}
    </button>
  );
}

interface ColorPickerProps {
  readonly icon: ReactNode;
  readonly title: string;
  readonly testId: string;
  readonly disabled: boolean;
  readonly onPick: (rrggbb: string) => void;
}

const COLOR_SWATCHES: ReadonlyArray<{ id: string; rrggbb: string }> = [
  { id: "black", rrggbb: "000000" },
  { id: "white", rrggbb: "FFFFFF" },
  { id: "red", rrggbb: "DC2626" },
  { id: "orange", rrggbb: "EA580C" },
  { id: "yellow", rrggbb: "FACC15" },
  { id: "green", rrggbb: "16A34A" },
  { id: "blue", rrggbb: "2563EB" },
  { id: "violet", rrggbb: "7C3AED" },
];

/**
 * Tiny inline colour picker. Renders the trigger as a button with
 * the same visual rhythm as the toggles, then a small grid of
 * swatches in a `<details>` so we don't need a portal/popover stack
 * for the P0 toolbar. Each swatch carries `data-testid` so Playwright
 * can pick a deterministic colour.
 */
function ColorPicker(props: ColorPickerProps): ReactNode {
  const { icon, title, testId, disabled, onPick } = props;
  return (
    <details
      data-testid={testId}
      className="relative"
      onToggle={(e) => {
        // Auto-close after a swatch click — kept minimal.
        void e;
      }}
    >
      <summary
        title={title}
        aria-label={title}
        className={cn(
          "list-none inline-flex h-7 w-7 cursor-pointer items-center justify-center rounded text-foreground hover:bg-hover",
          disabled && "pointer-events-none opacity-50"
        )}
        onMouseDown={(e) => e.preventDefault()}
      >
        {icon}
      </summary>
      <div
        role="menu"
        className="absolute left-0 top-full z-50 mt-1 grid w-[160px] grid-cols-4 gap-1 rounded-md border border-divider bg-surface p-2 shadow-md"
      >
        {COLOR_SWATCHES.map((s) => (
          <button
            key={s.id}
            type="button"
            data-testid={`${testId}-${s.id}`}
            aria-label={`Color ${s.id}`}
            onMouseDown={(e) => e.preventDefault()}
            onClick={(e) => {
              onPick(s.rrggbb);
              const det = e.currentTarget.closest("details");
              if (det) det.removeAttribute("open");
            }}
            style={{ background: `#${s.rrggbb}` }}
            className="h-6 w-6 rounded border border-divider"
          />
        ))}
      </div>
    </details>
  );
}
