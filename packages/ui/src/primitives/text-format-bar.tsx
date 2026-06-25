"use client";

import { type ReactNode } from "react";
import { Bold, Italic, Underline, Strikethrough, Palette, Highlighter } from "../sonaloop-icons";
import { isOnTruthy, type ActiveTextFormat, type TextFormatProvider } from "@officeai/text-formatting";
import { ColorPicker } from "./color-picker";
import { FontFamilyPicker } from "./font-family-picker";
import { FontSizePicker } from "./font-size-picker";
import { FormatToggle } from "./format-toggle";
import { HighlightPicker } from "./highlight-picker";

export interface TextFormatBarProps {
  provider: TextFormatProvider;
  /**
   * Pre-computed active format. Caller is expected to recompute
   * this on every render via `provider.getActive()` so React
   * rerenders pick it up — passing it as a prop avoids the bar
   * having to subscribe to format-specific state.
   */
  active: ActiveTextFormat;
  /** Disable everything (e.g. agent not ready, no document loaded). */
  disabled?: boolean;
  /** Optional `data-testid` prefix; children get `${prefix}-bold`, etc. */
  testIdPrefix?: string;
}

/**
 * The shared text-formatting bar embedded in every editor's toolbar.
 * Hides controls that the provider's `capabilities` flag as
 * unsupported.
 */
export function TextFormatBar({
  provider,
  active,
  disabled = false,
  testIdPrefix,
}: TextFormatBarProps): ReactNode {
  const caps = provider.capabilities;
  const tid = (suffix: string) => (testIdPrefix ? `${testIdPrefix}-${suffix}` : undefined);

  const isDisabled = disabled || !provider.hasSelection();

  return (
    <div className="inline-flex flex-wrap items-center gap-1">
      {caps.fontFamily && (
        <FontFamilyPicker
          value={active.fontFamily}
          onChange={(family) => provider.apply({ fontFamily: family })}
          disabled={isDisabled}
        />
      )}
      {caps.fontSize && (
        <FontSizePicker
          value={active.fontSizePt}
          onChange={(pt) => provider.apply({ fontSizePt: pt })}
          disabled={isDisabled}
        />
      )}
      {(caps.fontFamily || caps.fontSize) && <Divider />}
      <FormatToggle
        label="Bold"
        value={active.bold}
        disabled={isDisabled}
        onClick={() => provider.apply({ bold: !isOnTruthy(active.bold) })}
        data-testid={tid("bold")}
      >
        <Bold size={14} />
      </FormatToggle>
      <FormatToggle
        label="Italic"
        value={active.italic}
        disabled={isDisabled}
        onClick={() => provider.apply({ italic: !isOnTruthy(active.italic) })}
        data-testid={tid("italic")}
      >
        <Italic size={14} />
      </FormatToggle>
      <FormatToggle
        label="Underline"
        value={active.underline as never}
        disabled={isDisabled}
        onClick={() =>
          provider.apply({
            underline: !(active.underline === true || typeof active.underline === "string"),
          })
        }
        data-testid={tid("underline")}
      >
        <Underline size={14} />
      </FormatToggle>
      {caps.strike && (
        <FormatToggle
          label="Strike"
          value={active.strike}
          disabled={isDisabled}
          onClick={() => provider.apply({ strike: !isOnTruthy(active.strike) })}
          data-testid={tid("strike")}
        >
          <Strikethrough size={14} />
        </FormatToggle>
      )}
      <Divider />
      <ColorPicker
        label="Font color"
        icon={<Palette size={14} />}
        value={active.color}
        onChange={(rrggbb) => provider.apply({ color: rrggbb })}
        onClear={() => provider.apply({ color: "" })}
        disabled={isDisabled}
      />
      {caps.highlight !== "none" && (
        <HighlightPicker
          label={caps.highlight === "fill-fallback" ? "Cell fill" : "Highlight"}
          icon={<Highlighter size={14} />}
          value={active.highlight}
          onChange={(rrggbb) => provider.apply({ highlight: rrggbb })}
          onClear={() => provider.apply({ highlight: "" })}
          disabled={isDisabled}
        />
      )}
    </div>
  );
}

function Divider(): ReactNode {
  return <span aria-hidden className="mx-1 h-5 w-px bg-divider" />;
}
