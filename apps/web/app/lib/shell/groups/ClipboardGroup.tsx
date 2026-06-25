"use client";

/**
 * Shared "Clipboard" ribbon group used by every editor (DOCX, XLSX,
 * PPTX). Mirrors Office's leftmost group on the Start tab:
 *
 *   ┌──────────────────────┐
 *   │ Paste │ Cut Copy FP   │   ← Format Painter (FP) toggle
 *   └──────────────────────┘
 *
 * The group is presentational — each editor passes a typed callback
 * bag. Cut/Copy/Paste fall back to the browser-native clipboard
 * APIs when the editor doesn't override them; Format Painter is
 * opt-in (only XLSX has it today, the others can wire it later).
 *
 * Format Painter UX:
 *   - single click   → arm one-shot mode
 *   - double click   → arm sticky mode (active until cleared)
 *   - active state   → aria-pressed + accent fill
 */

import type { ReactNode } from "react";
import { Clipboard, ClipboardCopy, ClipboardPaste, Paintbrush, Scissors } from "@officeai/ui/sonaloop-icons";
import { cn } from "@officeai/ui";

export interface ClipboardGroupProps {
  readonly disabled?: boolean;
  readonly onCut?: () => void;
  readonly onCopy?: () => void;
  readonly onPaste?: () => void;
  readonly onPasteSpecial?: () => void;
  readonly canPaste?: boolean;
  readonly formatPainter?: {
    readonly active: boolean;
    readonly onActivate: (sticky: boolean) => void;
  };
  /** Test-id prefix, e.g. `"docx"` produces `docx-clipboard-cut`. */
  readonly testIdPrefix: string;
}

export function ClipboardGroup(props: ClipboardGroupProps): ReactNode {
  const {
    disabled = false,
    onCut,
    onCopy,
    onPaste,
    onPasteSpecial,
    canPaste = true,
    formatPainter,
    testIdPrefix,
  } = props;

  return (
    <div className="flex items-center gap-0.5">
      {/* Big paste button on the left — Office convention */}
      {onPaste ? (
        <button
          type="button"
          data-testid={`${testIdPrefix}-clipboard-paste`}
          title="Paste (⌘V)"
          aria-label="Paste"
          disabled={disabled || !canPaste}
          onMouseDown={(e) => e.preventDefault()}
          onClick={onPaste}
          onDoubleClick={onPasteSpecial}
          className="inline-flex h-8 w-8 items-center justify-center rounded text-foreground hover:bg-hover disabled:opacity-40"
        >
          <ClipboardPaste size={16} />
        </button>
      ) : null}
      <div className="flex flex-col items-center gap-0.5">
        {onCut ? (
          <ClipboardMiniBtn
            icon={<Scissors size={11} />}
            label="Cut (⌘X)"
            testId={`${testIdPrefix}-clipboard-cut`}
            disabled={disabled}
            onClick={onCut}
          />
        ) : null}
        {onCopy ? (
          <ClipboardMiniBtn
            icon={<ClipboardCopy size={11} />}
            label="Copy (⌘C)"
            testId={`${testIdPrefix}-clipboard-copy`}
            disabled={disabled}
            onClick={onCopy}
          />
        ) : null}
        {formatPainter ? (
          <button
            type="button"
            data-testid={`${testIdPrefix}-clipboard-format-painter`}
            title="Format Painter (double-click for sticky)"
            aria-label="Format Painter"
            aria-pressed={formatPainter.active}
            disabled={disabled}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => formatPainter.onActivate(false)}
            onDoubleClick={() => formatPainter.onActivate(true)}
            className={cn(
              "inline-flex h-4 w-4 items-center justify-center rounded text-secondary hover:bg-hover hover:text-foreground disabled:opacity-40",
              formatPainter.active && "bg-accent-soft text-accent"
            )}
          >
            <Paintbrush size={11} />
          </button>
        ) : null}
        {!onCut && !onCopy && !formatPainter ? (
          <span className="inline-flex h-4 w-4 items-center justify-center text-tertiary/40">
            <Clipboard size={11} />
          </span>
        ) : null}
      </div>
    </div>
  );
}

interface ClipboardMiniBtnProps {
  readonly icon: ReactNode;
  readonly label: string;
  readonly testId: string;
  readonly disabled: boolean;
  readonly onClick: () => void;
}

function ClipboardMiniBtn(props: ClipboardMiniBtnProps): ReactNode {
  const { icon, label, testId, disabled, onClick } = props;
  return (
    <button
      type="button"
      data-testid={testId}
      title={label}
      aria-label={label}
      disabled={disabled}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
      className="inline-flex h-4 w-4 items-center justify-center rounded text-secondary hover:bg-hover hover:text-foreground disabled:opacity-40"
    >
      {icon}
    </button>
  );
}
