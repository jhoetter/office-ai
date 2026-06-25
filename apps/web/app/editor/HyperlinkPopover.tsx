"use client";

import { useEffect, useRef, useState, type ReactNode } from "react";
import { Link as LinkIcon, Trash2, X } from "@officeai/ui/sonaloop-icons";
import { Button } from "@officeai/ui";

/**
 * B5 — Hyperlink popover.
 *
 * Replaces the previous `window.prompt` flow for inserting links.
 * Renders a Word-style anchored popover with two fields (URL and
 * optional display text) plus actions Apply / Remove / Cancel.
 *
 * Pure presentational: anchor coords + initial values come from the
 * parent so this component never reads PM state directly. The parent
 * funnels Apply / Remove through the existing
 * `docx:insert-hyperlink` / `docx:remove-hyperlink` commands so the
 * model stays the source of truth and the OOXML round-trip is
 * unaffected.
 */

export interface HyperlinkPopoverProps {
  readonly anchor: { left: number; top: number; bottom: number } | null;
  readonly initialUrl: string;
  readonly initialText: string;
  /** When set, an existing hyperlink id is being edited (Remove enabled). */
  readonly existingHyperlinkId: string | null;
  readonly onApply: (next: { url: string; text: string }) => void;
  readonly onRemove?: () => void;
  readonly onCancel: () => void;
}

export function HyperlinkPopover(props: HyperlinkPopoverProps): ReactNode {
  const [url, setUrl] = useState(props.initialUrl);
  const [text, setText] = useState(props.initialText);
  const urlRef = useRef<HTMLInputElement | null>(null);
  const valid = isLikelyUrl(url);

  useEffect(() => {
    setUrl(props.initialUrl);
    setText(props.initialText);
    requestAnimationFrame(() => urlRef.current?.focus());
  }, [props.initialUrl, props.initialText]);

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        e.preventDefault();
        props.onCancel();
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [props]);

  const submit = () => {
    if (!valid) return;
    props.onApply({ url: url.trim(), text: text.trim() });
  };

  const style: React.CSSProperties = props.anchor
    ? {
        position: "fixed",
        left: Math.max(
          8,
          Math.min(props.anchor.left, (typeof window !== "undefined" ? window.innerWidth : 1024) - 360)
        ),
        top: props.anchor.bottom + 6,
      }
    : {
        position: "fixed",
        left: "50%",
        top: "30%",
        transform: "translate(-50%,-30%)",
      };

  return (
    <div
      data-testid="hyperlink-popover"
      role="dialog"
      aria-label="Insert hyperlink"
      className="z-50 flex w-[340px] flex-col gap-2 rounded-md border border-divider bg-surface p-3 shadow-lg"
      style={style}
      onClick={(e) => e.stopPropagation()}
    >
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-1.5 text-xs font-medium uppercase tracking-wide text-secondary">
          <LinkIcon size={12} />
          {props.existingHyperlinkId ? "Edit link" : "Insert link"}
        </div>
        <button
          type="button"
          aria-label="Cancel"
          onClick={props.onCancel}
          className="rounded p-1 text-secondary hover:bg-divider/50"
        >
          <X size={12} />
        </button>
      </div>

      <label className="flex flex-col gap-1 text-xs">
        <span className="text-secondary">Address</span>
        <input
          ref={urlRef}
          type="url"
          inputMode="url"
          spellCheck={false}
          autoComplete="off"
          value={url}
          onChange={(e) => setUrl(e.currentTarget.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter") {
              e.preventDefault();
              submit();
            }
          }}
          placeholder="https://example.com"
          className="w-full rounded border border-divider bg-background px-2 py-1.5 text-xs text-foreground placeholder:text-tertiary focus:border-accent focus:outline-none"
          data-testid="hyperlink-url-input"
        />
      </label>

      <label className="flex flex-col gap-1 text-xs">
        <span className="text-secondary">Display text</span>
        <input
          type="text"
          value={text}
          onChange={(e) => setText(e.currentTarget.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter") {
              e.preventDefault();
              submit();
            }
          }}
          placeholder="Optional — uses selection if blank"
          className="w-full rounded border border-divider bg-background px-2 py-1.5 text-xs text-foreground placeholder:text-tertiary focus:border-accent focus:outline-none"
          data-testid="hyperlink-text-input"
        />
      </label>

      <div className="mt-1 flex items-center justify-between gap-1.5">
        {props.existingHyperlinkId && props.onRemove ? (
          <button
            type="button"
            onClick={props.onRemove}
            className="inline-flex items-center gap-1 rounded border border-divider px-2 py-1 text-xs text-error hover:bg-hover"
            data-testid="hyperlink-remove"
          >
            <Trash2 size={12} /> Remove
          </button>
        ) : (
          <span />
        )}
        <div className="flex items-center gap-1.5">
          <Button variant="ghost" size="sm" onClick={props.onCancel}>
            Cancel
          </Button>
          <Button variant="accent" size="sm" onClick={submit} disabled={!valid} data-testid="hyperlink-apply">
            {props.existingHyperlinkId ? "Update" : "Insert"}
          </Button>
        </div>
      </div>
      {!valid && url.length > 0 && (
        <p className="text-[11px] text-error" role="status">
          Enter a URL like https://… or mailto:…
        </p>
      )}
    </div>
  );
}

function isLikelyUrl(s: string): boolean {
  const v = s.trim();
  if (v.length === 0) return false;
  if (/^https?:\/\//i.test(v)) return true;
  if (/^mailto:/i.test(v)) return true;
  if (/^tel:/i.test(v)) return true;
  if (/^#/.test(v)) return true; // bookmark/anchor
  // Bare domain like "example.com/path" — treated as https://.
  if (/^[a-z0-9.-]+\.[a-z]{2,}(\/.*)?$/i.test(v)) return true;
  return false;
}
