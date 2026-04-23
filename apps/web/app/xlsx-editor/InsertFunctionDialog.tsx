"use client";

/**
 * Excel-style "Insert Function" wizard.
 *
 * Two-step flow:
 *   1. Pick — searchable, category-filterable list of common Excel
 *      functions. Arrow keys + Enter to accept.
 *   2. Args — one input per declared argument (with `range`/`value`
 *      hint, optional flag, and a one-line description). The "Insert"
 *      button assembles `=NAME(arg1, arg2, …)` and forwards it via
 *      `onInsertFormula` so the parent can dispatch
 *      `xlsx:set-cell-formula` against the active cell. If
 *      `onInsertFormula` is omitted (or the user clicks "Skip args"),
 *      we fall back to the legacy seed-the-formula-bar behaviour
 *      via `onPick`.
 */

import { useEffect, useMemo, useRef, useState } from "react";
import type { ReactNode } from "react";
import { Search, X } from "lucide-react";
import { useFocusTrap } from "@officeai/ui";

export interface InsertFunctionDialogProps {
  readonly open: boolean;
  readonly onClose: () => void;
  /** Receives the chosen function name (e.g. "SUM"). Caller seeds
   * the formula bar with `=NAME(` and focuses it; used when the
   * user opts out of the per-argument step. */
  readonly onPick: (functionName: string) => void;
  /**
   * If provided, the wizard advances to a per-argument step after
   * the user picks a function and dispatches the assembled formula
   * (e.g. `=SUMIF(A1:A10, ">0", B1:B10)`) through this callback.
   * The caller is responsible for routing it to the active cell
   * via `xlsx:set-cell-formula`.
   */
  readonly onInsertFormula?: (formula: string) => void;
  /**
   * Cell-ref label of the current selection (e.g. `"B7"`). Surfaces
   * in the args step so users see where the formula will land.
   */
  readonly targetCellLabel?: string;
}

interface ArgSpec {
  readonly name: string;
  readonly kind: "range" | "value" | "criterion" | "text";
  readonly optional?: boolean;
  readonly hint?: string;
}

interface FunctionEntry {
  readonly name: string;
  readonly category: string;
  readonly summary: string;
  /**
   * Optional argument schema. When present and the parent passed
   * `onInsertFormula`, the wizard's step-2 surfaces a form. When
   * absent we still let the user proceed (single free-form
   * argument input).
   */
  readonly args?: ReadonlyArray<ArgSpec>;
}

const ARG_RANGE: ArgSpec = { name: "range", kind: "range", hint: "e.g. A1:A10" };
const ARG_NUMBER: ArgSpec = { name: "number", kind: "value", hint: "e.g. 42 or B2" };
const ARG_CRITERION: ArgSpec = { name: "criterion", kind: "criterion", hint: 'e.g. ">0" or "Apple"' };
const ARG_TEXT: ArgSpec = { name: "text", kind: "text", hint: 'e.g. "hello"' };

const CATALOGUE: ReadonlyArray<FunctionEntry> = [
  // Math & Trig
  { name: "SUM", category: "Math & Trig", summary: "Adds numbers in a range.", args: [ARG_RANGE] },
  { name: "AVERAGE", category: "Math & Trig", summary: "Returns the arithmetic mean.", args: [ARG_RANGE] },
  { name: "MIN", category: "Math & Trig", summary: "Returns the smallest value.", args: [ARG_RANGE] },
  { name: "MAX", category: "Math & Trig", summary: "Returns the largest value.", args: [ARG_RANGE] },
  {
    name: "ROUND",
    category: "Math & Trig",
    summary: "Rounds a number to N digits.",
    args: [
      { name: "number", kind: "value", hint: "e.g. A1" },
      { name: "num_digits", kind: "value", hint: "e.g. 2" },
    ],
  },
  { name: "ABS", category: "Math & Trig", summary: "Absolute value.", args: [ARG_NUMBER] },
  {
    name: "INT",
    category: "Math & Trig",
    summary: "Rounds down to the nearest integer.",
    args: [ARG_NUMBER],
  },
  {
    name: "MOD",
    category: "Math & Trig",
    summary: "Returns the remainder after division.",
    args: [
      { name: "number", kind: "value", hint: "dividend" },
      { name: "divisor", kind: "value", hint: "divisor" },
    ],
  },
  {
    name: "POWER",
    category: "Math & Trig",
    summary: "Returns a number raised to a power.",
    args: [
      { name: "number", kind: "value", hint: "base" },
      { name: "power", kind: "value", hint: "exponent" },
    ],
  },
  { name: "SQRT", category: "Math & Trig", summary: "Square root.", args: [ARG_NUMBER] },
  {
    name: "SUMIF",
    category: "Math & Trig",
    summary: "Sums cells matching a criterion.",
    args: [
      ARG_RANGE,
      ARG_CRITERION,
      { name: "sum_range", kind: "range", optional: true, hint: "(optional)" },
    ],
  },
  {
    name: "SUMIFS",
    category: "Math & Trig",
    summary: "Sums cells matching multiple criteria.",
    args: [
      { name: "sum_range", kind: "range", hint: "values to sum" },
      { name: "criteria_range1", kind: "range" },
      { name: "criteria1", kind: "criterion" },
      { name: "criteria_range2", kind: "range", optional: true },
      { name: "criteria2", kind: "criterion", optional: true },
    ],
  },
  // Statistical
  { name: "COUNT", category: "Statistical", summary: "Counts numeric cells.", args: [ARG_RANGE] },
  { name: "COUNTA", category: "Statistical", summary: "Counts non-empty cells.", args: [ARG_RANGE] },
  {
    name: "COUNTIF",
    category: "Statistical",
    summary: "Counts cells matching a criterion.",
    args: [ARG_RANGE, ARG_CRITERION],
  },
  {
    name: "COUNTIFS",
    category: "Statistical",
    summary: "Counts cells matching multiple criteria.",
    args: [
      { name: "criteria_range1", kind: "range" },
      { name: "criteria1", kind: "criterion" },
      { name: "criteria_range2", kind: "range", optional: true },
      { name: "criteria2", kind: "criterion", optional: true },
    ],
  },
  { name: "MEDIAN", category: "Statistical", summary: "Returns the median value.", args: [ARG_RANGE] },
  {
    name: "STDEV",
    category: "Statistical",
    summary: "Estimates standard deviation (sample).",
    args: [ARG_RANGE],
  },
  // Logical
  {
    name: "IF",
    category: "Logical",
    summary: "Conditional value.",
    args: [
      { name: "logical_test", kind: "value", hint: "e.g. A1>0" },
      { name: "value_if_true", kind: "value" },
      { name: "value_if_false", kind: "value", optional: true },
    ],
  },
  { name: "IFS", category: "Logical", summary: "Multiple conditions, returns first true value." },
  { name: "AND", category: "Logical", summary: "TRUE iff all arguments are TRUE." },
  { name: "OR", category: "Logical", summary: "TRUE iff at least one argument is TRUE." },
  {
    name: "NOT",
    category: "Logical",
    summary: "Logical negation.",
    args: [{ name: "logical", kind: "value" }],
  },
  {
    name: "IFERROR",
    category: "Logical",
    summary: "Catches errors and returns a fallback.",
    args: [
      { name: "value", kind: "value" },
      { name: "value_if_error", kind: "value" },
    ],
  },
  // Lookup & reference
  {
    name: "VLOOKUP",
    category: "Lookup",
    summary: "Vertical lookup in a table.",
    args: [
      { name: "lookup_value", kind: "value" },
      { name: "table_array", kind: "range" },
      { name: "col_index_num", kind: "value", hint: "1-based column" },
      { name: "range_lookup", kind: "value", optional: true, hint: "FALSE for exact" },
    ],
  },
  {
    name: "HLOOKUP",
    category: "Lookup",
    summary: "Horizontal lookup in a table.",
    args: [
      { name: "lookup_value", kind: "value" },
      { name: "table_array", kind: "range" },
      { name: "row_index_num", kind: "value", hint: "1-based row" },
      { name: "range_lookup", kind: "value", optional: true, hint: "FALSE for exact" },
    ],
  },
  {
    name: "INDEX",
    category: "Lookup",
    summary: "Returns a value at a position in a range.",
    args: [
      { name: "array", kind: "range" },
      { name: "row_num", kind: "value" },
      { name: "column_num", kind: "value", optional: true },
    ],
  },
  {
    name: "MATCH",
    category: "Lookup",
    summary: "Finds the position of a value in a range.",
    args: [
      { name: "lookup_value", kind: "value" },
      { name: "lookup_array", kind: "range" },
      { name: "match_type", kind: "value", optional: true, hint: "0 = exact" },
    ],
  },
  {
    name: "XLOOKUP",
    category: "Lookup",
    summary: "Modern lookup (default + horizontal/vertical).",
    args: [
      { name: "lookup_value", kind: "value" },
      { name: "lookup_array", kind: "range" },
      { name: "return_array", kind: "range" },
      { name: "if_not_found", kind: "value", optional: true },
    ],
  },
  {
    name: "INDIRECT",
    category: "Lookup",
    summary: "Resolves a text reference into a range.",
    args: [ARG_TEXT],
  },
  // Text
  { name: "CONCAT", category: "Text", summary: "Concatenates strings." },
  {
    name: "TEXTJOIN",
    category: "Text",
    summary: "Joins strings with a delimiter.",
    args: [
      { name: "delimiter", kind: "text", hint: '","' },
      { name: "ignore_empty", kind: "value", hint: "TRUE / FALSE" },
      ARG_RANGE,
    ],
  },
  {
    name: "LEFT",
    category: "Text",
    summary: "Returns N leftmost characters.",
    args: [
      { name: "text", kind: "text" },
      { name: "num_chars", kind: "value", optional: true },
    ],
  },
  {
    name: "RIGHT",
    category: "Text",
    summary: "Returns N rightmost characters.",
    args: [
      { name: "text", kind: "text" },
      { name: "num_chars", kind: "value", optional: true },
    ],
  },
  {
    name: "MID",
    category: "Text",
    summary: "Returns characters from the middle.",
    args: [
      { name: "text", kind: "text" },
      { name: "start_num", kind: "value" },
      { name: "num_chars", kind: "value" },
    ],
  },
  { name: "LEN", category: "Text", summary: "Returns the number of characters.", args: [ARG_TEXT] },
  { name: "TRIM", category: "Text", summary: "Removes extra whitespace.", args: [ARG_TEXT] },
  { name: "UPPER", category: "Text", summary: "Uppercases a string.", args: [ARG_TEXT] },
  { name: "LOWER", category: "Text", summary: "Lowercases a string.", args: [ARG_TEXT] },
  {
    name: "SUBSTITUTE",
    category: "Text",
    summary: "Replaces text in a string.",
    args: [
      { name: "text", kind: "text" },
      { name: "old_text", kind: "text" },
      { name: "new_text", kind: "text" },
      { name: "instance_num", kind: "value", optional: true },
    ],
  },
  // Date & time
  { name: "TODAY", category: "Date & Time", summary: "Returns today's date.", args: [] },
  { name: "NOW", category: "Date & Time", summary: "Returns the current date and time.", args: [] },
  {
    name: "DATE",
    category: "Date & Time",
    summary: "Returns a date from y/m/d.",
    args: [
      { name: "year", kind: "value" },
      { name: "month", kind: "value" },
      { name: "day", kind: "value" },
    ],
  },
  {
    name: "YEAR",
    category: "Date & Time",
    summary: "Returns the year of a date.",
    args: [{ name: "date", kind: "value" }],
  },
  {
    name: "MONTH",
    category: "Date & Time",
    summary: "Returns the month of a date.",
    args: [{ name: "date", kind: "value" }],
  },
  {
    name: "DAY",
    category: "Date & Time",
    summary: "Returns the day of a date.",
    args: [{ name: "date", kind: "value" }],
  },
  {
    name: "WEEKDAY",
    category: "Date & Time",
    summary: "Returns the day of the week.",
    args: [
      { name: "date", kind: "value" },
      { name: "return_type", kind: "value", optional: true },
    ],
  },
];

const CATEGORIES = ["All", ...Array.from(new Set(CATALOGUE.map((c) => c.category)))];

export function InsertFunctionDialog(props: InsertFunctionDialogProps): ReactNode {
  const { open, onClose, onPick, onInsertFormula, targetCellLabel } = props;
  const [query, setQuery] = useState("");
  const [category, setCategory] = useState<string>("All");
  const [activeIdx, setActiveIdx] = useState(0);
  const [step, setStep] = useState<"pick" | "args">("pick");
  const [chosen, setChosen] = useState<FunctionEntry | null>(null);
  const [argValues, setArgValues] = useState<string[]>([]);
  const panelRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const firstArgRef = useRef<HTMLInputElement>(null);
  useFocusTrap(panelRef, { enabled: open, onEscape: onClose });

  useEffect(() => {
    if (open) {
      setQuery("");
      setCategory("All");
      setActiveIdx(0);
      setStep("pick");
      setChosen(null);
      setArgValues([]);
      requestAnimationFrame(() => inputRef.current?.focus());
    }
  }, [open]);

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    return CATALOGUE.filter((entry) => {
      if (category !== "All" && entry.category !== category) return false;
      if (q.length === 0) return true;
      return (
        entry.name.toLowerCase().includes(q) ||
        entry.summary.toLowerCase().includes(q) ||
        entry.category.toLowerCase().includes(q)
      );
    });
  }, [query, category]);

  if (!open) return null;

  const choose = (entry: FunctionEntry): void => {
    if (onInsertFormula && entry.args && entry.args.length > 0) {
      setChosen(entry);
      setArgValues(entry.args.map(() => ""));
      setStep("args");
      requestAnimationFrame(() => firstArgRef.current?.focus());
      return;
    }
    if (onInsertFormula && entry.args && entry.args.length === 0) {
      onInsertFormula(`=${entry.name}()`);
      onClose();
      return;
    }
    onPick(entry.name);
    onClose();
  };

  const submitArgs = (): void => {
    if (!chosen || !onInsertFormula) return;
    const specs = chosen.args ?? [];
    const parts: string[] = [];
    for (let i = 0; i < specs.length; i += 1) {
      const spec = specs[i];
      const raw = (argValues[i] ?? "").trim();
      if (raw === "") {
        if (spec.optional) continue;
        return;
      }
      parts.push(formatArg(raw, spec));
    }
    onInsertFormula(`=${chosen.name}(${parts.join(", ")})`);
    onClose();
  };

  const skipArgs = (): void => {
    if (!chosen) return;
    onPick(chosen.name);
    onClose();
  };

  const onKeyDown = (e: React.KeyboardEvent): void => {
    if (e.key === "ArrowDown") {
      e.preventDefault();
      setActiveIdx((i) => Math.min(i + 1, filtered.length - 1));
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      setActiveIdx((i) => Math.max(i - 1, 0));
    } else if (e.key === "Enter") {
      e.preventDefault();
      const pick = filtered[activeIdx];
      if (pick) choose(pick);
    }
  };

  const argRequiredCount = (chosen?.args ?? []).filter((a) => !a.optional).length;
  const filledRequiredCount = (chosen?.args ?? [])
    .map((a, i) => (!a.optional ? (argValues[i] ?? "").trim() !== "" : true))
    .filter(Boolean).length;
  const canInsert = step === "args" && filledRequiredCount === (chosen?.args ?? []).length;

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="xlsx-fx-title"
      className="fixed inset-0 z-[60] flex items-center justify-center bg-black/40 px-4 py-6"
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div
        ref={panelRef}
        tabIndex={-1}
        className="flex w-full max-w-md flex-col rounded-lg border border-divider bg-surface shadow-2xl outline-none"
        onClick={(e) => e.stopPropagation()}
        data-testid="xlsx-insert-function-dialog"
      >
        <header className="flex items-center justify-between gap-3 border-b border-divider px-5 py-3">
          <div className="flex items-baseline gap-2">
            <h2 id="xlsx-fx-title" className="text-base font-semibold">
              {step === "pick" ? "Insert function" : `Function arguments`}
            </h2>
            {step === "args" && chosen ? (
              <span className="text-xs text-secondary font-mono">{chosen.name}</span>
            ) : null}
            {targetCellLabel ? <span className="text-[11px] text-tertiary">→ {targetCellLabel}</span> : null}
          </div>
          <button
            type="button"
            aria-label="Close"
            onClick={onClose}
            className="rounded p-1 text-secondary transition-colors hover:bg-hover hover:text-default"
          >
            <X size={16} />
          </button>
        </header>
        {step === "pick" ? (
          <div className="flex flex-col gap-3 px-5 py-4">
            <div className="flex items-center gap-2">
              <div className="relative flex-1">
                <Search
                  size={12}
                  className="pointer-events-none absolute left-2 top-1/2 -translate-y-1/2 text-tertiary"
                />
                <input
                  ref={inputRef}
                  value={query}
                  onChange={(e) => {
                    setQuery(e.currentTarget.value);
                    setActiveIdx(0);
                  }}
                  onKeyDown={onKeyDown}
                  placeholder="Search a function (e.g. VLOOKUP)…"
                  className="h-7 w-full rounded-md border border-divider bg-background pl-6 pr-2 text-sm focus:border-[var(--accent)] focus:outline-none"
                  data-testid="xlsx-fx-search"
                  aria-label="Search functions"
                />
              </div>
              <select
                value={category}
                onChange={(e) => {
                  setCategory(e.currentTarget.value);
                  setActiveIdx(0);
                }}
                className="h-7 rounded-md border border-divider bg-background px-2 text-sm focus:border-[var(--accent)] focus:outline-none"
                data-testid="xlsx-fx-category"
                aria-label="Function category"
              >
                {CATEGORIES.map((c) => (
                  <option key={c} value={c}>
                    {c}
                  </option>
                ))}
              </select>
            </div>
            <ul
              className="max-h-72 min-h-32 overflow-auto rounded border border-divider bg-background"
              data-testid="xlsx-fx-list"
            >
              {filtered.length === 0 ? (
                <li className="px-3 py-4 text-center text-sm text-tertiary">No matches.</li>
              ) : (
                filtered.map((entry, i) => (
                  <li
                    key={entry.name}
                    className={`flex cursor-pointer items-baseline justify-between gap-3 px-3 py-1.5 text-sm hover:bg-hover ${
                      i === activeIdx ? "bg-hover" : ""
                    }`}
                    onMouseEnter={() => setActiveIdx(i)}
                    onClick={() => choose(entry)}
                    data-testid={`xlsx-fx-row-${entry.name}`}
                  >
                    <span className="font-mono text-foreground">{entry.name}</span>
                    <span className="truncate text-xs text-secondary">{entry.summary}</span>
                  </li>
                ))
              )}
            </ul>
          </div>
        ) : (
          <div className="flex flex-col gap-3 px-5 py-4" data-testid="xlsx-fx-args">
            {chosen ? <p className="text-xs text-secondary">{chosen.summary}</p> : null}
            {(chosen?.args ?? []).map((spec, i) => (
              <label key={`${spec.name}-${i}`} className="flex flex-col gap-1 text-xs">
                <span className="flex items-baseline gap-2">
                  <span className="font-mono text-foreground">{spec.name}</span>
                  <span className="text-[10px] uppercase tracking-wide text-tertiary">
                    {spec.kind}
                    {spec.optional ? " · optional" : ""}
                  </span>
                </span>
                <input
                  ref={i === 0 ? firstArgRef : undefined}
                  value={argValues[i] ?? ""}
                  onChange={(e) => {
                    const next = argValues.slice();
                    next[i] = e.currentTarget.value;
                    setArgValues(next);
                  }}
                  onKeyDown={(e) => {
                    if (e.key === "Enter" && canInsert) {
                      e.preventDefault();
                      submitArgs();
                    }
                  }}
                  placeholder={spec.hint ?? ""}
                  className="h-7 rounded-md border border-divider bg-background px-2 font-mono text-sm focus:border-[var(--accent)] focus:outline-none"
                  data-testid={`xlsx-fx-arg-${spec.name}`}
                />
              </label>
            ))}
            {chosen ? (
              <p
                className="rounded border border-dashed border-divider bg-background px-2 py-1 font-mono text-[11px] text-secondary"
                data-testid="xlsx-fx-preview"
              >
                ={chosen.name}(
                {(chosen.args ?? [])
                  .map((spec, i) => {
                    const raw = (argValues[i] ?? "").trim();
                    if (raw === "" && spec.optional) return null;
                    return raw === "" ? `<${spec.name}>` : formatArg(raw, spec);
                  })
                  .filter((x): x is string => x !== null)
                  .join(", ")}
                )
              </p>
            ) : null}
          </div>
        )}
        <footer className="flex items-center justify-between gap-2 border-t border-divider px-5 py-3 text-xs text-secondary">
          {step === "pick" ? (
            <>
              <span>↑ ↓ to navigate · Enter to insert · Esc to cancel</span>
              <button
                type="button"
                onClick={onClose}
                className="rounded border border-divider bg-background px-3 py-1.5 text-sm hover:bg-hover"
              >
                Cancel
              </button>
            </>
          ) : (
            <>
              <span>
                {filledRequiredCount}/{argRequiredCount} required filled · Enter to insert · Esc to cancel
              </span>
              <div className="flex items-center gap-1">
                <button
                  type="button"
                  onClick={() => {
                    setStep("pick");
                    setChosen(null);
                  }}
                  className="rounded border border-divider bg-background px-3 py-1.5 text-sm hover:bg-hover"
                  data-testid="xlsx-fx-args-back"
                >
                  Back
                </button>
                <button
                  type="button"
                  onClick={skipArgs}
                  className="rounded border border-divider bg-background px-3 py-1.5 text-sm hover:bg-hover"
                  data-testid="xlsx-fx-args-skip"
                  title="Seed the formula bar instead and finish editing inline"
                >
                  Edit in formula bar
                </button>
                <button
                  type="button"
                  onClick={submitArgs}
                  disabled={!canInsert}
                  className="rounded bg-[var(--accent)] px-3 py-1.5 text-sm font-medium text-white hover:opacity-90 disabled:cursor-not-allowed disabled:opacity-40"
                  data-testid="xlsx-fx-args-insert"
                >
                  Insert
                </button>
              </div>
            </>
          )}
        </footer>
      </div>
    </div>
  );
}

/**
 * Quote a literal text/criterion arg if it isn't already a cell ref,
 * range, number, boolean, or quoted string. The heuristic is loose
 * by design — Excel itself happily accepts unquoted strings that
 * look like cell refs and we don't want to over-quote.
 */
function formatArg(raw: string, spec: ArgSpec): string {
  if (spec.kind === "range") return raw;
  const looksLikeRef = /^[A-Z]+\d+(?::[A-Z]+\d+)?$/i.test(raw);
  const looksLikeNumber = /^-?\d+(?:\.\d+)?$/.test(raw);
  const looksLikeBoolean = /^(?:true|false)$/i.test(raw);
  const alreadyQuoted = /^".*"$/.test(raw);
  const looksLikeFunctionCall = /^[A-Z]+\(/i.test(raw);
  const looksLikeExpression = /[+\-*/<>=&]/.test(raw);
  if (
    looksLikeRef ||
    looksLikeNumber ||
    looksLikeBoolean ||
    alreadyQuoted ||
    looksLikeFunctionCall ||
    looksLikeExpression
  ) {
    return raw;
  }
  if (spec.kind === "text" || spec.kind === "criterion") {
    return `"${raw.replace(/"/g, '""')}"`;
  }
  return raw;
}
