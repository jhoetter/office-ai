"use client";

/**
 * C10 — Conditional Formatting dialog.
 *
 * Compact, modal authoring surface for the most common Excel CF
 * presets. Mirrors the "Home > Conditional Formatting" splitter in
 * spirit (Highlight Cells, Top/Bottom, Data Bars, Color Scales,
 * Duplicates) without trying to reproduce the full OOXML rule
 * grammar — anything we don't model here is preserved verbatim
 * through the opaque round-trip path in the parser/serializer.
 *
 * The dialog is intentionally read-light: it shows existing typed
 * rules on the active sheet plus a single "New rule…" form. Apply
 * dispatches one command per change so undo/redo line up with the
 * user's mental model ("undo my last new rule").
 */

import { useEffect, useMemo, useState } from "react";
import type { ReactNode } from "react";
import { Trash2, X } from "lucide-react";
import { cn } from "@officeai/ui";
import type { ConditionalFormat, ConditionalFormatOverlay } from "@officeai/xlsx";
import { useTranslator } from "@/lib/i18n";

export type CfRuleKind = "cellIs" | "top10" | "containsText" | "duplicate" | "colorScale" | "dataBar";

export interface ConditionalFormatDialogProps {
  readonly open: boolean;
  readonly onClose: () => void;
  /** Default range for new rules — usually the current selection (A1 form). */
  readonly defaultRange: string;
  readonly rules: ReadonlyArray<ConditionalFormat>;
  readonly onAddRule: (rule: ConditionalFormat) => void;
  readonly onRemoveRule: (id: string) => void;
  readonly onClearRules: () => void;
}

/**
 * Excel's "classic" preset palette. Each entry pairs a fill +
 * font colour so the swatch row in the dialog feels familiar.
 */
const PRESET_OVERLAYS: ReadonlyArray<{
  readonly id: string;
  readonly labelKey: string;
  readonly overlay: ConditionalFormatOverlay;
}> = [
  { id: "red", labelKey: "xlsx.conditionalFormat.presetRedFill", overlay: { fill: "FFC7CE", fontColor: "9C0006" } },
  { id: "yellow", labelKey: "xlsx.conditionalFormat.presetYellowFill", overlay: { fill: "FFEB9C", fontColor: "9C5700" } },
  { id: "green", labelKey: "xlsx.conditionalFormat.presetGreenFill", overlay: { fill: "C6EFCE", fontColor: "006100" } },
  { id: "redText", labelKey: "xlsx.conditionalFormat.presetRedText", overlay: { fontColor: "9C0006", bold: true } },
  { id: "redBorder", labelKey: "xlsx.conditionalFormat.presetRedBorder", overlay: { fontColor: "9C0006", italic: true } },
];

const RULE_KIND_KEY: Record<CfRuleKind, string> = {
  cellIs: "xlsx.conditionalFormat.ruleCellIs",
  top10: "xlsx.conditionalFormat.ruleTop10",
  containsText: "xlsx.conditionalFormat.ruleContainsText",
  duplicate: "xlsx.conditionalFormat.ruleDuplicate",
  colorScale: "xlsx.conditionalFormat.ruleColorScale",
  dataBar: "xlsx.conditionalFormat.ruleDataBar",
};

export function ConditionalFormatDialog(props: ConditionalFormatDialogProps): ReactNode {
  const { open, onClose, defaultRange, rules, onAddRule, onRemoveRule, onClearRules } = props;
  const { t } = useTranslator();

  const [kind, setKind] = useState<CfRuleKind>("cellIs");
  const [range, setRange] = useState<string>(defaultRange);
  const [presetIdx, setPresetIdx] = useState<number>(0);

  // cellIs
  const [op, setOp] = useState<"gt" | "ge" | "lt" | "le" | "eq" | "ne" | "between" | "notBetween">("gt");
  const [value, setValue] = useState<string>("");
  const [value2, setValue2] = useState<string>("");

  // top10
  const [bottom, setBottom] = useState<boolean>(false);
  const [percent, setPercent] = useState<boolean>(false);
  const [rank, setRank] = useState<string>("10");

  // containsText
  const [text, setText] = useState<string>("");
  const [contains, setContains] = useState<boolean>(true);

  // duplicate
  const [unique, setUnique] = useState<boolean>(false);

  // colorScale
  const [minColor, setMinColor] = useState<string>("F8696B");
  const [midColor, setMidColor] = useState<string>("FFEB84");
  const [maxColor, setMaxColor] = useState<string>("63BE7B");

  // dataBar
  const [barColor, setBarColor] = useState<string>("638EC6");

  useEffect(() => {
    if (!open) return;
    setRange(defaultRange);
  }, [open, defaultRange]);

  useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        e.preventDefault();
        onClose();
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [open, onClose]);

  const overlay = useMemo(
    () => PRESET_OVERLAYS[presetIdx]?.overlay ?? PRESET_OVERLAYS[0]!.overlay,
    [presetIdx]
  );

  if (!open) return null;

  const submit = () => {
    const id = `cf-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}`;
    const r = range.trim() || defaultRange;
    if (!r) return;
    let rule: ConditionalFormat | null = null;
    switch (kind) {
      case "cellIs": {
        const v = Number(value);
        const v2 = value2 === "" ? undefined : Number(value2);
        if (!Number.isFinite(v)) return;
        if ((op === "between" || op === "notBetween") && (v2 === undefined || !Number.isFinite(v2))) return;
        rule = { kind: "cellIs", id, range: r, op, value: v, value2: v2, overlay };
        break;
      }
      case "top10": {
        const n = Math.max(1, Math.floor(Number(rank) || 10));
        rule = { kind: "top10", id, range: r, bottom, percent, rank: n, overlay };
        break;
      }
      case "containsText": {
        if (!text.trim()) return;
        rule = { kind: "containsText", id, range: r, text, contains, overlay };
        break;
      }
      case "duplicate": {
        rule = { kind: "duplicate", id, range: r, unique, overlay };
        break;
      }
      case "colorScale": {
        rule = { kind: "colorScale", id, range: r, minColor, midColor, maxColor };
        break;
      }
      case "dataBar": {
        rule = { kind: "dataBar", id, range: r, color: barColor };
        break;
      }
      default: {
        const _exhaustive: never = kind;
        void _exhaustive;
      }
    }
    if (rule) onAddRule(rule);
  };

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/40"
      role="dialog"
      aria-modal="true"
      aria-label={t("xlsx.conditionalFormat.title")}
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div className="relative w-[520px] rounded-md border border-divider bg-background p-4 text-sm text-foreground shadow-xl">
        <div className="mb-3 flex items-center justify-between">
          <h2 className="text-base font-semibold">{t("xlsx.conditionalFormat.title")}</h2>
          <button
            type="button"
            aria-label={t("common.close")}
            className="inline-flex h-6 w-6 items-center justify-center rounded text-secondary hover:bg-hover"
            onClick={onClose}
          >
            <X size={14} />
          </button>
        </div>

        <div className="mb-4">
          <h3 className="mb-1 text-xs font-medium text-secondary">{t("xlsx.conditionalFormat.existingRules")}</h3>
          {rules.length === 0 ? (
            <p className="rounded border border-dashed border-divider px-3 py-2 text-xs text-secondary">
              {t("xlsx.conditionalFormat.noRules")}
            </p>
          ) : (
            <ul className="max-h-40 overflow-auto rounded border border-divider">
              {rules.map((r) => (
                <li
                  key={r.id}
                  className="flex items-center justify-between gap-2 border-b border-divider px-2 py-1 text-xs last:border-b-0"
                >
                  <span className="truncate">
                    <span className="font-medium">{t(RULE_KIND_KEY[r.kind])}</span>
                    <span className="ml-2 text-secondary">{r.range}</span>
                  </span>
                  <button
                    type="button"
                    aria-label={t("xlsx.conditionalFormat.removeRule")}
                    title={t("xlsx.conditionalFormat.removeRule")}
                    className="inline-flex h-6 w-6 items-center justify-center rounded text-secondary hover:bg-hover"
                    onClick={() => onRemoveRule(r.id)}
                  >
                    <Trash2 size={12} />
                  </button>
                </li>
              ))}
            </ul>
          )}
          {rules.length > 0 ? (
            <div className="mt-1 text-right">
              <button
                type="button"
                className="text-xs text-secondary hover:text-foreground hover:underline"
                onClick={onClearRules}
              >
                {t("xlsx.conditionalFormat.removeAllRules")}
              </button>
            </div>
          ) : null}
        </div>

        <div className="mb-3">
          <h3 className="mb-1 text-xs font-medium text-secondary">{t("xlsx.conditionalFormat.newRule")}</h3>
          <div className="flex gap-2">
            <select
              value={kind}
              onChange={(e) => setKind(e.target.value as CfRuleKind)}
              className="h-7 flex-1 rounded border border-divider bg-background px-1 text-xs"
              aria-label={t("xlsx.conditionalFormat.ruleKind")}
            >
              {(Object.keys(RULE_KIND_KEY) as CfRuleKind[]).map((k) => (
                <option key={k} value={k}>
                  {t(RULE_KIND_KEY[k])}
                </option>
              ))}
            </select>
            <input
              value={range}
              onChange={(e) => setRange(e.target.value)}
              placeholder={t("xlsx.conditionalFormat.rangePlaceholder")}
              aria-label={t("xlsx.conditionalFormat.range")}
              className="h-7 w-32 rounded border border-divider bg-background px-2 text-xs"
            />
          </div>
        </div>

        <div className="mb-3 space-y-2">
          {kind === "cellIs" ? (
            <div className="flex gap-2">
              <select
                value={op}
                onChange={(e) => setOp(e.target.value as typeof op)}
                className="h-7 rounded border border-divider bg-background px-1 text-xs"
                aria-label={t("xlsx.conditionalFormat.operator")}
              >
                <option value="gt">{t("xlsx.conditionalFormat.opGreaterThan")}</option>
                <option value="ge">{t("xlsx.conditionalFormat.opGreaterThanOrEqual")}</option>
                <option value="lt">{t("xlsx.conditionalFormat.opLessThan")}</option>
                <option value="le">{t("xlsx.conditionalFormat.opLessThanOrEqual")}</option>
                <option value="eq">{t("xlsx.conditionalFormat.opEqual")}</option>
                <option value="ne">{t("xlsx.conditionalFormat.opNotEqual")}</option>
                <option value="between">{t("xlsx.conditionalFormat.opBetween")}</option>
                <option value="notBetween">{t("xlsx.conditionalFormat.opNotBetween")}</option>
              </select>
              <input
                value={value}
                onChange={(e) => setValue(e.target.value)}
                placeholder="0"
                aria-label={t("xlsx.conditionalFormat.thresholdValue")}
                className="h-7 w-20 rounded border border-divider bg-background px-2 text-xs"
              />
              {op === "between" || op === "notBetween" ? (
                <input
                  value={value2}
                  onChange={(e) => setValue2(e.target.value)}
                  placeholder="100"
                  aria-label={t("xlsx.conditionalFormat.upperBound")}
                  className="h-7 w-20 rounded border border-divider bg-background px-2 text-xs"
                />
              ) : null}
            </div>
          ) : null}

          {kind === "top10" ? (
            <div className="flex gap-2">
              <select
                value={bottom ? "bottom" : "top"}
                onChange={(e) => setBottom(e.target.value === "bottom")}
                className="h-7 rounded border border-divider bg-background px-1 text-xs"
                aria-label={t("xlsx.conditionalFormat.direction")}
              >
                <option value="top">{t("xlsx.conditionalFormat.directionTop")}</option>
                <option value="bottom">{t("xlsx.conditionalFormat.directionBottom")}</option>
              </select>
              <input
                value={rank}
                onChange={(e) => setRank(e.target.value)}
                aria-label={t("xlsx.conditionalFormat.rank")}
                className="h-7 w-16 rounded border border-divider bg-background px-2 text-xs"
              />
              <label className="flex items-center gap-1 text-xs">
                <input type="checkbox" checked={percent} onChange={(e) => setPercent(e.target.checked)} />%
              </label>
            </div>
          ) : null}

          {kind === "containsText" ? (
            <div className="flex gap-2">
              <select
                value={contains ? "contains" : "not"}
                onChange={(e) => setContains(e.target.value === "contains")}
                className="h-7 rounded border border-divider bg-background px-1 text-xs"
                aria-label={t("xlsx.conditionalFormat.match")}
              >
                <option value="contains">{t("xlsx.conditionalFormat.matchContains")}</option>
                <option value="not">{t("xlsx.conditionalFormat.matchNotContains")}</option>
              </select>
              <input
                value={text}
                onChange={(e) => setText(e.target.value)}
                placeholder={t("xlsx.conditionalFormat.textPlaceholder")}
                aria-label={t("xlsx.conditionalFormat.textPlaceholder")}
                className="h-7 flex-1 rounded border border-divider bg-background px-2 text-xs"
              />
            </div>
          ) : null}

          {kind === "duplicate" ? (
            <div className="flex gap-2">
              <select
                value={unique ? "unique" : "duplicate"}
                onChange={(e) => setUnique(e.target.value === "unique")}
                className="h-7 rounded border border-divider bg-background px-1 text-xs"
                aria-label={t("xlsx.conditionalFormat.mode")}
              >
                <option value="duplicate">{t("xlsx.conditionalFormat.duplicateValues")}</option>
                <option value="unique">{t("xlsx.conditionalFormat.uniqueValues")}</option>
              </select>
            </div>
          ) : null}

          {kind === "colorScale" ? (
            <div className="flex items-center gap-2">
              <ColorField label={t("xlsx.conditionalFormat.stopMin")} value={minColor} onChange={setMinColor} />
              <ColorField label={t("xlsx.conditionalFormat.stopMid")} value={midColor} onChange={setMidColor} />
              <ColorField label={t("xlsx.conditionalFormat.stopMax")} value={maxColor} onChange={setMaxColor} />
            </div>
          ) : null}

          {kind === "dataBar" ? (
            <div className="flex items-center gap-2">
              <ColorField label={t("xlsx.conditionalFormat.stopBar")} value={barColor} onChange={setBarColor} />
            </div>
          ) : null}

          {kind === "cellIs" || kind === "top10" || kind === "containsText" || kind === "duplicate" ? (
            <div>
              <h4 className="mb-1 text-xs font-medium text-secondary">{t("xlsx.conditionalFormat.format")}</h4>
              <div className="flex flex-wrap gap-1">
                {PRESET_OVERLAYS.map((p, i) => (
                  <button
                    key={p.id}
                    type="button"
                    onClick={() => setPresetIdx(i)}
                    title={t(p.labelKey)}
                    aria-label={t(p.labelKey)}
                    aria-pressed={presetIdx === i}
                    className={cn(
                      "inline-flex h-6 items-center gap-1 rounded border px-2 text-[11px]",
                      presetIdx === i ? "border-accent" : "border-divider"
                    )}
                    style={{
                      background: p.overlay.fill ? `#${p.overlay.fill}` : "transparent",
                      color: p.overlay.fontColor ? `#${p.overlay.fontColor}` : undefined,
                      fontWeight: p.overlay.bold ? 700 : 400,
                      fontStyle: p.overlay.italic ? "italic" : "normal",
                    }}
                  >
                    {t("xlsx.conditionalFormat.sample")}
                  </button>
                ))}
              </div>
            </div>
          ) : null}
        </div>

        <div className="flex justify-end gap-2 border-t border-divider pt-3">
          <button
            type="button"
            onClick={onClose}
            className="h-7 rounded border border-divider bg-background px-3 text-xs hover:bg-hover"
          >
            {t("common.close")}
          </button>
          <button
            type="button"
            onClick={submit}
            className="h-7 rounded bg-accent px-3 text-xs font-medium text-white hover:bg-accent/90"
          >
            {t("xlsx.conditionalFormat.addRule")}
          </button>
        </div>
      </div>
    </div>
  );
}

function ColorField(props: {
  readonly label: string;
  readonly value: string;
  readonly onChange: (next: string) => void;
}): ReactNode {
  const { label, value, onChange } = props;
  const { t } = useTranslator();
  return (
    <label className="inline-flex items-center gap-1 text-xs">
      <span className="text-secondary">{label}</span>
      <input
        type="color"
        value={`#${value}`}
        onChange={(e) => onChange(e.target.value.replace(/^#/, "").toUpperCase())}
        className="h-6 w-7 cursor-pointer rounded border border-divider bg-transparent p-0"
        aria-label={t("xlsx.conditionalFormat.colorAriaLabel", { label })}
      />
    </label>
  );
}
