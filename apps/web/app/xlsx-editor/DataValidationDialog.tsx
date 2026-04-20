"use client";

/**
 * C11 — Data Validation dialog.
 *
 * Compact authoring surface for in-cell list validation (the most
 * common Excel data-validation use case). Other rule kinds (whole,
 * decimal, date, custom) round-trip opaquely from the parser but
 * aren't editable here yet.
 */

import { useEffect, useState } from "react";
import type { ReactNode } from "react";
import { Trash2, X } from "lucide-react";
import type { DataValidation } from "@officeai/xlsx";
import { useTranslator } from "@/lib/i18n";

export interface DataValidationDialogProps {
  readonly open: boolean;
  readonly onClose: () => void;
  readonly defaultRange: string;
  readonly rules: ReadonlyArray<DataValidation>;
  /** Set when the parser captured opaque non-list rules on this sheet. */
  readonly hasOpaqueRules: boolean;
  readonly onAddRule: (rule: DataValidation) => void;
  readonly onRemoveRule: (id: string) => void;
  readonly onClearRules: () => void;
}

export function DataValidationDialog(props: DataValidationDialogProps): ReactNode {
  const { open, onClose, defaultRange, rules, hasOpaqueRules, onAddRule, onRemoveRule, onClearRules } = props;
  const { t } = useTranslator();

  const [range, setRange] = useState<string>(defaultRange);
  const [mode, setMode] = useState<"literal" | "formula">("literal");
  const [source, setSource] = useState<string>("");
  const [showDropDown, setShowDropDown] = useState<boolean>(true);
  const [stopOnInvalid, setStopOnInvalid] = useState<boolean>(true);
  const [allowBlank, setAllowBlank] = useState<boolean>(true);

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

  if (!open) return null;

  const submit = () => {
    const r = range.trim() || defaultRange;
    if (!r) return;
    if (!source.trim()) return;
    const id = `dv-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}`;
    onAddRule({
      kind: "list",
      id,
      range: r,
      source: source.trim(),
      formula: mode === "formula",
      showDropDown,
      stopOnInvalid,
      allowBlank,
    });
    setSource("");
  };

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/40"
      role="dialog"
      aria-modal="true"
      aria-label={t("xlsx.dataValidation.title")}
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div className="relative w-[480px] rounded-md border border-divider bg-background p-4 text-sm text-foreground shadow-xl">
        <div className="mb-3 flex items-center justify-between">
          <h2 className="text-base font-semibold">{t("xlsx.dataValidation.title")}</h2>
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
          <h3 className="mb-1 text-xs font-medium text-secondary">{t("xlsx.dataValidation.existingRules")}</h3>
          {rules.length === 0 && !hasOpaqueRules ? (
            <p className="rounded border border-dashed border-divider px-3 py-2 text-xs text-secondary">
              {t("xlsx.dataValidation.noRules")}
            </p>
          ) : (
            <ul className="max-h-32 overflow-auto rounded border border-divider">
              {rules.map((r) => (
                <li
                  key={r.id}
                  className="flex items-center justify-between gap-2 border-b border-divider px-2 py-1 text-xs last:border-b-0"
                >
                  <span className="truncate">
                    <span className="font-medium">{t("xlsx.dataValidation.list")}</span>
                    <span className="ml-2 text-secondary">{r.range}</span>
                    <span className="ml-2 text-secondary">
                      {r.formula ? `→ ${r.source}` : `(${r.source})`}
                    </span>
                  </span>
                  <button
                    type="button"
                    aria-label={t("xlsx.dataValidation.removeRule")}
                    title={t("xlsx.dataValidation.removeRule")}
                    className="inline-flex h-6 w-6 items-center justify-center rounded text-secondary hover:bg-hover"
                    onClick={() => onRemoveRule(r.id)}
                  >
                    <Trash2 size={12} />
                  </button>
                </li>
              ))}
              {hasOpaqueRules ? (
                <li className="border-b border-divider px-2 py-1 text-xs italic text-secondary last:border-b-0">
                  {t("xlsx.dataValidation.preservedNonListRules")}
                </li>
              ) : null}
            </ul>
          )}
          {rules.length > 0 || hasOpaqueRules ? (
            <div className="mt-1 text-right">
              <button
                type="button"
                className="text-xs text-secondary hover:text-foreground hover:underline"
                onClick={onClearRules}
              >
                {t("xlsx.dataValidation.removeAllRules")}
              </button>
            </div>
          ) : null}
        </div>

        <div className="mb-3 space-y-2">
          <h3 className="mb-1 text-xs font-medium text-secondary">{t("xlsx.dataValidation.newListRule")}</h3>
          <div className="flex gap-2">
            <label className="flex flex-1 flex-col gap-0.5 text-xs">
              <span className="text-secondary">{t("xlsx.dataValidation.range")}</span>
              <input
                value={range}
                onChange={(e) => setRange(e.target.value)}
                placeholder={t("xlsx.dataValidation.rangePlaceholder")}
                className="h-7 rounded border border-divider bg-background px-2 text-xs"
              />
            </label>
            <label className="flex flex-col gap-0.5 text-xs">
              <span className="text-secondary">{t("xlsx.dataValidation.source")}</span>
              <select
                value={mode}
                onChange={(e) => setMode(e.target.value as "literal" | "formula")}
                className="h-7 rounded border border-divider bg-background px-1 text-xs"
              >
                <option value="literal">{t("xlsx.dataValidation.sourceLiteral")}</option>
                <option value="formula">{t("xlsx.dataValidation.sourceFormula")}</option>
              </select>
            </label>
          </div>
          <label className="flex flex-col gap-0.5 text-xs">
            <span className="text-secondary">
              {mode === "literal" ? t("xlsx.dataValidation.literalLabel") : t("xlsx.dataValidation.formulaLabel")}
            </span>
            <input
              value={source}
              onChange={(e) => setSource(e.target.value)}
              placeholder={mode === "literal" ? t("xlsx.dataValidation.literalPlaceholder") : t("xlsx.dataValidation.formulaPlaceholder")}
              className="h-7 rounded border border-divider bg-background px-2 text-xs"
            />
          </label>
          <div className="flex flex-wrap gap-3 pt-1 text-xs">
            <label className="flex items-center gap-1">
              <input
                type="checkbox"
                checked={showDropDown}
                onChange={(e) => setShowDropDown(e.target.checked)}
              />
              {t("xlsx.dataValidation.inCellDropdown")}
            </label>
            <label className="flex items-center gap-1">
              <input
                type="checkbox"
                checked={stopOnInvalid}
                onChange={(e) => setStopOnInvalid(e.target.checked)}
              />
              {t("xlsx.dataValidation.rejectInvalid")}
            </label>
            <label className="flex items-center gap-1">
              <input type="checkbox" checked={allowBlank} onChange={(e) => setAllowBlank(e.target.checked)} />
              {t("xlsx.dataValidation.allowBlank")}
            </label>
          </div>
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
            {t("xlsx.dataValidation.addRule")}
          </button>
        </div>
      </div>
    </div>
  );
}
