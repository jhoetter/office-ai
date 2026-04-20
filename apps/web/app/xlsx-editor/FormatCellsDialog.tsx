"use client";

import { useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import { X } from "lucide-react";
import { cn, useFocusTrap } from "@officeai/ui";
import type { CellFormatPatch, EffectiveStyle, StyleTable } from "@officeai/xlsx";
import { flattenCellXf } from "@officeai/xlsx";
import { NUMBER_FORMAT_PRESETS, numberFormatPresetKey, type NumberFormatPresetId } from "./styles";
import { useTranslator } from "@/lib/i18n";

/**
 * C5 — Excel-parity Format Cells dialog.
 *
 * Single modal with six tabs (Number / Alignment / Font / Border / Fill
 * / Protection) backed by the same {@link CellFormatPatch} surface the
 * toolbar uses. The dialog is intentionally thin: it derives initial
 * values from the *anchor* cell's effective style, lets the user tweak
 * any field, and emits a single sparse patch on Apply that the caller
 * dispatches over the *whole* selection range. Fields the user didn't
 * touch are left out of the patch so they don't blow away whatever
 * other cells in a multi-cell selection already had.
 *
 * Triggered from `Mod+1` (Excel parity) and from the right-click
 * context menu's "Format cells…" entry.
 */
export interface FormatCellsDialogProps {
  readonly open: boolean;
  readonly styles: StyleTable;
  readonly anchorStyleId: number | undefined;
  readonly onClose: () => void;
  readonly onApply: (patch: CellFormatPatch) => void;
  readonly initialTab?: TabId;
}

export type TabId = "number" | "alignment" | "font" | "border" | "fill" | "protection";

const TABS: ReadonlyArray<{ id: TabId; labelKey: string }> = [
  { id: "number", labelKey: "xlsx.formatCells.tabNumber" },
  { id: "alignment", labelKey: "xlsx.formatCells.tabAlignment" },
  { id: "font", labelKey: "xlsx.formatCells.tabFont" },
  { id: "border", labelKey: "xlsx.formatCells.tabBorder" },
  { id: "fill", labelKey: "xlsx.formatCells.tabFill" },
  { id: "protection", labelKey: "xlsx.formatCells.tabProtection" },
];

const FONT_FAMILIES = [
  "Calibri",
  "Arial",
  "Helvetica",
  "Times New Roman",
  "Georgia",
  "Courier New",
  "Verdana",
  "Tahoma",
];
const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48];

const BORDER_STYLES: ReadonlyArray<{
  id: NonNullable<NonNullable<CellFormatPatch["border"]>["top"]>["style"];
  labelKey: string;
}> = [
  { id: "none", labelKey: "xlsx.formatCells.borderStyleNone" },
  { id: "thin", labelKey: "xlsx.formatCells.borderStyleThin" },
  { id: "medium", labelKey: "xlsx.formatCells.borderStyleMedium" },
  { id: "thick", labelKey: "xlsx.formatCells.borderStyleThick" },
  { id: "double", labelKey: "xlsx.formatCells.borderStyleDouble" },
  { id: "dashed", labelKey: "xlsx.formatCells.borderStyleDashed" },
  { id: "dotted", labelKey: "xlsx.formatCells.borderStyleDotted" },
];

const FILL_SWATCHES = [
  "FFFFFF",
  "FEF3C7",
  "FCE7F3",
  "DBEAFE",
  "DCFCE7",
  "E0E7FF",
  "FFE4E6",
  "F3E8FF",
  "FFEDD5",
  "F0FDF4",
  "ECFEFF",
  "FCE7F3",
  "FFFBEB",
  "F1F5F9",
  "E5E7EB",
  "1F2937",
];

const FONT_COLOR_SWATCHES = [
  "000000",
  "111827",
  "374151",
  "6B7280",
  "1D4ED8",
  "0E7490",
  "047857",
  "B45309",
  "B91C1C",
  "9333EA",
  "0F766E",
  "A21CAF",
  "DC2626",
  "16A34A",
  "EA580C",
  "2563EB",
];

interface NumberPresetState {
  readonly id: NumberFormatPresetId | "custom";
  readonly customCode: string;
}

export function FormatCellsDialog(props: FormatCellsDialogProps): ReactNode {
  const { open, styles, anchorStyleId, onClose, onApply, initialTab } = props;
  const { t } = useTranslator();
  const effective: EffectiveStyle = useMemo(
    () => flattenCellXf(styles, anchorStyleId),
    [styles, anchorStyleId]
  );

  const [tab, setTab] = useState<TabId>(initialTab ?? "number");
  const panelRef = useRef<HTMLDivElement>(null);

  useFocusTrap(panelRef, { enabled: open, onEscape: onClose });

  // ── Number ────────────────────────────────────────────────────────
  const [numberPreset, setNumberPreset] = useState<NumberPresetState>({
    id: "general",
    customCode: "",
  });
  // ── Alignment ────────────────────────────────────────────────────
  const [horizontal, setHorizontal] = useState<NonNullable<CellFormatPatch["alignment"]>["horizontal"] | "">(
    ""
  );
  const [vertical, setVertical] = useState<NonNullable<CellFormatPatch["alignment"]>["vertical"] | "">("");
  const [wrapText, setWrapText] = useState(false);
  const [indent, setIndent] = useState(0);
  // ── Font ─────────────────────────────────────────────────────────
  const [fontFamily, setFontFamily] = useState("");
  const [fontSize, setFontSize] = useState<number>(11);
  const [bold, setBold] = useState(false);
  const [italic, setItalic] = useState(false);
  const [underline, setUnderline] = useState(false);
  const [strike, setStrike] = useState(false);
  const [fontColor, setFontColor] = useState("");
  // ── Border ───────────────────────────────────────────────────────
  type BorderStyle = NonNullable<NonNullable<CellFormatPatch["border"]>["top"]>["style"];
  const [borderTop, setBorderTop] = useState<BorderStyle>(undefined);
  const [borderRight, setBorderRight] = useState<BorderStyle>(undefined);
  const [borderBottom, setBorderBottom] = useState<BorderStyle>(undefined);
  const [borderLeft, setBorderLeft] = useState<BorderStyle>(undefined);
  const [borderColor, setBorderColor] = useState("000000");
  // ── Fill ─────────────────────────────────────────────────────────
  const [fillColor, setFillColor] = useState("");
  // ── Protection ───────────────────────────────────────────────────
  const [locked, setLocked] = useState(true);
  const [hidden, setHidden] = useState(false);

  // Reset on each open from the anchor's effective style. We track
  // *touched* fields separately via "" / undefined sentinels so the
  // emitted patch only includes things the user actually changed.
  useEffect(() => {
    if (!open) return;
    setTab(initialTab ?? "number");
    // Number — try to match an existing preset by code; otherwise
    // surface the existing custom code so the user can tweak it.
    const matched = NUMBER_FORMAT_PRESETS.find((p) => numFmtIdMatches(effective.numFmtId, p.code));
    setNumberPreset(
      matched
        ? { id: matched.id, customCode: matched.code }
        : { id: "custom", customCode: numFmtIdToCustomGuess(effective.numFmtId) }
    );
    const h = effective.alignment?.horizontal;
    setHorizontal(
      h === "left" || h === "center" || h === "right" || h === "justify" || h === "fill"
        ? (h as NonNullable<CellFormatPatch["alignment"]>["horizontal"])
        : ""
    );
    setVertical(
      ((): "top" | "middle" | "bottom" | "" => {
        const v = effective.alignment?.vertical;
        if (v === "top") return "top";
        if (v === "center") return "middle";
        if (v === "bottom") return "bottom";
        return "";
      })()
    );
    setWrapText(!!effective.alignment?.wrapText);
    setIndent(effective.alignment?.indent ?? 0);
    setFontFamily(effective.font.name ?? "");
    setFontSize(effective.font.size ?? 11);
    setBold(!!effective.font.bold);
    setItalic(!!effective.font.italic);
    setUnderline(!!effective.font.underline);
    setStrike(!!effective.font.strike);
    setFontColor(effective.font.color?.rgb ? normalizeRgb(effective.font.color.rgb) : "");
    setBorderTop(borderToStyle(effective.border.top?.style));
    setBorderRight(borderToStyle(effective.border.right?.style));
    setBorderBottom(borderToStyle(effective.border.bottom?.style));
    setBorderLeft(borderToStyle(effective.border.left?.style));
    setBorderColor(
      effective.border.top?.color?.rgb ? normalizeRgb(effective.border.top.color.rgb) : "000000"
    );
    setFillColor(
      effective.fill.kind === "pattern" &&
        effective.fill.patternType === "solid" &&
        effective.fill.fgColor?.rgb
        ? normalizeRgb(effective.fill.fgColor.rgb)
        : ""
    );
    setLocked(effective.protection?.locked ?? true);
    setHidden(effective.protection?.hidden ?? false);
  }, [open, effective, initialTab]);

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

  const apply = () => {
    type MutableSide = NonNullable<NonNullable<CellFormatPatch["border"]>["top"]>;
    type MutablePatch = {
      numberFormat?: string;
      alignment?: {
        horizontal?: NonNullable<CellFormatPatch["alignment"]>["horizontal"];
        vertical?: NonNullable<CellFormatPatch["alignment"]>["vertical"];
        wrapText?: boolean;
        indent?: number;
      };
      font?: {
        fontFamily?: string;
        size?: number;
        bold?: boolean;
        italic?: boolean;
        underline?: boolean;
        strike?: boolean;
        color?: string;
      };
      border?: {
        top?: MutableSide;
        right?: MutableSide;
        bottom?: MutableSide;
        left?: MutableSide;
      };
      fill?: { color?: string; pattern?: "solid" | "none" };
      protection?: { locked?: boolean; hidden?: boolean };
    };
    const patch: MutablePatch = {};
    // Number
    if (numberPreset.id === "custom") {
      const code = numberPreset.customCode.trim();
      if (code.length > 0) patch.numberFormat = code;
    } else {
      const preset = NUMBER_FORMAT_PRESETS.find((p) => p.id === numberPreset.id);
      if (preset) patch.numberFormat = preset.code;
    }
    // Alignment
    const alignment: NonNullable<MutablePatch["alignment"]> = {};
    if (horizontal) alignment.horizontal = horizontal;
    if (vertical) alignment.vertical = vertical;
    alignment.wrapText = wrapText;
    if (Number.isFinite(indent)) alignment.indent = Math.max(0, Math.min(15, indent));
    if (Object.keys(alignment).length > 0) patch.alignment = alignment;
    // Font
    const font: NonNullable<MutablePatch["font"]> = {};
    if (fontFamily) font.fontFamily = fontFamily;
    if (Number.isFinite(fontSize) && fontSize > 0) font.size = fontSize;
    font.bold = bold;
    font.italic = italic;
    font.underline = underline;
    font.strike = strike;
    if (fontColor) font.color = fontColor;
    if (Object.keys(font).length > 0) patch.font = font;
    // Border
    const border: NonNullable<MutablePatch["border"]> = {};
    const sideOf = (s: BorderStyle | undefined): MutableSide | undefined => {
      if (s === undefined) return undefined;
      if (s === "none") return { style: "none" };
      return { style: s, color: borderColor };
    };
    const t = sideOf(borderTop);
    if (t) border.top = t;
    const r = sideOf(borderRight);
    if (r) border.right = r;
    const b = sideOf(borderBottom);
    if (b) border.bottom = b;
    const l = sideOf(borderLeft);
    if (l) border.left = l;
    if (Object.keys(border).length > 0) patch.border = border;
    // Fill
    if (fillColor) {
      patch.fill = { color: fillColor, pattern: "solid" };
    }
    // Protection — always emit since the dialog gives a definite answer
    patch.protection = { locked, hidden };
    onApply(patch as CellFormatPatch);
    onClose();
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="format-cells-title"
      data-testid="format-cells-dialog"
      className="fixed inset-0 z-[60] flex items-center justify-center bg-black/40 px-4 py-6"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div
        ref={panelRef}
        tabIndex={-1}
        className="flex max-h-[85vh] w-full max-w-2xl flex-col rounded-lg border border-divider bg-surface shadow-2xl outline-none"
        onMouseDown={(e) => e.stopPropagation()}
      >
        <header className="flex items-center justify-between gap-3 border-b border-divider px-5 py-3">
          <div>
            <h2 id="format-cells-title" className="text-base font-semibold text-foreground">
              {t("xlsx.formatCells.title")}
            </h2>
            <p className="text-xs text-secondary">
              {t("xlsx.formatCells.hint")}
            </p>
          </div>
          <button
            type="button"
            aria-label={t("common.close")}
            onClick={onClose}
            className="rounded p-1 text-secondary transition-colors hover:bg-hover hover:text-foreground"
          >
            <X size={16} />
          </button>
        </header>

        <nav
          className="flex gap-1 border-b border-divider px-3 py-2"
          role="tablist"
          aria-label={t("xlsx.formatCells.tabsAriaLabel")}
        >
          {TABS.map((tabDef) => (
            <button
              key={tabDef.id}
              role="tab"
              aria-selected={tab === tabDef.id}
              data-testid={`format-tab-${tabDef.id}`}
              onClick={() => setTab(tabDef.id)}
              className={cn(
                "rounded-md px-3 py-1 text-xs font-medium transition-colors",
                tab === tabDef.id
                  ? "bg-[var(--ai-violet-light)] text-[var(--ai-violet)]"
                  : "text-secondary hover:bg-hover hover:text-foreground"
              )}
            >
              {t(tabDef.labelKey)}
            </button>
          ))}
        </nav>

        <div className="flex-1 overflow-y-auto px-5 py-4 text-sm text-foreground">
          {tab === "number" ? (
            <NumberTab state={numberPreset} setState={setNumberPreset} />
          ) : tab === "alignment" ? (
            <AlignmentTab
              horizontal={horizontal}
              vertical={vertical}
              wrapText={wrapText}
              indent={indent}
              setHorizontal={setHorizontal}
              setVertical={setVertical}
              setWrapText={setWrapText}
              setIndent={setIndent}
            />
          ) : tab === "font" ? (
            <FontTab
              family={fontFamily}
              size={fontSize}
              bold={bold}
              italic={italic}
              underline={underline}
              strike={strike}
              color={fontColor}
              setFamily={setFontFamily}
              setSize={setFontSize}
              setBold={setBold}
              setItalic={setItalic}
              setUnderline={setUnderline}
              setStrike={setStrike}
              setColor={setFontColor}
            />
          ) : tab === "border" ? (
            <BorderTab
              top={borderTop}
              right={borderRight}
              bottom={borderBottom}
              left={borderLeft}
              color={borderColor}
              setTop={setBorderTop}
              setRight={setBorderRight}
              setBottom={setBorderBottom}
              setLeft={setBorderLeft}
              setColor={setBorderColor}
            />
          ) : tab === "fill" ? (
            <FillTab color={fillColor} setColor={setFillColor} />
          ) : (
            <ProtectionTab locked={locked} hidden={hidden} setLocked={setLocked} setHidden={setHidden} />
          )}
        </div>

        <footer className="flex justify-end gap-2 border-t border-divider px-5 py-3">
          <button
            type="button"
            onClick={onClose}
            className="rounded-md border border-divider bg-surface px-3 py-1.5 text-xs font-medium text-foreground hover:bg-hover"
            data-testid="format-cells-cancel"
          >
            {t("common.cancel")}
          </button>
          <button
            type="button"
            onClick={apply}
            className="rounded-md bg-[var(--accent)] px-3 py-1.5 text-xs font-medium text-white hover:opacity-90"
            data-testid="format-cells-apply"
          >
            {t("common.apply")}
          </button>
        </footer>
      </div>
    </div>
  );
}

function NumberTab({
  state,
  setState,
}: {
  readonly state: NumberPresetState;
  readonly setState: (s: NumberPresetState) => void;
}): ReactNode {
  const { t } = useTranslator();
  return (
    <div className="flex flex-col gap-3">
      <div className="flex flex-col gap-1">
        <label className="text-xs font-medium text-secondary">{t("xlsx.formatCells.numberCategory")}</label>
        <div className="grid grid-cols-2 gap-1">
          {NUMBER_FORMAT_PRESETS.map((p) => (
            <button
              key={p.id}
              type="button"
              data-testid={`format-number-${p.id}`}
              onClick={() => setState({ id: p.id, customCode: p.code })}
              className={cn(
                "rounded-md border border-divider px-3 py-2 text-left text-xs hover:bg-hover",
                state.id === p.id && "border-[var(--accent)] bg-[var(--ai-violet-light)]"
              )}
            >
              {t(numberFormatPresetKey(p.id))}
            </button>
          ))}
          <button
            type="button"
            data-testid="format-number-custom"
            onClick={() => setState({ id: "custom", customCode: state.customCode || "0.00" })}
            className={cn(
              "rounded-md border border-divider px-3 py-2 text-left text-xs hover:bg-hover",
              state.id === "custom" && "border-[var(--accent)] bg-[var(--ai-violet-light)]"
            )}
          >
            {t("xlsx.formatCells.numberCustomFormat")}
          </button>
        </div>
      </div>
      {state.id === "custom" ? (
        <div className="flex flex-col gap-1">
          <label className="text-xs font-medium text-secondary">{t("xlsx.formatCells.numberCustomFormatCode")}</label>
          <input
            value={state.customCode}
            onChange={(e) => setState({ id: "custom", customCode: e.target.value })}
            className="rounded border border-divider bg-background px-2 py-1 font-mono text-xs text-foreground focus:border-[var(--accent)] focus:outline-none"
            placeholder={t("xlsx.formatCells.numberCustomFormatPlaceholder")}
            data-testid="format-number-custom-input"
          />
          <p className="text-[11px] text-tertiary">
            {t("xlsx.formatCells.numberCustomFormatHint")}
          </p>
        </div>
      ) : null}
    </div>
  );
}

function AlignmentTab({
  horizontal,
  vertical,
  wrapText,
  indent,
  setHorizontal,
  setVertical,
  setWrapText,
  setIndent,
}: {
  readonly horizontal: NonNullable<CellFormatPatch["alignment"]>["horizontal"] | "";
  readonly vertical: NonNullable<CellFormatPatch["alignment"]>["vertical"] | "";
  readonly wrapText: boolean;
  readonly indent: number;
  readonly setHorizontal: (h: NonNullable<CellFormatPatch["alignment"]>["horizontal"] | "") => void;
  readonly setVertical: (v: NonNullable<CellFormatPatch["alignment"]>["vertical"] | "") => void;
  readonly setWrapText: (b: boolean) => void;
  readonly setIndent: (n: number) => void;
}): ReactNode {
  const { t } = useTranslator();
  return (
    <div className="grid grid-cols-2 gap-4">
      <Field label={t("xlsx.formatCells.alignmentHorizontal")}>
        <select
          value={horizontal ?? ""}
          onChange={(e) => setHorizontal((e.target.value || "") as typeof horizontal)}
          className="rounded border border-divider bg-background px-2 py-1 text-xs text-foreground"
          data-testid="format-alignment-horizontal"
        >
          <option value="">{t("xlsx.formatCells.alignmentUnchanged")}</option>
          <option value="left">{t("xlsx.formatCells.alignmentLeft")}</option>
          <option value="center">{t("xlsx.formatCells.alignmentCenter")}</option>
          <option value="right">{t("xlsx.formatCells.alignmentRight")}</option>
          <option value="fill">{t("xlsx.formatCells.alignmentFill")}</option>
          <option value="justify">{t("xlsx.formatCells.alignmentJustify")}</option>
        </select>
      </Field>
      <Field label={t("xlsx.formatCells.alignmentVertical")}>
        <select
          value={vertical ?? ""}
          onChange={(e) => setVertical((e.target.value || "") as typeof vertical)}
          className="rounded border border-divider bg-background px-2 py-1 text-xs text-foreground"
          data-testid="format-alignment-vertical"
        >
          <option value="">{t("xlsx.formatCells.alignmentUnchanged")}</option>
          <option value="top">{t("xlsx.formatCells.alignmentTop")}</option>
          <option value="middle">{t("xlsx.formatCells.alignmentMiddle")}</option>
          <option value="bottom">{t("xlsx.formatCells.alignmentBottom")}</option>
        </select>
      </Field>
      <Field label={t("xlsx.formatCells.alignmentIndent")}>
        <input
          type="number"
          min={0}
          max={15}
          value={indent}
          onChange={(e) => setIndent(Number.parseInt(e.target.value, 10) || 0)}
          className="w-20 rounded border border-divider bg-background px-2 py-1 text-xs text-foreground"
          data-testid="format-alignment-indent"
        />
      </Field>
      <Field label=" ">
        <label className="inline-flex items-center gap-2 text-xs text-foreground">
          <input
            type="checkbox"
            checked={wrapText}
            onChange={(e) => setWrapText(e.target.checked)}
            data-testid="format-alignment-wrap"
          />
          {t("xlsx.formatCells.alignmentWrapText")}
        </label>
      </Field>
    </div>
  );
}

function FontTab(p: {
  readonly family: string;
  readonly size: number;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly underline: boolean;
  readonly strike: boolean;
  readonly color: string;
  readonly setFamily: (s: string) => void;
  readonly setSize: (n: number) => void;
  readonly setBold: (b: boolean) => void;
  readonly setItalic: (b: boolean) => void;
  readonly setUnderline: (b: boolean) => void;
  readonly setStrike: (b: boolean) => void;
  readonly setColor: (c: string) => void;
}): ReactNode {
  const { t } = useTranslator();
  return (
    <div className="flex flex-col gap-4">
      <div className="grid grid-cols-2 gap-4">
        <Field label={t("xlsx.formatCells.fontFamily")}>
          <select
            value={p.family}
            onChange={(e) => p.setFamily(e.target.value)}
            className="rounded border border-divider bg-background px-2 py-1 text-xs text-foreground"
            data-testid="format-font-family"
          >
            <option value="">{t("xlsx.formatCells.alignmentUnchanged")}</option>
            {FONT_FAMILIES.map((f) => (
              <option key={f} value={f}>
                {f}
              </option>
            ))}
          </select>
        </Field>
        <Field label={t("xlsx.formatCells.fontSize")}>
          <select
            value={p.size}
            onChange={(e) => p.setSize(Number.parseInt(e.target.value, 10) || 11)}
            className="w-24 rounded border border-divider bg-background px-2 py-1 text-xs text-foreground"
            data-testid="format-font-size"
          >
            {FONT_SIZES.map((s) => (
              <option key={s} value={s}>
                {s}
              </option>
            ))}
          </select>
        </Field>
      </div>
      <div className="flex flex-wrap items-center gap-3">
        <Toggle label={t("xlsx.formatCells.fontBold")} testId="format-font-bold" checked={p.bold} onChange={p.setBold} />
        <Toggle label={t("xlsx.formatCells.fontItalic")} testId="format-font-italic" checked={p.italic} onChange={p.setItalic} />
        <Toggle
          label={t("xlsx.formatCells.fontUnderline")}
          testId="format-font-underline"
          checked={p.underline}
          onChange={p.setUnderline}
        />
        <Toggle label={t("xlsx.formatCells.fontStrikethrough")} testId="format-font-strike" checked={p.strike} onChange={p.setStrike} />
      </div>
      <Field label={t("xlsx.formatCells.fontColor")}>
        <SwatchPicker
          color={p.color}
          onChange={p.setColor}
          swatches={FONT_COLOR_SWATCHES}
          allowClear
          testIdPrefix="format-font-color"
        />
      </Field>
    </div>
  );
}

function BorderTab(p: {
  readonly top: NonNullable<NonNullable<CellFormatPatch["border"]>["top"]>["style"] | undefined;
  readonly right: typeof p.top;
  readonly bottom: typeof p.top;
  readonly left: typeof p.top;
  readonly color: string;
  readonly setTop: (s: typeof p.top) => void;
  readonly setRight: (s: typeof p.top) => void;
  readonly setBottom: (s: typeof p.top) => void;
  readonly setLeft: (s: typeof p.top) => void;
  readonly setColor: (c: string) => void;
}): ReactNode {
  const { t } = useTranslator();
  const setAll = (style: typeof p.top) => {
    p.setTop(style);
    p.setRight(style);
    p.setBottom(style);
    p.setLeft(style);
  };
  return (
    <div className="flex flex-col gap-4">
      <div className="flex flex-wrap items-center gap-2">
        <span className="text-xs font-medium text-secondary">{t("xlsx.formatCells.borderApplyAllSides")}</span>
        <button
          type="button"
          onClick={() => setAll("thin")}
          className="rounded border border-divider bg-background px-2 py-1 text-xs text-foreground hover:bg-hover"
          data-testid="format-border-all"
        >
          {t("xlsx.formatCells.borderBox")}
        </button>
        <button
          type="button"
          onClick={() => setAll("none")}
          className="rounded border border-divider bg-background px-2 py-1 text-xs text-foreground hover:bg-hover"
          data-testid="format-border-none"
        >
          {t("xlsx.formatCells.borderStyleNone")}
        </button>
      </div>
      <div className="grid grid-cols-2 gap-4">
        <BorderSide label={t("xlsx.formatCells.borderTop")} testId="top" value={p.top} onChange={p.setTop} />
        <BorderSide label={t("xlsx.formatCells.borderRight")} testId="right" value={p.right} onChange={p.setRight} />
        <BorderSide label={t("xlsx.formatCells.borderBottom")} testId="bottom" value={p.bottom} onChange={p.setBottom} />
        <BorderSide label={t("xlsx.formatCells.borderLeft")} testId="left" value={p.left} onChange={p.setLeft} />
      </div>
      <Field label={t("xlsx.formatCells.borderColor")}>
        <SwatchPicker
          color={p.color}
          onChange={p.setColor}
          swatches={FONT_COLOR_SWATCHES}
          testIdPrefix="format-border-color"
        />
      </Field>
    </div>
  );
}

function BorderSide({
  label,
  testId,
  value,
  onChange,
}: {
  readonly label: string;
  readonly testId: string;
  readonly value: NonNullable<NonNullable<CellFormatPatch["border"]>["top"]>["style"] | undefined;
  readonly onChange: (
    s: NonNullable<NonNullable<CellFormatPatch["border"]>["top"]>["style"] | undefined
  ) => void;
}): ReactNode {
  const { t } = useTranslator();
  return (
    <Field label={label}>
      <select
        value={value ?? ""}
        onChange={(e) => onChange((e.target.value || undefined) as typeof value)}
        className="rounded border border-divider bg-background px-2 py-1 text-xs text-foreground"
        data-testid={`format-border-${testId}`}
      >
        <option value="">{t("xlsx.formatCells.alignmentUnchanged")}</option>
        {BORDER_STYLES.map((s) => (
          <option key={s.id} value={s.id}>
            {t(s.labelKey)}
          </option>
        ))}
      </select>
    </Field>
  );
}

function FillTab({
  color,
  setColor,
}: {
  readonly color: string;
  readonly setColor: (c: string) => void;
}): ReactNode {
  const { t } = useTranslator();
  return (
    <div className="flex flex-col gap-3">
      <Field label={t("xlsx.formatCells.fillBackgroundColor")}>
        <SwatchPicker
          color={color}
          onChange={setColor}
          swatches={FILL_SWATCHES}
          allowClear
          testIdPrefix="format-fill-color"
        />
      </Field>
      <p className="text-[11px] text-tertiary">{t("xlsx.formatCells.fillHint")}</p>
    </div>
  );
}

function ProtectionTab({
  locked,
  hidden,
  setLocked,
  setHidden,
}: {
  readonly locked: boolean;
  readonly hidden: boolean;
  readonly setLocked: (b: boolean) => void;
  readonly setHidden: (b: boolean) => void;
}): ReactNode {
  const { t } = useTranslator();
  return (
    <div className="flex flex-col gap-3">
      <Toggle label={t("xlsx.formatCells.protectionLockedLabel")} testId="format-protection-locked" checked={locked} onChange={setLocked} />
      <Toggle label={t("xlsx.formatCells.protectionHiddenLabel")} testId="format-protection-hidden" checked={hidden} onChange={setHidden} />
      <p className="text-[11px] text-tertiary">{t("xlsx.formatCells.protectionHint")}</p>
    </div>
  );
}

function Field({ label, children }: { readonly label: string; readonly children: ReactNode }): ReactNode {
  return (
    <label className="flex flex-col gap-1">
      <span className="text-xs font-medium text-secondary">{label}</span>
      {children}
    </label>
  );
}

function Toggle({
  label,
  testId,
  checked,
  onChange,
}: {
  readonly label: string;
  readonly testId: string;
  readonly checked: boolean;
  readonly onChange: (b: boolean) => void;
}): ReactNode {
  return (
    <label className="inline-flex items-center gap-2 text-xs text-foreground">
      <input
        type="checkbox"
        checked={checked}
        onChange={(e) => onChange(e.target.checked)}
        data-testid={testId}
      />
      {label}
    </label>
  );
}

function SwatchPicker({
  color,
  onChange,
  swatches,
  allowClear,
  testIdPrefix,
}: {
  readonly color: string;
  readonly onChange: (c: string) => void;
  readonly swatches: ReadonlyArray<string>;
  readonly allowClear?: boolean;
  readonly testIdPrefix: string;
}): ReactNode {
  const { t } = useTranslator();
  return (
    <div className="flex flex-wrap items-center gap-1.5">
      {allowClear ? (
        <button
          type="button"
          onClick={() => onChange("")}
          className={cn(
            "h-6 w-6 rounded border text-[10px]",
            color === "" ? "border-[var(--accent)]" : "border-divider"
          )}
          data-testid={`${testIdPrefix}-clear`}
          title={t("xlsx.formatCells.clear")}
        >
          ✕
        </button>
      ) : null}
      {swatches.map((hex) => (
        <button
          key={hex}
          type="button"
          onClick={() => onChange(hex)}
          className={cn(
            "h-6 w-6 rounded border",
            color.toUpperCase() === hex.toUpperCase()
              ? "border-[var(--accent)] ring-2 ring-[var(--accent)]"
              : "border-divider"
          )}
          style={{ background: `#${hex}` }}
          aria-label={`#${hex}`}
          title={`#${hex}`}
          data-testid={`${testIdPrefix}-${hex}`}
        />
      ))}
      <input
        type="text"
        value={color}
        onChange={(e) => onChange(normalizeRgb(e.target.value))}
        placeholder={t("xlsx.formatCells.rgbPlaceholder")}
        className="ml-1 w-20 rounded border border-divider bg-background px-1.5 py-0.5 font-mono text-[11px] text-foreground"
        data-testid={`${testIdPrefix}-input`}
      />
    </div>
  );
}

function normalizeRgb(input: string): string {
  const trimmed = input.trim().replace(/^#/, "").toUpperCase();
  if (/^[0-9A-F]{6}$/.test(trimmed)) return trimmed;
  if (/^[0-9A-F]{8}$/.test(trimmed)) return trimmed.slice(2);
  return input.replace(/^#/, "").toUpperCase();
}

function borderToStyle(
  ooxmlStyle: string | undefined
): NonNullable<NonNullable<CellFormatPatch["border"]>["top"]>["style"] | undefined {
  if (!ooxmlStyle) return undefined;
  switch (ooxmlStyle) {
    case "thin":
    case "medium":
    case "thick":
    case "double":
    case "dashed":
    case "dotted":
    case "none":
      return ooxmlStyle;
    default:
      return undefined;
  }
}

function numFmtIdMatches(id: number, code: string): boolean {
  // Only the built-in slots 0..49 have a stable id ↔ code mapping we can
  // reverse here. Custom numFmtIds (≥ 164) come from the workbook's own
  // table; the dialog falls through to "Custom format code…" for those.
  switch (id) {
    case 0:
      return code === "General";
    case 4:
      return code === "#,##0.00";
    case 9:
      return code === "0%";
    case 10:
      return code === "0.00%";
    case 14:
      return code === "yyyy-mm-dd" || code === "m/d/yy";
    default:
      return false;
  }
}

function numFmtIdToCustomGuess(id: number): string {
  switch (id) {
    case 0:
      return "General";
    case 1:
      return "0";
    case 2:
      return "0.00";
    case 3:
      return "#,##0";
    case 4:
      return "#,##0.00";
    case 9:
      return "0%";
    case 10:
      return "0.00%";
    case 14:
      return "m/d/yy";
    default:
      return "";
  }
}
