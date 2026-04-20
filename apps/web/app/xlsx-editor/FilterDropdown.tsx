"use client";

import { useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import {
  cellKey,
  flattenCellXf,
  type AutoFilter,
  type CustomFilterOp,
  type DynamicFilterType,
  type FilterColumn,
  type Sheet,
  type StyleTable,
} from "@officeai/xlsx";
import { formatCellValue } from "./styles";
import { useTranslator } from "@/lib/i18n";

/**
 * Excel-style header dropdown.
 *
 * Layout (top → bottom):
 *   1. Sort A→Z / Sort Z→A
 *   2. "Clear filter from <Column>"
 *   3. Filter type submenu (Text / Number / Date) — auto-detected by
 *      column data shape (date number-format → Date; numeric column →
 *      Number; otherwise Text). Each opens a small criterion editor.
 *   4. Color filter — distinct fill colours present in the column.
 *   5. Search box + virtualized value checklist with `(Select All)`
 *      and `(Blanks)` toggles.
 *   6. Apply / Cancel.
 *
 * The dropdown is fully presentational: callbacks raise the chosen
 * action and the parent dispatches the matching command. The parent
 * is also responsible for closing the dropdown after a confirmed
 * action.
 */
export interface FilterDropdownProps {
  readonly open: boolean;
  readonly sheet: Sheet;
  readonly styles: StyleTable;
  readonly autoFilter: AutoFilter;
  readonly colId: number;
  readonly anchor: DOMRect | null;
  readonly onClose: () => void;
  readonly onSort: (order: "asc" | "desc") => void;
  readonly onClear: () => void;
  readonly onApply: (criterion: FilterColumn) => void;
}

type Mode = "values" | "text" | "number" | "date" | "color";

const DATE_NUM_FMT_IDS = new Set([14, 15, 16, 17, 22]);

const DYNAMIC_LABELS: ReadonlyArray<{ readonly id: DynamicFilterType; readonly labelKey: string }> = [
  { id: "today", labelKey: "xlsx.filter.dynToday" },
  { id: "yesterday", labelKey: "xlsx.filter.dynYesterday" },
  { id: "tomorrow", labelKey: "xlsx.filter.dynTomorrow" },
  { id: "thisWeek", labelKey: "xlsx.filter.dynThisWeek" },
  { id: "lastWeek", labelKey: "xlsx.filter.dynLastWeek" },
  { id: "nextWeek", labelKey: "xlsx.filter.dynNextWeek" },
  { id: "thisMonth", labelKey: "xlsx.filter.dynThisMonth" },
  { id: "lastMonth", labelKey: "xlsx.filter.dynLastMonth" },
  { id: "nextMonth", labelKey: "xlsx.filter.dynNextMonth" },
  { id: "thisQuarter", labelKey: "xlsx.filter.dynThisQuarter" },
  { id: "lastQuarter", labelKey: "xlsx.filter.dynLastQuarter" },
  { id: "nextQuarter", labelKey: "xlsx.filter.dynNextQuarter" },
  { id: "thisYear", labelKey: "xlsx.filter.dynThisYear" },
  { id: "lastYear", labelKey: "xlsx.filter.dynLastYear" },
  { id: "nextYear", labelKey: "xlsx.filter.dynNextYear" },
  { id: "yearToDate", labelKey: "xlsx.filter.dynYearToDate" },
];

const TEXT_OPS: ReadonlyArray<{ readonly id: CustomFilterOp["operator"]; readonly labelKey: string }> = [
  { id: "equal", labelKey: "xlsx.filter.opEquals" },
  { id: "notEqual", labelKey: "xlsx.filter.opDoesNotEqual" },
];

const NUMBER_OPS: ReadonlyArray<{ readonly id: CustomFilterOp["operator"]; readonly labelKey: string }> = [
  { id: "equal", labelKey: "xlsx.filter.opEquals" },
  { id: "notEqual", labelKey: "xlsx.filter.opDoesNotEqual" },
  { id: "greaterThan", labelKey: "xlsx.filter.opGreaterThan" },
  { id: "greaterThanOrEqual", labelKey: "xlsx.filter.opGreaterThanOrEqual" },
  { id: "lessThan", labelKey: "xlsx.filter.opLessThan" },
  { id: "lessThanOrEqual", labelKey: "xlsx.filter.opLessThanOrEqual" },
];

interface ColumnSummary {
  readonly values: ReadonlyArray<{ readonly key: string; readonly count: number }>;
  readonly hasBlanks: boolean;
  readonly numericCount: number;
  readonly dateCount: number;
  readonly stringCount: number;
  readonly fillColors: ReadonlyArray<string>;
}

function summarizeColumn(
  sheet: Sheet,
  styles: StyleTable,
  autoFilter: AutoFilter,
  colId: number
): ColumnSummary {
  const counts = new Map<string, number>();
  const colors = new Set<string>();
  let hasBlanks = false;
  let numericCount = 0;
  let dateCount = 0;
  let stringCount = 0;
  const absCol = autoFilter.range.c1 + colId;
  for (let r = autoFilter.range.r1 + 1; r <= autoFilter.range.r2; r++) {
    const cell = sheet.cells.get(cellKey(r, absCol));
    if (!cell || cell.value === null || cell.value === "") {
      hasBlanks = true;
      continue;
    }
    const eff = flattenCellXf(styles, cell.styleId);
    if (typeof cell.value === "number") {
      numericCount += 1;
      if (DATE_NUM_FMT_IDS.has(eff.numFmtId)) dateCount += 1;
    } else if (typeof cell.value === "string") {
      stringCount += 1;
    }
    const display = formatCellValue(cell.value, eff.numFmtId);
    counts.set(display, (counts.get(display) ?? 0) + 1);
    if (eff.fill.kind === "pattern" && eff.fill.fgColor?.rgb) {
      colors.add(eff.fill.fgColor.rgb.toUpperCase());
    }
  }
  const values = [...counts.entries()]
    .map(([key, count]) => ({ key, count }))
    .sort((a, b) => a.key.localeCompare(b.key, undefined, { sensitivity: "base", numeric: true }));
  return {
    values,
    hasBlanks,
    numericCount,
    dateCount,
    stringCount,
    fillColors: [...colors].sort(),
  };
}

function detectMode(summary: ColumnSummary): Exclude<Mode, "values" | "color"> {
  if (summary.dateCount > 0 && summary.dateCount >= summary.stringCount) return "date";
  if (summary.numericCount >= summary.stringCount) return "number";
  return "text";
}

export function FilterDropdown(props: FilterDropdownProps): ReactNode {
  const { open, sheet, styles, autoFilter, colId, anchor, onClose, onSort, onClear, onApply } = props;
  const { t } = useTranslator();
  const ref = useRef<HTMLDivElement | null>(null);

  const summary = useMemo(
    () => (open ? summarizeColumn(sheet, styles, autoFilter, colId) : null),
    [open, sheet, styles, autoFilter, colId]
  );

  const existing = autoFilter.columns.get(colId);
  const initialChecked = useMemo<ReadonlySet<string>>(() => {
    if (!summary) return new Set();
    if (existing?.kind === "values") return new Set(existing.values);
    return new Set(summary.values.map((v) => v.key));
  }, [summary, existing]);

  const [checked, setChecked] = useState<ReadonlySet<string>>(initialChecked);
  const [includeBlanks, setIncludeBlanks] = useState<boolean>(
    existing?.kind === "values" ? existing.blank : true
  );
  const [search, setSearch] = useState<string>("");
  const [mode, setMode] = useState<Mode>("values");

  const detectedConditionMode = summary ? detectMode(summary) : "text";

  // Custom-filter editor state
  const [op1, setOp1] = useState<CustomFilterOp["operator"]>("equal");
  const [val1, setVal1] = useState<string>("");
  const [op2, setOp2] = useState<CustomFilterOp["operator"] | "">("");
  const [val2, setVal2] = useState<string>("");
  const [combine, setCombine] = useState<"and" | "or">("and");

  // Top10
  const [topMode, setTopMode] = useState<"top" | "bottom">("top");
  const [topN, setTopN] = useState<number>(10);
  const [topPercent, setTopPercent] = useState<boolean>(false);

  // Reset transient state whenever the dropdown reopens for a new column.
  useEffect(() => {
    if (!open) return;
    setChecked(initialChecked);
    setSearch("");
    setMode("values");
  }, [open, initialChecked]);

  // Click-outside / Escape to dismiss.
  useEffect(() => {
    if (!open) return;
    const onMouseDown = (e: MouseEvent) => {
      if (!ref.current) return;
      if (ref.current.contains(e.target as Node)) return;
      onClose();
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") onClose();
    };
    window.addEventListener("mousedown", onMouseDown);
    window.addEventListener("keydown", onKey);
    return () => {
      window.removeEventListener("mousedown", onMouseDown);
      window.removeEventListener("keydown", onKey);
    };
  }, [open, onClose]);

  if (!open || !summary || !anchor) return null;

  const filteredValues = search
    ? summary.values.filter((v) => v.key.toLowerCase().includes(search.toLowerCase()))
    : summary.values;
  const allShownChecked = filteredValues.length > 0 && filteredValues.every((v) => checked.has(v.key));
  const someShownChecked = filteredValues.some((v) => checked.has(v.key));

  const top = Math.min(window.innerHeight - 480, anchor.bottom + 4);
  const left = Math.min(window.innerWidth - 320, anchor.left);

  function applyValues(): void {
    onApply({ kind: "values", values: new Set(checked), blank: includeBlanks });
  }

  function applyCustom(): void {
    if (val1.length === 0) return;
    const c: FilterColumn = {
      kind: "custom",
      op1: { operator: op1, val: val1 },
      ...(op2 && val2.length > 0 ? { op2: { operator: op2, val: val2 } } : {}),
      combine,
    };
    onApply(c);
  }

  function applyTop10(): void {
    onApply({
      kind: "top10",
      top: topMode === "top",
      percent: topPercent,
      n: Math.max(1, Math.min(500, Math.floor(topN))),
      filterVal: 0,
    });
  }

  function applyDynamic(type: DynamicFilterType): void {
    onApply({ kind: "dynamic", type });
  }

  function applyColor(argb: string): void {
    onApply({ kind: "color", argb, isCellColor: true });
  }

  return (
    <div
      ref={ref}
      role="dialog"
      data-testid="filter-dropdown"
      style={{
        position: "fixed",
        top,
        left,
        width: 280,
        maxHeight: 460,
        zIndex: 50,
        background: "var(--surface)",
        border: "1px solid var(--divider)",
        borderRadius: 4,
        boxShadow: "0 4px 12px rgba(0,0,0,0.15)",
        display: "flex",
        flexDirection: "column",
        fontSize: 12,
      }}
    >
      <div style={{ display: "flex", flexDirection: "column", borderBottom: "1px solid var(--divider)" }}>
        <DropdownButton onClick={() => onSort("asc")} testId="filter-sort-asc">
          {t("xlsx.filter.sortAZ")}
        </DropdownButton>
        <DropdownButton onClick={() => onSort("desc")} testId="filter-sort-desc">
          {t("xlsx.filter.sortZA")}
        </DropdownButton>
        {existing ? (
          <DropdownButton onClick={onClear} testId="filter-clear">
            {t("xlsx.filter.clearFilter")}
          </DropdownButton>
        ) : null}
      </div>

      <div style={{ display: "flex", borderBottom: "1px solid var(--divider)" }}>
        <ModeTab id="values" current={mode} onClick={setMode} label={t("xlsx.filter.tabValues")} />
        <ModeTab
          id={detectedConditionMode}
          current={mode}
          onClick={setMode}
          label={t(modeLabelKey(detectedConditionMode))}
        />
        {summary.fillColors.length > 0 ? (
          <ModeTab id="color" current={mode} onClick={setMode} label={t("xlsx.filter.tabColor")} />
        ) : null}
      </div>

      <div style={{ flex: 1, overflowY: "auto", padding: 8 }}>
        {mode === "values" ? (
          <ValuesPanel
            values={filteredValues}
            checked={checked}
            allShown={allShownChecked}
            someShown={someShownChecked}
            includeBlanks={includeBlanks}
            hasBlanks={summary.hasBlanks}
            search={search}
            onSearch={setSearch}
            onToggleAll={() => {
              const next = new Set(checked);
              if (allShownChecked) {
                for (const v of filteredValues) next.delete(v.key);
              } else {
                for (const v of filteredValues) next.add(v.key);
              }
              setChecked(next);
            }}
            onToggleOne={(key) => {
              const next = new Set(checked);
              if (next.has(key)) next.delete(key);
              else next.add(key);
              setChecked(next);
            }}
            onToggleBlanks={() => setIncludeBlanks((b) => !b)}
          />
        ) : null}
        {mode === "text" || mode === "number" ? (
          <CustomPanel
            ops={mode === "text" ? TEXT_OPS : NUMBER_OPS}
            op1={op1}
            val1={val1}
            op2={op2}
            val2={val2}
            combine={combine}
            onChange={(p) => {
              if (p.op1 !== undefined) setOp1(p.op1);
              if (p.val1 !== undefined) setVal1(p.val1);
              if (p.op2 !== undefined) setOp2(p.op2);
              if (p.val2 !== undefined) setVal2(p.val2);
              if (p.combine !== undefined) setCombine(p.combine);
            }}
            top10Section={
              mode === "number" ? (
                <Top10Panel
                  topMode={topMode}
                  setTopMode={setTopMode}
                  topN={topN}
                  setTopN={setTopN}
                  topPercent={topPercent}
                  setTopPercent={setTopPercent}
                  onApply={applyTop10}
                />
              ) : null
            }
          />
        ) : null}
        {mode === "date" ? <DatePanel onApply={applyDynamic} /> : null}
        {mode === "color" ? <ColorPanel colors={summary.fillColors} onApply={applyColor} /> : null}
      </div>

      <div
        style={{
          display: "flex",
          justifyContent: "flex-end",
          gap: 6,
          padding: 8,
          borderTop: "1px solid var(--divider)",
        }}
      >
        <button type="button" onClick={onClose} data-testid="filter-cancel" style={btnStyle("secondary")}>
          {t("common.cancel")}
        </button>
        <button
          type="button"
          data-testid="filter-apply"
          style={btnStyle("primary")}
          onClick={() => {
            if (mode === "values") applyValues();
            else if (mode === "text" || mode === "number") applyCustom();
          }}
          disabled={(mode === "text" || mode === "number") && val1.length === 0}
        >
          {t("common.apply")}
        </button>
      </div>
    </div>
  );
}

function modeLabelKey(mode: Exclude<Mode, "values" | "color">): string {
  switch (mode) {
    case "text":
      return "xlsx.filter.tabText";
    case "number":
      return "xlsx.filter.tabNumber";
    case "date":
      return "xlsx.filter.tabDate";
  }
}

function DropdownButton(props: { children: ReactNode; onClick: () => void; testId?: string }): ReactNode {
  return (
    <button
      type="button"
      onClick={props.onClick}
      data-testid={props.testId}
      style={{
        textAlign: "left",
        padding: "6px 10px",
        background: "transparent",
        border: 0,
        cursor: "pointer",
        color: "var(--foreground)",
        fontSize: 12,
      }}
      onMouseEnter={(e) => (e.currentTarget.style.background = "var(--hover)")}
      onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
    >
      {props.children}
    </button>
  );
}

function ModeTab(props: { id: Mode; current: Mode; onClick: (m: Mode) => void; label: string }): ReactNode {
  const active = props.id === props.current;
  return (
    <button
      type="button"
      onClick={() => props.onClick(props.id)}
      data-testid={`filter-tab-${props.id}`}
      style={{
        flex: 1,
        padding: "6px 8px",
        background: active ? "var(--hover)" : "transparent",
        border: 0,
        borderBottom: active ? "2px solid var(--accent, #2563eb)" : "2px solid transparent",
        cursor: "pointer",
        fontSize: 11,
        color: "var(--foreground)",
      }}
    >
      {props.label}
    </button>
  );
}

interface ValuesPanelProps {
  readonly values: ReadonlyArray<{ readonly key: string; readonly count: number }>;
  readonly checked: ReadonlySet<string>;
  readonly allShown: boolean;
  readonly someShown: boolean;
  readonly includeBlanks: boolean;
  readonly hasBlanks: boolean;
  readonly search: string;
  readonly onSearch: (s: string) => void;
  readonly onToggleAll: () => void;
  readonly onToggleOne: (k: string) => void;
  readonly onToggleBlanks: () => void;
}

function ValuesPanel(props: ValuesPanelProps): ReactNode {
  const { t } = useTranslator();
  return (
    <div>
      <input
        type="text"
        placeholder={t("xlsx.filter.searchPlaceholder")}
        value={props.search}
        onChange={(e) => props.onSearch(e.target.value)}
        data-testid="filter-search"
        style={{
          width: "100%",
          height: 26,
          padding: "0 6px",
          border: "1px solid var(--divider)",
          borderRadius: 3,
          fontSize: 12,
          marginBottom: 6,
          background: "var(--background)",
          color: "var(--foreground)",
        }}
      />
      <label style={{ display: "flex", alignItems: "center", gap: 6, padding: "2px 4px" }}>
        <input
          type="checkbox"
          checked={props.allShown}
          ref={(el) => {
            if (el) el.indeterminate = !props.allShown && props.someShown;
          }}
          onChange={props.onToggleAll}
          data-testid="filter-select-all"
        />
        <span style={{ fontWeight: 500 }}>{t("xlsx.filter.selectAll")}</span>
      </label>
      {props.hasBlanks ? (
        <label style={{ display: "flex", alignItems: "center", gap: 6, padding: "2px 4px" }}>
          <input
            type="checkbox"
            checked={props.includeBlanks}
            onChange={props.onToggleBlanks}
            data-testid="filter-blanks"
          />
          <span style={{ fontStyle: "italic" }}>{t("xlsx.filter.blanks")}</span>
        </label>
      ) : null}
      <div style={{ maxHeight: 200, overflowY: "auto", marginTop: 4 }}>
        {props.values.map((v) => (
          <label
            key={v.key}
            style={{
              display: "flex",
              alignItems: "center",
              gap: 6,
              padding: "2px 4px",
              whiteSpace: "nowrap",
              overflow: "hidden",
              textOverflow: "ellipsis",
            }}
          >
            <input
              type="checkbox"
              checked={props.checked.has(v.key)}
              onChange={() => props.onToggleOne(v.key)}
              data-testid={`filter-value-${v.key}`}
            />
            <span style={{ overflow: "hidden", textOverflow: "ellipsis" }}>{v.key || "—"}</span>
            <span style={{ marginLeft: "auto", color: "var(--secondary)" }}>{v.count}</span>
          </label>
        ))}
      </div>
    </div>
  );
}

interface CustomPanelProps {
  readonly ops: ReadonlyArray<{ readonly id: CustomFilterOp["operator"]; readonly labelKey: string }>;
  readonly op1: CustomFilterOp["operator"];
  readonly val1: string;
  readonly op2: CustomFilterOp["operator"] | "";
  readonly val2: string;
  readonly combine: "and" | "or";
  readonly onChange: (p: {
    op1?: CustomFilterOp["operator"];
    val1?: string;
    op2?: CustomFilterOp["operator"] | "";
    val2?: string;
    combine?: "and" | "or";
  }) => void;
  readonly top10Section: ReactNode;
}

function CustomPanel(props: CustomPanelProps): ReactNode {
  const { t } = useTranslator();
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
      <OpRow
        ops={props.ops}
        op={props.op1}
        val={props.val1}
        onOpChange={(o) => props.onChange({ op1: o })}
        onValChange={(v) => props.onChange({ val1: v })}
        testIdPrefix="filter-op1"
      />
      <div style={{ display: "flex", gap: 8, alignItems: "center", padding: "0 4px" }}>
        <label style={{ display: "flex", alignItems: "center", gap: 4 }}>
          <input
            type="radio"
            checked={props.combine === "and"}
            onChange={() => props.onChange({ combine: "and" })}
          />
          {t("xlsx.filter.and")}
        </label>
        <label style={{ display: "flex", alignItems: "center", gap: 4 }}>
          <input
            type="radio"
            checked={props.combine === "or"}
            onChange={() => props.onChange({ combine: "or" })}
          />
          {t("xlsx.filter.or")}
        </label>
      </div>
      <OpRow
        ops={props.ops}
        op={props.op2}
        val={props.val2}
        onOpChange={(o) => props.onChange({ op2: o })}
        onValChange={(v) => props.onChange({ val2: v })}
        allowEmpty
        testIdPrefix="filter-op2"
      />
      {props.top10Section}
    </div>
  );
}

function OpRow(props: {
  ops: ReadonlyArray<{ readonly id: CustomFilterOp["operator"]; readonly labelKey: string }>;
  op: CustomFilterOp["operator"] | "";
  val: string;
  onOpChange: (o: CustomFilterOp["operator"]) => void;
  onValChange: (v: string) => void;
  allowEmpty?: boolean;
  testIdPrefix: string;
}): ReactNode {
  const { t } = useTranslator();
  return (
    <div style={{ display: "flex", gap: 4 }}>
      <select
        value={props.op}
        onChange={(e) => props.onOpChange(e.target.value as CustomFilterOp["operator"])}
        data-testid={`${props.testIdPrefix}-op`}
        style={{
          flex: 1,
          height: 26,
          fontSize: 11,
          background: "var(--background)",
          color: "var(--foreground)",
          border: "1px solid var(--divider)",
        }}
      >
        {props.allowEmpty ? <option value="">—</option> : null}
        {props.ops.map((o) => (
          <option key={o.id} value={o.id}>
            {t(o.labelKey)}
          </option>
        ))}
      </select>
      <input
        type="text"
        value={props.val}
        onChange={(e) => props.onValChange(e.target.value)}
        data-testid={`${props.testIdPrefix}-val`}
        style={{
          flex: 1,
          height: 26,
          fontSize: 12,
          padding: "0 6px",
          border: "1px solid var(--divider)",
          borderRadius: 3,
          background: "var(--background)",
          color: "var(--foreground)",
        }}
      />
    </div>
  );
}

function Top10Panel(props: {
  topMode: "top" | "bottom";
  setTopMode: (m: "top" | "bottom") => void;
  topN: number;
  setTopN: (n: number) => void;
  topPercent: boolean;
  setTopPercent: (p: boolean) => void;
  onApply: () => void;
}): ReactNode {
  const { t } = useTranslator();
  return (
    <div
      style={{
        display: "flex",
        flexDirection: "column",
        gap: 4,
        marginTop: 6,
        paddingTop: 6,
        borderTop: "1px solid var(--divider)",
      }}
    >
      <div style={{ fontWeight: 500, padding: "0 4px" }}>{t("xlsx.filter.top10")}</div>
      <div style={{ display: "flex", gap: 4, alignItems: "center", padding: "0 4px" }}>
        <select
          value={props.topMode}
          onChange={(e) => props.setTopMode(e.target.value as "top" | "bottom")}
          style={{
            height: 24,
            background: "var(--background)",
            color: "var(--foreground)",
            border: "1px solid var(--divider)",
          }}
          data-testid="filter-top-mode"
        >
          <option value="top">{t("xlsx.filter.top")}</option>
          <option value="bottom">{t("xlsx.filter.bottom")}</option>
        </select>
        <input
          type="number"
          value={props.topN}
          min={1}
          max={500}
          onChange={(e) => props.setTopN(Number(e.target.value))}
          data-testid="filter-top-n"
          style={{
            width: 60,
            height: 24,
            border: "1px solid var(--divider)",
            background: "var(--background)",
            color: "var(--foreground)",
            padding: "0 4px",
          }}
        />
        <select
          value={props.topPercent ? "percent" : "items"}
          onChange={(e) => props.setTopPercent(e.target.value === "percent")}
          style={{
            height: 24,
            background: "var(--background)",
            color: "var(--foreground)",
            border: "1px solid var(--divider)",
          }}
          data-testid="filter-top-unit"
        >
          <option value="items">{t("xlsx.filter.items")}</option>
          <option value="percent">{t("xlsx.filter.percent")}</option>
        </select>
        <button
          type="button"
          onClick={props.onApply}
          data-testid="filter-top-apply"
          style={btnStyle("primary")}
        >
          {t("common.apply")}
        </button>
      </div>
    </div>
  );
}

function DatePanel(props: { onApply: (t: DynamicFilterType) => void }): ReactNode {
  const { t } = useTranslator();
  return (
    <div style={{ display: "flex", flexDirection: "column" }}>
      {DYNAMIC_LABELS.map((d) => (
        <DropdownButton key={d.id} testId={`filter-dynamic-${d.id}`} onClick={() => props.onApply(d.id)}>
          {t(d.labelKey)}
        </DropdownButton>
      ))}
    </div>
  );
}

function ColorPanel(props: { colors: ReadonlyArray<string>; onApply: (argb: string) => void }): ReactNode {
  const { t } = useTranslator();
  if (props.colors.length === 0) {
    return <div style={{ color: "var(--secondary)", padding: 6 }}>{t("xlsx.filter.noColors")}</div>;
  }
  return (
    <div style={{ display: "flex", flexWrap: "wrap", gap: 6, padding: 4 }}>
      {props.colors.map((c) => (
        <button
          key={c}
          type="button"
          data-testid={`filter-color-${c}`}
          onClick={() => props.onApply(c)}
          aria-label={t("xlsx.filter.byColor", { argb: c })}
          style={{
            width: 24,
            height: 24,
            borderRadius: 3,
            border: "1px solid var(--divider)",
            background: `#${c.length === 8 ? c.slice(2) : c}`,
            cursor: "pointer",
          }}
        />
      ))}
    </div>
  );
}

function btnStyle(variant: "primary" | "secondary"): React.CSSProperties {
  if (variant === "primary") {
    return {
      height: 26,
      padding: "0 10px",
      fontSize: 12,
      fontWeight: 500,
      borderRadius: 3,
      border: 0,
      background: "var(--ai-violet)",
      color: "var(--on-accent)",
      cursor: "pointer",
    };
  }
  return {
    height: 26,
    padding: "0 10px",
    fontSize: 12,
    borderRadius: 3,
    border: "1px solid var(--divider)",
    background: "var(--background)",
    color: "var(--foreground)",
    cursor: "pointer",
  };
}
