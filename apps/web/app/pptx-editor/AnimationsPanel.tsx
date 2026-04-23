"use client";

import * as React from "react";
import {
  ChevronDown,
  ChevronUp,
  Clipboard,
  Copy,
  CornerDownLeft,
  CornerUpLeft,
  Hand,
  MousePointerClick,
  Play,
  Plus,
  Sparkles,
  Trash2,
  Zap,
} from "lucide-react";
import {
  presetsByCategory,
  type AnimationCategory,
  type AnimationDirection,
  type AnimationPreset,
  type AnimationTrigger,
  type PresetSpec,
  type PptxSnapshot,
  type Shape,
  type ShapeAnimation,
  type Slide,
  type SlideTransition,
  type TransitionKind,
  type TransitionSpeed,
} from "@officeai/pptx";
import { useTranslator } from "@/lib/i18n";

/**
 * F4 v2 — Animations panel mounted in the right rail when the active
 * product is PPTX.
 *
 * Three sections:
 *   1. Slide transition (unchanged from the original panel)
 *   2. Preset gallery — category tabs + tile picker driven entirely by
 *      `presetsByCategory()` so adding a new preset to the registry
 *      shows up here for free
 *   3. Animations on this slide — list with reorder / remove buttons
 *      and an inline Effect Options popover for editing direction,
 *      duration, delay and trigger of an existing animation
 *
 * The panel uses the typed `(category, preset, …)` payload of
 * `pptx:add-shape-animation` and the new `pptx:set-shape-animation`
 * command. The legacy `effect: EntranceEffect` shortcut is gone.
 */
export interface AnimationsPanelProps {
  readonly snapshot: PptxSnapshot;
  readonly activeIndex: number;
  readonly selectedShape: Shape | null;
  readonly disabled: boolean;
  readonly onSetTransition: (kind: TransitionKind, speed: TransitionSpeed | null) => void;
  readonly onAddAnimation: (params: AddAnimationParams) => void;
  readonly onSetAnimation: (params: SetAnimationParams) => void;
  readonly onRemoveAnimation: (animationId: string) => void;
  readonly onReorderAnimations: (orderIds: ReadonlyArray<string>) => void;
  /** Optional Phase 4 hook: play the given animation (or full sequence when omitted). */
  readonly onPreviewAnimation?: (animationId: string | null) => void;
}

export interface AddAnimationParams {
  readonly category: AnimationCategory;
  readonly preset: AnimationPreset;
  readonly direction?: AnimationDirection;
  readonly trigger?: AnimationTrigger;
  readonly durationMs?: number;
  readonly delayMs?: number;
}

export interface SetAnimationParams {
  readonly animationId: string;
  readonly category?: AnimationCategory;
  readonly preset?: AnimationPreset;
  readonly direction?: AnimationDirection | null;
  readonly trigger?: AnimationTrigger;
  readonly durationMs?: number | null;
  readonly delayMs?: number | null;
}

const TRANSITION_LABEL_KEY: Record<"none" | "fade" | "push" | "wipe" | "split" | "cut", string> = {
  none: "transitionNone",
  fade: "transitionFade",
  push: "transitionPush",
  wipe: "transitionWipe",
  split: "transitionSplit",
  cut: "transitionCut",
};

const TRANSITION_SPEED_KEY: Record<TransitionSpeed, string> = {
  slow: "speedSlow",
  med: "speedMedium",
  fast: "speedFast",
};

const CAT_KEY: Record<AnimationCategory, string> = {
  entrance: "categoryEntrance",
  emphasis: "categoryEmphasis",
  exit: "categoryExit",
  motionPath: "categoryMotionPath",
};

// Hue accents per category. Mirrored in `SlideCanvas` Phase 3 so the
// badge colour and the picker tile colour reinforce each other.
const CATEGORY_ACCENT: Record<AnimationCategory, { dot: string; bg: string; border: string }> = {
  entrance: { dot: "bg-emerald-500", bg: "bg-emerald-50", border: "border-emerald-200" },
  emphasis: { dot: "bg-amber-500", bg: "bg-amber-50", border: "border-amber-200" },
  exit: { dot: "bg-rose-500", bg: "bg-rose-50", border: "border-rose-200" },
  motionPath: { dot: "bg-sky-500", bg: "bg-sky-50", border: "border-sky-200" },
};

const TRIGGER_KEY: Record<AnimationTrigger, string> = {
  onClick: "triggerOnClick",
  withPrevious: "triggerWithPrevious",
  afterPrevious: "triggerAfterPrevious",
};

const TRIGGER_ICONS: Record<AnimationTrigger, React.ComponentType<{ size?: number }>> = {
  onClick: MousePointerClick,
  withPrevious: CornerUpLeft,
  afterPrevious: CornerDownLeft,
};

const DIRECTION_KEY: Record<AnimationDirection, string> = {
  left: "directionLeft",
  right: "directionRight",
  up: "directionUp",
  down: "directionDown",
  in: "directionIn",
  out: "directionOut",
  horizontal: "directionHorizontal",
  vertical: "directionVertical",
  clockwise: "directionClockwise",
  counterclockwise: "directionCounterclockwise",
};

const PRESET_KEY: Record<AnimationPreset, string> = {
  appear: "presetAppear",
  fade: "presetFade",
  flyIn: "presetFlyIn",
  floatIn: "presetFloatIn",
  split: "presetSplit",
  wipe: "presetWipe",
  shape: "presetShape",
  wheel: "presetWheel",
  randomBars: "presetRandomBars",
  growAndTurn: "presetGrowAndTurn",
  zoom: "presetZoom",
  swivel: "presetSwivel",
  bounce: "presetBounce",
  pulse: "presetPulse",
  colorPulse: "presetColorPulse",
  teeter: "presetTeeter",
  spin: "presetSpin",
  growShrink: "presetGrowShrink",
  desaturate: "presetDesaturate",
  fontColor: "presetFontColor",
  lineColor: "presetLineColor",
  disappear: "presetDisappear",
  flyOut: "presetFlyOut",
  floatOut: "presetFloatOut",
  line: "presetLine",
  arc: "presetArc",
  turn: "presetTurn",
  loops: "presetLoops",
  custom: "presetCustom",
};

function animT(
  t: (k: string, vars?: Readonly<Record<string, string | number>>) => string,
  key: string
): string {
  return t(`pptx.animations.${key}`);
}

function labelForPreset(
  t: (k: string, vars?: Readonly<Record<string, string | number>>) => string,
  preset: AnimationPreset
): string {
  return animT(t, PRESET_KEY[preset]);
}

/**
 * Phase 9c §5c — Animation Painter buffer.
 *
 * Mirrors PowerPoint's "Animation Painter" tool: copy every animation
 * targeting the currently-selected shape, then paste them onto another
 * shape (selected later). Implemented as a UI-only feature that
 * composes existing `pptx:add-shape-animation` calls — no new backend
 * command is required since we already round-trip every preset listed
 * in the typed registry.
 *
 * The buffer is React state inside the panel rather than module-level
 * state so a remount (closing/opening the rail) clears it; if persistence
 * across rail open/close is desired later, lift this into a context.
 */
interface PaintedAnimation {
  readonly category: AnimationCategory;
  readonly preset: AnimationPreset;
  readonly trigger: AnimationTrigger;
  readonly direction?: AnimationDirection;
  readonly durationMs?: number;
  readonly delayMs?: number;
}

export function AnimationsPanel(props: AnimationsPanelProps): React.ReactElement {
  const { t } = useTranslator();
  const slide: Slide | undefined = props.snapshot.root.slides[props.activeIndex];
  const [activeCategory, setActiveCategory] = React.useState<AnimationCategory>("entrance");
  const [editingId, setEditingId] = React.useState<string | null>(null);
  const [painterBuffer, setPainterBuffer] = React.useState<ReadonlyArray<PaintedAnimation>>([]);
  const [painterSourceCNvPrId, setPainterSourceCNvPrId] = React.useState<number | null>(null);

  const animations = slide?.animations;

  // Auto-close the editor when the underlying animation disappears (e.g.
  // user removed it via the trash button or a remote agent edit). Hook
  // must run unconditionally — bail inside.
  React.useEffect(() => {
    if (!animations) return;
    if (editingId && !animations.some((a) => a.id === editingId)) {
      setEditingId(null);
    }
  }, [animations, editingId]);

  // Animation Painter wiring. The buffer captures every animation
  // currently targeting the selected shape (reading from the live
  // animation list rather than a separate copy so we always paint
  // the latest state, not a stale snapshot from when Copy was
  // pressed but before the user tweaked direction/duration).
  // Hooks below MUST run unconditionally before any early return,
  // hence the `?? []` fallback when animations is undefined.
  const selectedCNvPrId =
    props.selectedShape && "cNvPrId" in props.selectedShape
      ? (props.selectedShape as { cNvPrId: number }).cNvPrId
      : null;
  const selectedShapeAnimations = React.useMemo(
    () =>
      selectedCNvPrId === null ? [] : (animations ?? []).filter((a) => a.targetCNvPrId === selectedCNvPrId),
    [animations, selectedCNvPrId]
  );
  const onCopyPainter = React.useCallback((): void => {
    if (selectedCNvPrId === null) return;
    const copied: PaintedAnimation[] = selectedShapeAnimations.map((a) => ({
      category: a.category,
      preset: a.preset,
      trigger: a.trigger,
      ...(a.direction ? { direction: a.direction } : {}),
      ...(a.durationMs !== undefined ? { durationMs: a.durationMs } : {}),
      ...(a.delayMs !== undefined ? { delayMs: a.delayMs } : {}),
    }));
    setPainterBuffer(copied);
    setPainterSourceCNvPrId(selectedCNvPrId);
  }, [selectedCNvPrId, selectedShapeAnimations]);
  const onPastePainter = React.useCallback((): void => {
    if (selectedCNvPrId === null) return;
    if (painterBuffer.length === 0) return;
    if (painterSourceCNvPrId === selectedCNvPrId) return;
    for (const a of painterBuffer) {
      props.onAddAnimation({
        category: a.category,
        preset: a.preset,
        trigger: a.trigger,
        ...(a.direction ? { direction: a.direction } : {}),
        ...(a.durationMs !== undefined ? { durationMs: a.durationMs } : {}),
        ...(a.delayMs !== undefined ? { delayMs: a.delayMs } : {}),
      });
    }
  }, [painterBuffer, painterSourceCNvPrId, selectedCNvPrId, props]);

  if (!slide || !animations) {
    return <EmptyState message={t("pptx.animations.noSlide")} />;
  }
  const transition = slide.transition;
  const selectedShapeName = props.selectedShape?.name ?? null;
  const canAnimateSelected = canAnimate(props.selectedShape);
  const canCopyPainter = !props.disabled && selectedCNvPrId !== null && selectedShapeAnimations.length > 0;
  const canPastePainter =
    !props.disabled &&
    selectedCNvPrId !== null &&
    painterBuffer.length > 0 &&
    selectedCNvPrId !== painterSourceCNvPrId;

  const animationListSubtitle =
    animations.length === 0
      ? t("pptx.animations.noAnimationsHint")
      : animations.length === 1
        ? t("pptx.animations.animationCount", { count: 1 })
        : t("pptx.animations.animationCountPlural", { count: animations.length });

  return (
    <div className="flex min-h-0 flex-1 flex-col overflow-y-auto p-3 text-xs">
      <Section
        title={t("pptx.animations.slideTransition")}
        subtitle={t("pptx.animations.transitionSubtitle")}
        icon={<Sparkles size={14} />}
      >
        <TransitionEditor
          transition={transition}
          disabled={props.disabled}
          onChange={(kind, speed) => props.onSetTransition(kind, speed)}
        />
      </Section>

      <Section
        title={t("pptx.animations.addAnimation")}
        subtitle={
          selectedShapeName
            ? t("pptx.animations.selectedShape", { name: selectedShapeName })
            : t("pptx.animations.selectShapeHint")
        }
        icon={<Zap size={14} />}
        action={
          <span className="inline-flex items-center gap-1">
            <button
              type="button"
              disabled={!canCopyPainter}
              onClick={onCopyPainter}
              className="inline-flex items-center gap-1 rounded border border-divider bg-surface px-1.5 py-0.5 text-[10px] font-medium text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
              data-testid="pptx-anim-painter-copy"
              title={t("pptx.animations.painterCopyTitle")}
            >
              <Copy size={10} />
              {t("pptx.animations.painterCopy")}
            </button>
            <button
              type="button"
              disabled={!canPastePainter}
              onClick={onPastePainter}
              className="inline-flex items-center gap-1 rounded border border-divider bg-surface px-1.5 py-0.5 text-[10px] font-medium text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
              data-testid="pptx-anim-painter-paste"
              title={
                painterBuffer.length === 0
                  ? t("pptx.animations.painterPasteEmpty")
                  : t("pptx.animations.painterPasteTitle", { count: painterBuffer.length })
              }
            >
              <Clipboard size={10} />
              {t("pptx.animations.painterPaste")}
              {painterBuffer.length > 0 ? (
                <span className="ml-0.5 rounded bg-foreground/10 px-1 text-[9px] tabular-nums">
                  {painterBuffer.length}
                </span>
              ) : null}
            </button>
          </span>
        }
      >
        <CategoryTabs active={activeCategory} onChange={setActiveCategory} />
        <PresetGallery
          category={activeCategory}
          disabled={props.disabled || !canAnimateSelected}
          onPick={(spec) =>
            props.onAddAnimation({
              category: spec.category,
              preset: spec.preset,
              ...(spec.defaultDirection ? { direction: spec.defaultDirection } : {}),
            })
          }
        />
      </Section>

      <Section
        title={t("pptx.animations.animationsOnSlide")}
        subtitle={animationListSubtitle}
        icon={<Play size={14} />}
        action={
          props.onPreviewAnimation && animations.length > 0 ? (
            <button
              type="button"
              disabled={props.disabled}
              onClick={() => props.onPreviewAnimation?.(null)}
              className="inline-flex items-center gap-1 rounded border border-divider bg-surface px-2 py-0.5 text-[11px] font-medium text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
              data-testid="pptx-anim-preview-all"
              title={t("pptx.animations.previewAll")}
            >
              <Play size={11} /> {t("pptx.animations.preview")}
            </button>
          ) : null
        }
      >
        <AnimationList
          animations={animations}
          shapeNamesByCNvPrId={shapeNamesByCNvPrId(slide.shapes, t)}
          disabled={props.disabled}
          editingId={editingId}
          onToggleEdit={(id) => setEditingId((cur) => (cur === id ? null : id))}
          onRemove={props.onRemoveAnimation}
          onMove={(animationId, direction) => {
            const next = moveAnimation(animations, animationId, direction);
            if (next) props.onReorderAnimations(next.map((a) => a.id));
          }}
          onReorder={props.onReorderAnimations}
          onSet={props.onSetAnimation}
          onPreview={props.onPreviewAnimation}
        />
      </Section>
    </div>
  );
}

interface SectionProps {
  readonly title: string;
  readonly subtitle?: string;
  readonly icon?: React.ReactNode;
  readonly action?: React.ReactNode;
  readonly children: React.ReactNode;
}

function Section(props: SectionProps): React.ReactElement {
  return (
    <section className="mb-5 last:mb-0">
      <header className="mb-2 flex items-center gap-1.5 text-[11px] font-semibold uppercase tracking-wide text-secondary">
        {props.icon}
        <span className="flex-1">{props.title}</span>
        {props.action}
      </header>
      {props.subtitle ? (
        <p className="mb-2 text-[11px] leading-snug text-secondary">{props.subtitle}</p>
      ) : null}
      {props.children}
    </section>
  );
}

interface TransitionEditorProps {
  readonly transition: SlideTransition | undefined;
  readonly disabled: boolean;
  readonly onChange: (kind: TransitionKind, speed: TransitionSpeed | null) => void;
}

function TransitionEditor(props: TransitionEditorProps): React.ReactElement {
  const { t } = useTranslator();
  const kind: TransitionKind =
    props.transition?.kind === "unsupported" ? "none" : (props.transition?.kind ?? "none");
  const speed: TransitionSpeed = props.transition?.speed ?? "med";

  return (
    <div className="space-y-2">
      <label className="block">
        <span className="mb-1 block text-[11px] text-secondary">{t("pptx.animations.effect")}</span>
        <select
          value={kind}
          disabled={props.disabled}
          onChange={(e) => {
            const next = e.target.value as TransitionKind;
            props.onChange(next, next === "none" ? null : speed);
          }}
          className="w-full rounded border border-divider bg-surface px-2 py-1 text-xs text-foreground"
          data-testid="pptx-anim-transition-kind"
        >
          {(Object.keys(TRANSITION_LABEL_KEY) as Array<keyof typeof TRANSITION_LABEL_KEY>).map((k) => (
            <option key={k} value={k}>
              {t(`pptx.animations.${TRANSITION_LABEL_KEY[k]}`)}
            </option>
          ))}
        </select>
      </label>
      <label className="block">
        <span className="mb-1 block text-[11px] text-secondary">{t("pptx.animations.speed")}</span>
        <select
          value={speed}
          disabled={props.disabled || kind === "none"}
          onChange={(e) => {
            const nextSpeed = e.target.value as TransitionSpeed;
            props.onChange(kind === "none" ? "fade" : kind, nextSpeed);
          }}
          className="w-full rounded border border-divider bg-surface px-2 py-1 text-xs text-foreground disabled:opacity-50"
          data-testid="pptx-anim-transition-speed"
        >
          {(Object.keys(TRANSITION_SPEED_KEY) as TransitionSpeed[]).map((s) => (
            <option key={s} value={s}>
              {t(`pptx.animations.${TRANSITION_SPEED_KEY[s]}`)}
            </option>
          ))}
        </select>
      </label>
    </div>
  );
}

interface CategoryTabsProps {
  readonly active: AnimationCategory;
  readonly onChange: (next: AnimationCategory) => void;
}

const CATEGORY_ORDER: ReadonlyArray<AnimationCategory> = ["entrance", "emphasis", "exit", "motionPath"];

function CategoryTabs(props: CategoryTabsProps): React.ReactElement {
  const { t } = useTranslator();
  return (
    <div
      className="mb-2 grid grid-cols-4 gap-1 rounded border border-divider bg-surface p-0.5"
      role="tablist"
      aria-label={t("pptx.animations.categoryAriaLabel")}
    >
      {CATEGORY_ORDER.map((c) => {
        const isActive = c === props.active;
        const accent = CATEGORY_ACCENT[c];
        return (
          <button
            key={c}
            type="button"
            role="tab"
            aria-selected={isActive}
            onClick={() => props.onChange(c)}
            data-testid={`pptx-anim-tab-${c}`}
            className={`flex flex-col items-center gap-0.5 rounded px-1 py-1 text-[10px] font-medium transition-colors ${
              isActive
                ? `${accent.bg} text-foreground`
                : "text-secondary hover:bg-hover hover:text-foreground"
            }`}
          >
            <span className={`h-1.5 w-1.5 rounded-full ${accent.dot}`} />
            <span>{t(`pptx.animations.${CAT_KEY[c]}`)}</span>
          </button>
        );
      })}
    </div>
  );
}

interface PresetGalleryProps {
  readonly category: AnimationCategory;
  readonly disabled: boolean;
  readonly onPick: (spec: PresetSpec) => void;
}

function PresetGallery(props: PresetGalleryProps): React.ReactElement {
  const { t } = useTranslator();
  const presets = React.useMemo(() => presetsByCategory()[props.category], [props.category]);
  const categoryName = t(`pptx.animations.${CAT_KEY[props.category]}`);
  return (
    <div
      className="grid grid-cols-3 gap-1"
      role="group"
      aria-label={t("pptx.animations.addCategoryAriaLabel", { category: categoryName })}
    >
      {presets.map((spec) => (
        <button
          key={`${spec.category}-${spec.preset}`}
          type="button"
          disabled={props.disabled}
          onClick={() => props.onPick(spec)}
          title={presetTitle(spec, t)}
          data-testid={`pptx-anim-preset-${spec.category}-${spec.preset}`}
          className="group flex min-h-[42px] flex-col items-center justify-center gap-0.5 rounded border border-divider bg-surface px-1 py-1 text-[10px] font-medium text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
        >
          <Plus size={10} className="text-secondary group-disabled:opacity-50" />
          <span className="truncate">{labelForPreset(t, spec.preset)}</span>
        </button>
      ))}
    </div>
  );
}

function presetTitle(
  spec: PresetSpec,
  t: (k: string, vars?: Readonly<Record<string, string | number>>) => string
): string {
  const parts: string[] = [
    `${t(`pptx.animations.${CAT_KEY[spec.category]}`)} · ${labelForPreset(t, spec.preset)}`,
  ];
  if (spec.directions && spec.directions.length > 0) {
    const list = spec.directions.map((d) => t(`pptx.animations.${DIRECTION_KEY[d]}`)).join(", ");
    parts.push(t("pptx.animations.presetTooltipDirections", { list }));
  }
  parts.push(t("pptx.animations.presetTooltipDefaultDuration", { ms: spec.defaultDurationMs }));
  return parts.join("\n");
}

interface AnimationListProps {
  readonly animations: ReadonlyArray<ShapeAnimation>;
  readonly shapeNamesByCNvPrId: ReadonlyMap<number, string>;
  readonly disabled: boolean;
  readonly editingId: string | null;
  readonly onToggleEdit: (id: string) => void;
  readonly onRemove: (id: string) => void;
  readonly onMove: (id: string, direction: -1 | 1) => void;
  /**
   * Drag-and-drop reorder: receives the new id order in full. The
   * panel forwards this straight to `pptx:reorder-shape-animations`.
   */
  readonly onReorder: (orderIds: ReadonlyArray<string>) => void;
  readonly onSet: (params: SetAnimationParams) => void;
  readonly onPreview?: (id: string | null) => void;
}

function AnimationList(props: AnimationListProps): React.ReactElement {
  const { t } = useTranslator();
  const dragIdRef = React.useRef<string | null>(null);
  const [dropAt, setDropAt] = React.useState<number | null>(null);

  const finishDrag = (): void => {
    dragIdRef.current = null;
    setDropAt(null);
  };

  const commitDrop = (): void => {
    const moving = dragIdRef.current;
    const target = dropAt;
    finishDrag();
    if (!moving || target === null) return;
    const order = props.animations.map((a) => a.id);
    const fromIdx = order.indexOf(moving);
    if (fromIdx < 0) return;
    const next = order.slice();
    next.splice(fromIdx, 1);
    const insertAt = target > fromIdx ? target - 1 : target;
    next.splice(Math.max(0, Math.min(next.length, insertAt)), 0, moving);
    if (sameOrder(order, next)) return;
    props.onReorder(next);
  };

  if (props.animations.length === 0) {
    return (
      <div className="rounded border border-dashed border-divider p-3 text-center text-xs text-secondary">
        {t("pptx.animations.noAnimations")}
      </div>
    );
  }
  return (
    <ol
      className="flex flex-col gap-1"
      data-testid="pptx-anim-list"
      aria-label={t("pptx.animations.listAriaLabel")}
      onDragOver={(e) => {
        if (dragIdRef.current === null) return;
        e.preventDefault();
        // Falling outside any item but inside the list — treat as
        // "drop at the end" so the item can be moved to the bottom
        // even when there's no row beneath the pointer.
        if (e.target === e.currentTarget) {
          setDropAt(props.animations.length);
        }
      }}
      onDrop={(e) => {
        if (dragIdRef.current === null) return;
        e.preventDefault();
        commitDrop();
      }}
    >
      {props.animations.map((a, idx) => {
        const accent = CATEGORY_ACCENT[a.category];
        const isOpen = props.editingId === a.id;
        const TriggerIcon = TRIGGER_ICONS[a.trigger];
        const shapeName =
          props.shapeNamesByCNvPrId.get(a.targetCNvPrId) ??
          t("pptx.animations.shapeFallback", { id: a.targetCNvPrId });
        return (
          <React.Fragment key={a.id}>
            {dropAt === idx ? (
              <div
                className="mx-1 h-0.5 rounded bg-accent"
                aria-hidden
                data-testid="pptx-anim-drop-indicator"
              />
            ) : null}
            <li
              className={`overflow-hidden rounded border ${accent.border}`}
              data-testid={`pptx-anim-row-${a.id}`}
              draggable={!props.disabled}
              onDragStart={(e) => {
                dragIdRef.current = a.id;
                e.dataTransfer.effectAllowed = "move";
                e.dataTransfer.setData("text/plain", a.id);
              }}
              onDragOver={(e) => {
                if (dragIdRef.current === null) return;
                e.preventDefault();
                const rect = (e.currentTarget as HTMLElement).getBoundingClientRect();
                const before = e.clientY < rect.top + rect.height / 2;
                setDropAt(before ? idx : idx + 1);
              }}
              onDragEnd={() => finishDrag()}
              onDrop={(e) => {
                e.preventDefault();
                commitDrop();
              }}
            >
              <div className="flex items-center gap-2 bg-surface px-2 py-1.5">
                <span
                  className="inline-flex h-5 w-5 items-center justify-center rounded-full bg-hover text-[10px] font-semibold tabular-nums"
                  title={`${t(`pptx.animations.${CAT_KEY[a.category]}`)} · ${labelForPreset(t, a.preset)}`}
                >
                  {idx + 1}
                </span>
                <span className={`h-1.5 w-1.5 rounded-full ${accent.dot}`} aria-hidden />
                <button
                  type="button"
                  disabled={props.disabled}
                  onClick={() => props.onToggleEdit(a.id)}
                  className="min-w-0 flex-1 text-left disabled:cursor-not-allowed disabled:opacity-50"
                  data-testid={`pptx-anim-edit-${a.id}`}
                  aria-expanded={isOpen}
                >
                  <div className="truncate text-xs font-medium text-foreground">
                    {labelForPreset(t, a.preset)}
                    {a.direction ? (
                      <span className="ml-1 text-[10px] font-normal text-secondary">
                        · {t(`pptx.animations.${DIRECTION_KEY[a.direction]}`)}
                      </span>
                    ) : null}
                  </div>
                  <div className="flex items-center gap-1 truncate text-[10px] text-secondary">
                    <TriggerIcon size={10} />
                    <span>{t(`pptx.animations.${TRIGGER_KEY[a.trigger]}`)}</span>
                    <span aria-hidden>·</span>
                    <span className="truncate">{shapeName}</span>
                  </div>
                </button>
                <div className="flex items-center gap-0.5">
                  {props.onPreview ? (
                    <button
                      type="button"
                      disabled={props.disabled}
                      onClick={() => props.onPreview?.(a.id)}
                      className="inline-flex h-6 w-6 items-center justify-center rounded text-secondary hover:bg-hover disabled:cursor-not-allowed disabled:opacity-30"
                      title={t("pptx.animations.previewStep")}
                      aria-label={t("pptx.animations.previewStep")}
                      data-testid={`pptx-anim-preview-${a.id}`}
                    >
                      <Play size={11} />
                    </button>
                  ) : null}
                  <button
                    type="button"
                    disabled={props.disabled || idx === 0}
                    onClick={() => props.onMove(a.id, -1)}
                    className="inline-flex h-6 w-6 items-center justify-center rounded text-secondary hover:bg-hover disabled:cursor-not-allowed disabled:opacity-30"
                    title={t("pptx.animations.moveEarlier")}
                    aria-label={t("pptx.animations.moveEarlier")}
                  >
                    <ChevronUp size={12} />
                  </button>
                  <button
                    type="button"
                    disabled={props.disabled || idx === props.animations.length - 1}
                    onClick={() => props.onMove(a.id, 1)}
                    className="inline-flex h-6 w-6 items-center justify-center rounded text-secondary hover:bg-hover disabled:cursor-not-allowed disabled:opacity-30"
                    title={t("pptx.animations.moveLater")}
                    aria-label={t("pptx.animations.moveLater")}
                  >
                    <ChevronDown size={12} />
                  </button>
                  <button
                    type="button"
                    disabled={props.disabled}
                    onClick={() => props.onRemove(a.id)}
                    className="inline-flex h-6 w-6 items-center justify-center rounded text-secondary hover:bg-hover hover:text-destructive disabled:cursor-not-allowed disabled:opacity-30"
                    title={t("pptx.animations.removeAnimation")}
                    aria-label={t("pptx.animations.removeAnimation")}
                    data-testid={`pptx-anim-remove-${a.id}`}
                  >
                    <Trash2 size={12} />
                  </button>
                </div>
              </div>
              {isOpen ? (
                <EffectOptions
                  animation={a}
                  disabled={props.disabled}
                  onSet={props.onSet}
                  onPreview={props.onPreview}
                />
              ) : null}
            </li>
            {dropAt !== null && idx === props.animations.length - 1 && dropAt === props.animations.length ? (
              <div
                className="mx-1 h-0.5 rounded bg-accent"
                aria-hidden
                data-testid="pptx-anim-drop-indicator"
              />
            ) : null}
          </React.Fragment>
        );
      })}
    </ol>
  );
}

interface EffectOptionsProps {
  readonly animation: ShapeAnimation;
  readonly disabled: boolean;
  readonly onSet: (params: SetAnimationParams) => void;
  readonly onPreview?: (id: string | null) => void;
}

function EffectOptions(props: EffectOptionsProps): React.ReactElement {
  const { t } = useTranslator();
  const a = props.animation;
  const spec = React.useMemo(() => {
    const cat = presetsByCategory()[a.category];
    return cat.find((p) => p.preset === a.preset) ?? null;
  }, [a.category, a.preset]);
  const directions = spec?.directions ?? [];
  const accent = CATEGORY_ACCENT[a.category];
  return (
    <div
      className={`space-y-2 border-t ${accent.border} ${accent.bg} px-2 py-2`}
      data-testid={`pptx-anim-options-${a.id}`}
    >
      <div className="grid grid-cols-2 gap-2">
        <label className="block">
          <span className="mb-0.5 block text-[10px] uppercase tracking-wide text-secondary">
            {t("pptx.animations.trigger")}
          </span>
          <select
            value={a.trigger}
            disabled={props.disabled}
            onChange={(e) =>
              props.onSet({
                animationId: a.id,
                trigger: e.target.value as AnimationTrigger,
              })
            }
            className="w-full rounded border border-divider bg-surface px-2 py-1 text-xs text-foreground"
            data-testid={`pptx-anim-trigger-${a.id}`}
          >
            {(["onClick", "withPrevious", "afterPrevious"] as const).map((tr) => (
              <option key={tr} value={tr}>
                {t(`pptx.animations.${TRIGGER_KEY[tr]}`)}
              </option>
            ))}
          </select>
        </label>
        {directions.length > 0 ? (
          <label className="block">
            <span className="mb-0.5 block text-[10px] uppercase tracking-wide text-secondary">
              {t("pptx.animations.direction")}
            </span>
            <select
              value={a.direction ?? ""}
              disabled={props.disabled}
              onChange={(e) => {
                const v = e.target.value;
                props.onSet({
                  animationId: a.id,
                  direction: v ? (v as AnimationDirection) : null,
                });
              }}
              className="w-full rounded border border-divider bg-surface px-2 py-1 text-xs text-foreground"
              data-testid={`pptx-anim-direction-${a.id}`}
            >
              <option value="">—</option>
              {directions.map((d) => (
                <option key={d} value={d}>
                  {t(`pptx.animations.${DIRECTION_KEY[d]}`)}
                </option>
              ))}
            </select>
          </label>
        ) : (
          <div />
        )}
        <label className="block">
          <span className="mb-0.5 block text-[10px] uppercase tracking-wide text-secondary">
            {t("pptx.animations.duration")}
          </span>
          <input
            type="number"
            min={0}
            step={50}
            value={a.durationMs ?? ""}
            disabled={props.disabled}
            onChange={(e) => {
              const v = e.target.value;
              props.onSet({
                animationId: a.id,
                durationMs: v === "" ? null : Math.max(0, Number(v)),
              });
            }}
            className="w-full rounded border border-divider bg-surface px-2 py-1 text-xs text-foreground"
            data-testid={`pptx-anim-duration-${a.id}`}
          />
        </label>
        <label className="block">
          <span className="mb-0.5 block text-[10px] uppercase tracking-wide text-secondary">
            {t("pptx.animations.delay")}
          </span>
          <input
            type="number"
            min={0}
            step={50}
            value={a.delayMs ?? ""}
            disabled={props.disabled}
            onChange={(e) => {
              const v = e.target.value;
              props.onSet({
                animationId: a.id,
                delayMs: v === "" ? null : Math.max(0, Number(v)),
              });
            }}
            className="w-full rounded border border-divider bg-surface px-2 py-1 text-xs text-foreground"
            data-testid={`pptx-anim-delay-${a.id}`}
          />
        </label>
      </div>
      {props.onPreview ? (
        <button
          type="button"
          disabled={props.disabled}
          onClick={() => props.onPreview?.(a.id)}
          className="inline-flex w-full items-center justify-center gap-1 rounded border border-divider bg-surface px-2 py-1 text-[11px] font-medium text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
          data-testid={`pptx-anim-preview-row-${a.id}`}
        >
          <Hand size={11} /> {t("pptx.animations.preview")}
        </button>
      ) : null}
    </div>
  );
}

function EmptyState({ message }: { readonly message: string }): React.ReactElement {
  return <div className="p-4 text-sm text-secondary">{message}</div>;
}

function canAnimate(shape: Shape | null): boolean {
  if (!shape) return false;
  return shape.cNvPrId > 0;
}

function shapeNamesByCNvPrId(
  shapes: ReadonlyArray<Shape>,
  t: (key: string, vars?: Readonly<Record<string, string | number>>) => string
): Map<number, string> {
  const out = new Map<number, string>();
  walk(shapes, out, t);
  return out;
}

function walk(
  shapes: ReadonlyArray<Shape>,
  out: Map<number, string>,
  t: (key: string, vars?: Readonly<Record<string, string | number>>) => string
): void {
  for (const s of shapes) {
    if (s.cNvPrId != null) {
      out.set(s.cNvPrId, s.name || t("pptx.animations.shapeFallback", { id: s.cNvPrId }));
    }
    if (s.kind === "group") walk(s.children, out, t);
  }
}

function sameOrder(a: ReadonlyArray<string>, b: ReadonlyArray<string>): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i += 1) {
    if (a[i] !== b[i]) return false;
  }
  return true;
}

function moveAnimation(
  animations: ReadonlyArray<ShapeAnimation>,
  id: string,
  direction: -1 | 1
): ReadonlyArray<ShapeAnimation> | null {
  const idx = animations.findIndex((a) => a.id === id);
  if (idx < 0) return null;
  const target = idx + direction;
  if (target < 0 || target >= animations.length) return null;
  const next = animations.slice();
  const [removed] = next.splice(idx, 1);
  if (!removed) return null;
  next.splice(target, 0, removed);
  return next;
}
