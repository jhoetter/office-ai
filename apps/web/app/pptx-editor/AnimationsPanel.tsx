"use client";

import * as React from "react";
import {
  ChevronDown,
  ChevronUp,
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

const TRANSITION_KINDS: ReadonlyArray<{ readonly value: TransitionKind; readonly label: string }> = [
  { value: "none", label: "None" },
  { value: "fade", label: "Fade" },
  { value: "push", label: "Push" },
  { value: "wipe", label: "Wipe" },
  { value: "split", label: "Split" },
  { value: "cut", label: "Cut" },
];

const TRANSITION_SPEEDS: ReadonlyArray<{ readonly value: TransitionSpeed; readonly label: string }> = [
  { value: "slow", label: "Slow" },
  { value: "med", label: "Medium" },
  { value: "fast", label: "Fast" },
];

const CATEGORY_LABELS: Record<AnimationCategory, string> = {
  entrance: "Entrance",
  emphasis: "Emphasis",
  exit: "Exit",
  motionPath: "Motion paths",
};

// Hue accents per category. Mirrored in `SlideCanvas` Phase 3 so the
// badge colour and the picker tile colour reinforce each other.
const CATEGORY_ACCENT: Record<AnimationCategory, { dot: string; bg: string; border: string }> = {
  entrance: { dot: "bg-emerald-500", bg: "bg-emerald-50", border: "border-emerald-200" },
  emphasis: { dot: "bg-amber-500", bg: "bg-amber-50", border: "border-amber-200" },
  exit: { dot: "bg-rose-500", bg: "bg-rose-50", border: "border-rose-200" },
  motionPath: { dot: "bg-sky-500", bg: "bg-sky-50", border: "border-sky-200" },
};

const TRIGGER_LABELS: Record<AnimationTrigger, string> = {
  onClick: "On click",
  withPrevious: "With previous",
  afterPrevious: "After previous",
};

const TRIGGER_ICONS: Record<AnimationTrigger, React.ComponentType<{ size?: number }>> = {
  onClick: MousePointerClick,
  withPrevious: CornerUpLeft,
  afterPrevious: CornerDownLeft,
};

const DIRECTION_LABELS: Record<AnimationDirection, string> = {
  left: "Left",
  right: "Right",
  up: "Up",
  down: "Down",
  in: "In",
  out: "Out",
  horizontal: "Horizontal",
  vertical: "Vertical",
  clockwise: "Clockwise",
  counterclockwise: "Counter-clockwise",
};

const PRESET_LABELS: Record<string, string> = {
  // Entrance
  appear: "Appear",
  fade: "Fade",
  flyIn: "Fly in",
  floatIn: "Float in",
  split: "Split",
  wipe: "Wipe",
  shape: "Shape",
  wheel: "Wheel",
  randomBars: "Random bars",
  growAndTurn: "Grow & turn",
  zoom: "Zoom",
  swivel: "Swivel",
  bounce: "Bounce",
  // Emphasis
  pulse: "Pulse",
  colorPulse: "Color pulse",
  teeter: "Teeter",
  spin: "Spin",
  growShrink: "Grow / shrink",
  desaturate: "Desaturate",
  fontColor: "Font color",
  lineColor: "Line color",
  // Exit
  disappear: "Disappear",
  flyOut: "Fly out",
  floatOut: "Float out",
  // Motion paths
  line: "Line",
  arc: "Arc",
  turn: "Turn",
  loops: "Loops",
  custom: "Custom",
};

export function AnimationsPanel(props: AnimationsPanelProps): React.ReactElement {
  const slide: Slide | undefined = props.snapshot.root.slides[props.activeIndex];
  const [activeCategory, setActiveCategory] = React.useState<AnimationCategory>("entrance");
  const [editingId, setEditingId] = React.useState<string | null>(null);

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

  if (!slide || !animations) {
    return <EmptyState message="No slide selected." />;
  }
  const transition = slide.transition;
  const selectedShapeName = props.selectedShape?.name ?? null;
  const canAnimateSelected = canAnimate(props.selectedShape);

  return (
    <div className="flex min-h-0 flex-1 flex-col overflow-y-auto p-3 text-xs">
      <Section title="Slide transition" subtitle="Plays as this slide enters." icon={<Sparkles size={14} />}>
        <TransitionEditor
          transition={transition}
          disabled={props.disabled}
          onChange={(kind, speed) => props.onSetTransition(kind, speed)}
        />
      </Section>

      <Section
        title="Add animation"
        subtitle={
          selectedShapeName
            ? `Selected shape: ${selectedShapeName}`
            : "Select a shape on the slide to add an animation."
        }
        icon={<Zap size={14} />}
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
        title="Animations on this slide"
        subtitle={
          animations.length === 0
            ? "Pick a preset above to add the first one."
            : `${animations.length} animation${animations.length === 1 ? "" : "s"} in sequence.`
        }
        icon={<Play size={14} />}
        action={
          props.onPreviewAnimation && animations.length > 0 ? (
            <button
              type="button"
              disabled={props.disabled}
              onClick={() => props.onPreviewAnimation?.(null)}
              className="inline-flex items-center gap-1 rounded border border-divider bg-surface px-2 py-0.5 text-[11px] font-medium text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
              data-testid="pptx-anim-preview-all"
              title="Preview the full sequence"
            >
              <Play size={11} /> Preview
            </button>
          ) : null
        }
      >
        <AnimationList
          animations={animations}
          shapeNamesByCNvPrId={shapeNamesByCNvPrId(slide.shapes)}
          disabled={props.disabled}
          editingId={editingId}
          onToggleEdit={(id) => setEditingId((cur) => (cur === id ? null : id))}
          onRemove={props.onRemoveAnimation}
          onMove={(animationId, direction) => {
            const next = moveAnimation(animations, animationId, direction);
            if (next) props.onReorderAnimations(next.map((a) => a.id));
          }}
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
  const kind: TransitionKind =
    props.transition?.kind === "unsupported" ? "none" : (props.transition?.kind ?? "none");
  const speed: TransitionSpeed = props.transition?.speed ?? "med";

  return (
    <div className="space-y-2">
      <label className="block">
        <span className="mb-1 block text-[11px] text-secondary">Effect</span>
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
          {TRANSITION_KINDS.map((k) => (
            <option key={k.value} value={k.value}>
              {k.label}
            </option>
          ))}
        </select>
      </label>
      <label className="block">
        <span className="mb-1 block text-[11px] text-secondary">Speed</span>
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
          {TRANSITION_SPEEDS.map((s) => (
            <option key={s.value} value={s.value}>
              {s.label}
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
  return (
    <div
      className="mb-2 grid grid-cols-4 gap-1 rounded border border-divider bg-surface p-0.5"
      role="tablist"
      aria-label="Animation category"
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
            <span>{CATEGORY_LABELS[c]}</span>
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
  const presets = React.useMemo(() => presetsByCategory()[props.category], [props.category]);
  return (
    <div
      className="grid grid-cols-3 gap-1"
      role="group"
      aria-label={`Add ${CATEGORY_LABELS[props.category].toLowerCase()} animation`}
    >
      {presets.map((spec) => (
        <button
          key={`${spec.category}-${spec.preset}`}
          type="button"
          disabled={props.disabled}
          onClick={() => props.onPick(spec)}
          title={presetTitle(spec)}
          data-testid={`pptx-anim-preset-${spec.category}-${spec.preset}`}
          className="group flex min-h-[42px] flex-col items-center justify-center gap-0.5 rounded border border-divider bg-surface px-1 py-1 text-[10px] font-medium text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
        >
          <Plus size={10} className="text-secondary group-disabled:opacity-50" />
          <span className="truncate">{labelFor(spec.preset)}</span>
        </button>
      ))}
    </div>
  );
}

function presetTitle(spec: PresetSpec): string {
  const parts: string[] = [`${CATEGORY_LABELS[spec.category]} · ${labelFor(spec.preset)}`];
  if (spec.directions && spec.directions.length > 0) {
    parts.push(`Directions: ${spec.directions.map((d) => DIRECTION_LABELS[d]).join(", ")}`);
  }
  parts.push(`Default duration: ${spec.defaultDurationMs} ms`);
  return parts.join("\n");
}

function labelFor(preset: AnimationPreset): string {
  return PRESET_LABELS[preset] ?? preset;
}

interface AnimationListProps {
  readonly animations: ReadonlyArray<ShapeAnimation>;
  readonly shapeNamesByCNvPrId: ReadonlyMap<number, string>;
  readonly disabled: boolean;
  readonly editingId: string | null;
  readonly onToggleEdit: (id: string) => void;
  readonly onRemove: (id: string) => void;
  readonly onMove: (id: string, direction: -1 | 1) => void;
  readonly onSet: (params: SetAnimationParams) => void;
  readonly onPreview?: (id: string | null) => void;
}

function AnimationList(props: AnimationListProps): React.ReactElement {
  if (props.animations.length === 0) {
    return (
      <div className="rounded border border-dashed border-divider p-3 text-center text-xs text-secondary">
        No animations on this slide yet.
      </div>
    );
  }
  return (
    <ol className="flex flex-col gap-1" data-testid="pptx-anim-list" aria-label="Animations on this slide">
      {props.animations.map((a, idx) => {
        const accent = CATEGORY_ACCENT[a.category];
        const isOpen = props.editingId === a.id;
        const TriggerIcon = TRIGGER_ICONS[a.trigger];
        const shapeName = props.shapeNamesByCNvPrId.get(a.targetCNvPrId) ?? `Shape #${a.targetCNvPrId}`;
        return (
          <li
            key={a.id}
            className={`overflow-hidden rounded border ${accent.border}`}
            data-testid={`pptx-anim-row-${a.id}`}
          >
            <div className="flex items-center gap-2 bg-surface px-2 py-1.5">
              <span
                className="inline-flex h-5 w-5 items-center justify-center rounded-full bg-hover text-[10px] font-semibold tabular-nums"
                title={`${CATEGORY_LABELS[a.category]} · ${labelFor(a.preset)}`}
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
                  {labelFor(a.preset)}
                  {a.direction ? (
                    <span className="ml-1 text-[10px] font-normal text-secondary">
                      · {DIRECTION_LABELS[a.direction]}
                    </span>
                  ) : null}
                </div>
                <div className="flex items-center gap-1 truncate text-[10px] text-secondary">
                  <TriggerIcon size={10} />
                  <span>{TRIGGER_LABELS[a.trigger]}</span>
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
                    title="Preview this step"
                    aria-label="Preview this step"
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
                  title="Move earlier"
                  aria-label="Move earlier"
                >
                  <ChevronUp size={12} />
                </button>
                <button
                  type="button"
                  disabled={props.disabled || idx === props.animations.length - 1}
                  onClick={() => props.onMove(a.id, 1)}
                  className="inline-flex h-6 w-6 items-center justify-center rounded text-secondary hover:bg-hover disabled:cursor-not-allowed disabled:opacity-30"
                  title="Move later"
                  aria-label="Move later"
                >
                  <ChevronDown size={12} />
                </button>
                <button
                  type="button"
                  disabled={props.disabled}
                  onClick={() => props.onRemove(a.id)}
                  className="inline-flex h-6 w-6 items-center justify-center rounded text-secondary hover:bg-hover hover:text-destructive disabled:cursor-not-allowed disabled:opacity-30"
                  title="Remove animation"
                  aria-label="Remove animation"
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
          <span className="mb-0.5 block text-[10px] uppercase tracking-wide text-secondary">Trigger</span>
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
            {(["onClick", "withPrevious", "afterPrevious"] as const).map((t) => (
              <option key={t} value={t}>
                {TRIGGER_LABELS[t]}
              </option>
            ))}
          </select>
        </label>
        {directions.length > 0 ? (
          <label className="block">
            <span className="mb-0.5 block text-[10px] uppercase tracking-wide text-secondary">Direction</span>
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
                  {DIRECTION_LABELS[d]}
                </option>
              ))}
            </select>
          </label>
        ) : (
          <div />
        )}
        <label className="block">
          <span className="mb-0.5 block text-[10px] uppercase tracking-wide text-secondary">
            Duration (ms)
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
          <span className="mb-0.5 block text-[10px] uppercase tracking-wide text-secondary">Delay (ms)</span>
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
          <Hand size={11} /> Preview
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

function shapeNamesByCNvPrId(shapes: ReadonlyArray<Shape>): Map<number, string> {
  const out = new Map<number, string>();
  walk(shapes, out);
  return out;
}

function walk(shapes: ReadonlyArray<Shape>, out: Map<number, string>): void {
  for (const s of shapes) {
    if (s.cNvPrId != null) out.set(s.cNvPrId, s.name || `Shape #${s.cNvPrId}`);
    if (s.kind === "group") walk(s.children, out);
  }
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
