"use client";

import * as React from "react";
import { ChevronDown, ChevronUp, Plus, Sparkles, Trash2, Zap } from "lucide-react";
import type {
  EntranceAnimation,
  EntranceEffect,
  PptxSnapshot,
  Shape,
  Slide,
  SlideTransition,
  TransitionKind,
  TransitionSpeed,
} from "@officeai/pptx";

/**
 * D11 — Animations panel mounted in the right rail when the active
 * product is PPTX.
 *
 * The panel is a thin shell over four already-implemented commands:
 *   • `pptx:set-slide-transition`
 *   • `pptx:add-shape-animation`
 *   • `pptx:remove-shape-animation`
 *   • `pptx:reorder-shape-animations`
 *
 * Scope is deliberately narrow — we surface only the typed
 * primitives. Anything that doesn't round-trip cleanly (custom
 * motion paths, advanced sequences, build options) lives in
 * `timingTailRaw` / `transition.raw` and is never edited here.
 */
export interface AnimationsPanelProps {
  readonly snapshot: PptxSnapshot;
  readonly activeIndex: number;
  readonly selectedShape: Shape | null;
  readonly disabled: boolean;
  readonly onSetTransition: (kind: TransitionKind, speed: TransitionSpeed | null) => void;
  readonly onAddAnimation: (effect: EntranceEffect) => void;
  readonly onRemoveAnimation: (animationId: string) => void;
  readonly onReorderAnimations: (orderIds: ReadonlyArray<string>) => void;
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

const ENTRANCE_EFFECTS: ReadonlyArray<{ readonly value: EntranceEffect; readonly label: string }> = [
  { value: "appear", label: "Appear" },
  { value: "fade", label: "Fade" },
  { value: "fly-in", label: "Fly in" },
  { value: "wipe", label: "Wipe" },
];

export function AnimationsPanel(props: AnimationsPanelProps): React.ReactElement {
  const slide: Slide | undefined = props.snapshot.root.slides[props.activeIndex];
  if (!slide) {
    return <EmptyState message="No slide selected." />;
  }
  const animations = slide.animations;
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
        title="Entrance animations"
        subtitle={
          selectedShapeName
            ? `Selected shape: ${selectedShapeName}`
            : "Select a shape on the slide to add an entrance animation."
        }
        icon={<Zap size={14} />}
      >
        <AddAnimationButtons disabled={props.disabled || !canAnimateSelected} onAdd={props.onAddAnimation} />
        <AnimationList
          animations={animations}
          shapeNamesByCNvPrId={shapeNamesByCNvPrId(slide.shapes)}
          disabled={props.disabled}
          onRemove={props.onRemoveAnimation}
          onMove={(animationId, direction) => {
            const next = moveAnimation(animations, animationId, direction);
            if (next) props.onReorderAnimations(next.map((a) => a.id));
          }}
        />
      </Section>
    </div>
  );
}

interface SectionProps {
  readonly title: string;
  readonly subtitle?: string;
  readonly icon?: React.ReactNode;
  readonly children: React.ReactNode;
}

function Section(props: SectionProps): React.ReactElement {
  return (
    <section className="mb-5 last:mb-0">
      <header className="mb-2 flex items-center gap-1.5 text-[11px] font-semibold uppercase tracking-wide text-secondary">
        {props.icon}
        <span>{props.title}</span>
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
  // We intentionally hide the "unsupported" kind from the user — it
  // exists only as a parser landing pad and is round-tripped via the
  // verbatim raw blob.
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

interface AddAnimationButtonsProps {
  readonly disabled: boolean;
  readonly onAdd: (effect: EntranceEffect) => void;
}

function AddAnimationButtons(props: AddAnimationButtonsProps): React.ReactElement {
  return (
    <div className="mb-3 grid grid-cols-2 gap-1" role="group" aria-label="Add entrance animation">
      {ENTRANCE_EFFECTS.map((e) => (
        <button
          key={e.value}
          type="button"
          disabled={props.disabled}
          onClick={() => props.onAdd(e.value)}
          className="inline-flex items-center justify-center gap-1 rounded border border-divider bg-surface px-2 py-1 text-xs text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
          data-testid={`pptx-anim-add-${e.value}`}
        >
          <Plus size={12} /> {e.label}
        </button>
      ))}
    </div>
  );
}

interface AnimationListProps {
  readonly animations: ReadonlyArray<EntranceAnimation>;
  readonly shapeNamesByCNvPrId: ReadonlyMap<number, string>;
  readonly disabled: boolean;
  readonly onRemove: (id: string) => void;
  readonly onMove: (id: string, direction: -1 | 1) => void;
}

function AnimationList(props: AnimationListProps): React.ReactElement {
  if (props.animations.length === 0) {
    return (
      <div className="rounded border border-dashed border-divider p-3 text-center text-xs text-secondary">
        No entrance animations on this slide yet.
      </div>
    );
  }
  return (
    <ol className="flex flex-col gap-1" data-testid="pptx-anim-list" aria-label="Entrance animations">
      {props.animations.map((a, idx) => (
        <li
          key={a.id}
          className="flex items-center gap-2 rounded border border-divider bg-surface px-2 py-1.5"
          data-testid={`pptx-anim-row-${a.id}`}
        >
          <span className="inline-flex h-5 w-5 items-center justify-center rounded-full bg-hover text-[10px] font-semibold tabular-nums">
            {idx + 1}
          </span>
          <div className="min-w-0 flex-1">
            <div className="truncate text-xs font-medium capitalize text-foreground">
              {a.effect.replace("-", " ")}
            </div>
            <div className="truncate text-[10px] text-secondary">
              {props.shapeNamesByCNvPrId.get(a.targetCNvPrId) ?? `Shape #${a.targetCNvPrId}`}
            </div>
          </div>
          <div className="flex items-center gap-0.5">
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
        </li>
      ))}
    </ol>
  );
}

function EmptyState({ message }: { readonly message: string }): React.ReactElement {
  return <div className="p-4 text-sm text-secondary">{message}</div>;
}

function canAnimate(shape: Shape | null): boolean {
  if (!shape) return false;
  // The command handler rejects shapes whose `cNvPrId` is non-positive
  // (synthetic / placeholder); mirror that constraint here so the
  // Add buttons are disabled before the user even tries.
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
  animations: ReadonlyArray<EntranceAnimation>,
  id: string,
  direction: -1 | 1
): ReadonlyArray<EntranceAnimation> | null {
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
