"use client";

import { type CSSProperties } from "react";
import { cn } from "../lib/cn";

/**
 * Display shape consumed by {@link PresenceStack}. Mirrors the
 * {@link import("@officeai/realtime").AwarenessState} but kept
 * decoupled so this primitive can also be storybook-fed without
 * pulling Yjs into the renderer tree.
 */
export interface PresencePeer {
  readonly id: string;
  readonly name: string;
  readonly color: string;
  /** ISO product slug — drives the optional product hint glyph. */
  readonly product?: "docx" | "xlsx" | "pptx";
  /** Wall-clock when the peer last published awareness; "active 12s ago". */
  readonly lastSeen?: number;
}

export interface PresenceStackProps {
  readonly peers: ReadonlyArray<PresencePeer>;
  /** Local user — when present, rendered first with a "you" marker. */
  readonly self?: PresencePeer;
  /** Connection status pill colour ring. */
  readonly status?: "disabled" | "disconnected" | "connecting" | "connected";
  readonly className?: string;
  readonly maxVisible?: number;
}

const RING_BY_STATUS: Record<NonNullable<PresenceStackProps["status"]>, string> = {
  disabled: "ring-divider",
  disconnected: "ring-[color:var(--error)]",
  connecting: "ring-[color:var(--warning)]",
  connected: "ring-[color:var(--success,#10b981)]",
};

function initialsOf(name: string): string {
  const parts = name.trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return "?";
  if (parts.length === 1) return parts[0]!.slice(0, 2).toUpperCase();
  return (parts[0]![0]! + parts[1]![0]!).toUpperCase();
}

function Avatar({
  peer,
  isSelf,
  ring,
}: {
  readonly peer: PresencePeer;
  readonly isSelf?: boolean;
  readonly ring?: string;
}): React.ReactNode {
  const style: CSSProperties = {
    backgroundColor: peer.color,
    color: "#ffffff",
  };
  return (
    <span
      title={`${peer.name}${isSelf ? " (you)" : ""}`}
      className={cn(
        "relative inline-flex h-6 w-6 items-center justify-center rounded-full text-[10px] font-semibold leading-none",
        "ring-2 ring-background",
        ring
      )}
      style={style}
      data-testid={`presence-avatar-${peer.id}`}
    >
      {initialsOf(peer.name)}
      {isSelf ? (
        <span
          className="absolute -bottom-0.5 -right-0.5 inline-flex h-2 w-2 rounded-full bg-[var(--accent,#3b82f6)] ring-1 ring-background"
          aria-hidden
        />
      ) : null}
    </span>
  );
}

/**
 * Compact horizontal stack of avatar circles for everyone currently
 * collaborating in the room. The local user (when given) leads the
 * row, then up to `maxVisible` remote peers, then a "+N" overflow
 * chip.
 */
export function PresenceStack({
  peers,
  self,
  status = "disabled",
  className,
  maxVisible = 5,
}: PresenceStackProps): React.ReactNode {
  if (status === "disabled" && peers.length === 0 && !self) return null;
  const ring = RING_BY_STATUS[status];
  const visible = peers.slice(0, maxVisible);
  const overflow = Math.max(0, peers.length - visible.length);

  return (
    <div
      className={cn("inline-flex items-center -space-x-1.5", className)}
      role="group"
      aria-label="Active collaborators"
      data-testid="presence-stack"
    >
      {self ? <Avatar peer={self} isSelf ring={ring} /> : null}
      {visible.map((peer) => (
        <Avatar key={peer.id} peer={peer} ring={ring} />
      ))}
      {overflow > 0 ? (
        <span
          className={cn(
            "relative inline-flex h-6 min-w-6 items-center justify-center rounded-full bg-hover px-1.5 text-[10px] font-semibold text-secondary",
            "ring-2 ring-background",
            ring
          )}
          title={`+${overflow} more`}
          data-testid="presence-overflow"
        >
          +{overflow}
        </span>
      ) : null}
    </div>
  );
}
