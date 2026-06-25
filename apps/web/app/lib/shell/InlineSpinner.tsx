"use client";

import type { ReactNode } from "react";
import { Loader2 } from "@officeai/ui/sonaloop-icons";
import { cn } from "@officeai/ui";

/**
 * The single inline spinner the shell uses anywhere it needs to say
 * "this short async thing is in flight" — save state pill, status
 * bar mirror, export dialog confirm button, and (under the hood) the
 * `LoadingScreen` itself.
 *
 * Why a shared one: every product had its own `Loader2` invocation
 * with subtly different sizes and margins, and the save/export
 * states had no spinner at all. Centralising it means every
 * "in-flight" beat in the app spins at the same speed, the same
 * size, and the same colour, which is the whole point of the
 * shared shell.
 *
 * Spinner is decorative by default — the surrounding text ("Saving…",
 * "Exporting…", "Loading…") already conveys the state to assistive
 * tech, so we hide the icon from screen readers to avoid duplicate
 * announcements. Pass `decorative={false}` only when there's no
 * adjacent label.
 */
export interface InlineSpinnerProps {
  /** Pixels. Defaults to 12 (matches `text-xs` x-height for inline
   * usage in pills / buttons). */
  readonly size?: number;
  readonly className?: string;
  readonly decorative?: boolean;
}

export function InlineSpinner({ size = 12, className, decorative = true }: InlineSpinnerProps): ReactNode {
  return (
    <Loader2
      size={size}
      className={cn("animate-spin", className)}
      aria-hidden={decorative ? "true" : undefined}
      role={decorative ? undefined : "status"}
    />
  );
}
