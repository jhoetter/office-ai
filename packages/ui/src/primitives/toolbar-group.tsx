"use client";

import { type ReactNode } from "react";
import { cn } from "../lib/cn";

export interface ToolbarGroupProps {
  /**
   * Optional caption rendered below the buttons in tiny tertiary text
   * (Word/Excel/PowerPoint convention). Omit when the icons alone are
   * unambiguous (e.g. the leading file-ops cluster in the top bar).
   */
  readonly label?: string;
  /**
   * Buttons / pickers / dropdowns that belong to this functional
   * group. The component does not impose its own layout on children;
   * it merely wraps them in a flex strip with consistent spacing so
   * adjacent groups read as separate clusters rather than one long
   * row.
   */
  readonly children: ReactNode;
  /** Additional classes for the outer wrapper. */
  readonly className?: string;
  /** Test hook — typically `docx-group-clipboard` or similar. */
  readonly testId?: string;
}

/**
 * Word-style toolbar group: a cluster of related buttons with an
 * optional small caption underneath. Designed to be dropped into the
 * existing `ToolbarRow` (which still drives the row's overall scroll
 * and height) so editors can adopt grouping incrementally.
 *
 * The label is intentionally tiny (10 px) and rendered in `tertiary`
 * tone so it never competes with the actual icons. We deliberately
 * do NOT render a vertical divider on either side — the row's outer
 * `<Divider />` between groups is still the right separator. Pulling
 * the divider into this primitive would couple the cluster to its
 * neighbour (the leading-edge group should not have a left divider,
 * the trailing-edge group should not have a right one), and that
 * coupling is more painful than the duplication of writing
 * `<Divider />` between groups in the call-site.
 */
export function ToolbarGroup({ label, children, className, testId }: ToolbarGroupProps): ReactNode {
  return (
    <div
      className={cn("flex flex-col items-center justify-center", className)}
      data-testid={testId}
      role="group"
      aria-label={label}
    >
      <div className="flex items-center gap-0.5">{children}</div>
      {label ? (
        <span className="mt-0.5 select-none text-[9px] font-medium uppercase tracking-wide text-tertiary">
          {label}
        </span>
      ) : null}
    </div>
  );
}
