import type { CommandDiagnostic, CommandPolicyMode, CommandSurface } from "./lifecycle.js";

export interface ReviewPolicyInput {
  readonly operation: string;
  readonly requestedMode?: CommandPolicyMode;
  readonly requestedRequiresReview?: boolean;
  readonly actionRequiresReview?: boolean;
  readonly sourceSurface?: CommandSurface;
}

export interface ReviewPolicyResolution {
  readonly mode: CommandPolicyMode;
  readonly requiresReview: boolean;
  readonly diagnostics: ReadonlyArray<CommandDiagnostic>;
}

const DESTRUCTIVE_OPERATION_RE =
  /(^|[:.-])(delete|remove|reject|reset|flatten|clear|purge|drop|erase|wipe)([:.-]|$)/i;

export function operationLooksDestructive(operation: string): boolean {
  return DESTRUCTIVE_OPERATION_RE.test(operation);
}

export function resolveReviewPolicy(input: ReviewPolicyInput): ReviewPolicyResolution {
  const requestedMode = input.requestedMode ?? "pending";
  const destructive = operationLooksDestructive(input.operation);
  const catalogRequiresReview = input.actionRequiresReview === true;
  const explicitReview = input.requestedRequiresReview === true;
  const requiresReview =
    requestedMode === "pending" || explicitReview || catalogRequiresReview || destructive;
  const diagnostics: CommandDiagnostic[] = [];

  if (destructive) {
    diagnostics.push({
      level: "warning",
      code: "destructive-command-review-required",
      message: `${input.operation} is destructive and requires explicit review before approval.`,
    });
  } else if (catalogRequiresReview) {
    diagnostics.push({
      level: "warning",
      code: "catalog-review-required",
      message: `${input.operation} is marked requiresReview in the action catalogue.`,
    });
  }

  if (input.requestedRequiresReview === false && requiresReview) {
    diagnostics.push({
      level: "warning",
      code: "review-opt-out-ignored",
      message: `${input.operation} cannot opt out of review under the current policy.`,
    });
  }

  if (requestedMode === "dry_run") {
    return {
      mode: "dry_run",
      requiresReview: false,
      diagnostics,
    };
  }

  if (requestedMode === "auto_apply" && requiresReview) {
    diagnostics.push({
      level: "warning",
      code: "auto-apply-downgraded-to-pending",
      message: `${input.operation} was requested as auto_apply but will stay pending for review.`,
    });
    return {
      mode: "pending",
      requiresReview: true,
      diagnostics,
    };
  }

  return {
    mode: requestedMode,
    requiresReview,
    diagnostics,
  };
}
