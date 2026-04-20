"use client";

import type { ReactNode } from "react";
import { CommentsSidebar as SharedCommentsSidebar } from "@officeai/ui";
import type { CommentsProvider } from "@officeai/comments";
import { useTranslator } from "@/lib/i18n";

/**
 * Thin DOCX-shell wrapper around the format-agnostic
 * `CommentsSidebar` from `@officeai/ui`.
 *
 * The shared sidebar drives all three editors (DOCX / XLSX / PPTX) off
 * the same `CommentsProvider` interface so add / reply / resolve /
 * delete look identical to the user across products. This file used to
 * own a Word-specific implementation (DOCX threads + reply input UI);
 * that logic now lives in `@officeai/ui` and `@officeai/comments`,
 * keeping this module as a tiny re-export so existing imports from
 * `DocxEditor.tsx` keep working without churn.
 */
export interface CommentsSidebarProps {
  provider: CommentsProvider;
  /**
   * DOCX-side scroll-to-comment side-effect (highlights the anchor and
   * scrolls the page into view). The provider's own `onScrollTo` is
   * used as a fallback so callers can wire it once at construction
   * time.
   */
  onScrollTo?: (commentId: string) => void;
}

export function CommentsSidebar(props: CommentsSidebarProps): ReactNode {
  const { t } = useTranslator();
  return (
    <SharedCommentsSidebar
      provider={props.provider}
      onScrollTo={props.onScrollTo}
      emptyHint={t("docx.comments.emptyHint")}
    />
  );
}
