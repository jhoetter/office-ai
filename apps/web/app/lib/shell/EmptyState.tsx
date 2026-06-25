"use client";

import type { ReactNode } from "react";
import { FolderOpen, MousePointer2 } from "@officeai/ui/sonaloop-icons";
import { Button } from "@officeai/ui";
import { useTranslator } from "@/lib/i18n";

export interface EmptyStateProps {
  readonly product: "docx" | "xlsx" | "pptx";
  readonly onOpen: () => void;
}

const PRODUCT_LABEL_KEY: Record<EmptyStateProps["product"], string> = {
  docx: "common.kindDocx",
  xlsx: "common.kindXlsx",
  pptx: "common.kindPptx",
};

const PRODUCT_EXT: Record<EmptyStateProps["product"], string> = {
  docx: ".docx",
  xlsx: ".xlsx",
  pptx: ".pptx",
};

/**
 * The "no file open" surface — used by every editor when there's no
 * document loaded. Promotes only Open + drag-drop. No recent files,
 * no AI prompts; the shell is a frame for the document.
 */
export function EmptyState({ product, onOpen }: EmptyStateProps): ReactNode {
  const { t } = useTranslator();
  const label = t(PRODUCT_LABEL_KEY[product]);
  const ext = PRODUCT_EXT[product];
  return (
    <div
      className="flex h-full w-full flex-col items-center justify-center gap-6 px-6 text-center"
      data-testid="empty-state"
    >
      <div className="flex flex-col items-center gap-2">
        <div className="inline-flex h-12 w-12 items-center justify-center rounded-full bg-hover text-secondary">
          <FolderOpen size={22} />
        </div>
        <h2 className="text-lg font-semibold text-foreground">
          {t("common.openProduct", { product: label })}
        </h2>
        <p className="max-w-md text-sm text-secondary">{t("common.openProductHint", { ext })}</p>
      </div>
      <Button variant="primary" size="md" onClick={onOpen}>
        <FolderOpen size={14} />
        {t("common.open")}
      </Button>
      <div className="flex items-center gap-2 text-xs text-tertiary">
        <MousePointer2 size={12} />
        <span>{t("common.orDropHere", { ext })}</span>
      </div>
    </div>
  );
}
