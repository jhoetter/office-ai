"use client";

import { useMemo, type ReactNode } from "react";
import { SonaloopCommandPalette, type SonaloopCommandGroup } from "@officeai/ui/sonaloop-command";
import { useTranslator } from "@/lib/i18n";
import type { PaletteCommand } from "./types";

export interface CommandPaletteProps {
  readonly open: boolean;
  readonly onClose: () => void;
  readonly commands: ReadonlyArray<PaletteCommand>;
}

const RECENT_KEY = "officeai.palette.recent";
const RECENT_LIMIT = 6;

/**
 * Office-AI adapter over the shared Sonaloop command palette.
 *
 * Editors still hand the shell typed `PaletteCommand[]`; this component groups
 * and decorates those commands for `sonaloop-design`, preserving recent-command
 * ordering and the existing command IDs used by tests and telemetry.
 */
export function CommandPalette({ open, onClose, commands }: CommandPaletteProps): ReactNode {
  const { t } = useTranslator();
  const recents = useMemo(() => loadRecents(), [open]);

  const groups = useMemo(() => {
    const enabled = commands.filter((c) => c.enabled !== false);
    const byId = new Map(enabled.map((cmd) => [cmd.id, cmd]));
    const recentCommands = recents
      .map((id) => byId.get(id))
      .filter((cmd): cmd is PaletteCommand => Boolean(cmd));
    const recentSet = new Set(recentCommands.map((cmd) => cmd.id));
    const grouped = new Map<string, PaletteCommand[]>();

    for (const cmd of enabled) {
      if (recentSet.has(cmd.id)) continue;
      const section = cmd.section || t("common.commands");
      const list = grouped.get(section) ?? [];
      list.push(cmd);
      grouped.set(section, list);
    }

    const out: SonaloopCommandGroup[] = [];
    if (recentCommands.length > 0) {
      out.push({
        key: "recent",
        label: t("common.recent"),
        items: recentCommands.map(toCommandItem),
      });
    }
    for (const [section, items] of grouped) {
      out.push({
        key: section.toLowerCase().replace(/[^a-z0-9]+/g, "-") || "commands",
        label: section,
        items: items.map(toCommandItem),
      });
    }
    return out;
  }, [commands, recents, t]);

  return (
    <SonaloopCommandPalette
      open={open}
      onOpenChange={(next) => {
        if (!next) onClose();
      }}
      groups={groups}
      placeholder={t("common.commandPalettePlaceholder")}
      emptyMessage={t("common.commandPaletteEmpty")}
      hotkey={false}
    />
  );
}

function toCommandItem(cmd: PaletteCommand) {
  return {
    title: cmd.label,
    subtitle: cmd.shortcut ? [cmd.hint, cmd.shortcut].filter(Boolean).join(" · ") : cmd.hint,
    keywords: [cmd.id, cmd.hint, cmd.section, cmd.shortcut].filter(Boolean).join(" "),
    onSelect: () => {
      rememberRecent(cmd.id);
      void cmd.run();
    },
  };
}

function loadRecents(): string[] {
  if (typeof window === "undefined") return [];
  try {
    const raw = window.localStorage.getItem(RECENT_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed.filter((x): x is string => typeof x === "string") : [];
  } catch {
    return [];
  }
}

function rememberRecent(id: string): void {
  if (typeof window === "undefined") return;
  try {
    const current = loadRecents();
    const next = [id, ...current.filter((x) => x !== id)].slice(0, RECENT_LIMIT);
    window.localStorage.setItem(RECENT_KEY, JSON.stringify(next));
  } catch {
    // localStorage may be unavailable (private mode); silently ignore.
  }
}
