"use client";

import { useEffect, useState } from "react";

/**
 * Stable per-tab id that survives reloads via `sessionStorage` but
 * is unique across browser tabs/windows. Used as the fallback room
 * id for blank documents (no `?src=`) so two new docs in two tabs
 * don't accidentally land in the same collaboration room.
 */
export function useStableTabId(scope: string): string {
  const [id, setId] = useState<string>(() => "");

  useEffect(() => {
    if (typeof window === "undefined") return;
    const key = `officeai.tabId.${scope}`;
    let existing = "";
    try {
      existing = window.sessionStorage.getItem(key) ?? "";
    } catch {
      /* noop */
    }
    if (!existing) {
      existing =
        typeof crypto !== "undefined" && "randomUUID" in crypto
          ? crypto.randomUUID()
          : Math.random().toString(36).slice(2);
      try {
        window.sessionStorage.setItem(key, existing);
      } catch {
        /* noop */
      }
    }
    setId(existing);
  }, [scope]);

  return id;
}
