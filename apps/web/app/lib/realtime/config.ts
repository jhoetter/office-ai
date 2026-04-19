/**
 * Resolve the realtime websocket endpoint. Order of precedence:
 *
 *   1. `NEXT_PUBLIC_OAI_REALTIME_URL` (full ws:// or wss:// URL).
 *   2. Same-host `ws://<location.hostname>:1234`.
 *   3. `ws://localhost:1234` (SSR / tests).
 *
 * Kept tiny so it's safe to call from the room client without
 * pulling in a config layer.
 */
export function resolveRealtimeUrl(): string {
  const fromEnv = (process.env.NEXT_PUBLIC_OAI_REALTIME_URL ?? "").trim();
  if (fromEnv) return fromEnv;
  if (typeof window !== "undefined" && window.location && window.location.hostname) {
    const proto = window.location.protocol === "https:" ? "wss:" : "ws:";
    return `${proto}//${window.location.hostname}:1234`;
  }
  return "ws://localhost:1234";
}

/**
 * Stable room id derived from the editor's source URL (for sample
 * files) or a synthetic anonymous slug for blank docs. Same input
 * → same room across reloads, so two browsers opening the same URL
 * land in the same room without configuration.
 */
export function roomIdForSource(args: {
  readonly product: "docx" | "xlsx" | "pptx";
  readonly src: string | null | undefined;
  /**
   * Stable per-tab id used as a fallback when there is no `?src=`.
   * Persisted in sessionStorage by the caller (so a refresh keeps
   * the room) but unique across tabs (so two new docs don't collide).
   */
  readonly tabFallback: string;
}): string {
  const { product, src, tabFallback } = args;
  if (src && src.length > 0) {
    return `oai/${product}/src/${encodeURIComponent(src)}`;
  }
  return `oai/${product}/local/${tabFallback}`;
}
