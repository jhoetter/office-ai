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
 *
 * Resolution order:
 *
 *   1. Explicit `?room=<id>` query parameter — the override used for
 *      multi-user testing on fresh docs. Two browsers visiting
 *      `/pptx-editor?room=demo` collide intentionally and edit live.
 *   2. `?src=<url>` — sample / external file. The URL is the room.
 *   3. Per-tab `tabFallback`. Unique across tabs by design so two
 *      unrelated blank decks don't accidentally merge.
 */
export function roomIdForSource(args: {
  readonly product: "docx" | "xlsx" | "pptx" | "pdf";
  /**
   * Source URL for the document (sample file, external URL, etc.). Optional
   * so callers using a conditional spread (`...(x ? { src: x.url } : {})`)
   * compile under `exactOptionalPropertyTypes`. Absent / null / undefined
   * all mean "no source", and the function falls through to `tabFallback`.
   */
  readonly src?: string | null | undefined;
  /**
   * Stable per-tab id used as a fallback when there is no `?src=`
   * and no `?room=` override. Persisted in sessionStorage by the
   * caller (so a refresh keeps the room) but unique across tabs.
   */
  readonly tabFallback: string;
  /**
   * Explicit room override (typically read from `?room=…`). Wins over
   * both `src` and `tabFallback` so multi-user demos can opt into a
   * shared room without serving the same file.
   */
  readonly explicitRoom?: string | null | undefined;
}): string {
  const { product, src, tabFallback, explicitRoom } = args;
  if (explicitRoom && explicitRoom.length > 0) {
    return `oai/${product}/room/${encodeURIComponent(explicitRoom)}`;
  }
  if (src && src.length > 0) {
    return `oai/${product}/src/${encodeURIComponent(src)}`;
  }
  return `oai/${product}/local/${tabFallback}`;
}

/**
 * Read the `?room=<id>` query parameter from the current URL, if
 * any. Returns `null` outside the browser. Kept here (rather than
 * inlined in editors) so the resolution rule lives next to
 * {@link roomIdForSource}.
 */
export function readExplicitRoomFromUrl(): string | null {
  if (typeof window === "undefined") return null;
  try {
    const sp = new URLSearchParams(window.location.search);
    const v = sp.get("room");
    return v && v.length > 0 ? v : null;
  } catch {
    return null;
  }
}
