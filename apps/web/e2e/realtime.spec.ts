import { expect, test, type BrowserContext, type Page } from "@playwright/test";

/**
 * D11 — multi-tab Playwright e2e for the realtime pipeline.
 *
 * Two browser contexts open the XLSX editor in the same realtime
 * room. We then assert three things end-to-end:
 *
 *  1. Each peer eventually appears in the other's `realtime-presence-list`.
 *  2. A typed cell value in tab A is mirrored into tab B's grid.
 *  3. After tab A moves its selection, tab B's presence list reports
 *     a non-empty `Sheet!Range` summary (i.e. awareness flows through
 *     the y-websocket roundtrip, not just the Yjs update channel).
 *
 * Prerequisites:
 *   - `playwright.config.ts` starts both `next start` and the
 *     y-websocket server (`@officeai/realtime-server`) via its
 *     `webServer` array. If you run with `E2E_BASE_URL` pointing at a
 *     pre-launched stack, you must also have `ws://localhost:1234`
 *     reachable; the `beforeAll` hook below probes `/health` and
 *     fails loudly if not.
 *
 * Why two `browser.newContext()` instead of two `page.newPage()` on
 * the same context: presence/awareness identity is keyed by the
 * `RoomClient` constructor (`useStableTabId` → `sessionStorage`),
 * which is shared per-context. Distinct contexts give us distinct
 * peers, which is what the realtime stack is designed for.
 */

const RT_PORT = Number(process.env.OAI_RT_PORT ?? 1234);
const RT_HEALTH_URL = `http://localhost:${RT_PORT}/health`;

test.describe.serial("realtime collab @realtime", () => {
  test.beforeAll(async () => {
    let ok = false;
    try {
      const res = await fetch(RT_HEALTH_URL, { method: "GET" });
      ok = res.ok;
    } catch {
      ok = false;
    }
    if (!ok) {
      throw new Error(
        `realtime server not running on ws://localhost:${RT_PORT} — ` +
          `start it with \`pnpm --filter @officeai/realtime-server dev\` ` +
          `or rely on the playwright.config.ts webServer entry.`
      );
    }
  });

  test("edit in tab A appears in tab B and presence updates in both tabs @realtime", async ({
    browser,
  }) => {
    // Same `tabFallback` in both contexts → same `roomIdForSource`
    // result → both peers join the same Yjs room without needing a
    // hosted .xlsx URL. The key (`officeai.tabId.xlsx`) mirrors the
    // shape used by `useStableTabId("xlsx")` in `XlsxEditor.tsx`.
    const sharedTabId = `e2e-realtime-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;

    const ctxA: BrowserContext = await browser.newContext();
    const ctxB: BrowserContext = await browser.newContext();
    try {
      for (const ctx of [ctxA, ctxB]) {
        await ctx.addInitScript((id: string) => {
          try {
            window.sessionStorage.setItem("officeai.tabId.xlsx", id);
          } catch {
            /* noop — quota / private mode */
          }
        }, sharedTabId);
      }

      const a: Page = await ctxA.newPage();
      const b: Page = await ctxB.newPage();

      await Promise.all([a.goto("/xlsx-editor"), b.goto("/xlsx-editor")]);

      // Same wait the rest of the XLSX suite uses — the seeded sample
      // workbook's B1 header lands once `XlsxAgent.fromBuffer` has
      // mounted.
      await Promise.all([
        expect(a.getByTestId("cell-B1")).toContainText("Score", { timeout: 20_000 }),
        expect(b.getByTestId("cell-B1")).toContainText("Score", { timeout: 20_000 }),
      ]);

      // Both editors must observe ≥1 remote peer before we drive any
      // interaction. The list element only renders when there's at
      // least one peer (see `RemotePresenceList.tsx`), so a 0 count
      // here also covers "list not yet mounted".
      await expect
        .poll(() => peerCardCount(a), { timeout: 20_000, intervals: [200, 500, 1000] })
        .toBeGreaterThan(0);
      await expect
        .poll(() => peerCardCount(b), { timeout: 20_000, intervals: [200, 500, 1000] })
        .toBeGreaterThan(0);

      // Sentinel value typed via the same path `xlsx-typing.spec.ts`
      // uses (printable key on a focused cell → live draft → Enter
      // commits through the command bus). Using digits + lowercase
      // keeps things safe across keyboard layouts and avoids any
      // formula-mode interpretation.
      const sentinel = `e2e${Date.now().toString(36)}`;
      await a.getByTestId("cell-B5").click();
      // Confirm the click registered as the active selection — the
      // formula bar input is the canonical surface mirror in
      // `XlsxEditor.tsx`. Empty B5 ⇒ empty formula input.
      await expect(a.getByTestId("formula-input")).toHaveValue("");
      await a.keyboard.type(sentinel, { delay: 25 });
      await a.keyboard.press("Enter");

      // Local sanity in tab A: the value committed locally before the
      // remote-mirror assertion fires. If this fails we know the bug
      // is in the editor, not the realtime pipe.
      await expect(a.getByTestId("cell-B5")).toContainText(sentinel, { timeout: 5_000 });

      // The actual realtime assertion: the same sentinel surfaces in
      // tab B's grid. Generous timeout because the round-trip goes
      // through Yjs sync + agent applyCommand + React render.
      await expect
        .poll(
          async () => (await b.getByTestId("cell-B5").textContent()) ?? "",
          { timeout: 10_000, intervals: [200, 500, 1000] }
        )
        .toContain(sentinel);

      // Move the selection in tab A so its presence carries a fresh,
      // non-empty range. We don't reuse the post-Enter caret position
      // because that's "wherever Enter dropped us" and varies with
      // the editor's commit-then-advance behaviour.
      await a.getByTestId("cell-D7").click();

      // Tab B's presence list contains a row whose summary matches
      // `is on <Sheet>!<Range>` (see `describePeer` in
      // `RemotePresenceList.tsx`). We don't pin to "D7" specifically
      // because the editor may have committed the typed value into a
      // different active cell first; what we care about is that *some*
      // non-empty cell-range reaches the peer.
      await expect
        .poll(
          async () => {
            const list = b.getByTestId("realtime-presence-list");
            if ((await list.count()) === 0) return "";
            return (await list.textContent()) ?? "";
          },
          { timeout: 10_000, intervals: [200, 500, 1000] }
        )
        .toMatch(/is on \S+!\S+/);
    } finally {
      // Explicit close so the y-websocket server promptly drops the
      // peers (otherwise the next test's "≥1 peer" wait would race
      // against this run's lingering awareness state).
      await ctxA.close();
      await ctxB.close();
    }
  });
});

async function peerCardCount(page: Page): Promise<number> {
  const list = page.getByTestId("realtime-presence-list");
  if ((await list.count()) === 0) return 0;
  return list.locator("> div").count();
}
