import { mkdirSync, readFileSync } from "node:fs";
import { join, resolve } from "node:path";
import { expect, test, type Page } from "@playwright/test";

type OfficeFormat = "docx" | "xlsx" | "pptx" | "pdf";

const FORMATS: readonly OfficeFormat[] = ["docx", "xlsx", "pptx", "pdf"];
const OUT_DIR = resolve(process.cwd(), "test-results/sonaloop-visual-parity");

const FIXTURE_PATHS: Record<OfficeFormat, string> = {
  docx: resolve(process.cwd(), "../../fixtures/docx/synthetic/01-plain-paragraphs.docx"),
  xlsx: resolve(process.cwd(), "../../fixtures/xlsx/synthetic/01-single-sheet-numbers.xlsx"),
  pptx: resolve(process.cwd(), "../../fixtures/pptx/synthetic/02-title-only.pptx"),
  pdf: resolve(process.cwd(), "../../fixtures/pdf/simple-text-1page.pdf"),
};

const EDITOR_PATHS: Record<OfficeFormat, string> = {
  docx: "/editor",
  xlsx: "/xlsx-editor",
  pptx: "/pptx-editor",
  pdf: "/pdf-viewer",
};

/**
 * Sonaloop app visual parity gate.
 *
 * This is a deterministic screenshot review workflow rather than a
 * pixel-baseline suite: it fails on missing Sonaloop shell/semantic
 * classes and writes review images to test-results/sonaloop-visual-parity.
 */
test.describe("sonaloop app visual parity", () => {
  test.setTimeout(90_000);

  test.beforeEach(async ({ page }) => {
    await page.route("**/api/sample-files", async (route) => {
      await route.fulfill({ json: { files: [] } });
    });
    await mockVisualWorkspace(page);
  });

  test("captures home, session browser and inspector review surfaces", async ({ page }) => {
    await page.goto("/");

    await expect(page.getByRole("heading", { name: "Local workspace" })).toBeVisible();
    await expect(page.getByText("Visual DOCX review")).toBeVisible();
    await expect(page.locator(".sl-app-main")).toBeVisible();
    await expect(page.locator(".sl-app-topbar")).toBeVisible();
    await expect(page.locator(".sl-card").first()).toBeVisible();
    await expect(page.locator(".sl-table").first()).toBeVisible();
    await expect(page.locator(".sl-badge").first()).toBeVisible();
    await capture(page, "home-session-browser");

    await page.goto("/sessions/doc_visual_docx");

    await expect(page.getByRole("heading", { name: "visual.docx" })).toBeVisible();
    await expect(page.getByText("docx.replace-text: tighten executive summary.")).toBeVisible();
    await expect(page.locator(".sl-app-inspector").first()).toBeVisible();
    await expect(page.locator(".sl-app-inspector__header").first()).toBeVisible();
    await expect(page.locator(".sl-app-inspector__body").first()).toBeVisible();
    await capture(page, "inspector-review");
  });

  test("captures every format editor in the Sonaloop shell", async ({ page }) => {
    for (const format of FORMATS) {
      await page.goto(`${EDITOR_PATHS[format]}?session=doc_visual_${format}`);
      await waitForSessionEditor(page, format);

      await expect(page.locator(".sl-app-main")).toBeVisible();
      await expect(page.locator(".sl-app-topbar")).toBeVisible();
      await expect(page.locator(".sl-app-editor__toolbar")).toBeVisible();
      await expect(page.getByTestId("shell-save")).toBeVisible();
      await expect(page.getByTestId("shell-export")).toBeVisible();
      await capture(page, `editor-${format}`);
    }
  });

  test("captures CMD+K and export dialog Sonaloop surfaces", async ({ page }) => {
    await page.goto("/editor?session=doc_visual_docx");
    await waitForSessionEditor(page, "docx");

    await page.keyboard.press("ControlOrMeta+k");
    await expect(page.locator(".sl-cmdk")).toBeVisible();
    await expect(page.locator(".sl-cmdk-panel")).toBeVisible();
    await expect(page.locator(".sl-cmdk-input")).toBeFocused();
    await capture(page, "cmdk-docx");

    await page.keyboard.press("Escape");
    await expect(page.locator(".sl-cmdk")).toHaveCount(0);

    await page.getByTestId("shell-export").click();
    await page.getByTestId("shell-export-pdf").click();

    await expect(page.getByTestId("shell-export-dialog")).toBeVisible();
    await expect(page.locator(".sl-modal.oa-export-modal")).toBeVisible();
    await expect(page.locator(".sl-modal__panel")).toBeVisible();
    await expect(page.locator(".sl-modal__head")).toBeVisible();
    await expect(page.locator(".sl-modal__body")).toBeVisible();
    await expect(page.locator(".sl-modal__foot")).toBeVisible();
    await capture(page, "export-dialog-docx");
  });
});

async function capture(page: Page, name: string): Promise<void> {
  mkdirSync(OUT_DIR, { recursive: true });
  const image = await page.screenshot({
    path: join(OUT_DIR, `${name}.png`),
    fullPage: true,
    animations: "disabled",
  });
  expect(image.length).toBeGreaterThan(10_000);
}

async function mockVisualWorkspace(page: Page): Promise<void> {
  for (const format of FORMATS) {
    await mockSessionBytes(page, {
      documentId: `doc_visual_${format}`,
      format,
      filename: `visual.${format}`,
      bytes: readFileSync(FIXTURE_PATHS[format]),
      revision: 7,
    });
    await page.route(`**/api/sessions/doc_visual_${format}`, async (route) => {
      await route.fulfill({
        json: {
          schema: "office-ai/web-document@1",
          session: visualSession(format),
          document: visualDocument(format, true),
        },
      });
    });
  }

  await page.route("**/api/sessions", async (route) => {
    await route.fulfill({
      json: {
        schema: "office-ai/web-sessions@1",
        sessions: FORMATS.map(visualSession),
        documents: FORMATS.map((format) => visualDocument(format, false)),
      },
    });
  });
}

function visualSession(format: OfficeFormat) {
  return {
    sessionId: `session_visual_${format}`,
    title: `Visual ${format.toUpperCase()} review`,
    createdAt: "2026-06-25T07:00:00.000Z",
    updatedAt: "2026-06-25T07:03:00.000Z",
    documentCount: 1,
  };
}

function visualDocument(format: OfficeFormat, detail: boolean) {
  const hasReview = format === "docx";
  return {
    documentId: `doc_visual_${format}`,
    sessionId: `session_visual_${format}`,
    format,
    name: `visual.${format}`,
    status: "ready",
    createdAt: "2026-06-25T07:00:00.000Z",
    updatedAt: "2026-06-25T07:03:00.000Z",
    revision: 7,
    diagnostics: hasReview
      ? [{ level: "info", code: "visual-review", message: "Pending review for Sonaloop shell parity." }]
      : [],
    exportCount: format === "pdf" ? 0 : 1,
    pendingChangeCount: hasReview ? 1 : 0,
    commandLogCount: hasReview ? 2 : 1,
    artifacts: { hasOriginal: true, hasWorking: true },
    ...(detail
      ? {
          exports: format === "pdf" ? [] : [{ bytes: 4096, exportedAt: "2026-06-25T07:02:00.000Z" }],
          pendingChanges: hasReview
            ? [
                {
                  id: "mut_visual_1",
                  operation: "docx.replace-text",
                  status: "pending",
                  source: "mcp",
                  actorId: "assistant",
                  timestamp: 1782370980000,
                  hasDiff: true,
                  diffSummary: "docx.replace-text: tighten executive summary.",
                },
              ]
            : [],
          commandLog: [
            {
              id: "log_visual_1",
              operation: hasReview ? "docx.replace-text" : `${format}.open`,
              status: hasReview ? "pending" : "applied",
              stage: hasReview ? "previewed" : "opened",
              source: "mcp",
              actorId: "assistant",
              recordedAt: "2026-06-25T07:01:00.000Z",
              hasDiff: hasReview,
              diagnostics: [],
            },
          ],
        }
      : {}),
  };
}

async function mockSessionBytes(
  page: Page,
  args: {
    readonly documentId: string;
    readonly format: OfficeFormat;
    readonly filename: string;
    readonly bytes: Buffer;
    readonly revision: number;
  }
): Promise<void> {
  await page.route(`**/api/sessions/${args.documentId}/bytes`, async (route) => {
    if (route.request().method() !== "GET") {
      await route.fulfill({ status: 405, json: { message: "Method not allowed" } });
      return;
    }
    await route.fulfill({
      status: 200,
      headers: sessionByteHeaders({
        documentId: args.documentId,
        sessionId: `session_${args.documentId}`,
        format: args.format,
        filename: args.filename,
        revision: args.revision,
      }),
      body: args.bytes,
    });
  });
}

function sessionByteHeaders(args: {
  readonly documentId: string;
  readonly sessionId: string;
  readonly format: OfficeFormat;
  readonly filename: string;
  readonly revision: number;
}): Record<string, string> {
  return {
    "content-type": mimeForFormat(args.format),
    "content-disposition": `attachment; filename="${args.filename}"`,
    "cache-control": "no-store",
    etag: `"officeai:${args.documentId}:${args.revision}"`,
    "x-officeai-document-id": args.documentId,
    "x-officeai-session-id": args.sessionId,
    "x-officeai-format": args.format,
    "x-officeai-filename": encodeURIComponent(args.filename),
    "x-officeai-revision": String(args.revision),
  };
}

function mimeForFormat(format: OfficeFormat): string {
  switch (format) {
    case "docx":
      return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    case "xlsx":
      return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    case "pptx":
      return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
    case "pdf":
      return "application/pdf";
  }
}

async function waitForSessionEditor(page: Page, format: OfficeFormat): Promise<void> {
  switch (format) {
    case "docx":
      await expect(page.locator(".ProseMirror").first()).toBeVisible({ timeout: 20_000 });
      return;
    case "xlsx":
      await expect(page.getByTestId("cell-B2")).toBeVisible({ timeout: 20_000 });
      return;
    case "pptx":
      await expect(page.getByTestId("pptx-sidebar")).toBeVisible({ timeout: 20_000 });
      await expect(page.getByTestId("pptx-sidebar").getByRole("button").first()).toBeVisible({
        timeout: 20_000,
      });
      return;
    case "pdf":
      await expect(page.getByTestId("pdf-canvas")).toBeVisible({ timeout: 20_000 });
      return;
  }
}
