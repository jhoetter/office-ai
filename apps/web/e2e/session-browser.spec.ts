import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { expect, test } from "@playwright/test";

type OfficeFormat = "docx" | "xlsx" | "pptx" | "pdf";

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

test.beforeEach(async ({ page }) => {
  await page.route("**/api/sample-files", async (route) => {
    await route.fulfill({ json: { files: [] } });
  });
});

test("home page shows an empty local workspace state", async ({ page }) => {
  await page.route("**/api/sessions", async (route) => {
    await route.fulfill({
      json: {
        schema: "office-ai/web-sessions@1",
        sessions: [],
        documents: [],
      },
    });
  });

  await page.goto("/");

  await expect(page.getByRole("heading", { name: "Local workspace" })).toBeVisible();
  await expect(page.getByText("No local sessions yet")).toBeVisible();
});

test("home page lists local sessions, documents, pending changes and diagnostics", async ({ page }) => {
  let changeApproved = false;
  const sessionBytes = await mockSessionBytes(page, {
    documentId: "doc_1",
    format: "docx",
    filename: "proposal.docx",
    bytes: fixtureBytes("docx"),
    revision: 4,
  });
  await page.route("**/api/sessions/doc_1/export", async (route) => {
    expect(route.request().method()).toBe("POST");
    await route.fulfill({
      status: 200,
      headers: {
        "content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "content-disposition": 'attachment; filename="proposal.docx"',
      },
      body: Buffer.from("exported-docx"),
    });
  });
  await page.route("**/api/sessions/doc_1/changes/mut_1", async (route) => {
    expect(route.request().method()).toBe("POST");
    expect(route.request().postDataJSON()).toEqual({ decision: "approve" });
    changeApproved = true;
    await route.fulfill({
      json: {
        schema: "office-ai/web-change-review@1",
        decision: "approved",
        mutationId: "mut_1",
      },
    });
  });
  await page.route("**/api/sessions/doc_1", async (route) => {
    await route.fulfill({
      json: {
        schema: "office-ai/web-document@1",
        session: {
          sessionId: "session_1",
          title: "MCP review",
          createdAt: "2026-06-24T10:00:00.000Z",
          updatedAt: "2026-06-24T10:03:00.000Z",
          documentCount: 2,
        },
        document: {
          documentId: "doc_1",
          sessionId: "session_1",
          format: "docx",
          name: "proposal.docx",
          status: "ready",
          createdAt: "2026-06-24T10:00:00.000Z",
          updatedAt: "2026-06-24T10:03:00.000Z",
          revision: 4,
          diagnostics: changeApproved
            ? [
                { level: "warning", code: "needs-review", message: "Review layout." },
                { level: "info", code: "change-approved", message: "Approved change mut_1." },
              ]
            : [{ level: "warning", code: "needs-review", message: "Review layout." }],
          exportCount: 1,
          pendingChangeCount: changeApproved ? 0 : 1,
          commandLogCount: changeApproved ? 5 : 4,
          artifacts: { hasOriginal: true, hasWorking: true },
          exports: [{ bytes: 4096, exportedAt: "2026-06-24T10:03:00.000Z" }],
          pendingChanges: [
            {
              id: "mut_1",
              operation: "docx.replace-text",
              status: changeApproved ? "approved" : "pending",
              source: "agent",
              actorId: "assistant",
              timestamp: 1782295380000,
              hasDiff: true,
              diffSummary: "docx.replace-text: 1 DOCX change; low review risk.",
            },
          ],
          commandLog: [
            {
              id: "log_1",
              commandId: "cmd_1",
              operation: "docx.replace-text",
              status: "pending",
              stage: "previewed",
              source: "agent",
              actorId: "assistant",
              recordedAt: "2026-06-24T10:02:00.000Z",
              hasDiff: true,
              diagnostics: [{ level: "info", code: "preview-ready", message: "Preview is ready." }],
            },
            ...(changeApproved
              ? [
                  {
                    id: "log_approved",
                    operation: "docx.replace-text",
                    status: "approved",
                    stage: "reviewed",
                    source: "web",
                    actorId: "assistant",
                    recordedAt: "2026-06-24T10:04:00.000Z",
                    hasDiff: false,
                    diagnostics: [
                      { level: "info", code: "change-approved", message: "Approved change mut_1." },
                    ],
                  },
                ]
              : []),
          ],
        },
      },
    });
  });
  await page.route("**/api/sessions", async (route) => {
    await route.fulfill({
      json: {
        schema: "office-ai/web-sessions@1",
        sessions: [
          {
            sessionId: "session_1",
            title: "MCP review",
            createdAt: "2026-06-24T10:00:00.000Z",
            updatedAt: "2026-06-24T10:03:00.000Z",
            documentCount: 2,
          },
        ],
        documents: [
          {
            documentId: "doc_1",
            sessionId: "session_1",
            format: "docx",
            name: "proposal.docx",
            status: "ready",
            createdAt: "2026-06-24T10:00:00.000Z",
            updatedAt: "2026-06-24T10:03:00.000Z",
            revision: 4,
            diagnostics: [{ level: "info", code: "command-pending", message: "Pending review." }],
            exportCount: 1,
            pendingChangeCount: 1,
            commandLogCount: 4,
            artifacts: { hasOriginal: true, hasWorking: true },
          },
          {
            documentId: "doc_2",
            sessionId: "session_1",
            format: "pdf",
            name: "appendix.pdf",
            status: "ready",
            createdAt: "2026-06-24T10:01:00.000Z",
            updatedAt: "2026-06-24T10:02:00.000Z",
            revision: 1,
            diagnostics: [],
            exportCount: 0,
            pendingChangeCount: 0,
            commandLogCount: 1,
            artifacts: { hasOriginal: true, hasWorking: true },
          },
        ],
      },
    });
  });

  await page.goto("/");

  await expect(page.getByText("MCP review")).toBeVisible();
  await expect(page.getByText("proposal.docx")).toBeVisible();
  await expect(page.getByText("appendix.pdf")).toBeVisible();
  await expect(page.getByText("Pending: 1")).toBeVisible();
  await expect(page.getByText("command-pending")).toBeVisible();
  await expect(page.getByText("web-parity-pdf-review-only")).toBeVisible();

  await page.getByRole("link", { name: "proposal.docx" }).click();

  await expect(page).toHaveURL(/\/editor\?session=doc_1/);
  await expect(page.locator(".ProseMirror").first()).toBeVisible({ timeout: 20_000 });
  expect(sessionBytes.getCount()).toBeGreaterThan(0);

  await page.goto("/");
  await page.locator("tr", { hasText: "proposal.docx" }).getByRole("link", { name: "Inspector" }).click();

  await expect(page.getByRole("heading", { name: "proposal.docx" })).toBeVisible();
  await expect(page.getByText("MCP review")).toBeVisible();
  await expect(page.getByText("Web format parity")).toBeVisible();
  await expect(page.getByText("DOCX web editing")).toBeVisible();
  await expect(page.getByText("docx.replace-text · previewed")).toBeVisible();
  await expect(page.getByText("needs-review")).toBeVisible();
  await expect(page.getByText("web-parity-docx-partial-edit")).toBeVisible();
  await expect(page.getByText("docx.replace-text: 1 DOCX change; low review risk.")).toBeVisible();
  await expect(page.getByText("4.0 KB")).toBeVisible();

  await page.getByRole("button", { name: "Approve" }).click();
  await expect(page.getByText("change-approved")).toBeVisible();
  await expect(page.getByText("approved", { exact: true }).first()).toBeVisible();
  expect(changeApproved).toBe(true);

  const downloadPromise = page.waitForEvent("download");
  await page.getByRole("button", { name: "Export" }).click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toBe("proposal.docx");
});

test("document detail rejects a pending change and can undo the review decision", async ({ page }) => {
  let reviewState: "pending" | "rejected" = "pending";
  let reviewUndone = false;
  await page.route("**/api/sessions/doc_reject/changes/mut_reject", async (route) => {
    expect(route.request().method()).toBe("POST");
    if (reviewState === "pending") {
      expect(route.request().postDataJSON()).toEqual({ decision: "reject" });
      reviewState = "rejected";
      await route.fulfill({
        json: {
          schema: "office-ai/web-change-review@1",
          decision: "rejected",
          mutationId: "mut_reject",
        },
      });
      return;
    }

    expect(route.request().postDataJSON()).toEqual({ decision: "undo" });
    reviewState = "pending";
    reviewUndone = true;
    await route.fulfill({
      json: {
        schema: "office-ai/web-change-review@1",
        decision: "pending",
        mutationId: "mut_reject",
      },
    });
  });
  await page.route("**/api/sessions/doc_reject", async (route) => {
    await route.fulfill({
      json: {
        schema: "office-ai/web-document@1",
        session: {
          sessionId: "session_reject",
          title: "Rejected edits",
          createdAt: "2026-06-24T14:00:00.000Z",
          updatedAt: "2026-06-24T14:00:00.000Z",
          documentCount: 1,
        },
        document: {
          documentId: "doc_reject",
          sessionId: "session_reject",
          format: "docx",
          name: "reject.docx",
          status: "ready",
          createdAt: "2026-06-24T14:00:00.000Z",
          updatedAt: "2026-06-24T14:00:00.000Z",
          revision: 2,
          diagnostics:
            reviewState === "rejected"
              ? [{ level: "info", code: "change-rejected", message: "Rejected change mut_reject." }]
              : reviewUndone
                ? [
                    {
                      level: "info",
                      code: "change-review-undone",
                      message: "Moved change mut_reject back to pending review.",
                    },
                  ]
                : [],
          exportCount: 0,
          pendingChangeCount: reviewState === "pending" ? 1 : 0,
          commandLogCount: reviewUndone ? 3 : reviewState === "rejected" ? 2 : 1,
          artifacts: { hasOriginal: true, hasWorking: true },
          exports: [],
          pendingChanges: [
            {
              id: "mut_reject",
              operation: "docx.replace-text",
              status: reviewState,
              source: "mcp",
              actorId: "assistant",
              timestamp: 1782309600000,
              hasDiff: true,
              diffSummary: "docx.replace-text: 1 DOCX change; low review risk.",
              ...(reviewState === "rejected"
                ? { rejection: { code: "human-rejected", message: "Rejected in web review." } }
                : {}),
            },
          ],
          commandLog: [
            {
              id: "log_preview",
              operation: "docx.replace-text",
              status: "pending",
              stage: "previewed",
              source: "mcp",
              actorId: "assistant",
              recordedAt: "2026-06-24T14:00:00.000Z",
              hasDiff: true,
              diagnostics: [],
            },
            ...(reviewState === "rejected" || reviewUndone
              ? [
                  {
                    id: "log_rejected",
                    operation: "docx.replace-text",
                    status: "rejected",
                    stage: "reviewed",
                    source: "web",
                    actorId: "assistant",
                    recordedAt: "2026-06-24T14:01:00.000Z",
                    hasDiff: false,
                    diagnostics: [
                      { level: "info", code: "change-rejected", message: "Rejected change mut_reject." },
                    ],
                  },
                ]
              : []),
            ...(reviewUndone
              ? [
                  {
                    id: "log_undone",
                    operation: "docx.replace-text",
                    status: "pending",
                    stage: "review-undone",
                    source: "web",
                    actorId: "assistant",
                    recordedAt: "2026-06-24T14:02:00.000Z",
                    hasDiff: false,
                    diagnostics: [
                      {
                        level: "info",
                        code: "change-review-undone",
                        message: "Moved change mut_reject back to pending review.",
                      },
                    ],
                  },
                ]
              : []),
          ],
        },
      },
    });
  });

  await page.goto("/sessions/doc_reject");

  await expect(page.getByRole("heading", { name: "reject.docx" })).toBeVisible();
  await expect(page.getByRole("button", { name: "Reject" })).toBeVisible();

  await page.getByRole("button", { name: "Reject" }).click();

  await expect(page.getByText("change-rejected")).toBeVisible();
  await expect(page.getByText("Rejection: Rejected in web review.")).toBeVisible();
  await expect(page.getByText("rejected", { exact: true }).first()).toBeVisible();
  expect(reviewState).toBe("rejected");

  await page.getByRole("button", { name: "Undo review" }).click();

  await expect(page.getByText("change-review-undone")).toBeVisible();
  await expect(page.getByRole("button", { name: "Reject" })).toBeVisible();
  await expect(page.getByText("pending", { exact: true }).first()).toBeVisible();
  await expect(page.getByText("Rejection: Rejected in web review.")).toHaveCount(0);
  expect(reviewState).toBe("pending");
  expect(reviewUndone).toBe(true);
});

test("document detail exposes web format parity for every format", async ({ page }) => {
  for (const format of ["docx", "xlsx", "pptx", "pdf"] as const) {
    await page.route(`**/api/sessions/doc_${format}`, async (route) => {
      await route.fulfill({
        json: {
          schema: "office-ai/web-document@1",
          session: {
            sessionId: `session_${format}`,
            title: `${format.toUpperCase()} parity`,
            createdAt: "2026-06-24T15:00:00.000Z",
            updatedAt: "2026-06-24T15:01:00.000Z",
            documentCount: 1,
          },
          document: {
            documentId: `doc_${format}`,
            sessionId: `session_${format}`,
            format,
            name: `parity.${format}`,
            status: "ready",
            createdAt: "2026-06-24T15:00:00.000Z",
            updatedAt: "2026-06-24T15:01:00.000Z",
            revision: 1,
            diagnostics: [],
            exportCount: 0,
            pendingChangeCount: 0,
            commandLogCount: 0,
            artifacts: { hasOriginal: true, hasWorking: true },
            exports: [],
            pendingChanges: [],
            commandLog: [],
          },
        },
      });
    });

    await page.goto(`/sessions/doc_${format}`);
    await expect(page.getByRole("heading", { name: `parity.${format}` })).toBeVisible();
    await expect(page.getByText("Web format parity")).toBeVisible();
    await expect(
      page.getByText(format === "pdf" ? "PDF web review" : `${format.toUpperCase()} web editing`)
    ).toBeVisible();
    await expect(
      page.getByText(format === "pdf" ? "web-parity-pdf-review-only" : `web-parity-${format}-partial-edit`)
    ).toBeVisible();
    await page.screenshot({ path: `test-results/format-parity-${format}.png`, fullPage: true });
    await page.unroute(`**/api/sessions/doc_${format}`);
  }
});

for (const format of ["docx", "xlsx", "pptx", "pdf"] as const) {
  test(`fixture session opens, edits, saves and reopens ${format}`, async ({ page }) => {
    const documentId = `doc_matrix_${format}`;
    const filename = `matrix.${format}`;
    const sessionBytes = await mockSessionBytes(page, {
      documentId,
      format,
      filename,
      bytes: fixtureBytes(format),
      revision: 0,
    });
    await page.route("**/api/sessions", async (route) => {
      await route.fulfill({
        json: {
          schema: "office-ai/web-sessions@1",
          sessions: [
            {
              sessionId: `session_matrix_${format}`,
              title: `${format.toUpperCase()} matrix`,
              createdAt: "2026-06-25T06:00:00.000Z",
              updatedAt: "2026-06-25T06:00:00.000Z",
              documentCount: 1,
            },
          ],
          documents: [
            {
              documentId,
              sessionId: `session_matrix_${format}`,
              format,
              name: filename,
              status: "ready",
              createdAt: "2026-06-25T06:00:00.000Z",
              updatedAt: "2026-06-25T06:00:00.000Z",
              revision: sessionBytes.revision(),
              diagnostics: [],
              exportCount: 0,
              pendingChangeCount: 0,
              commandLogCount: sessionBytes.saveCount(),
              artifacts: { hasOriginal: true, hasWorking: true },
            },
          ],
        },
      });
    });

    await page.goto("/");
    await page.getByRole("link", { name: filename }).click();
    await expect(page).toHaveURL(new RegExp(`${escapeRegExp(EDITOR_PATHS[format])}\\?session=${documentId}`));

    await waitForSessionEditor(page, format);
    await applyFormatEdit(page, format);
    await expect(page.getByTestId("shell-save-state-modified")).toBeVisible({ timeout: 10_000 });

    await page.getByTestId("shell-save").click();
    await expect.poll(() => sessionBytes.saveCount(), { timeout: 15_000 }).toBe(1);
    await expect(page.getByTestId("shell-save-state-saved")).toBeVisible({ timeout: 15_000 });

    await page.reload();
    await waitForSessionEditor(page, format);
    await assertFormatEditSurvived(page, format);
  });
}

test("session editor shows a clear error when the bytes endpoint reports another format", async ({
  page,
}) => {
  await page.route("**/api/sessions/doc_wrong/bytes", async (route) => {
    await route.fulfill({
      status: 200,
      headers: sessionByteHeaders({
        documentId: "doc_wrong",
        sessionId: "session_wrong",
        format: "xlsx",
        filename: "wrong.xlsx",
        revision: 0,
      }),
      body: fixtureBytes("xlsx"),
    });
  });

  await page.goto("/editor?session=doc_wrong");

  await expect(page.getByRole("heading", { name: "Couldn't open the session document" })).toBeVisible();
  await expect(page.getByText("not docx")).toBeVisible();
});

test("home page imports an uploaded document into the local workspace", async ({ page }) => {
  let imported = false;
  const uploadedPayload = {
    schema: "office-ai/web-sessions@1",
    sessions: [
      {
        sessionId: "session_upload",
        title: "Web imports",
        createdAt: "2026-06-24T11:00:00.000Z",
        updatedAt: "2026-06-24T11:00:00.000Z",
        documentCount: 1,
      },
    ],
    documents: [
      {
        documentId: "doc_upload",
        sessionId: "session_upload",
        format: "docx",
        name: "upload.docx",
        status: "ready",
        createdAt: "2026-06-24T11:00:00.000Z",
        updatedAt: "2026-06-24T11:00:00.000Z",
        revision: 0,
        diagnostics: [{ level: "info", code: "imported", message: "Imported upload.docx as docx." }],
        exportCount: 0,
        pendingChangeCount: 0,
        commandLogCount: 1,
        artifacts: { hasOriginal: true, hasWorking: true },
      },
    ],
  };

  await page.route("**/api/sessions/import", async (route) => {
    expect(route.request().method()).toBe("POST");
    imported = true;
    await route.fulfill({
      status: 201,
      json: {
        schema: "office-ai/web-import@1",
        session: uploadedPayload.sessions[0],
        document: {
          ...uploadedPayload.documents[0],
          exports: [],
          pendingChanges: [],
          commandLog: [
            {
              id: "log_upload",
              operation: "import_document",
              status: "applied",
              stage: "imported",
              source: "web",
              recordedAt: "2026-06-24T11:00:00.000Z",
              hasDiff: false,
              diagnostics: [{ level: "info", code: "imported", message: "Imported upload.docx as docx." }],
            },
          ],
        },
      },
    });
  });
  await page.route("**/api/sessions", async (route) => {
    await route.fulfill({
      json: imported
        ? uploadedPayload
        : {
            schema: "office-ai/web-sessions@1",
            sessions: [],
            documents: [],
          },
    });
  });

  await page.goto("/");
  await expect(page.getByText("No local sessions yet")).toBeVisible();

  const chooserPromise = page.waitForEvent("filechooser");
  await page.getByRole("button", { name: "Import document" }).click();
  const chooser = await chooserPromise;
  await chooser.setFiles({
    name: "upload.docx",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: Buffer.from("uploaded-docx"),
  });

  await expect(page.getByText("upload.docx")).toBeVisible();
  await expect(page.getByText("Web imports")).toBeVisible();
  expect(imported).toBe(true);
});

test("home page creates a persisted blank workspace document", async ({ page }) => {
  let created = false;
  const createdPayload = {
    schema: "office-ai/web-sessions@1",
    sessions: [
      {
        sessionId: "session_create",
        title: "Web creates",
        createdAt: "2026-06-24T11:30:00.000Z",
        updatedAt: "2026-06-24T11:30:00.000Z",
        documentCount: 1,
      },
    ],
    documents: [
      {
        documentId: "doc_create",
        sessionId: "session_create",
        format: "docx",
        name: "untitled.docx",
        status: "ready",
        createdAt: "2026-06-24T11:30:00.000Z",
        updatedAt: "2026-06-24T11:30:00.000Z",
        revision: 0,
        diagnostics: [{ level: "info", code: "created", message: "Created blank docx document." }],
        exportCount: 0,
        pendingChangeCount: 0,
        commandLogCount: 1,
        artifacts: { hasOriginal: false, hasWorking: true },
      },
    ],
  };

  await page.route("**/api/sessions/create", async (route) => {
    expect(route.request().method()).toBe("POST");
    expect(route.request().postDataJSON()).toEqual({ format: "docx" });
    created = true;
    await route.fulfill({
      status: 201,
      json: {
        schema: "office-ai/web-create@1",
        session: createdPayload.sessions[0],
        document: {
          ...createdPayload.documents[0],
          exports: [],
          pendingChanges: [],
          commandLog: [
            {
              id: "log_create",
              operation: "create_document",
              status: "applied",
              stage: "created",
              source: "web",
              recordedAt: "2026-06-24T11:30:00.000Z",
              hasDiff: false,
              diagnostics: [{ level: "info", code: "created", message: "Created blank docx document." }],
            },
          ],
        },
      },
    });
  });
  await page.route("**/api/sessions", async (route) => {
    await route.fulfill({
      json: created
        ? createdPayload
        : {
            schema: "office-ai/web-sessions@1",
            sessions: [],
            documents: [],
          },
    });
  });

  await page.goto("/");
  await expect(page.getByText("No local sessions yet")).toBeVisible();

  await page.getByRole("button", { name: "DOCX" }).click();

  await expect(page.getByText("untitled.docx")).toBeVisible();
  await expect(page.getByText("Web creates")).toBeVisible();
  expect(created).toBe(true);
});

test("home page keeps the local workspace readable on mobile width", async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await page.route("**/api/sessions", async (route) => {
    await route.fulfill({
      json: {
        schema: "office-ai/web-sessions@1",
        sessions: [
          {
            sessionId: "session_mobile",
            title: "Mobile review",
            createdAt: "2026-06-24T12:00:00.000Z",
            updatedAt: "2026-06-24T12:01:00.000Z",
            documentCount: 1,
          },
        ],
        documents: [
          {
            documentId: "doc_mobile",
            sessionId: "session_mobile",
            format: "pdf",
            name: "mobile.pdf",
            status: "ready",
            createdAt: "2026-06-24T12:00:00.000Z",
            updatedAt: "2026-06-24T12:01:00.000Z",
            revision: 1,
            diagnostics: [],
            exportCount: 0,
            pendingChangeCount: 1,
            commandLogCount: 1,
            artifacts: { hasOriginal: true, hasWorking: true },
          },
        ],
      },
    });
  });

  await page.goto("/");

  await expect(page.getByRole("heading", { name: "Local workspace" })).toBeVisible();
  await expect(page.getByText("Mobile review")).toBeVisible();
  await expect(page.getByText("mobile.pdf")).toBeVisible();
  await expect(page.getByText("Pending: 1")).toBeVisible();
});

function fixtureBytes(format: OfficeFormat): Buffer {
  return readFileSync(FIXTURE_PATHS[format]);
}

async function mockSessionBytes(
  page: import("@playwright/test").Page,
  args: {
    readonly documentId: string;
    readonly format: OfficeFormat;
    readonly filename: string;
    readonly bytes: Buffer;
    readonly revision: number;
  }
): Promise<{
  readonly getCount: () => number;
  readonly saveCount: () => number;
  readonly revision: () => number;
}> {
  let bytes = Buffer.from(args.bytes);
  let revision = args.revision;
  let gets = 0;
  let saves = 0;
  await page.route(`**/api/sessions/${args.documentId}/bytes`, async (route) => {
    const request = route.request();
    if (request.method() === "GET") {
      gets += 1;
      await route.fulfill({
        status: 200,
        headers: sessionByteHeaders({
          documentId: args.documentId,
          sessionId: `session_${args.documentId}`,
          format: args.format,
          filename: args.filename,
          revision,
        }),
        body: bytes,
      });
      return;
    }
    if (request.method() === "PUT") {
      expect(request.headers()["if-match"]).toBe(revisionEtag(args.documentId, revision));
      const next = request.postDataBuffer();
      expect(next).not.toBeNull();
      bytes = Buffer.from(next ?? []);
      revision += 1;
      saves += 1;
      await route.fulfill({
        status: 200,
        headers: { etag: revisionEtag(args.documentId, revision) },
        json: {
          schema: "office-ai/session-bytes-save@1",
          etag: revisionEtag(args.documentId, revision),
          document: {
            documentId: args.documentId,
            sessionId: `session_${args.documentId}`,
            format: args.format,
            name: args.filename,
            status: "ready",
            createdAt: "2026-06-25T06:00:00.000Z",
            updatedAt: "2026-06-25T06:01:00.000Z",
            revision,
            diagnostics: [{ level: "info", code: "web-editor-save", message: "Saved from E2E." }],
            exportCount: 0,
            pendingChangeCount: 0,
            commandLogCount: saves,
            artifacts: { hasOriginal: true, hasWorking: true },
            exports: [],
            pendingChanges: [],
            commandLog: [
              {
                id: `log_save_${saves}`,
                operation: "save_document",
                status: "applied",
                stage: "saved",
                source: "web",
                recordedAt: "2026-06-25T06:01:00.000Z",
                hasDiff: false,
                diagnostics: [{ level: "info", code: "web-editor-save", message: "Saved from E2E." }],
              },
            ],
          },
        },
      });
      return;
    }
    await route.fulfill({ status: 405, json: { message: "Method not allowed" } });
  });
  return {
    getCount: () => gets,
    saveCount: () => saves,
    revision: () => revision,
  };
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
    etag: revisionEtag(args.documentId, args.revision),
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

function revisionEtag(documentId: string, revision: number): string {
  return `"officeai:${documentId}:${revision}"`;
}

async function waitForSessionEditor(
  page: import("@playwright/test").Page,
  format: OfficeFormat
): Promise<void> {
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

async function applyFormatEdit(page: import("@playwright/test").Page, format: OfficeFormat): Promise<void> {
  switch (format) {
    case "docx": {
      const editor = page.locator(".ProseMirror").first();
      await editor.click();
      await page.keyboard.type("M1 DOCX ");
      await expect(editor).toContainText("M1 DOCX");
      return;
    }
    case "xlsx":
      await page.getByTestId("cell-B2").click();
      await page.keyboard.type("89");
      await page.keyboard.press("Enter");
      await expect(page.getByTestId("cell-B2")).toContainText("89");
      return;
    case "pptx":
      await page.getByRole("button", { name: "Add slide" }).click();
      await expect(page.getByTestId("pptx-sidebar").getByRole("button")).toHaveCount(2);
      return;
    case "pdf": {
      await page.getByTestId("pdf-open-metadata").click();
      const dialog = page.getByTestId("pdf-metadata-dialog");
      await expect(dialog).toBeVisible();
      await dialog.getByTestId("pdf-meta-title").locator("input").fill("M1 PDF");
      await page.getByTestId("pdf-metadata-save").click();
      await expect(dialog).toHaveCount(0);
      return;
    }
  }
}

async function assertFormatEditSurvived(
  page: import("@playwright/test").Page,
  format: OfficeFormat
): Promise<void> {
  switch (format) {
    case "docx":
      await expect(page.locator(".ProseMirror").first()).toContainText("M1 DOCX");
      return;
    case "xlsx":
      await expect(page.getByTestId("cell-B2")).toContainText("89");
      return;
    case "pptx":
      await expect(page.getByTestId("pptx-sidebar").getByRole("button")).toHaveCount(2);
      return;
    case "pdf": {
      await page.getByTestId("pdf-open-metadata").click();
      await expect(page.getByTestId("pdf-meta-title").locator("input")).toHaveValue("M1 PDF");
      await page.getByTestId("pdf-metadata-cancel").click();
      return;
    }
  }
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
