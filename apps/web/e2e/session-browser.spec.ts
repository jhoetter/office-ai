import { expect, test } from "@playwright/test";

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
          diagnostics: [{ level: "warning", code: "needs-review", message: "Review layout." }],
          exportCount: 1,
          pendingChangeCount: 2,
          commandLogCount: 4,
          artifacts: { hasOriginal: true, hasWorking: true },
          exports: [{ bytes: 4096, exportedAt: "2026-06-24T10:03:00.000Z" }],
          pendingChanges: [
            {
              id: "mut_1",
              operation: "docx.replace-text",
              status: "pending",
              source: "agent",
              actorId: "assistant",
              timestamp: 1782295380000,
              hasDiff: true,
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
            pendingChangeCount: 2,
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
  await expect(page.getByText("Pending: 2")).toBeVisible();
  await expect(page.getByText("command-pending")).toBeVisible();

  await page.getByRole("link", { name: "proposal.docx" }).click();

  await expect(page.getByRole("heading", { name: "proposal.docx" })).toBeVisible();
  await expect(page.getByText("MCP review")).toBeVisible();
  await expect(page.getByText("docx.replace-text · previewed")).toBeVisible();
  await expect(page.getByText("needs-review")).toBeVisible();
  await expect(page.getByText("4.0 KB")).toBeVisible();

  const downloadPromise = page.waitForEvent("download");
  await page.getByRole("button", { name: "Export" }).click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toBe("proposal.docx");
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
