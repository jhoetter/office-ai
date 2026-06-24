import { describe, expect, it } from "vitest";
import type { ActionDescriptor } from "@officeai/core";
import { docxActions } from "@officeai/docx";
import { isCliAutoBindableAction } from "./actions-to-cli.js";
import { isMcpBindableAction } from "./actions-to-mcp.js";

function makeAction(patch: Partial<ActionDescriptor>): ActionDescriptor {
  return {
    id: "docx.test-action",
    commandType: "docx:test-action",
    format: "docx",
    agentCallable: false,
    webCallable: false,
    cliCallable: false,
    requiresReview: true,
    supportsDryRun: false,
    supportsDiff: true,
    commandSchema: "catalogue-args",
    label: "Test action",
    description: "Synthetic binding metadata test action.",
    section: "Test",
    surfaces: [],
    args: [],
    buildPayload: () => ({}),
    ...patch,
  };
}

describe("action binding metadata", () => {
  it("allows an MCP-only action without creating a generated CLI subcommand", () => {
    const action = makeAction({ agentCallable: true, cliCallable: false });

    expect(isMcpBindableAction(action)).toBe(true);
    expect(isCliAutoBindableAction(action)).toBe(false);
  });

  it("allows a generated CLI-only action without creating an MCP tool", () => {
    const action = makeAction({ agentCallable: false, cliCallable: true });

    expect(isCliAutoBindableAction(action)).toBe(true);
    expect(isMcpBindableAction(action)).toBe(false);
  });

  it("keeps terminal read conveniences out of MCP auto-binding", () => {
    const read = docxActions.find((action) => action.id === "docx.read");
    expect(read).toBeDefined();
    expect(read?.cliCallable).toBe(true);
    expect(read?.agentCallable).toBe(false);
    expect(isMcpBindableAction(read as ActionDescriptor)).toBe(false);
    expect(isCliAutoBindableAction(read as ActionDescriptor)).toBe(false);
  });
});
