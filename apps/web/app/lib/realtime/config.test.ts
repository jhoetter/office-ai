import { describe, expect, it } from "vitest";
import { shouldReplayExistingCommandsForRoom } from "./config";

describe("realtime room command replay", () => {
  it("disables historical command replay for session-backed rooms", () => {
    expect(shouldReplayExistingCommandsForRoom("session:doc_123")).toBe(false);
  });

  it("keeps historical command replay for explicit collaboration rooms", () => {
    expect(shouldReplayExistingCommandsForRoom("demo-room")).toBe(true);
    expect(shouldReplayExistingCommandsForRoom(null)).toBe(true);
    expect(shouldReplayExistingCommandsForRoom(undefined)).toBe(true);
  });
});
