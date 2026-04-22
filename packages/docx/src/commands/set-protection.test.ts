import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { DocxAgent } from "../agent/agent.js";
import { makeSyntheticDocx, plainDocxXml } from "../test-utils/synthetic.js";

const SETTINGS_PART = "word/settings.xml";

const SAMPLE_SETTINGS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="720"/>
</w:settings>`;

async function loadAgent(opts?: { withSettings?: boolean }): Promise<DocxAgent> {
  const documentXml = plainDocxXml([{ text: "hello" }]);
  const extra = opts?.withSettings === false ? undefined : { [SETTINGS_PART]: SAMPLE_SETTINGS_XML };
  const buf = await makeSyntheticDocx({ documentXml, extra });
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

describe("docx:set-protection", () => {
  it("inserts <w:documentProtection> when enabled", async () => {
    const agent = await loadAgent();
    await agent.applyCommand({
      type: "docx:set-protection",
      payload: { enabled: true, edit: "readOnly" },
    });
    const settings = agent.getSnapshot().root.settingsXml ?? "";
    expect(settings).toMatch(/<w:documentProtection\b/);
    expect(settings).toMatch(/w:edit="readOnly"/);
    expect(settings).toMatch(/w:enforcement="1"/);
    expect(settings).toMatch(/<w:zoom w:percent="100"\/>/);
  });

  it("removes <w:documentProtection> when disabled", async () => {
    const agent = await loadAgent();
    await agent.applyCommand({
      type: "docx:set-protection",
      payload: { enabled: true },
    });
    await agent.applyCommand({
      type: "docx:set-protection",
      payload: { enabled: false },
    });
    const settings = agent.getSnapshot().root.settingsXml ?? "";
    expect(settings).not.toMatch(/<w:documentProtection\b/);
    expect(settings).toMatch(/<w:zoom w:percent="100"\/>/);
  });

  it("synthesises a settings.xml when none exists", async () => {
    const agent = await loadAgent({ withSettings: false });
    expect(agent.getSnapshot().root.settingsXml).toBeUndefined();

    await agent.applyCommand({
      type: "docx:set-protection",
      payload: { enabled: true, edit: "comments" },
    });
    const settings = agent.getSnapshot().root.settingsXml ?? "";
    expect(settings).toMatch(/<w:settings\b/);
    expect(settings).toMatch(/<w:documentProtection\b/);
    expect(settings).toMatch(/w:edit="comments"/);
  });

  it("rejects plaintext passwords on the payload", async () => {
    const agent = await loadAgent();
    const result = await agent.applyCommand({
      type: "docx:set-protection",
      payload: { enabled: true, password: "secret" } as never,
    });
    expect(result.status).toBe("rejected");
    if (result.status !== "rejected") return;
    expect(result.rejection.message).toMatch(/Plaintext passwords/);
  });

  it("survives a save/reload round-trip", async () => {
    const agent = await loadAgent();
    await agent.applyCommand({
      type: "docx:set-protection",
      payload: { enabled: true, edit: "trackedChanges", formatting: true },
    });
    const bytes = await agent.exportFile();
    const reopened = await DocxAgent.fromBuffer(bytes, { idMinter: deterministicIdMinter() });
    const settings = reopened.getSnapshot().root.settingsXml ?? "";
    expect(settings).toMatch(/w:edit="trackedChanges"/);
    expect(settings).toMatch(/w:formatting="1"/);
  });
});
