import { CommandError, type CommandHandler } from "@officeai/core";
import type { DocxDocument, DocxSnapshot } from "../model/types.js";
import { buildDiff, buildDiffMulti, evolveSnapshot } from "./helpers.js";
import type { SetProtectionPayload } from "./payloads.js";

/**
 * `docx:set-protection` — toggle Word's "Restrict Editing" / document
 * protection. Patches the `<w:documentProtection>` element inside
 * `word/settings.xml`.
 *
 * The verbatim settings.xml lives on `DocxDocument.settingsXml`
 * (lifted at parse time, written back on `dirty.settings`). We do
 * regex surgery on just the protection element so every other
 * setting (default font, view mode, compat flags) survives untouched.
 *
 * If the document has no settings.xml part the handler synthesises a
 * minimal one carrying just the `<w:settings>` root and the requested
 * protection element. Mirrors what Word does the first time the user
 * enables protection on a doc that was authored without a settings
 * part.
 *
 * Spec: ECMA-376 Part 1, §17.15.1.29 (documentProtection).
 */
export const setProtectionHandler: CommandHandler<SetProtectionPayload, DocxSnapshot> = {
  type: "docx:set-protection",
  apply(snapshot, payload) {
    if ((payload as { password?: unknown }).password !== undefined) {
      throw new CommandError(
        "invalid-payload",
        "Plaintext passwords are not accepted; pass a precomputed passwordHash + algorithmName + saltValue + spinCount instead."
      );
    }

    const current = snapshot.root.settingsXml;
    const next = mutateSettingsXml(current, payload);
    if (next === current) {
      return { next: snapshot, diff: buildDiffMulti(snapshot.revision, snapshot.revision, []) };
    }

    const root: DocxDocument = { ...snapshot.root, settingsXml: next };
    const evolved = evolveSnapshot(snapshot, root, { settings: true });
    return {
      next: evolved,
      diff: buildDiff(snapshot.revision, evolved.revision, {
        kind: "node-updated",
        nodeId: snapshot.root.id,
        path: ["settings", "protection"],
        field: "protection",
        summary: payload.enabled ? `enabled (${payload.edit ?? "readOnly"})` : "disabled",
      }),
    };
  },
};

function mutateSettingsXml(current: string | undefined, payload: SetProtectionPayload): string {
  let xml = current ?? defaultSettingsXml();

  // Always strip an existing protection element first.
  const protectionRe = /<w:documentProtection\b[^>]*?(?:\/>|>[\s\S]*?<\/w:documentProtection>)/;
  xml = xml.replace(protectionRe, "");

  if (!payload.enabled) {
    return xml;
  }

  const block = buildProtectionElement(payload);
  // Splice the protection element into <w:settings>. It belongs near
  // the top — before <w:writeProtection>'s peers — but Word is
  // lenient about ordering. We insert immediately after `<w:settings>`
  // so the result is consistent across runs.
  const settingsOpenRe = /<w:settings\b[^>]*>/;
  const m = settingsOpenRe.exec(xml);
  if (m) {
    const idx = m.index + m[0].length;
    return xml.slice(0, idx) + block + xml.slice(idx);
  }
  // Fallback — settings root could not be located, append before the
  // closing `</w:settings>` tag.
  if (xml.includes("</w:settings>")) {
    return xml.replace("</w:settings>", `${block}</w:settings>`);
  }
  return xml + block;
}

function buildProtectionElement(payload: SetProtectionPayload): string {
  const attrs: string[] = [];
  attrs.push(`w:edit="${escapeXmlAttr(payload.edit ?? "readOnly")}"`);
  if (payload.enforce ?? true) attrs.push(`w:enforcement="1"`);
  if (payload.formatting !== undefined) {
    attrs.push(`w:formatting="${payload.formatting ? "1" : "0"}"`);
  }
  if (payload.algorithmName) {
    attrs.push(`w:cryptProviderType="rsaAES"`);
    attrs.push(`w:cryptAlgorithmClass="hash"`);
    attrs.push(`w:cryptAlgorithmType="typeAny"`);
    attrs.push(`w:cryptAlgorithmSid="${escapeXmlAttr(payload.algorithmName)}"`);
    if (payload.spinCount !== undefined) {
      attrs.push(`w:cryptSpinCount="${Math.max(0, Math.floor(payload.spinCount))}"`);
    }
  }
  if (payload.passwordHash) {
    attrs.push(`w:hash="${escapeXmlAttr(payload.passwordHash)}"`);
  }
  if (payload.saltValue) {
    attrs.push(`w:salt="${escapeXmlAttr(payload.saltValue)}"`);
  }
  return `<w:documentProtection ${attrs.join(" ")}/>`;
}

function defaultSettingsXml(): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:settings>`
  );
}

function escapeXmlAttr(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
