/**
 * PPTX flavor of the LLM bridge client. Mirrors `llm-client.ts` but
 * targets the `pptx:*` command surface and uses `PptxAgent.toMarkdown()`
 * as the document projection.
 *
 * Behaviour:
 *  - POST `/api/llm` with `{ prompt, snapshotMarkdown, format: "pptx" }`.
 *  - HTTP 200 → return the (already filtered) `pptx:*` commands as
 *    `agent`-sourced commands the host can queue via `agent.applyCommands`.
 *  - HTTP 501 ("LLM bridge not configured" — no `OPENAI_API_KEY`) →
 *    silently fall back to the offline intent parser so the editor stays
 *    usable without any env vars.
 *  - Network / parse errors → also fall back, with the failure surfaced
 *    in `note` so the host can show a toast.
 */

import type { CommandLite } from "@officeai/core";
import type { PptxAgent } from "@officeai/pptx";
import type { TextShape } from "@officeai/pptx";

export interface DispatchResult {
  commands: CommandLite[];
  rationale: string;
  /** When set, the helper used the offline fallback. UI may surface this. */
  note?: string;
}

const AGENT_ID = "web-llm-bridge-pptx";
const ENDPOINT = "/api/llm";

export async function dispatchToLlmPptx(
  prompt: string,
  agent: PptxAgent,
  activeSlideIndex: number
): Promise<DispatchResult> {
  const trimmed = prompt.trim();
  if (!trimmed) return { commands: [], rationale: "Empty prompt; nothing to do." };

  const snapshotMarkdown = agent.toMarkdown();

  let response: Response;
  try {
    response = await fetch(ENDPOINT, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        prompt: trimmed,
        snapshotMarkdown,
        format: "pptx",
      }),
    });
  } catch (err) {
    return offlineFallback(
      trimmed,
      agent,
      activeSlideIndex,
      `Network error reaching ${ENDPOINT}: ${errorMessage(err)}`
    );
  }

  if (response.status === 501) {
    return offlineFallback(trimmed, agent, activeSlideIndex);
  }
  if (!response.ok) {
    return offlineFallback(
      trimmed,
      agent,
      activeSlideIndex,
      `LLM bridge returned HTTP ${response.status}`
    );
  }

  let parsed: { commands?: Array<{ type: string; payload: unknown }>; rationale?: string };
  try {
    parsed = (await response.json()) as typeof parsed;
  } catch (err) {
    return offlineFallback(
      trimmed,
      agent,
      activeSlideIndex,
      `LLM bridge returned non-JSON: ${errorMessage(err)}`
    );
  }

  const rawCommands = Array.isArray(parsed.commands) ? parsed.commands : [];
  const commands: CommandLite[] = rawCommands.map((c) => ({
    type: c.type,
    payload: c.payload,
    source: "agent",
    agentId: AGENT_ID,
  }));
  const rationale = typeof parsed.rationale === "string" ? parsed.rationale : "";
  return { commands, rationale };
}

/**
 * Demo intent parser, the same one PptxEditor used pre-bridge:
 *   "add slide"          → pptx:add-slide
 *   "delete slide"       → pptx:delete-slide(activeSlideIndex)
 *   "bold|italic|under-" → pptx:format-text on the first text shape
 *   anything else        → pptx:add-text-box with the prompt as content
 */
function offlineFallback(
  prompt: string,
  agent: PptxAgent,
  activeSlideIndex: number,
  note?: string
): DispatchResult {
  const lower = prompt.toLowerCase();
  const commands: CommandLite[] = [];
  let rationale = "";

  if (/^add (a )?slide/.test(lower)) {
    commands.push({
      type: "pptx:add-slide",
      payload: {},
      source: "agent",
      agentId: AGENT_ID,
    });
    rationale = "Offline fallback: append a blank slide.";
  } else if (/^delete (the )?slide/.test(lower)) {
    commands.push({
      type: "pptx:delete-slide",
      payload: { slideIndex: activeSlideIndex },
      source: "agent",
      agentId: AGENT_ID,
    });
    rationale = "Offline fallback: delete the active slide.";
  } else if (/bold|italic|underline/.test(lower)) {
    const slide = agent.getSnapshot().root.slides[activeSlideIndex];
    const ts = slide?.shapes.find((s): s is TextShape => s.kind === "text");
    if (ts && ts.txBody.paragraphs.length > 0) {
      const p = ts.txBody.paragraphs[0];
      const flatLen = p.runs.reduce(
        (acc, r) => acc + (r.isLineBreak ? 0 : r.text.length),
        0
      );
      const fmt: { bold?: boolean; italic?: boolean; underline?: boolean } = {};
      if (lower.includes("bold")) fmt.bold = true;
      if (lower.includes("italic")) fmt.italic = true;
      if (lower.includes("underline")) fmt.underline = true;
      commands.push({
        type: "pptx:format-text",
        payload: {
          slideIndex: activeSlideIndex,
          shapeId: ts.id,
          range: { paragraph: 0, start: 0, end: flatLen },
          format: fmt,
        },
        source: "agent",
        agentId: AGENT_ID,
      });
      rationale = "Offline fallback: toggle formatting on the first text shape.";
    } else {
      rationale = "Offline fallback: no text shape on the active slide; nothing to do.";
    }
  } else {
    commands.push({
      type: "pptx:add-text-box",
      payload: {
        slideIndex: activeSlideIndex,
        text: prompt,
        x: 500_000,
        y: 4_500_000,
        width: 6_000_000,
        height: 800_000,
      },
      source: "agent",
      agentId: AGENT_ID,
    });
    rationale = "Offline fallback: drop the prompt as a new text box.";
  }

  return {
    commands,
    rationale,
    ...(note ? { note } : {}),
  };
}

function errorMessage(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}
