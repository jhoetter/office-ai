/**
 * Isomorphic LLM client helper.
 *
 * `dispatchToLlm(prompt, agent)` returns an array of `Command`-shaped
 * objects ready to be queued via `agent.applyCommands(...)`. The helper
 * intentionally does NOT call `applyCommand` itself — that's the host
 * component's call so it can wire pending-staging UI around the dispatch.
 *
 * Behaviour:
 *  - First it POSTs to the local `/api/llm` route with the prompt and the
 *    agent's current Markdown projection (`agent.toMarkdown()`).
 *  - If the route returns 200 with a JSON `{ commands, rationale }` body,
 *    those commands are returned as agent-sourced commands.
 *  - If the route returns 501 (the documented "LLM bridge not configured"
 *    body — happens when `OPENAI_API_KEY` is unset), we fall back to the
 *    same `[AI] ` + `add-comment` recipe the editor used before W6 so the
 *    e2e specs and offline dev flow keep working.
 *  - Any other failure (network, parse) also falls back, so the demo never
 *    breaks because of a transient API hiccup. The error is returned via
 *    the `note` field for the caller to surface as a toast.
 */

import type { CommandLite } from "@officeai/core";
import type { DocxAgent } from "@officeai/docx";

export interface DispatchResult {
  commands: CommandLite[];
  rationale: string;
  /** When set, the helper used the offline fallback. UI may surface this. */
  note?: string;
}

const AGENT_ID = "web-llm-bridge";
const ENDPOINT = "/api/llm";

export async function dispatchToLlm(prompt: string, agent: DocxAgent): Promise<DispatchResult> {
  const trimmed = prompt.trim();
  if (!trimmed) return { commands: [], rationale: "Empty prompt; nothing to do." };

  const snapshotMarkdown = agent.toMarkdown();

  let response: Response;
  try {
    response = await fetch(ENDPOINT, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ prompt: trimmed, snapshotMarkdown }),
    });
  } catch (err) {
    return offlineFallback(trimmed, `Network error reaching ${ENDPOINT}: ${errorMessage(err)}`);
  }

  if (response.status === 501) {
    // Documented "no API key configured" branch — silently fall back.
    return offlineFallback(trimmed);
  }
  if (!response.ok) {
    return offlineFallback(trimmed, `LLM bridge returned HTTP ${response.status}`);
  }

  let parsed: { commands?: Array<{ type: string; payload: unknown }>; rationale?: string };
  try {
    parsed = (await response.json()) as typeof parsed;
  } catch (err) {
    return offlineFallback(trimmed, `LLM bridge returned non-JSON: ${errorMessage(err)}`);
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
 * The hard-coded recipe that used to live in `DocxEditor.tsx`. Two commands:
 * prefix the first paragraph with `[AI] ` and add a comment carrying the
 * user's prompt. Keeps the demo and e2e specs functional with no API key.
 */
function offlineFallback(prompt: string, note?: string): DispatchResult {
  const commands: CommandLite[] = [
    {
      type: "docx:insert-text",
      payload: { at: { paragraph: 0, run: 0, offset: 0 }, text: "[AI] " },
      source: "agent",
      agentId: AGENT_ID,
    },
    {
      type: "docx:add-comment",
      payload: {
        range: {
          start: { paragraph: 0, run: 0, offset: 0 },
          end: { paragraph: 0, run: 0, offset: 5 },
        },
        text: prompt,
        author: AGENT_ID,
        initials: "AI",
      },
      source: "agent",
      agentId: AGENT_ID,
    },
  ];
  return {
    commands,
    rationale: "Offline fallback: prefix `[AI] ` and attach the prompt as a comment.",
    ...(note ? { note } : {}),
  };
}

function errorMessage(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}
