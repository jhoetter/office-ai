/**
 * Isomorphic LLM client helper.
 *
 * `dispatchToLlm(prompt, agent, ctx?)` returns an array of `Command`-shaped
 * objects ready to be queued via `agent.applyCommands(...)`. The helper
 * intentionally does NOT call `applyCommand` itself — that's the host
 * component's call so it can wire pending-staging UI around the dispatch.
 *
 * Behaviour:
 *  - First it POSTs to the local `/api/llm` route with the prompt, the
 *    agent's current Markdown projection (`agent.toMarkdown()`), and any
 *    selection context the caller provided so the model can target the
 *    user's actual selection (P2.5/W25).
 *  - If the route returns 200 with a JSON `{ commands, rationale }` body,
 *    those commands are returned as agent-sourced commands. Commands the
 *    server somehow let through are still re-validated against the
 *    canonical `DOCX_COMMAND_TYPES` allow-list (P2.5/W26) — defence in
 *    depth in case a future server tweak forgets the filter.
 *  - If the route returns 501 (the documented "LLM bridge not configured"
 *    body — happens when `OPENAI_API_KEY` is unset) we fall back to a
 *    deliberately-honest offline recipe (P2.5/W27): attach a single
 *    comment carrying the user's prompt verbatim. The fallback never
 *    edits document text any more — pretending to "edit" without an LLM
 *    was actively misleading. The caller surfaces a toast via `note`.
 *  - Any other failure (network, parse) also falls back, with the error
 *    explained in `note`.
 */

import type { CommandLite } from "@officeai/core";
import { DOCX_COMMAND_TYPES, type DocxAgent, type DocxCommandType } from "@officeai/docx";

export interface DispatchSelectionContext {
  /** Plain-text snippet of the user's current selection. */
  text: string;
  /** Body-paragraph index that contains the selection's start. */
  paragraph: number;
  /** Number of runs the selection spans (informational; helps the model). */
  runs?: number;
  /**
   * Full DOCX selection range. Used by the offline fallback so the
   * comment is anchored to what the user actually selected, instead of
   * an arbitrary first-paragraph slice.
   */
  range?: {
    start: { paragraph: number; run: number; offset: number };
    end: { paragraph: number; run: number; offset: number };
  };
}

export interface DispatchResult {
  commands: CommandLite[];
  rationale: string;
  /** When set, the helper used the offline fallback. UI may surface this. */
  note?: string;
}

const AGENT_ID = "web-llm-bridge";
const ENDPOINT = "/api/llm";
const ALLOWED_TYPES: ReadonlySet<DocxCommandType> = new Set(DOCX_COMMAND_TYPES);

export async function dispatchToLlm(
  prompt: string,
  agent: DocxAgent,
  ctx?: DispatchSelectionContext
): Promise<DispatchResult> {
  const trimmed = prompt.trim();
  if (!trimmed) return { commands: [], rationale: "Empty prompt; nothing to do." };

  const snapshotMarkdown = agent.toMarkdown();
  const selection = ctx
    ? {
        text: ctx.text,
        paragraph: ctx.paragraph,
        ...(ctx.runs !== undefined ? { runs: ctx.runs } : {}),
      }
    : undefined;

  let response: Response;
  try {
    response = await fetch(ENDPOINT, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ prompt: trimmed, snapshotMarkdown, selection }),
    });
  } catch (err) {
    return offlineFallback(trimmed, ctx, `Network error reaching ${ENDPOINT}: ${errorMessage(err)}`);
  }

  if (response.status === 501) {
    // Documented "no API key configured" branch — silently fall back.
    return offlineFallback(trimmed, ctx);
  }
  if (!response.ok) {
    return offlineFallback(trimmed, ctx, `LLM bridge returned HTTP ${response.status}`);
  }

  let parsed: { commands?: Array<{ type: string; payload: unknown }>; rationale?: string };
  try {
    parsed = (await response.json()) as typeof parsed;
  } catch (err) {
    return offlineFallback(trimmed, ctx, `LLM bridge returned non-JSON: ${errorMessage(err)}`);
  }

  const rawCommands = Array.isArray(parsed.commands) ? parsed.commands : [];
  const commands: CommandLite[] = [];
  for (const c of rawCommands) {
    if (!ALLOWED_TYPES.has(c.type as DocxCommandType)) continue;
    commands.push({ type: c.type, payload: c.payload, source: "agent", agentId: AGENT_ID });
  }
  const rationale = typeof parsed.rationale === "string" ? parsed.rationale : "";
  return { commands, rationale };
}

/**
 * Honest offline fallback (P2.5/W27).
 *
 * Without an API key there is no AI. We therefore do NOT touch the
 * document text. Instead we attach a single comment carrying the user's
 * prompt verbatim — anchored to the user's selection when one exists,
 * otherwise to the first paragraph as a best-effort placeholder. The
 * `rationale` makes the offline-ness explicit so the editor toast can
 * surface it.
 */
function offlineFallback(
  prompt: string,
  ctx: DispatchSelectionContext | undefined,
  note?: string
): DispatchResult {
  const range = ctx?.range ?? {
    start: { paragraph: 0, run: 0, offset: 0 },
    end: { paragraph: 0, run: 0, offset: 0 },
  };
  const commands: CommandLite[] = [
    {
      type: "docx:add-comment",
      payload: {
        range,
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
    rationale:
      "Offline fallback: no OPENAI_API_KEY is configured, so no LLM was called. " +
      "Attaching the prompt as a comment instead of pretending to edit the document.",
    ...(note ? { note } : {}),
  };
}

function errorMessage(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}
