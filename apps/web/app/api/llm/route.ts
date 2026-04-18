/**
 * LLM bridge route.
 *
 * POST /api/llm — body: {
 *     prompt: string;
 *     snapshotMarkdown: string;
 *     format?: "docx" | "pptx";
 *     selection?: { text: string; paragraph: number; runs?: number };
 *   }
 *   → { commands: Array<{ type: string; payload: unknown }>; rationale: string }
 *
 * `format` defaults to "docx" for backward compatibility.
 *
 * Behind `process.env.OPENAI_API_KEY`:
 *  - When unset: returns HTTP 501 with `{ error: "LLM bridge not configured" }`.
 *    The client helpers (`apps/web/app/lib/llm-client.ts` and
 *    `apps/web/app/lib/llm-client-pptx.ts`) interpret that response as a
 *    signal to use the offline fallback. The DOCX fallback (P2.5/W27) is
 *    honest — it attaches a comment carrying the user's prompt verbatim
 *    instead of pretending to "edit" the doc. The PPTX fallback dispatches
 *    a deterministic intent-parsed command set so the e2e specs (and the
 *    no-key dev experience) keep working.
 *  - When set: forwards the prompt to OpenAI's chat-completions endpoint via
 *    raw `fetch` (no SDK). The model is asked to respond with a strict JSON
 *    object listing the agent commands it wants the human to queue. The
 *    route does NOT execute the commands — it returns them so the human
 *    review surface (the staging area in the editor) can approve them per
 *    `prompt.md` "AI proposes, human approves".
 *
 * The system prompt enumerates the public `{docx|pptx}:*` command surface
 * so the model can't invent one. The DOCX list is sourced from the canonical
 * `DOCX_COMMAND_TYPES` export in `@officeai/docx` (P2.5/W26) so the route,
 * the client helper, and the agent itself stay in lockstep. Model defaults
 * to `gpt-4.1` and is overridable via `OPENAI_MODEL`.
 */

import { DOCX_COMMAND_TYPES, type DocxCommandType } from "@officeai/docx";

/**
 * The subset of the typed DOCX command surface the LLM is allowed to
 * propose. Mutating commands like `docx:accept-change` /
 * `docx:reject-change` are deliberately excluded — those represent
 * human review decisions on tracked changes, not authoring intents.
 */
const DOCX_ALLOWED_COMMAND_TYPES: ReadonlyArray<DocxCommandType> = DOCX_COMMAND_TYPES.filter(
  (t) => t !== "docx:accept-change" && t !== "docx:reject-change"
);

const PPTX_ALLOWED_COMMAND_TYPES = [
  "pptx:add-slide",
  "pptx:delete-slide",
  "pptx:duplicate-slide",
  "pptx:move-slide",
  "pptx:set-text",
  "pptx:set-position",
  "pptx:set-size",
  "pptx:format-text",
  "pptx:add-text-box",
] as const;

type Format = "docx" | "pptx";

interface LlmSelectionContext {
  /** Plain-text snippet of the user's current selection. */
  text: string;
  /** Body-paragraph index containing the selection's start. */
  paragraph: number;
  /** Number of runs the selection spans (informational; helps the model). */
  runs?: number;
}

interface LlmRequest {
  prompt: string;
  snapshotMarkdown: string;
  /** "docx" (default) or "pptx". Selects the system prompt + command allow-list. */
  format?: Format;
  /** Present when the user had a non-empty selection at dispatch time (DOCX only). */
  selection?: LlmSelectionContext;
}

interface LlmCommand {
  type: string;
  payload: unknown;
}

export interface LlmResponse {
  commands: LlmCommand[];
  rationale: string;
}

const DOCX_SYSTEM_PROMPT = [
  "You are an editing co-pilot for a Word-compatible document editor.",
  "You will receive (a) a human prompt, (b) the document as Markdown, and",
  "optionally (c) the user's current text selection (snippet + body-paragraph index).",
  "When a selection is provided, default to operating on it instead of guessing.",
  "Respond with a single JSON object — no prose, no code fences — of the form:",
  '  { "rationale": string, "commands": Array<{ "type": string, "payload": object }> }',
  "Each command must use one of these types:",
  DOCX_ALLOWED_COMMAND_TYPES.map((t) => `  - ${t}`).join("\n"),
  "Payloads follow the docx command spec (positions = { paragraph, run, offset }).",
  "Prefer small, targeted commands — the human will review and approve each one.",
  "If the prompt is unclear, return an empty `commands` array and explain why in `rationale`.",
].join("\n");

const PPTX_SYSTEM_PROMPT = [
  "You are an editing co-pilot for a PowerPoint-compatible slide editor.",
  "You will receive (a) a human prompt and (b) the presentation as Markdown.",
  "The Markdown lists each slide with its shapes; every text/picture/group/opaque shape",
  "carries a stable `id` like `nodeId-…` and an EMU bounding box `@ (x, y) cx×cy`.",
  "Address shapes by `(slideIndex, shapeId)` — slideIndex is the 0-based ordinal.",
  "All EMU values are integers; 914400 EMU = 1 inch. Slide size sits at the top of the doc.",
  "Respond with a single JSON object — no prose, no code fences — of the form:",
  '  { "rationale": string, "commands": Array<{ "type": string, "payload": object }> }',
  "Each command must use one of these types:",
  PPTX_ALLOWED_COMMAND_TYPES.map((t) => `  - ${t}`).join("\n"),
  "Payloads follow the pptx command spec. Notable payloads:",
  '  - pptx:add-slide → { at?: number; layoutPartPath?: string }',
  '  - pptx:delete-slide → { slideIndex: number }',
  '  - pptx:duplicate-slide → { slideIndex: number }',
  '  - pptx:move-slide → { from: number; to: number }',
  '  - pptx:set-text → { slideIndex: number; shapeId: string; text: string }  (use "\\n" for paragraph breaks)',
  '  - pptx:set-position → { slideIndex: number; shapeId: string; x: number; y: number }  (EMU)',
  '  - pptx:set-size → { slideIndex: number; shapeId: string; width: number; height: number }  (EMU, both > 0)',
  '  - pptx:format-text → { slideIndex: number; shapeId: string; range: { paragraph: number; start: number; end: number }; format: { bold?: boolean; italic?: boolean; underline?: boolean; color?: string; fontFamily?: string; fontSizeHundredths?: number } }',
  '  - pptx:add-text-box → { slideIndex: number; text: string; x: number; y: number; width: number; height: number }  (EMU)',
  "Prefer small, targeted commands — the human will review and approve each one.",
  "If the prompt is unclear, return an empty `commands` array and explain why in `rationale`.",
].join("\n");

function systemPromptFor(format: Format): string {
  switch (format) {
    case "docx":
      return DOCX_SYSTEM_PROMPT;
    case "pptx":
      return PPTX_SYSTEM_PROMPT;
  }
}

function allowedTypesFor(format: Format): ReadonlyArray<string> {
  switch (format) {
    case "docx":
      return DOCX_ALLOWED_COMMAND_TYPES;
    case "pptx":
      return PPTX_ALLOWED_COMMAND_TYPES;
  }
}

export async function POST(request: Request): Promise<Response> {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    return jsonResponse({ error: "LLM bridge not configured" }, 501);
  }

  let body: LlmRequest;
  try {
    body = (await request.json()) as LlmRequest;
  } catch {
    return jsonResponse({ error: "Invalid JSON body" }, 400);
  }
  if (typeof body?.prompt !== "string" || typeof body?.snapshotMarkdown !== "string") {
    return jsonResponse(
      {
        error:
          'Body must be { prompt: string, snapshotMarkdown: string, format?: "docx"|"pptx", selection?: { text, paragraph, runs? } }',
      },
      400
    );
  }

  const format: Format = body.format === "pptx" ? "pptx" : "docx";

  const model = process.env.OPENAI_MODEL || "gpt-4.1";
  const sections: string[] = [
    `# Human request\n\n${body.prompt.trim()}`,
    `# Document (Markdown projection)\n\n${body.snapshotMarkdown}`,
  ];
  // Selection context only applies to DOCX (where positions are paragraph-indexed).
  // For PPTX, shape addressing happens by stable shape IDs rendered in the markdown.
  if (format === "docx") {
    const sel = normalizeSelection(body.selection);
    if (sel) {
      sections.push(
        `# Current selection\n\nparagraph index: ${sel.paragraph}\n` +
          (sel.runs !== undefined ? `runs: ${sel.runs}\n` : "") +
          `\nselected text:\n\n${sel.text}`
      );
    }
  }
  const userMessage = sections.join("\n\n");

  let openaiResponse: Response;
  try {
    openaiResponse = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify({
        model,
        response_format: { type: "json_object" },
        temperature: 0.2,
        messages: [
          { role: "system", content: systemPromptFor(format) },
          { role: "user", content: userMessage },
        ],
      }),
    });
  } catch (err) {
    return jsonResponse(
      { error: "Failed to reach OpenAI", detail: err instanceof Error ? err.message : String(err) },
      502
    );
  }

  if (!openaiResponse.ok) {
    const text = await openaiResponse.text().catch(() => "");
    return jsonResponse(
      { error: "OpenAI returned an error", status: openaiResponse.status, detail: text },
      502
    );
  }

  let payload: unknown;
  try {
    payload = await openaiResponse.json();
  } catch (err) {
    return jsonResponse(
      { error: "OpenAI returned non-JSON", detail: err instanceof Error ? err.message : String(err) },
      502
    );
  }

  const content = extractAssistantContent(payload);
  if (!content) {
    return jsonResponse({ error: "OpenAI response missing message content" }, 502);
  }

  const parsed = parseAssistantContent(content, allowedTypesFor(format));
  if (!parsed) {
    return jsonResponse(
      { error: "Could not parse OpenAI response as the expected JSON shape", raw: content },
      502
    );
  }

  return jsonResponse(parsed satisfies LlmResponse, 200);
}

function jsonResponse(value: unknown, status: number): Response {
  return new Response(JSON.stringify(value), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

function extractAssistantContent(payload: unknown): string | null {
  if (!isObj(payload)) return null;
  const choices = (payload as { choices?: unknown }).choices;
  if (!Array.isArray(choices) || choices.length === 0) return null;
  const first = choices[0];
  if (!isObj(first)) return null;
  const message = (first as { message?: unknown }).message;
  if (!isObj(message)) return null;
  const content = (message as { content?: unknown }).content;
  return typeof content === "string" ? content : null;
}

function parseAssistantContent(
  content: string,
  allowed: ReadonlyArray<string>
): LlmResponse | null {
  let json: unknown;
  try {
    json = JSON.parse(content);
  } catch {
    return null;
  }
  if (!isObj(json)) return null;
  const rationaleRaw = (json as { rationale?: unknown }).rationale;
  const commandsRaw = (json as { commands?: unknown }).commands;
  const rationale = typeof rationaleRaw === "string" ? rationaleRaw : "";
  if (!Array.isArray(commandsRaw)) return { commands: [], rationale };
  const commands: LlmCommand[] = [];
  for (const c of commandsRaw) {
    if (!isObj(c)) continue;
    const type = (c as { type?: unknown }).type;
    const payload = (c as { payload?: unknown }).payload;
    if (typeof type !== "string") continue;
    if (!allowed.includes(type)) continue;
    commands.push({ type, payload: payload ?? {} });
  }
  return { commands, rationale };
}

function isObj(v: unknown): v is Record<string, unknown> {
  return typeof v === "object" && v !== null;
}

function normalizeSelection(value: unknown): LlmSelectionContext | null {
  if (!isObj(value)) return null;
  const text = (value as { text?: unknown }).text;
  const paragraph = (value as { paragraph?: unknown }).paragraph;
  const runs = (value as { runs?: unknown }).runs;
  if (typeof text !== "string" || text.length === 0) return null;
  if (typeof paragraph !== "number" || !Number.isFinite(paragraph)) return null;
  const out: LlmSelectionContext = { text, paragraph };
  if (typeof runs === "number" && Number.isFinite(runs)) out.runs = runs;
  return out;
}
