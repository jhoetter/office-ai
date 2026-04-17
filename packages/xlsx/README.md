# @officeai/xlsx

The XLSX (Microsoft Excel) editor. Mirrors the architecture of
[`@officeai/docx`](../docx): a headless parser, in-memory model, command
bus integration, formula engine, headless agent, and a browser
virtualized-grid renderer.

Status: **active build** — see [`spec/xlsx/`](../../spec/xlsx/) for the
contract and [`docs/build-log/xlsx.md`](../../docs/build-log/xlsx.md) for
shipped work, deferred items, and the validation summary.

## Layout

```
src/
  parser/        OOXML → XlsxWorkbook (opaque-blob preservation for parts we don't model)
  serializer/    XlsxWorkbook → OOXML (byte-preserves untouched parts via partHashes)
  model/         XlsxWorkbook, Sheet, Cell, RangeRef, CellFormat, ConditionalFormatRule
  commands/      One handler per `xlsx:*` command, all dispatched through the shared CommandBus
  formula/       Lexer, Pratt parser, AST, evaluator, dependency graph, ~150 functions
  agent/         XlsxAgent — headless agent matching the shared `DocumentAgent` contract
  renderer/      Virtualized DOM grid, sheet tabs, formula bar (browser-only)
```

## Runtime dependencies

| Package           | License    | Why                                                                                 |
| ----------------- | ---------- | ----------------------------------------------------------------------------------- |
| `@officeai/core`  | MIT (this) | CommandBus, OOXML utils, MutationStore — format-agnostic                            |
| `xlsx` (SheetJS)  | Apache 2.0 | Battle-tested OOXML parse/serialize fallback for parts we don't model first-class   |
| `fast-xml-parser` | MIT        | Order- and namespace-faithful XML parser for the parts we own                       |
| `jszip`           | MIT        | OOXML zip container handling (re-exported from `@officeai/core`'s `OoxmlContainer`) |

## Quick start (post-build)

```ts
import { XlsxAgent } from "@officeai/xlsx/agent";
import { readFile, writeFile } from "node:fs/promises";

const agent = await XlsxAgent.fromBuffer(await readFile("data.xlsx"));
await agent.applyCommand({
  type: "xlsx:set-cell-value",
  payload: { sheet: "Sheet1", ref: "B2", value: 42 },
  source: "agent",
  agentId: "demo",
});
agent.getPendingMutations().forEach((m) => agent.approveMutation(m.id));
await writeFile("data-edited.xlsx", Buffer.from(await agent.exportFile()));
```
