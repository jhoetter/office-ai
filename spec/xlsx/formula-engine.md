# XLSX — Formula Engine

> The pure calculation core for `@officeai/xlsx`. Lexer → shunting-yard →
> AST → tree-walking evaluator, on top of a forward dependency graph
> with an interval-indexed range overlay. Sized for the ~150-function
> priority list in [`prompt.md`](../../prompt.md) lines 247–264 — no
> WASM, no plugin DI, no iterative calc, no async function path.
>
> This document is the implementation contract for `packages/xlsx/src/formula/`.
> Companion specs:
>
> - [`feature-scope.md`](feature-scope.md) — what ships in P0 vs P1.
> - [`edge-cases.md`](edge-cases.md) — formula edge cases `EC-F*`,
>   reference-handling edge cases `EC-R*`. Cited inline below.
> - [`analysis-univer-formula.md`](analysis-univer-formula.md) §12
>   "Suggested architecture" — the reference-repo pattern this engine
>   simplifies.
> - [`analysis.md`](analysis.md) — synthesis decisions (engine layout,
>   what we keep / differ on).
> - [`agent-commands.md`](agent-commands.md) — how
>   `xlsx:set-cell-formula` and friends drive the recalc loop.

---

## 1. Goal

Implement Excel-compatible formula evaluation for the **~150 priority
functions** enumerated in [`prompt.md`](../../prompt.md) lines 247–264.
The engine must:

1. Parse every well-formed Excel formula string the lexer/parser
   accepts (with the deferrals listed in §2).
2. Evaluate it against a pluggable cell-data context, producing an
   Excel-faithful value or error.
3. Track dependencies so that an edit to cell `X` triggers re-evaluation
   of every cell whose formula transitively reads `X` — and only those.
4. Surface circular references as a structured `#REF!` (per
   [`edge-cases.md`](edge-cases.md) `EC-F1`) without attempting
   iterative convergence.
5. Hit the perf budget in §17: recalc on a sheet with 10,000 dependent
   formulas after a single edit completes in **< 100 ms** on Apple
   Silicon (mirrors [`acceptance-criteria.md`](acceptance-criteria.md)
   `G5`).

What "Excel-compatible" means in practice: the table-driven conformance
suite in `__tests__/conformance/` matches Excel's published behaviour
on every documented input, and we record any known deltas (e.g.
`DATEDIF` undocumented quirks, locale-sensitive number parsing) in
`docs/build-log/xlsx.md` rather than silently diverging.

## 2. Non-goals (P0)

Explicit deferrals — the parser may accept these surface forms but the
evaluator returns `#NAME?` (or, for the structural ones, the parser
itself rejects them). Each is cross-referenced to the source decision.

| Deferral                                                       | Source                                                   |
| -------------------------------------------------------------- | -------------------------------------------------------- |
| Async / remote function path (`WEBSERVICE`, custom connectors) | [`edge-cases.md`](edge-cases.md) `EC-F5`, this doc §15.5 |
| Iterative calculation                                          | `EC-F1` — circular refs surface as `#REF!`               |
| `LAMBDA`, `LET`, `MAP`, `REDUCE`, `BYROW`, `BYCOL`             | `analysis-univer-formula.md` §11 item 7–8                |
| Structured (table) references `Table[[#All],[Col]]`            | `analysis-univer-formula.md` §11 item 10                 |
| External workbook refs `[Book.xlsx]Sheet1!A1`                  | `analysis-univer-formula.md` §11 item 11                 |
| 3D refs `Sheet1:Sheet3!A1` (preserved verbatim, not re-parsed) | `EC-R5`                                                  |
| Cube / database / engineering / web functions                  | `analysis-univer-formula.md` §11 items 5–6               |
| Implicit-intersection `@` operator                             | this doc §5 (parser rejects)                             |
| Single-space range-intersection operator                       | this doc §5 (parser rejects with hint)                   |
| Dynamic-array spill (`UNIQUE`/`SORT`/`FILTER`/`SEQUENCE`)      | "if time" — this doc §16.6                               |
| Date/Time + Finance categories                                 | P1 — registered as `#NAME?` until shipped (§16.4–§16.5)  |

The list above is **the** authoritative deferral set. Anything not
listed here ships in P0.

---

## 3. File layout

Exact tree under `packages/xlsx/src/formula/`. One concept per file;
no megafiles, no DI container.

```text
packages/xlsx/src/formula/
  tokens.ts              # token type constants (TokenType enum)
  errors.ts              # CellError type + helpers + interned singletons
  values.ts              # Value type (number|string|bool|error|range);
                         # arithmetic helpers with error short-circuit
  references.ts          # CellRef, RangeRef, A1↔R1C1, absolute-ref
                         # adjust on insert/delete row/column
  lexer.ts               # source string → flat token stream
  shunting-yard.ts       # token stream → RPN (postfix) sequence
  ast.ts                 # AST node types (discriminated union)
  parser.ts              # token stream → AST (Pratt-style; precedence /
                         # associativity table; consumes shunting-yard)
  evaluator.ts           # AST + EvalContext → Value
  dependency-graph.ts    # forward dep graph + range-interval index
  recalc.ts              # public API: recalc.run(dirty) → Map<cellRef, Value>
  function-registry.ts   # name → { fn, arity, returnType, volatile? }
  functions/
    math.ts              # SUM, AVERAGE, COUNT, ROUND, ABS, MOD, …
    logic.ts             # IF, IFS, AND, OR, NOT, XOR, IFERROR, …
    info.ts              # ISBLANK, ISNUMBER, ISTEXT, ISERROR, TYPE, N, …
    lookup.ts            # VLOOKUP, HLOOKUP, INDEX, MATCH, OFFSET, …
    text.ts              # CONCATENATE, LEFT, RIGHT, MID, LEN, TRIM, …
    date.ts              # P1 — registered as #NAME? in P0
    finance.ts           # P1 — registered as #NAME? in P0
    array.ts             # deferred ("if time")
  index.ts               # public surface (re-exports)
```

Tests live alongside in `__tests__/`:

- `__tests__/conformance/<category>/<function>.test.ts` — table-driven
  per-function correctness.
- `__tests__/lexer.test.ts`, `parser.test.ts`, `evaluator.test.ts`,
  `dependency-graph.test.ts`, `recalc.test.ts` — unit tests per layer.
- `__tests__/perf/` — perf-budget assertions (§17).
- `__tests__/round-trip.test.ts` — `parse(text) → serialize(ast) → text'`
  must equal the original modulo whitespace.

---

## 4. Token types

`tokens.ts` exports a single `TokenType` enum / const. Every value in
the lexer's output stream carries one of these tags. The catalogue is
deliberately **flat** — there is no nested tree at the token layer
(unlike Univer's `LexerNode` tree); structure is reconstructed by the
shunting-yard pass and the parser.

```typescript
export const TokenType = {
  NUMBER: "NUMBER", // 42, 3.14, 1e5, .5
  STRING: "STRING", // "hello" (doubled "" → literal ")
  BOOL: "BOOL", // TRUE, FALSE (case-insensitive)
  ERROR: "ERROR", // #VALUE!, #REF!, #NAME?, …  (literal error tokens)
  REF: "REF", // A1, $A$1, Sheet1!A1, 'My Sheet'!$B$2
  RANGE_REF: "RANGE_REF", // A1:B5, Sheet1!A:A, $A$1:$B$5
  NAME: "NAME", // defined-name candidate (resolved at parse time)
  FUNCTION: "FUNCTION", // identifier immediately followed by '('
  OPERATOR: "OPERATOR", // +, -, *, /, ^, &, =, <>, <, >, <=, >=
  LPAREN: "LPAREN", // (
  RPAREN: "RPAREN", // )
  LBRACE: "LBRACE", // {  (array literal start)
  RBRACE: "RBRACE", // }  (array literal end)
  COMMA: "COMMA", // ,  (function-arg separator; row separator inside {})
  COLON: "COLON", // :  (range operator)
  SEMICOLON: "SEMICOLON", // ;  (intersection — deferred; rejected by parser)
  PERCENT: "PERCENT", // %  (postfix)
  EOF: "EOF", // sentinel
} as const;
export type TokenType = (typeof TokenType)[keyof typeof TokenType];

export interface Token {
  readonly type: TokenType;
  readonly text: string; // the raw source slice
  readonly start: number; // 0-based char offset for diagnostics
  readonly end: number;
  /**
   * Pre-computed payload for value-bearing tokens. Filled by the lexer
   * when the conversion is cheap and side-effect-free (numeric parse,
   * boolean canonicalisation, error enum mapping). Refs/ranges are
   * parsed lazily by the parser via `references.ts` so the lexer stays
   * regex-light.
   */
  readonly value?: number | string | boolean | CellError;
}
```

Notes:

- `BOOL` matches the bare identifiers `TRUE`/`FALSE` (case-insensitive,
  but **not** the function calls `TRUE()`/`FALSE()` — those tokenize as
  `FUNCTION` and are evaluated normally).
- `ERROR` literal-tokens (e.g. `=IFERROR(#REF!, 0)`) are recognised by
  `errors.ts.parseErrorLiteral(text)` against the enum in §6.
- `REF` vs `RANGE_REF` is decided by the presence of `:` in the
  detected lexeme. The lexer first runs the broader `RANGE_REF` regex;
  if no colon is present, it downgrades to `REF`.
- `FUNCTION` is identifier-followed-by-`(`. The lexer emits both the
  `FUNCTION` token and the `LPAREN` separately so the parser can pop a
  matched-pair without a peek.
- `SEMICOLON` exists only so we can issue a precise "intersection
  operator not supported in P0" error (per §5 deferral) instead of a
  generic "unexpected token". It is **never** valid inside a P0
  formula.
- `EOF` is appended unconditionally; it has zero length and lets the
  parser write `match(TokenType.EOF)` instead of an end-of-input check
  on every loop.

---

## 5. Operator precedence

Single source of truth for the parser's precedence climbing. Lower
"level" number binds tighter (matches Univer's convention; see
`analysis-univer-formula.md` §1.2). The table is consulted by both the
shunting-yard pass (§11) and any future Pratt-style direct parse
implementation.

| Precedence | Operator(s)                       | Arity      | Associativity | Notes                                                                                                                |
| ---------- | --------------------------------- | ---------- | ------------- | -------------------------------------------------------------------------------------------------------------------- |
| 0          | `:` (range)                       | binary     | left          | `A1:B5` builds a `RANGE_REF` from two `REF`s. Tightest binding.                                                      |
| 1          | _space_ (intersection)            | binary     | left          | **DEFERRED**. Parser rejects `A1:B5 C1:C9` with `intersection-operator-not-supported` and a hint to use `INDEX`.     |
| 2          | `,` (union)                       | n-ary      | n/a           | Handled at the function-call level only (i.e. inside function `(...)`). A bare `(A1, B1)` outside a call is invalid. |
| 3          | `%` (postfix)                     | unary post | n/a           | `50%` → `0.5`. Wraps its operand in `PercentNode`.                                                                   |
| 4          | `^`                               | binary     | right         | Excel's exponent is right-associative: `2^3^2 = 2^9 = 512`.                                                          |
| 5          | `*`, `/`                          | binary     | left          |                                                                                                                      |
| 6          | `+`, `-` (binary)                 | binary     | left          |                                                                                                                      |
| 7          | `&`                               | binary     | left          | String concatenation. Coerces operands via §7.                                                                       |
| 8          | `=`, `<>`, `<`, `>`, `<=`, `>=`   | binary     | left          | Lowest-binding numeric/string comparisons.                                                                           |
| —          | unary `-`, unary `+` (prefix)     | unary pre  | right         | Higher than `^` per Excel quirk (`-2^2 == 4`, not `-4`). The parser wires this in directly when it sees a prefix.    |
| —          | `@` (implicit intersection)       | unary pre  | n/a           | **DEFERRED**. Parser rejects with `implicit-intersection-not-supported`.                                             |
| —          | `#` (spilled-range / `A1#`)       | unary post | n/a           | **DEFERRED** (depends on dynamic arrays).                                                                            |
| —          | `[…]` (structured-ref subscripts) | n/a        | n/a           | **DEFERRED**.                                                                                                        |

Excel's unary-minus precedence is the standard "gotcha". The table
encodes it correctly: `-2^2` parses as `(-(2))^2 = 4`. Tested in
`__tests__/parser.test.ts` against the published Excel behaviour.

---

## 6. Errors

`errors.ts` is the single source of truth for Excel error values. We
ship the **seven canonical Excel errors** as fully-supported runtime
values; the four "modern" errors (`#SPILL!`, `#CALC!`, `#CYCLE!`,
`#GETTING_DATA`) are recognised as literal tokens for round-trip but
the engine does not internally generate them in P0 (per
`analysis-univer-formula.md` §11 item 17).

```typescript
export const enum CellErrorKind {
  DIV0 = "#DIV/0!",
  NAME = "#NAME?",
  VALUE = "#VALUE!",
  NUM = "#NUM!",
  NA = "#N/A",
  REF = "#REF!",
  NULL = "#NULL!",
  // Recognised on import/round-trip only — engine never emits in P0:
  SPILL = "#SPILL!",
  CALC = "#CALC!",
  CYCLE = "#CYCLE!",
  GETTING_DATA = "#GETTING_DATA",
}

export interface CellError {
  readonly kind: CellErrorKind;
  /** Optional extra context used by the diff layer (cycle path, etc.). */
  readonly meta?: Readonly<Record<string, unknown>>;
}

/** Interned singleton per kind for cheap propagation. */
export const Errors: Readonly<Record<CellErrorKind, CellError>>;

export function parseErrorLiteral(text: string): CellError | undefined;
export function isError(v: Value): v is Extract<Value, { kind: "e" }>;
```

`Errors` is the interned cache: `Errors[CellErrorKind.VALUE]` returns
the same object every time. Keeps GC pressure low when a sheet has
many error cells (`analysis-univer-formula.md` §7.3).

`meta` is reserved for two known callers:

- `Errors.refWithCycle(["Sheet1!A1", "Sheet1!B1"])` — produces a
  `CellError` of kind `REF` whose `meta = { cycle: [...] }`. Surfaced
  in the mutation diff per `EC-F1`.
- `Errors.refWithDeletedTarget("Sheet2!A1")` — `meta = { deletedRef }`,
  surfaced per `EC-R2` / `EC-F4`.

The error mapping at handler boundaries is:

| Trigger                                                 | Error                                                    |
| ------------------------------------------------------- | -------------------------------------------------------- |
| Division by zero, `MOD(x, 0)`                           | `#DIV/0!`                                                |
| Unknown function name (post-parse)                      | `#NAME?` — `EC-F2`                                       |
| Unparseable formula                                     | `#NAME?` cached, original string preserved — `EC-F3`     |
| Wrong type to a function (after coercion)               | `#VALUE!`                                                |
| Numeric overflow / domain error (`SQRT(-1)`, `LOG(-1)`) | `#NUM!`                                                  |
| Lookup miss (`VLOOKUP`, `MATCH`, `XLOOKUP` exact)       | `#N/A`                                                   |
| Reference to deleted cell / sheet (after structural op) | `#REF!` — `EC-R2` / `EC-R3` / `EC-F4`                    |
| Cross-workbook reference                                | `#N/A` (preserved as opaque per `EC-O5`)                 |
| Circular reference                                      | `#REF!` with `meta.cycle` — `EC-F1`                      |
| Empty intersection (would-be `#NULL!`)                  | `#NULL!` — only emitted if/when intersection ships (P1+) |

---

## 7. Values, coercion, and arithmetic

`values.ts` defines the runtime value type used at every layer above
the lexer. We use a discriminated union (single-character `kind` tags
for compact debugger output and cheap dispatch) rather than the class
hierarchy in Univer — see `analysis-univer-formula.md` §11 item 1 for
why.

```typescript
import type { CellError } from "./errors";

export type Value =
  | { readonly kind: "n"; readonly v: number } //  number
  | { readonly kind: "s"; readonly v: string } //  string
  | { readonly kind: "b"; readonly v: boolean } //  boolean
  | { readonly kind: "e"; readonly v: CellError } //  error
  | { readonly kind: "r"; readonly v: Range2D }; //  range/array

export type Range2D = ReadonlyArray<ReadonlyArray<Value>>;

/** Sentinel for an empty cell. `blank.kind === "n"` with `v: 0`
 *  inside arithmetic; `blank.kind === "s"` with `v: ""` inside `&`;
 *  the §7.3 helpers do the right thing. */
export const Blank: Value;
```

### 7.1 Type coercion (Excel parity)

Coercion happens at **operator boundaries** and inside function
arguments where the function declares a coercion. Functions that
accept "any cell value" (`IF`, `ISNUMBER`, …) skip coercion.

| Source   | Target | Rule                                                                                                                            |
| -------- | ------ | ------------------------------------------------------------------------------------------------------------------------------- |
| `bool`   | number | `TRUE → 1`, `FALSE → 0`                                                                                                         |
| `bool`   | string | `TRUE → "TRUE"`, `FALSE → "FALSE"` (uppercase, locale-independent)                                                              |
| `string` | number | parse as Excel numeric: integer / decimal / leading sign / `e`-exponent / surrounding whitespace; on failure → `#VALUE!`        |
| `string` | bool   | exact case-insensitive match `"TRUE"` / `"FALSE"`; else `#VALUE!`                                                               |
| `number` | string | Excel's general format (no number-format applied here — that lives in the cell-display layer; the engine does plain `toString`) |
| `number` | bool   | `0 → FALSE`, anything else → `TRUE`                                                                                             |
| `error`  | _any_  | propagates unchanged (see §7.4)                                                                                                 |
| `blank`  | number | `0`                                                                                                                             |
| `blank`  | string | `""`                                                                                                                            |
| `blank`  | bool   | `FALSE`                                                                                                                         |
| `range`  | scalar | a 1×1 range collapses to its single cell; otherwise the function must accept ranges (e.g. `SUM`); else `#VALUE!`                |

Coercion helpers live in `functions/helpers/coerce.ts`:

```typescript
export function toNumber(v: Value): { kind: "n" | "e"; v: number | CellError };
export function toString(v: Value): { kind: "s" | "e"; v: string | CellError };
export function toBoolean(v: Value): { kind: "b" | "e"; v: boolean | CellError };
```

Each returns either the coerced value or a propagated `CellError` —
**never throws**. This is what makes function bodies short and
linear.

### 7.2 Comparison rules

Excel compares values across types deterministically. The order is:
`number < string < bool` (so `1 < "a" < FALSE`). Within a type:

- numbers: standard `<` / `>` (with `-0 === 0`).
- strings: case-**insensitive** lexicographic compare (`"a" === "A"`).
- booleans: `FALSE < TRUE`.
- errors: any comparison with an error returns the error (§7.4).

`functions/helpers/compare.ts` exports a single `compare(a, b)` that
returns `-1 | 0 | 1 | CellError`. The `=`/`<>` operators thin-wrap it.

### 7.3 Arithmetic helpers

```typescript
export function add(a: Value, b: Value): Value;
export function sub(a: Value, b: Value): Value;
export function mul(a: Value, b: Value): Value;
export function div(a: Value, b: Value): Value; //  → #DIV/0! on b===0
export function pow(a: Value, b: Value): Value;
export function neg(a: Value): Value;
export function pct(a: Value): Value; //  →  v / 100
export function concat(a: Value, b: Value): Value; //  &  operator
```

Each helper:

1. Short-circuits on `kind === "e"` (§7.4).
2. Coerces operands via the §7.1 rules.
3. Returns a `Value` — never throws.

Range operands: arithmetic on a range that contains a single cell
collapses to that cell. Arithmetic on a multi-cell range is a
`#VALUE!` (we don't ship implicit array calculation in P0; that's a
dynamic-array concern).

### 7.4 Error short-circuit

The single most important invariant in this layer. **Any arithmetic
helper, comparison, coercion, or function call that receives a
`Value` of `kind === "e"` returns that exact same error object.**

```typescript
export function add(a: Value, b: Value): Value {
  if (a.kind === "e") return a;
  if (b.kind === "e") return b;
  // …
}
```

This is the cleanest way to get correct error propagation across the
whole engine without try/catch in every function (Univer uses the
same pattern; see `analysis-univer-formula.md` §7.2). The interning
in §6 means error short-circuit is an object-identity preserve: a
chain of operations on `#DIV/0!` returns the same `#DIV/0!` object
the division produced.

`IFERROR` and `IFNA` are the only functions that **catch** errors;
they do so by inspecting `args[0].kind === "e"` directly before any
coercion would propagate it.

---

## 8. References

`references.ts` is the bridge between the lexer's textual ref tokens
and the rest of the engine. It owns A1 ↔ R1C1 conversion, absolute-ref
adjustment under insert/delete, and sheet-prefix handling.

```typescript
export const enum AbsoluteRef {
  NONE = 0, // A1
  ROW = 1, // A$1
  COLUMN = 2, // $A1
  ALL = 3, // $A$1
}

export interface CellRef {
  readonly sheet: string; // canonical sheet name (resolves quoting)
  readonly row: number; // 0-based internal
  readonly col: number; // 0-based internal
  readonly abs: AbsoluteRef;
}

export interface RangeRef {
  readonly sheet: string;
  readonly r0: number;
  readonly c0: number;
  readonly r1: number; // inclusive
  readonly c1: number; // inclusive
  readonly abs0: AbsoluteRef; // top-left absoluteness
  readonly abs1: AbsoluteRef; // bottom-right absoluteness
}

export function parseA1(text: string, defaultSheet: string): CellRef | undefined;
export function parseA1Range(text: string, defaultSheet: string): RangeRef | undefined;
export function serializeCellRef(ref: CellRef, anchor?: { sheet: string }): string;
export function serializeRangeRef(ref: RangeRef, anchor?: { sheet: string }): string;

/** Convert A1 ↔ R1C1 for display. Engine internals never use R1C1. */
export function a1ToR1C1(ref: CellRef, anchor: CellRef): string;
export function r1c1ToA1(text: string, anchor: CellRef): CellRef | undefined;

/** Insert/delete adjustment. Returns either an adjusted ref or a
 *  CellError of kind REF when the ref is invalidated. */
export function adjustForInsertRow(
  ref: CellRef | RangeRef,
  sheet: string,
  at: number, //  0-based row index where insertion happens
  count: number
): CellRef | RangeRef | CellError;

export function adjustForDeleteRow(
  ref: CellRef | RangeRef,
  sheet: string,
  at: number,
  count: number
): CellRef | RangeRef | CellError;

// adjustForInsertColumn / adjustForDeleteColumn: mirrors of the row
// versions. Sheet-rename is a separate pure rewrite handled in the
// command handler, not here, because it touches formula text.
```

### 8.1 Absolute-ref adjust on insert/delete

Per `EC-R1`: insertions shift refs at or below the insertion point by
`count` rows, **regardless of the absolute flag** (Excel parity —
absolute refs DO shift on structural ops; only relative refs shift on
copy/paste, which is a different code path entirely).

Per `EC-R2`: a delete that consumes the cell a ref points to returns
`Errors[CellErrorKind.REF]` with `meta.deletedRef` set. The command
handler (`xlsx:delete-row` — see [`agent-commands.md`](agent-commands.md))
collects these into the mutation diff.

Per `EC-R3`: defined names referencing deleted ranges are similarly
marked (the defined-name service calls into `references.ts` via the
same path).

Per `EC-R4`: sheet rename is **not** done in `references.ts` — it's a
formula-text rewrite owned by the rename-sheet handler. The handler
re-tokenizes every formula whose token stream contains a sheet
prefix matching the renamed sheet, swaps the prefix, and re-emits.
Quoted forms (`'My Sheet'!A1`) and bare forms are handled identically
because the lexer normalises them at parse time.

Per `EC-R5`: 3D references (`Sheet1:Sheet3!A1`) are detected at the
lexer level and emitted as a single opaque `REF` token whose `text`
is preserved verbatim. `references.ts` does not parse them; the rename
handler skips them with a single warning logged on first occurrence.
This is a documented deferral.

### 8.2 A1 ↔ R1C1

R1C1 is **display-only**. The internal representation is always
`{ row, col, abs }` — direction-agnostic numeric. `a1ToR1C1` /
`r1c1ToA1` exist solely so the renderer's "show formulas in R1C1"
toggle works without touching the engine. The lexer accepts A1 only;
the parser and AST never see R1C1.

### 8.3 Cross-sheet refs

Public API surface uses `Sheet1!A1` strings throughout (per
`analysis-agent-patterns.md` §9 — A1 with required sheet prefix).
Single-sheet refs in formula text default to the formula's anchor
sheet — that's what `defaultSheet` does in the parse helpers.

---

## 9. AST node types

`ast.ts` exports the discriminated union the parser builds and the
evaluator walks. Each node carries the full source span (`start`,
`end`) for diagnostics; the evaluator is otherwise type-driven (one
`switch` over `kind`, exhaustive per `typescript-exhaustive-switch.mdc`).

```typescript
import type { Value } from "./values";
import type { CellRef, RangeRef } from "./references";

export type AstNode =
  | LiteralNode
  | RefNode
  | RangeRefNode
  | NameNode
  | BinaryNode
  | UnaryNode
  | PercentNode
  | CallNode
  | ArrayLitNode;

export interface NodeBase {
  readonly start: number; //  0-based char offset into formula text
  readonly end: number;
}

export interface LiteralNode extends NodeBase {
  readonly kind: "lit";
  readonly value: Value; // numbers, strings, bools, error literals
}

export interface RefNode extends NodeBase {
  readonly kind: "ref";
  readonly ref: CellRef;
}

export interface RangeRefNode extends NodeBase {
  readonly kind: "range";
  readonly ref: RangeRef;
}

export interface NameNode extends NodeBase {
  readonly kind: "name";
  readonly name: string; //  defined-name text, lookup deferred to eval
}

export interface BinaryNode extends NodeBase {
  readonly kind: "binary";
  readonly op: BinaryOp;
  readonly left: AstNode;
  readonly right: AstNode;
}

export interface UnaryNode extends NodeBase {
  readonly kind: "unary";
  readonly op: UnaryOp; // "+" | "-"
  readonly operand: AstNode;
}

export interface PercentNode extends NodeBase {
  readonly kind: "pct";
  readonly operand: AstNode;
}

export interface CallNode extends NodeBase {
  readonly kind: "call";
  readonly name: string; //  upper-cased function name
  readonly args: ReadonlyArray<AstNode>;
}

export interface ArrayLitNode extends NodeBase {
  readonly kind: "array";
  readonly rows: ReadonlyArray<ReadonlyArray<AstNode>>;
  // {1,2;3,4}  →  rows = [[1,2], [3,4]]
}

export type BinaryOp = "+" | "-" | "*" | "/" | "^" | "&" | "=" | "<>" | "<" | ">" | "<=" | ">=";

export type UnaryOp = "+" | "-";
```

`Formula` is the evaluated wrapper the dependency graph and recalc
loop pass around:

```typescript
export interface Formula {
  readonly text: string; //  original source, anchor-relative
  readonly ast: AstNode;
  readonly anchor: CellRef; //  the cell this formula was parsed for
  readonly dependencies: ReadonlyArray<CellRef | RangeRef>;
  readonly volatile: boolean; //  any CallNode whose name is volatile
}

export class FormulaParseError extends Error {
  readonly span: { start: number; end: number };
  readonly hint?: string;
}
```

`dependencies` is collected in a single AST walk by the parser so the
dependency graph never has to re-walk (§13).

---

## 10. Lexer

`lexer.ts` exports a single function:

```typescript
export function lex(text: string): ReadonlyArray<Token>;
```

Algorithm — single forward scan, no backtracking:

1. Strip an optional leading `=` (formulas reach the lexer with or
   without it; the cell layer normalises but the lexer accepts both).
2. Loop over characters, holding state for:
   - inside-string-quote (between two `"`),
   - inside-sheet-quote (between two `'`),
   - inside-array-literal (counter for `{ … }`).
3. On entering one of those modes, accumulate raw characters into the
   current segment; exit on the matching closer (with `""` and `''`
   doubling rules).
4. Outside of quote/array modes, emit a token whenever the next
   character ends the current segment. The recognition order is:
   - `(`/`)`/`{`/`}`/`,`/`:`/`;`/`%` → single-char tokens.
   - `<>` / `<=` / `>=` (two-char compares; lookahead 1 char).
   - `=` / `<` / `>` / `+` / `-` / `*` / `/` / `^` / `&` → operator.
   - `#…` → either an error literal (lookahead via the regex below) or
     a deferred spilled-range marker (rejected at parse time).
   - digit / `.` followed by digit → number; we accept Excel's
     scientific notation `1.5e-3`.
   - `'` → enter sheet-quote mode; the closing `'` plus a following
     `!` indicates a sheet prefix; combine with the next ref/range
     lexeme to form a single `REF` / `RANGE_REF` token.
   - `"` → string literal.
   - identifier-or-ref start (`A-Z` / `a-z` / `_`):
     - run the **range regex** (§10.1) — if it matches, emit
       `RANGE_REF` (or downgrade to `REF` if no `:`).
     - else if the identifier is `TRUE` / `FALSE` (case-insensitive
       and not followed by `(`), emit `BOOL`.
     - else if followed by `(`, emit `FUNCTION` (then on the next
       iteration the `(` becomes `LPAREN`).
     - else emit `NAME` (defined-name candidate; resolved at parse).
5. Append `EOF`.

### 10.1 Reference recognition

References are recognised by a **small, audited regex catalogue**
rather than a hand-rolled DFA — same approach as Univer
(`analysis-univer-formula.md` §1.4). The full set, in match order:

```typescript
// Quoted sheet prefix: 'My Sheet'!  or  'It''s'!
const SHEET_QUOTED = /^'((?:[^']|'')*)'!/;
// Bare sheet prefix:  Sheet1!
const SHEET_BARE = /^([A-Za-z_][A-Za-z0-9_.]*)!/;

// Cell or range body (after optional sheet prefix):
//   $?A$?1                  cell with optional absolutes
//   $?A$?1 : $?B$?2         range
//   $?A : $?B               whole-column range
//   $?1 : $?2               whole-row range
const CELL = /^(\$?)([A-Z]{1,3})(\$?)(\d+)/i;
const RANGE = new RegExp(CELL.source + "(?::" + CELL.source + ")?", "i");
const COL_RANGE = /^(\$?)([A-Z]{1,3}):(\$?)([A-Z]{1,3})/i;
const ROW_RANGE = /^(\$?)(\d+):(\$?)(\d+)/i;

// 3D ref (preserved verbatim, never re-parsed — EC-R5):
const REF_3D = /^([A-Za-z_][A-Za-z0-9_.]*):([A-Za-z_][A-Za-z0-9_.]*)!/;
```

`#…` errors are matched against the kind enum directly:

```typescript
const ERROR_LITERALS = [
  "#DIV/0!",
  "#NAME?",
  "#VALUE!",
  "#NUM!",
  "#N/A",
  "#REF!",
  "#NULL!",
  "#SPILL!",
  "#CALC!",
  "#CYCLE!",
  "#GETTING_DATA",
] as const;
```

### 10.2 Lexer caching

A small `Map<string, Token[]>` keyed by formula text caches the most
recent N (default 1024) parses. On a fill-down of `=A2+1` across a
column the same formula text recurs once per cell — caching saves
hundreds of microseconds per row. The cache is invalidated on any
function-registry mutation (rare).

---

## 11. Shunting-yard

`shunting-yard.ts` exports:

```typescript
export function toPostfix(tokens: ReadonlyArray<Token>): ReadonlyArray<Token>;
```

A standard two-stack shunting-yard, operating on the **flat** token
stream from §10. The output is the same tokens reordered into postfix
(RPN); function-call boundaries are preserved by emitting a synthetic
`FUNCTION_END` token (with `argCount` recorded on the function token's
runtime metadata) so the parser can reconstruct n-ary calls without
ambiguity.

Algorithm sketch:

```text
out = []
ops = []
for t in tokens:
  if t.type in {NUMBER, STRING, BOOL, ERROR, REF, RANGE_REF, NAME}:
    out.push(t)
  elif t.type == FUNCTION:
    ops.push(t); t.argCount = 0
    expectArg = true
  elif t.type == LPAREN:
    ops.push(t)
  elif t.type == COMMA:
    while ops.top.type != LPAREN and not isFunction(ops.top):
      out.push(ops.pop())
    incrementTopFunctionArgCount()
  elif t.type == OPERATOR or t.type == COLON or t.type == PERCENT:
    while ops not empty and prec(ops.top) <= prec(t)
                       and not (t.assoc == "right" and prec(ops.top) == prec(t)):
      out.push(ops.pop())
    ops.push(t)
  elif t.type == RPAREN:
    while ops.top.type != LPAREN:
      out.push(ops.pop())
    ops.pop()  // discard LPAREN
    if ops.top is FUNCTION:
      f = ops.pop()
      out.push(f)  // function token now carries final argCount
  elif t.type == LBRACE:
    // array-literal scope; track row/col separators
    ...
while ops not empty: out.push(ops.pop())
return out
```

The `argCount` bookkeeping lets the parser turn a postfix
`FUNCTION` token into a `CallNode` without re-counting commas.

`SEMICOLON` outside an array literal short-circuits with a
`FormulaParseError("intersection-operator-not-supported")`.

---

## 12. Parser

`parser.ts` exports:

```typescript
export interface ParseOptions {
  readonly anchor: CellRef; //  drives default-sheet for unqualified refs
  readonly definedNames?: ReadonlyMap<string, RangeRef | CellRef>;
}

export function parse(text: string, opts: ParseOptions): Formula;
```

The parser is the thin RPN-to-AST conversion described in
`analysis-univer-formula.md` §2.1 (Appendix B.3). Because the
shunting-yard pass already linearises the operator tree, the parser
is a single stack walk:

```text
parse(text, opts):
  tokens = lex(text)
  rpn = toPostfix(tokens)
  stack: AstNode[] = []
  for t in rpn:
    if t.type in {NUMBER, STRING, BOOL, ERROR}:
      stack.push(LiteralNode(value))
    elif t.type == REF:
      ref = parseA1(t.text, opts.anchor.sheet)  // throws if invalid
      stack.push(RefNode(ref))
    elif t.type == RANGE_REF:
      stack.push(RangeRefNode(parseA1Range(t.text, opts.anchor.sheet)))
    elif t.type == NAME:
      stack.push(NameNode(t.text))
    elif t.type == OPERATOR (binary):
      r = stack.pop(); l = stack.pop()
      stack.push(BinaryNode(op, l, r))
    elif t.type == OPERATOR (unary):
      o = stack.pop()
      stack.push(UnaryNode(op, o))
    elif t.type == PERCENT:
      o = stack.pop()
      stack.push(PercentNode(o))
    elif t.type == FUNCTION:
      args = stack.popN(t.argCount)
      stack.push(CallNode(name.toUpperCase(), args))
    elif t.type == LBRACE:
      // collect array-literal body
      ...
  if stack.length != 1: throw FormulaParseError("malformed-expression")
  ast = stack[0]
  deps = collectDependencies(ast)
  volatile = walkContainsVolatileFunction(ast)
  return { text, ast, anchor: opts.anchor, dependencies: deps, volatile }
```

### 12.1 Why "Pratt-style" shows up in the file layout

The header comment (and the file-layout note in the brief) says
"Pratt-style". In practice, with the shunting-yard pass already
flattening the operator tree, the parser does **not** need a separate
recursive-descent / Pratt body for expressions. We keep `parser.ts`
small (the snippet above is ~80 LOC inflated to ~150 with diagnostics
and the array-literal sub-routine) and reserve the "Pratt-style"
naming for one specific case: **inside an array literal**, `LBRACE`
opens a sub-scope where commas separate columns and semicolons
separate rows. That sub-scope is parsed with a tiny precedence-climbing
helper (`parseArrayElement`) rather than re-running shunting-yard on
each cell — that helper is genuinely Pratt-shaped.

### 12.2 Defined-name resolution

`NameNode` lookups happen at parse time when `opts.definedNames` is
supplied: the name is resolved to a `RangeRef`/`CellRef` and the
resulting node is rewritten as `RefNode` / `RangeRefNode` directly.
Otherwise the `NameNode` survives into the AST and the evaluator
returns `#NAME?` (per `EC-F2`'s sibling case for unknown names).

### 12.3 Dependency collection

A single post-order walk after parse populates `formula.dependencies`:
every `RefNode` and `RangeRefNode` contributes one entry. `CallNode`
arguments are walked recursively; `INDIRECT(text)` and `OFFSET(base, …)`
contribute the **base** ref only (the dynamic resolution layer in
§13.4 handles the rest). The list is de-duplicated and stored in
parse-source order.

### 12.4 Volatile detection

`formula.volatile = true` if any `CallNode.name` in the AST is in the
volatile set: `RAND`, `RANDBETWEEN`, `NOW`, `TODAY`. The recalc loop
uses this flag to force-dirty the cell on every recalc (§14, §15.4).

### 12.5 Error handling

The parser **never throws on a recognised but unsupported construct**;
it throws `FormulaParseError` with a structured `kind` so callers can
choose to preserve-and-flag (per `EC-F3`) or reject (per the
`xlsx:set-cell-formula` precheck — see [`agent-commands.md`](agent-commands.md)
§2.2):

| `FormulaParseError.kind`                | Trigger                                           |
| --------------------------------------- | ------------------------------------------------- |
| `"empty-formula"`                       | Input is `""` or only `=`                         |
| `"unexpected-token"`                    | Generic "I have no idea what this is"             |
| `"unbalanced-paren"`                    | RPN stack underflow                               |
| `"intersection-operator-not-supported"` | Bare space between two ranges, `;` outside `{}`   |
| `"implicit-intersection-not-supported"` | `@` outside `[@col]` table syntax (also deferred) |
| `"structured-ref-not-supported"`        | `Table1[Col]`                                     |
| `"external-ref-not-supported"`          | `[Book.xlsx]Sheet1!A1`                            |
| `"3d-ref-not-supported"`                | `Sheet1:Sheet3!A1` outside a function call        |
| `"malformed-expression"`                | Final stack length ≠ 1                            |

---

## 13. Evaluator

`evaluator.ts` exports:

```typescript
export interface EvalContext {
  /** Cell-data accessor; returns Blank for empty cells. */
  getCell(ref: CellRef): Value;
  /** Range materialisation; returns a 2-D array of Values. */
  getRange(ref: RangeRef): Range2D;
  /** Function registry. */
  registry: FunctionRegistry;
  /** Defined-name resolver (post-parse fallback). */
  resolveName(name: string): RangeRef | CellRef | undefined;
  /** Anchor for relative-time semantics (volatile fns). */
  now: () => number;
  /** RNG for RAND / RANDBETWEEN; deterministic in tests. */
  random: () => number;
}

export function evaluate(node: AstNode, ctx: EvalContext): Value;
```

A single post-order tree walk:

```typescript
export function evaluate(node: AstNode, ctx: EvalContext): Value {
  switch (node.kind) {
    case "lit":
      return node.value;
    case "ref":
      return ctx.getCell(node.ref);
    case "range":
      return { kind: "r", v: ctx.getRange(node.ref) };
    case "name": {
      const resolved = ctx.resolveName(node.name);
      if (!resolved) return { kind: "e", v: Errors[CellErrorKind.NAME] };
      return "row" in resolved ? ctx.getCell(resolved) : { kind: "r", v: ctx.getRange(resolved) };
    }
    case "binary": {
      const l = evaluate(node.left, ctx);
      const r = evaluate(node.right, ctx);
      return BINARY_OPS[node.op](l, r);
    }
    case "unary": {
      const o = evaluate(node.operand, ctx);
      return node.op === "-" ? neg(o) : o; //  unary + is identity
    }
    case "pct":
      return pct(evaluate(node.operand, ctx));
    case "call": {
      const impl = ctx.registry.get(node.name);
      if (!impl) return { kind: "e", v: Errors[CellErrorKind.NAME] };
      if (!impl.arity.accepts(node.args.length)) return { kind: "e", v: Errors[CellErrorKind.NA] };
      const args = node.args.map((a) => evaluate(a, ctx));
      return impl.eval(args, ctx);
    }
    case "array":
      return { kind: "r", v: node.rows.map((row) => row.map((c) => evaluate(c, ctx))) };
  }
}
```

The switch is **exhaustive** per `typescript-exhaustive-switch.mdc`.

### 13.1 Error short-circuit at the AST boundary

The evaluator does **not** short-circuit AST walks on errors — it
defers to the §7.4 helpers inside operators and to the function impl
inside `CallNode`. This is intentional: `IF` and `IFERROR` need their
non-condition branches evaluated lazily / not-at-all, and the only way
to give them that control is to evaluate args inside the function impl.

In practice:

- `BinaryNode` / `UnaryNode` / `PercentNode` evaluate **both**
  operands first, then the helper short-circuits (cheap, predictable).
- `CallNode` evaluates **all** args before dispatch by default. Lazy
  functions (`IF`, `IFS`, `SWITCH`, `IFERROR`, `IFNA`) declare
  `lazyArgs: true` in their registry entry and the evaluator skips
  the eager arg eval, passing un-evaluated AST nodes plus the eval
  context. The function impl drives evaluation order itself.

### 13.2 Volatile evaluation

`RAND`, `RANDBETWEEN`, `NOW`, `TODAY` are normal functions in the
registry. The `volatile: true` flag on the registry entry is used by
the dependency graph (§14) to force the cell into the dirty set on
every recalc, not by the evaluator.

`RAND()` / `RANDBETWEEN()` use `ctx.random()` (deterministic in
tests). `NOW()` / `TODAY()` use `ctx.now()` so a single recalc has a
single coherent timestamp across the workbook.

### 13.3 Cycle handling

Cycle detection is structural — done in the dependency graph (§14) —
not in the evaluator. By the time the evaluator runs, the recalc
orchestrator (§15) has already partitioned cells into "topo-sortable"
and "in a cycle"; cells in a cycle are short-circuited to
`Errors.refWithCycle(...)` without the evaluator being invoked.

### 13.4 INDIRECT / OFFSET dynamic refs

`INDIRECT(text)` and `OFFSET(base, rowOffset, colOffset, [h], [w])`
return refs that depend on runtime values. We handle them as follows:

- At parse time, the function call contributes only the **literal
  args' refs** to `formula.dependencies`. For `INDIRECT("A1")` that's
  nothing; for `OFFSET(B2, …)` that's `B2`.
- At eval time, the function impl computes the resolved `CellRef` /
  `RangeRef` and calls `ctx.getCell` / `ctx.getRange` directly.
- Because the resolved refs are **not** registered as dependencies,
  edits to the resolved cell do **not** dirty the formula. This is a
  documented Excel-parity behaviour (Excel marks `INDIRECT`/`OFFSET`
  as volatile-equivalent for exactly this reason). We mirror the
  behaviour by adding `INDIRECT` and `OFFSET` to the `volatile` set
  in the registry.

---

## 14. Dependency graph

`dependency-graph.ts` exports:

```typescript
export type CellKey = string; //  `${sheet}!${row}:${col}`  (0-based internal)

export interface DepGraph {
  /** Add or replace the formula at `ref`; updates edges. */
  addCell(ref: CellRef, formula: Formula | null, value: Value | null): void;
  /** Remove a cell from the graph (delete edges). */
  removeCell(ref: CellRef): void;
  /** Mark the cell at `ref` (and all its dependents) dirty. */
  markDirty(ref: CellRef): void;
  /** Bulk dirty for an inserted/deleted range. */
  markRangeDirty(ref: RangeRef): void;
  /** Drain the dirty set; returns cells in safe topological order. */
  drainTopological(): { order: CellKey[]; cycles: CellKey[][] };
  /** Read the cached value (post-recalc). */
  getCachedValue(ref: CellRef): Value | undefined;
  /** Total cell count (for perf assertions). */
  size: number;
}

export function createDepGraph(): DepGraph;
```

### 14.1 Internal data structures

```typescript
interface FormulaCell {
  formula: Formula;
  cachedValue: Value;
}

interface ValueCell {
  cachedValue: Value;
}

interface State {
  cells: Map<CellKey, FormulaCell | ValueCell>;
  /** Forward edges: cell → cells that read it (its dependents). */
  forward: Map<CellKey, Set<CellKey>>;
  /** Reverse edges: cell → cells it reads (its precedents). */
  reverse: Map<CellKey, Set<CellKey>>;
  /** Range index: per sheet, an interval index over rows; values are
   *  the formula-cell keys whose dependencies overlap that row. */
  rangeIndex: Map<string, IntervalIndex<CellKey>>;
  /** Volatile cells (RAND/NOW/TODAY/INDIRECT/OFFSET): force-dirty
   *  every recalc. */
  volatile: Set<CellKey>;
  /** Dirty queue accumulated since the last drain. */
  dirty: Set<CellKey>;
}
```

The **range index** is the key data-structure decision. For a
formula `=SUM(A1:A100000)`, naïvely listing every cell key in that
range as a dependency is O(rows) per formula; instead we register
the **range** in an interval index keyed by `(sheet, rowMin..rowMax)`
and resolve hits at `markRangeDirty` time. See
`analysis-univer-formula.md` §10 item 4 (R-tree → interval index
simplification for our scope). v1 is a flat `Map<RangeKey, Set<CellKey>>`
plus a sequential scan for "ranges intersecting `(sheet, row, col)`";
we revisit if profiling shows >10ms per recalc on the §17 fixture.

### 14.2 `addCell` algorithm

```text
addCell(ref, formula, value):
  key = cellKey(ref)
  removeCell(ref)               // clear stale edges
  if formula:
    cells[key] = { formula, cachedValue: BLANK }
    for dep in formula.dependencies:
      if dep is CellRef:
        forward[depKey].add(key)
        reverse[key].add(depKey)
      else:
        rangeIndex[sheet].insert(dep, key)
        // we still need a reverse pointer for unwinding on remove:
        reverse[key].add(rangeKey(dep))
    if formula.volatile: volatile.add(key)
  else:
    cells[key] = { cachedValue: value ?? BLANK }
  markDirty(ref)
```

### 14.3 `markDirty` / dependent expansion

```text
markDirty(ref):
  key = cellKey(ref)
  if key in dirty: return
  dirty.add(key)
  for downstream in forward[key]:
    markDirty(parseKey(downstream))
  for cellKey in rangeIndex[ref.sheet].query(ref.row):
    markDirty(parseKey(cellKey))
```

This is a classic transitive-closure walk. The `dirty.add` guard
makes it linear in the size of the affected set.

### 14.4 `drainTopological` / cycle detection

```text
drainTopological():
  affected = collect every cell in dirty (∪ volatile)
  in_degree[k] = |reverse[k] ∩ affected|  // count only edges within set
  ready = [k for k in affected if in_degree[k] == 0]
  order = []
  while ready:
    k = ready.pop()
    order.push(k)
    for d in forward[k] ∩ affected:
      in_degree[d] -= 1
      if in_degree[d] == 0: ready.push(d)
  if order.length < affected.size:
    cycles = findCycles(affected - set(order))
  else:
    cycles = []
  dirty.clear()
  return { order, cycles }
```

Cycle finding uses Tarjan's SCC on the residual subgraph. Each SCC of
size ≥ 1 with at least one self-loop or back-edge is a cycle; the
recalc loop assigns `Errors.refWithCycle(...)` to every cell in the
SCC (per `EC-F1`).

---

## 15. Recalc orchestration

`recalc.ts` exports the public entry point invoked by every command
handler that touches cell values, formulas, or structure:

```typescript
export interface FormulaEngine {
  /** Parse a formula text relative to its anchor cell.
   *  Throws FormulaParseError on unrecoverable malformedness. */
  parse(text: string, anchor: CellRef): Formula;

  /** Add or replace a cell (formula OR value, never both). */
  addCell(ref: CellRef, formula: Formula | null, value: Value | null): void;

  /** Drop a cell (delete-row / delete-column / clear). */
  removeCell(ref: CellRef): void;

  /** Notify the engine that `ref` has changed. Marks dependents dirty. */
  onCellChanged(ref: CellRef): void;

  /** Drain the dirty set, evaluate in topological order, return
   *  affected cells' new values (and any cycles surfaced). */
  recalc(): RecalcResult;

  /** Read the cached value of a cell without triggering recalc. */
  getCachedValue(ref: CellRef): Value | undefined;

  /** Volatile-cell drain helper (used by xlsx_save and the renderer's
   *  "force-recalc" UI affordance). */
  recalcAll(): RecalcResult;
}

export interface RecalcResult {
  /** Cell key → new value, for every cell that was recalced. */
  values: Map<CellKey, Value>;
  /** SCCs of cells that participated in a cycle (each gets #REF!). */
  cycles: CellKey[][];
  /** Wall-clock time in ms (for the perf-budget assertion). */
  elapsedMs: number;
}

export function createFormulaEngine(): FormulaEngine;
```

### 15.1 Recalc loop

```text
recalc():
  start = now()
  // Volatile cells are always dirty:
  for k in volatile: dirty.add(k)
  { order, cycles } = depGraph.drainTopological()
  values = Map()
  // Cycles first — short-circuit before evaluator runs:
  for scc in cycles:
    cycleErr = Errors.refWithCycle(scc.map(refForKey))
    for k in scc:
      cells[k].cachedValue = { kind: "e", v: cycleErr }
      values.set(k, cells[k].cachedValue)
  // Topo-sorted evaluation:
  for k in order:
    cell = cells[k]
    if cell is FormulaCell:
      v = evaluate(cell.formula.ast, evalCtx(refForKey(k)))
      cell.cachedValue = v
      values.set(k, v)
  return { values, cycles, elapsedMs: now() - start }
```

### 15.2 Synchronous, single-pass

There is **no** async path. Every function in the registry returns a
`Value` synchronously. A single `recalc()` call drains the entire
dirty set and returns. This is the §2 deferral (no async functions, no
iterative calc) made concrete.

### 15.3 Affected-set scope

Per `EC-F1`: the affected set is the transitive closure from each
dirty cell along forward edges. We do **not** re-evaluate
non-dependent cells. This is what keeps the §17 perf budget tractable
on a 10k-formula workbook with a single edit.

### 15.4 Volatile cells

Per `EC-F5`: every recalc adds the volatile set to the dirty queue
unconditionally. On a 50,000-row sheet with `=RAND()` in every cell,
this costs O(50k) plus the topological-sort overhead — we explicitly
do **not** assert the §17 perf budget on this case (see `EC-F5`'s
"agent should batch" note).

### 15.5 Async / iterative calc — explicitly NOT supported

Re-stating the §2 deferral in a place callers will look:

- The engine has no `Promise`-typed evaluator. There is no
  `awaitCalculationResultApplied()` analogue.
- Circular references are detected once and surfaced as `#REF!`. There
  is no "iterate to convergence" knob (no `maxIterations`, no
  `maxChange`, no enable-iterative-calc setting).
- Per `EC-F1`: the cycle's SCC is recorded in `RecalcResult.cycles`
  and the mutation diff includes `{ kind: "circular", cycle: [...] }`.
  Agents that want convergence must implement it externally by
  iterating `xlsx:set-cell-value` themselves and tracking convergence.

---

## 16. Function-list breakdown

The full inventory below mirrors [`prompt.md`](../../prompt.md)
lines 247–264 verbatim. P0 / P1 / deferred markers reflect
[`feature-scope.md`](feature-scope.md) §"P0 vs P1 within in-scope".

P1 categories are **registered** in the function-registry as stubs
that return `Errors[CellErrorKind.NAME]` — this means a workbook
containing `=PMT(0.05/12, 60, -25000)` parses, round-trips, and
displays `#NAME?` in the cached-value column without rejecting the
file at load time (per `EC-F2` and `EC-S5`). When the P1 category
ships, swapping the stub for the real impl is a one-line registry
diff — no parser or AST change required.

### 16.1 Math / Statistics (P0)

| Function      | Min args | Max args | Returns | Category        |
| ------------- | -------- | -------- | ------- | --------------- |
| `SUM`         | 1        | 255      | number  | math            |
| `AVERAGE`     | 1        | 255      | number  | math            |
| `COUNT`       | 1        | 255      | number  | math            |
| `COUNTA`      | 1        | 255      | number  | math            |
| `COUNTBLANK`  | 1        | 1        | number  | math            |
| `MIN`         | 1        | 255      | number  | math            |
| `MAX`         | 1        | 255      | number  | math            |
| `SUMIF`       | 2        | 3        | number  | math            |
| `SUMIFS`      | 3        | 255      | number  | math            |
| `COUNTIF`     | 2        | 2        | number  | math            |
| `COUNTIFS`    | 2        | 254      | number  | math            |
| `AVERAGEIF`   | 2        | 3        | number  | math            |
| `AVERAGEIFS`  | 3        | 255      | number  | math            |
| `ROUND`       | 2        | 2        | number  | math            |
| `ROUNDUP`     | 2        | 2        | number  | math            |
| `ROUNDDOWN`   | 2        | 2        | number  | math            |
| `INT`         | 1        | 1        | number  | math            |
| `ABS`         | 1        | 1        | number  | math            |
| `MOD`         | 2        | 2        | number  | math            |
| `POWER`       | 2        | 2        | number  | math            |
| `SQRT`        | 1        | 1        | number  | math            |
| `CEILING`     | 1        | 2        | number  | math            |
| `FLOOR`       | 1        | 2        | number  | math            |
| `RAND`        | 0        | 0        | number  | math (volatile) |
| `RANDBETWEEN` | 2        | 2        | number  | math (volatile) |
| `LARGE`       | 2        | 2        | number  | math            |
| `SMALL`       | 2        | 2        | number  | math            |
| `RANK`        | 2        | 3        | number  | math            |
| `MEDIAN`      | 1        | 255      | number  | math            |
| `STDEV`       | 1        | 255      | number  | math            |
| `VAR`         | 1        | 255      | number  | math            |
| `PRODUCT`     | 1        | 255      | number  | math            |
| `SUMPRODUCT`  | 1        | 255      | number  | math            |

Key edge cases (see [`edge-cases.md`](edge-cases.md) for the broader
catalogue):

- `SUM(A1:A10)` where some cells contain text returns **only the
  numeric sum**; text cells are silently skipped (Excel parity).
  `SUM(A1, A2, A3)` where `A1 = "5"` (literal text), in contrast,
  coerces the text to number per §7.1 — this distinction matters and
  is tested.
- `MOD(x, 0)` returns `#DIV/0!` — explicit per Excel even though `MOD`
  isn't division.
- `SUMIF` / `COUNTIF` / `AVERAGEIF` accept Excel-style criteria
  expressions: `">5"`, `"<>0"`, `"apple*"`, `"=foo"`. The criteria
  parser lives in `functions/helpers/criteria.ts` and is shared by
  every `*IF` / `*IFS` family member.
- `CEILING(x, 1)` rounds away from zero per the modern Excel
  semantics; the legacy "round-toward-positive" is **not** supported
  (`CEILING.MATH` would be the modern alias and is not in the
  priority list).
- `RAND` / `RANDBETWEEN` use `ctx.random()` — see §13.2 for
  determinism in tests.

### 16.2 Logic (P0)

| Function  | Min | Max | Notes                                             |
| --------- | --- | --- | ------------------------------------------------- |
| `IF`      | 2   | 3   | lazy args (skip false-branch on truthy cond)      |
| `IFS`     | 2   | 254 | lazy args; pairs of (cond, value); `#N/A` on miss |
| `AND`     | 1   | 255 | short-circuit on first FALSE                      |
| `OR`      | 1   | 255 | short-circuit on first TRUE                       |
| `NOT`     | 1   | 1   |                                                   |
| `XOR`     | 1   | 255 | true iff odd number of TRUE args                  |
| `IFERROR` | 2   | 2   | lazy args; catches every CellError except `#N/A`  |
| `IFNA`    | 2   | 2   | lazy args; catches `#N/A` only                    |
| `SWITCH`  | 3   | 254 | lazy args                                         |
| `TRUE`    | 0   | 0   | constant                                          |
| `FALSE`   | 0   | 0   | constant                                          |

Edge cases:

- `IF(TRUE, "yes")` → `"yes"` (third arg defaults to `FALSE`, **not**
  to blank — Excel parity).
- `IFERROR(#N/A, 0)` → `0` (Excel: `IFERROR` catches `#N/A` too;
  `IFNA` is the more specific subset).
- `IFS` with no matching condition → `#N/A`.
- `SWITCH(expr, c1, v1, c2, v2, …, default)` — the trailing default
  is optional; if omitted and no match, return `#N/A`.

### 16.3 Lookup (P0)

| Function   | Min | Max | Notes                                                    |
| ---------- | --- | --- | -------------------------------------------------------- |
| `VLOOKUP`  | 3   | 4   | 4th arg defaults to `TRUE` (approximate); see edge cases |
| `HLOOKUP`  | 3   | 4   | mirror of VLOOKUP for rows                               |
| `INDEX`    | 2   | 4   | both array and reference forms                           |
| `MATCH`    | 2   | 3   | match-type defaults to 1 (largest ≤ lookup; sorted)      |
| `XLOOKUP`  | 3   | 6   | "if time" — included in P0 priority list                 |
| `CHOOSE`   | 2   | 254 |                                                          |
| `OFFSET`   | 3   | 5   | volatile; see §13.4                                      |
| `INDIRECT` | 1   | 2   | volatile; see §13.4                                      |
| `ROW`      | 0   | 1   | omit arg → caller's row                                  |
| `ROWS`     | 1   | 1   | takes a range; returns its row count                     |
| `COLUMN`   | 0   | 1   | omit arg → caller's column                               |
| `COLUMNS`  | 1   | 1   |                                                          |

Edge cases:

- `VLOOKUP(key, table, 2, FALSE)` (exact): linear scan, returns
  `#N/A` on miss.
- `VLOOKUP(key, table, 2, TRUE)` (approximate): requires the key
  column be **sorted ascending**; returns the largest value ≤ key.
  Out-of-order data produces incorrect results — Excel parity, do
  not validate input order.
- `INDEX(range, 0, n)` returns the entire `n`th column as an array;
  `INDEX(range, n, 0)` returns the entire `n`th row. Used heavily
  with `MATCH` for "real" lookups.
- `OFFSET(A1, -1, 0)` produces an out-of-bounds ref → `#REF!`.
- `INDIRECT("A" & ROW())` rebuilds a ref from a string at eval time.
  Per §13.4 it is registered as volatile so the caller dirties on
  every recalc.

### 16.4 Text (P0)

| Function      | Min | Max | Notes                                                |
| ------------- | --- | --- | ---------------------------------------------------- |
| `CONCATENATE` | 1   | 255 | scalar args only; ranges must be flattened by caller |
| `CONCAT`      | 1   | 255 | accepts ranges; flattens row-major                   |
| `TEXTJOIN`    | 3   | 254 | `(delim, ignoreEmpty, …text)`                        |
| `LEFT`        | 1   | 2   | numChars defaults to 1                               |
| `RIGHT`       | 1   | 2   | numChars defaults to 1                               |
| `MID`         | 3   | 3   |                                                      |
| `LEN`         | 1   | 1   |                                                      |
| `TRIM`        | 1   | 1   | collapses internal runs of spaces                    |
| `UPPER`       | 1   | 1   |                                                      |
| `LOWER`       | 1   | 1   |                                                      |
| `PROPER`      | 1   | 1   |                                                      |
| `FIND`        | 2   | 3   | case-sensitive; no wildcards; `#VALUE!` on miss      |
| `SEARCH`      | 2   | 3   | case-insensitive; supports `?` and `*` wildcards     |
| `SUBSTITUTE`  | 3   | 4   | optional Nth-occurrence arg                          |
| `REPLACE`     | 4   | 4   | by position+length                                   |
| `REPT`        | 2   | 2   | `#VALUE!` if result > 32,767 chars                   |
| `TEXT`        | 2   | 2   | format string per number-format engine (see §16.7)   |
| `VALUE`       | 1   | 1   | parse text → number; `#VALUE!` on failure            |
| `NUMBERVALUE` | 1   | 3   | locale-aware decimal/group separators                |
| `CHAR`        | 1   | 1   | `1..255`; `#VALUE!` outside                          |
| `CODE`        | 1   | 1   | first char's code; `#VALUE!` on empty                |
| `EXACT`       | 2   | 2   | case-sensitive equality                              |
| `T`           | 1   | 1   | returns text-or-empty (passes text, blanks numbers)  |

Edge cases:

- `LEFT("abc", -1)` → `#VALUE!`.
- `RIGHT("abc", 100)` → `"abc"` (clamps to length, not error).
- `FIND("a", "banana", 0)` → `#VALUE!` (start position must be ≥ 1).
- `SEARCH("a*c", "abc")` → `1` (wildcard match).
- `REPT("a", 32768)` → `#VALUE!` (max length).
- `TEXT(1234.5, "$#,##0.00")` → `"$1,234.50"`. `TEXT` is the bridge
  to the number-format engine; see §16.7.

### 16.5 Date / Time (P1 — registered as `#NAME?` in P0)

| Function      | Min | Max | Status                                    |
| ------------- | --- | --- | ----------------------------------------- |
| `TODAY`       | 0   | 0   | P1                                        |
| `NOW`         | 0   | 0   | P1                                        |
| `DATE`        | 3   | 3   | P1                                        |
| `TIME`        | 3   | 3   | P1                                        |
| `YEAR`        | 1   | 1   | P1                                        |
| `MONTH`       | 1   | 1   | P1                                        |
| `DAY`         | 1   | 1   | P1                                        |
| `HOUR`        | 1   | 1   | P1                                        |
| `MINUTE`      | 1   | 1   | P1                                        |
| `SECOND`      | 1   | 1   | P1                                        |
| `WEEKDAY`     | 1   | 2   | P1                                        |
| `WEEKNUM`     | 1   | 2   | P1                                        |
| `EOMONTH`     | 2   | 2   | P1                                        |
| `EDATE`       | 2   | 2   | P1                                        |
| `DATEDIF`     | 3   | 3   | P1 — Excel quirks documented when shipped |
| `NETWORKDAYS` | 2   | 3   | P1                                        |
| `WORKDAY`     | 2   | 3   | P1                                        |
| `DATEVALUE`   | 1   | 1   | P1                                        |
| `TIMEVALUE`   | 1   | 1   | P1                                        |

Each is registered with `eval = () => ({ kind: "e", v: Errors[CellErrorKind.NAME] })`
in P0 so the parser succeeds and the cached value renders `#NAME?`
without a flag.

When this category ships, the Excel epoch (1900-01-01 with the
"1900 is a leap year" bug) becomes the reference. Pick correctness
over compatibility (`analysis-univer-formula.md` §13). Document the
choice in `docs/build-log/xlsx.md`.

### 16.6 Finance (P1 — registered as `#NAME?` in P0)

| Function | Min | Max | Status |
| -------- | --- | --- | ------ |
| `PMT`    | 3   | 5   | P1     |
| `PV`     | 3   | 5   | P1     |
| `FV`     | 3   | 5   | P1     |
| `RATE`   | 3   | 6   | P1     |
| `NPER`   | 3   | 5   | P1     |
| `NPV`    | 2   | 255 | P1     |
| `IRR`    | 1   | 2   | P1     |
| `SLN`    | 3   | 3   | P1     |

Same registration story as §16.5.

### 16.7 Info (P0)

| Function   | Min | Max |
| ---------- | --- | --- |
| `ISBLANK`  | 1   | 1   |
| `ISNUMBER` | 1   | 1   |
| `ISTEXT`   | 1   | 1   |
| `ISERROR`  | 1   | 1   |
| `ISNA`     | 1   | 1   |
| `ISODD`    | 1   | 1   |
| `ISEVEN`   | 1   | 1   |
| `TYPE`     | 1   | 1   |
| `N`        | 1   | 1   |
| `NA`       | 0   | 0   |

Edge cases:

- `ISBLANK` is `TRUE` only for the **`Blank` sentinel**. A cell
  containing `""` (empty string) is **not** blank — Excel parity.
- `ISERROR` catches every kind of error including `#N/A`; `ISNA`
  catches only `#N/A`.
- `TYPE` returns 1 (number), 2 (text), 4 (logical), 16 (error), 64
  (array). The empty/blank case returns 1 (consistent with `N(blank)
= 0` being numeric).
- `NA()` always returns `Errors[CellErrorKind.NA]`.

### 16.8 Array (deferred — "if time")

| Function   | Status                           |
| ---------- | -------------------------------- |
| `UNIQUE`   | deferred — needs spill semantics |
| `SORT`     | deferred — needs spill semantics |
| `FILTER`   | deferred — needs spill semantics |
| `SEQUENCE` | deferred — needs spill semantics |

Dynamic-array spill requires changes to the cell model (anchor cell
holds the formula, neighbouring cells hold references back), the
serializer (`<f t="array" ref="A1:C3">` per `EC-F7`), and the
recalc loop (a second pass after normal recalc to allocate spill
ranges). None of those ship in P0. The functions are **not**
registered as stubs because they would need to spill values into
sibling cells to be useful, and the P0 cell model has no concept of
that. A formula `=UNIQUE(A1:A10)` parses (the parser doesn't know
which functions are array functions) and evaluates to `#NAME?` at
runtime via the standard "unknown function" path.

---

## 17. Performance budget

Single hard target carried into [`acceptance-criteria.md`](acceptance-criteria.md)
`G5`:

> Recalc on a sheet with **10,000 dependent formulas** after a single
> edit: **< 100 ms** on Apple Silicon.

The fixture: column `A` holds 10,000 numeric values; column `B` holds
`=A1+1`, `=A2+1`, …, `=A10000+1`; column `C` holds `=B1+B2`,
`=B2+B3`, …, `=B9999+B10000`. Editing `A1` dirties `B1`, which
dirties `C1`. Editing `A5000` dirties `B5000`, which dirties `C4999`
and `C5000`. We assert recalc completes in <100ms after **any single
edit**.

We explicitly do **not** assert the budget on:

- `EC-F5`-style 50,000-row volatile sheets.
- Recalc-all (`recalcAll()` after a workbook open) — a separate
  budget covers parse+evaluate-all in `G5` (50k × 10 cell parse <
  1.5s, serialize < 1.0s).
- Lookup-heavy sheets where every formula does a `VLOOKUP` over a
  10k-row table (that's a `O(N²)` workload by construction).

`__tests__/perf/recalc.test.ts` uses Vitest's `bench` plus a hard
assertion (`expect(elapsedMs).toBeLessThan(100)`) so a regression
fails CI, not just the bench output.

---

## 18. Interaction with the command bus

The engine is **passive**. It does not own state or react to
mutations on its own; every command handler that touches cells calls
into `FormulaEngine` explicitly. The contract:

| Command                 | Engine call sequence                                                |
| ----------------------- | ------------------------------------------------------------------- |
| `xlsx:set-cell-value`   | `addCell(ref, null, value)` → `recalc()`                            |
| `xlsx:set-cell-formula` | `parse(text, ref)` → `addCell(ref, formula, null)` → `recalc()`     |
| `xlsx:set-range-values` | `addCell(...)` per cell → single `recalc()` at end                  |
| `xlsx:insert-row`       | `adjustForInsertRow` over every formula → `recalc()`                |
| `xlsx:insert-column`    | `adjustForInsertColumn` over every formula → `recalc()`             |
| `xlsx:delete-row`       | `adjustForDeleteRow` (collect `#REF!` casualties) → `recalc()`      |
| `xlsx:delete-column`    | `adjustForDeleteColumn` (collect `#REF!` casualties) → `recalc()`   |
| `xlsx:rename-sheet`     | rewrite formula text per `EC-R4` → re-parse affected → `recalc()`   |
| `xlsx:add-sheet`        | no engine call (no formulas reference a non-existent sheet by name) |
| `xlsx:merge-cells`      | no engine call (merges affect display, not values)                  |
| `xlsx:unmerge-cells`    | no engine call                                                      |
| `xlsx:set-cell-format`  | no engine call                                                      |
| `xlsx:add-comment`      | no engine call                                                      |

The single `recalc()` call per command is the §17 perf hot path.

The `RecalcResult.cycles` and the `Errors.refWith*` payloads bubble
into the mutation diff via the handler, not via the engine — this
keeps `dependency-graph.ts` purely structural.

---

## 19. Round-trip discipline

Per [`feature-scope.md`](feature-scope.md) "Formulas":

- Formulas we **can** evaluate: `formula.text` is preserved verbatim
  (we do not re-emit a canonicalised form); the cached value is
  recomputed and written to `<v>` on serialize.
- Formulas we **cannot** parse (per `EC-F3`): the original string
  survives in the model's `formula.raw`; cached value is `#NAME?`;
  the cell is flagged so the renderer can show a tooltip.
- Formulas referencing unknown functions (per `EC-F2`): parse
  succeeds (`NAME` registered as stub), evaluator returns `#NAME?`,
  the original text round-trips.
- Shared formulas (per `EC-F6`): expanded to per-cell formulas on
  import; on serialize the cell layer optionally re-shares
  contiguous identical-shape groups.
- Array formulas (per `EC-F7`): preserved as opaque on the anchor
  cell; not evaluated in P0.
- 3D refs (per `EC-R5`): preserved verbatim; the lexer emits a
  single `REF` token whose `text` carries the full `Sheet1:Sheet3!A1`
  string, untouched by `references.ts`.

A property test in `__tests__/round-trip.test.ts` asserts: for every
parseable formula in the conformance suite, `parse(text) →
serialize(ast) → text'` produces a string that re-parses to an
isomorphic AST (whitespace differences are tolerated; absolute-ref
markers and shared-formula-id metadata are not).

---

## 20. Open questions (carried to `docs/build-log/xlsx.md`)

These do not block P0 closure but should be revisited before P1:

1. **Persistence of cached value on import.** Excel writes `<v>` for
   every formula cell; per `EC-S5` we recompute on import for cells
   missing `<v>`. Should we **trust** the cached `<v>` for cells
   that have one, or always recompute? Recommendation: trust on
   import, mark every formula cell dirty in a background pass, and
   replace the `<v>` if recompute disagrees by more than ε. Filed as
   an open question in `analysis-univer-formula.md` §13.
2. **Range-index data structure.** v1 is a `Map<RangeKey, Set<CellKey>>`
   plus sequential scan. If profiling on the §17 fixture exceeds
   10ms in the index alone, switch to `rbush` (MIT-licensed). Tracked.
3. **Number-format aware functions.** `TEXT` and `NUMBERVALUE` need
   the number-format engine (lives in `packages/xlsx/src/format/`,
   shared with the cell-display layer). API surface: `parseFormat`,
   `applyFormat` — exposed via `functions/helpers/format.ts`.
4. **Sheet-rename invalidation signal.** The recalc loop currently
   accepts only `dirty` ranges; sheet rename happens by re-parsing
   affected formulas and calling `addCell`. If the rename count is
   large (>1000 formulas) we may need a bulk-rewrite path.
5. **Defined-name authoring (P1).** Per `feature-scope.md`, defined
   names are read/preserved in P0 but not authored. When authoring
   ships, the engine needs a `definedNames` mutation listener; the
   §12.2 surface (`opts.definedNames` at parse time) is the
   integration point.
