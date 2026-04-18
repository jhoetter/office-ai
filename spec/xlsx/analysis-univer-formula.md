# Univer Formula Engine — Analysis (clean-room)

> Conceptual study of `@univerjs/engine-formula` and `@univerjs/sheets-formula` from
> the upstream repository `dream-num/univer` (Apache 2.0), commit
> **`3d6f4182f2d2f0f9be155592d336b39eff06fdd7`** (2026-04-17).
>
> This document records _concepts and architecture only_. No source code from
> Univer has been copied. All snippets that appear below are short pseudocode
> sketches written from scratch to explain ideas; they are not transcriptions.
> File paths refer to Univer's tree at the commit above and are quoted purely
> for orientation — they help us locate the equivalent of a concept, not to
> mirror line-for-line.

The goal is to inform a much smaller `packages/xlsx/src/formula/` module
sized for the ~150 functions enumerated in `prompt.md` line 247, with no
collaborative editing, no WASM, no plugin DI container, no iterative-calc
toggle.

---

## 0. High-level shape of Univer's engine

Univer separates the formula engine into two packages:

- `packages/engine-formula/` — the pure calculation core. Lexer, AST, AST
  evaluator, dependency graph, value/reference object model, the ~400+ Excel
  functions. No UI.
- `packages/sheets-formula/` — wires the engine into the sheet workbook
  controller, registers function descriptions for the formula bar, owns the
  range-tracking that adjusts formulas when rows/columns are inserted, and
  runs the calc loop in response to mutations.

The core directory tree under `packages/engine-formula/src/`:

- `basics/` — tokens, error enum, regex catalogue, common interfaces.
- `engine/analysis/` — `lexer.ts`, `lexer-tree-builder.ts`, `parser.ts`.
- `engine/ast-node/` — one file per AST node type plus a node factory.
- `engine/dependency/` — dependency tree + the topological run-list builder.
- `engine/interpreter/` — a tiny tree-walking executor.
- `engine/reference-object/` — Cell/Row/Column/Range/MultiArea/Table refs.
- `engine/value-object/` — Number/String/Bool/Error/Array/Cube/Lambda values.
- `functions/` — sub-folders per category, each with a `function-map.ts`
  registry and one folder per function (`sum/index.ts`, `vlookup/index.ts`…).
- `services/` — DI services for runtime state, current dataset, dependency
  manager, function registry, defined names, super tables, etc.
- `models/formula-data.model.ts` — persistent per-workbook formula state.
- `controllers/` — registers everything with the Univer DI container.

Because Univer uses dependency injection (`@Inject`, `createIdentifier`),
the calculation pipeline is composed of services that talk through
interfaces. We do **not** plan to mirror that DI ceremony.

---

## 1. Lexer

### 1.1 What "lexer" means in Univer

Univer's lexer is more than a flat token stream — it builds a _lexer tree_.
The class `Lexer` (`packages/engine-formula/src/engine/analysis/lexer.ts`) is
a thin façade over `LexerTreeBuilder`
(`packages/engine-formula/src/engine/analysis/lexer-tree-builder.ts`,
~1990 lines). The output is a `LexerNode`
(`packages/engine-formula/src/engine/analysis/lexer-node.ts`), a node with
a `token` string and a list of children that may themselves be `LexerNode`s
or raw string tokens. Each node has start/end indices into the original
formula string for accurate error/highlighting.

So "lexing" in Univer fuses tokenisation with bracket-grouping and a
preliminary classification: by the time the parser sees the tree, function
calls and parenthesised subexpressions are already grouped, parameter-comma
boundaries are encoded, and array literals are enclosed in their own
subtree.

### 1.2 Token catalogue

`packages/engine-formula/src/basics/token.ts` defines the recognised raw
character tokens:

- **Operators** (with precedence in `OPERATOR_TOKEN_PRIORITY`):
  - Comparison `=`, `<>`, `<`, `<=`, `>`, `>=` (priority 4 — lowest).
  - Concatenation `&` (priority 3).
  - Additive `+`, `-` (priority 2).
  - Multiplicative `*`, `/` (priority 1).
  - Exponent `^` (priority 0 — _highest_; lower number = higher).
- **Match tokens**: `(`, `)`, `,`, `'`, `"`, `{`, `}`, `:`, `[`, `]`.
- **Suffix tokens**: `%` (percent), `#` (spilled-range / structured-ref
  marker).
- **Prefix tokens**: `@` (implicit-intersection / `[@col]` style),
  unary `-`, unary `+`.
- A literal space token. Spaces are also Excel's _intersection operator_
  between two ranges.

### 1.3 Lexer state machine

`LexerTreeBuilder` walks the formula one character at a time, holding
several pieces of state:

- `_bracketState[]` — a stack with element type `NORMAL | FUNCTION |
LAMBDA` so opening/closing brackets remember context. A "function"
  bracket is opened by `<identifier>(` and turns the new lexer node into
  a function call node whose token is the function name.
- `_squareBracketState`, `_bracesState`, `_singleQuotationState`,
  `_doubleQuotationState`, `_colonState`, `_lambdaState`,
  `_tableBracketState`. Each is a counter or boolean; while inside a
  string, an array literal `{…}`, a quoted sheet name `'My Sheet'!`, or a
  table structured ref `Table[[#All],[Col1]]`, normal token recognition is
  suspended.
- `_segment` — the buffer accumulating the current token text.
- `_currentLexerNode` — the cursor into the tree being built; descends on
  `(` and `{`, ascends on the matching close.

Each "commit" of `_segment` either emits a string token (operator, match
character, comma, etc.), a value-like leaf (number, quoted string,
reference text, defined-name text), or pushes/pops a child node when a
bracket is opened/closed.

### 1.4 Specific cases

- **String literals** — between two `"` characters. Doubled `""` inside
  represents a literal quote. While `_doubleQuotationState !== 0`, the
  lexer treats every character as part of the string and does not break
  on operators or commas.
- **Error tokens** — `#VALUE!`, `#NAME?`, `#REF!`, `#DIV/0!`, `#NUM!`,
  `#N/A`, `#NULL!`, `#SPILL!`, `#CALC!`, `#GETTING_DATA`, `#CYCLE!` are
  enumerated in `packages/engine-formula/src/basics/error-type.ts`.
  `ERROR_TYPE_SET` and `ERROR_TYPE_COUNT_ARRAY` (the set of distinct
  lengths) let the lexer detect them with a constant-size lookahead when
  it sees a `#`.
- **References** — recognised by regex post-tokenisation, _not_ by a
  hand-rolled DFA. The regex catalogue is in
  `packages/engine-formula/src/basics/regex.ts`. Notable ones:
  - `REFERENCE_SINGLE_RANGE_REGEX` — `[Book.xlsx]Sheet1!$A$1` style with
    optional `#` spill suffix.
  - `REFERENCE_MULTIPLE_RANGE_REGEX` — `Sheet1!$A$1:$B$2`.
  - `REFERENCE_REGEX_ROW` / `REFERENCE_REGEX_COLUMN` — full-row `1:5` and
    full-column `A:C` references.
  - The unit prefix `[Book-1.xlsx]` and the optional sheet name with
    quoting `'My Sheet'!` are part of `UNIT_NAME_SHEET_NAME_REGEX`.
- **3D references** (`Sheet1:Sheet3!A1`) — the lexer treats the first
  colon-joined identifier as a sheet _range_; the regex catalogue
  recognises the syntactic shape, and a dedicated reference-object type
  fans out the read across the listed sheets at evaluation time.
  Univer's tree builder accepts the syntax; resolution lives downstream in
  the reference-object layer.
- **Array literals** `{1,2;3,4}` — the lexer enters "braces mode" on `{`,
  collects all characters (with `,` / `;` as element separators) and emits
  a value node carrying the full literal. The parser's `ValueNodeFactory`
  later parses the literal into a 2-D array of primitives.
- **Structured (table) references** — `Table3[[#All],[Col1]:[Col2]]` is
  detected via four dedicated regexes (`REFERENCE_TABLE_ALL_COLUMN_REGEX`,
  `REFERENCE_TABLE_SINGLE_COLUMN_REGEX`,
  `REFERENCE_TABLE_MULTIPLE_COLUMN_REGEX`,
  `REFERENCE_TABLE_TITLE_ONLY_ANY_HASH_REGEX`). The lexer's
  `_tableBracketState` keeps brackets nested correctly because table syntax
  uses `[[…]]` patterns that would otherwise look like overlapping range
  syntax.
- **Defined names** — when a token is not a number, string, error, or
  recognisable reference, the lexer marks the node as carrying a
  defined-name candidate; at parse time the defined-name service is
  consulted to expand it to a stored expression (and a "has defined
  names" flag propagates up the tree so the dependency graph knows to
  watch defined-name dirtiness).

### 1.5 Caching

Two LRU caches are exported (`FormulaLexerNodeCache`,
`FormulaSequenceNodeCache`) with capacity 2000. Repeated re-parsing of the
same formula text in different cells (very common after fill-down) is
deduplicated. Univer also keeps a separate AST cache
(`FORMULA_AST_CACHE` in `engine/utils/generate-ast-node.ts`).

---

## 2. Parser / AST

### 2.1 Algorithm

Two passes. First the `LexerTreeBuilder` produces an infix lexer tree.
Then the same builder rewrites each node's child list **into postfix
order** by running a Shunting-yard pass (`_suffixExpressionHandler`
in `lexer-tree-builder.ts` around the 880-line mark): two stacks
(`baseStack`, `symbolStack`) consume operators by precedence. After this
pass each node's children are an RPN sequence, which is exactly what the
AST builder wants.

The actual AST construction is in `packages/engine-formula/src/engine/
analysis/parser.ts` (`AstTreeBuilder._parse`). Because the children are
already in RPN, the code is a classic stack-based RPN-to-tree conversion:

```text
for each child token in postfix order:
  if it is an operator: pop two operands from stack, attach as children
  else: push as a leaf
final stack top is the AST root
```

This is _not_ a Pratt parser and _not_ a recursive-descent grammar in the
academic sense — Univer leans on the combination of (a) regex-based
classification of leaves and (b) Shunting-yard inside the lexer to keep
the parser itself short.

### 2.2 AST node types

Defined in `packages/engine-formula/src/engine/ast-node/node-type.ts`:

- `REFERENCE` — a cell, range, row, column, or table reference.
- `VALUE` — a literal number / string / boolean / array / error.
- `OPERATOR` — binary infix operators.
- `FUNCTION` — a named function call.
- `LAMBDA` / `LAMBDA_PARAMETER` — `LAMBDA(x,y, x+y)` and friends.
- `PREFIX` — unary `-`, `+`, `@`.
- `SUFFIX` — postfix `%`, `#`.
- `UNION` — the colon and space operators producing combined ranges.
- `ROOT` — top-level wrapper.
- `NULL`, `ERROR`, `BASE` — sentinels.

There is also a `NODE_ORDER_MAP` that gives each node-type a "z-index"
ordering. The AST builder iterates a _sorted_ list of node factories and
the first factory whose `checkAndCreateNodeType` accepts the lexer node
wins. This is how Univer disambiguates a token like `A1`: the
`ReferenceNodeFactory` sits at z-index 7, beating the `ValueNodeFactory`
(z-index 9) only because reference must be tested first.

Each AST node class extends `BaseAstNode`
(`engine/ast-node/base-ast-node.ts`) and exposes:

- `nodeType` getter.
- `getChildren()`, `setParent()`, `getValue()`, `setValue()`.
- `execute()` (sync) and `executeAsync()`.
- A flag for "forced recalculation" (see §4.4).

`OperatorNode`, `FunctionNode`, `ReferenceNode`, `ValueNode`,
`LambdaNode`, `LambdaParameterNode`, `PrefixNode`, `SuffixNode`,
`UnionNode`, `AstRootNode`, `NullNode` each have their own file, plus a
factory that decides whether a given lexer node should construct that
type.

### 2.3 Operator precedence and associativity

Precedence is encoded as the `OPERATOR_TOKEN_PRIORITY` map in
`basics/token.ts` (lower number = higher precedence). Associativity is
_left_ for everything; you can read this off the comparison
`charPriority >= lastSymbolPriority` in the Shunting-yard loop in
`_suffixExpressionHandler` — operators of equal priority pop
left-to-right.

Note: Univer's table puts `^` at priority 0 (the **highest** binding),
which matches Excel's behaviour where `2^3*4` is `(2^3)*4 = 32`. Excel
itself treats `-2^2` as `(-2)^2 = 4` because unary minus binds tighter
than `^`; Univer uses a `PrefixNode` for unary minus, separately from the
binary operator table.

### 2.4 Where this lives

- `engine/analysis/lexer.ts` — façade.
- `engine/analysis/lexer-tree-builder.ts` — character-level lexing,
  bracket grouping, Shunting-yard pass.
- `engine/analysis/parser.ts` — RPN-to-AST, node-factory dispatch,
  lambda parameter handling, `LET`-to-`LAMBDA` rewrite.
- `engine/ast-node/*.ts` — node implementations.

---

## 3. Reference / dependency graph

### 3.1 Direction and shape

Univer maintains a **forward** graph: each `FormulaDependencyTree`
(`packages/engine-formula/src/engine/dependency/dependency-tree.ts`) keeps
a `parents: Set<number>` (the formulas this one depends on, i.e. those
that must be calculated before it) and a `children: Set<number>` (the
formulas that depend on this one). So both directions are stored, but the
canonical direction the engine traverses for ordering is _parents_.

A "tree" in Univer's vocabulary is really a single node — one cell's
formula — annotated with the set of `IUnitRange` it reads
(`rangeList`), its position, its row/column counts, and the AST root.

### 3.2 Range membership lookup — R-tree

The interesting part is how Univer answers _"which formulas depend on
range X?"_ without scanning all formulas.
`packages/engine-formula/src/services/dependency-manager.service.ts`
keeps an `RTree` (imported from `@univerjs/core`) keyed by `(unitId,
sheetId, range)`. For each formula tree, every range it reads is inserted
as an R-tree item carrying that tree's id. When a dirty range is reported,
the manager bulk-searches the R-tree for tree-ids whose insertion range
intersects, giving an O(log n) lookup over potentially huge sheets.

This handles the `SUM(A1:A1000)` case naturally: the formula registers a
single R-tree item covering the 1000-cell range, regardless of how many
non-empty cells exist inside it. There is no per-cell expansion.

### 3.3 What counts as a dependency

In `formula-dependency.ts` (`FormulaDependencyGenerator`), a formula's
dependencies are computed from its AST root by:

1. Walking `REFERENCE` nodes — each contributes its static range.
2. For "address functions" like `OFFSET`, `INDIRECT`, `INDEX`,
   `CHOOSE`, the engine **pre-evaluates** them (see §3.5 below) to get a
   concrete `IUnitRange`, then inserts those ranges into the R-tree as
   well, into a separate `_dependencyRTreeCacheForAddressFunction`.
3. Defined-name expansions are inlined first (the lexer carries a
   "hasDefinedNames" flag), so a name resolves to its own ranges.
4. Volatile functions (RAND/NOW/TODAY/RANDBETWEEN) are not "dependencies"
   per se but get marked so they are forced dirty every recalc.

### 3.4 Topological run order

`FormulaDependencyGenerator._calculateRunList` produces the run order via
**iterative DFS with state tags** (not Kahn's algorithm). Each tree has
a small `FDtreeStateType` (`DEFAULT | ADDED | SKIP`):

- Pop a tree from the stack. If `SKIP`, drop it. If `ADDED` (it was
  re-pushed under its parents), it now belongs at the end of the run
  list — append it and mark `SKIP`.
- Otherwise, look up its parents (the things it depends on) and the R-tree
  intersections of address-function dependencies; if there are unfinished
  parents, mark this node `ADDED` and re-push it together with the
  parents.
- If no parents, append directly and mark `SKIP`.

The output is the run list in _bottom-up_ order (deepest dependencies
first). The calculate-formula service then `.reverse()`s it before
walking, giving a valid topological execution order. Cycle detection
falls out of this scheme — if a cycle exists, the marker bookkeeping
flags it and the runtime sets `_isCycleDependency = true`.

### 3.5 Pre-calculation of address functions

The trick that lets Univer compute dependencies _before_ evaluation is:
walk the AST and find any `OFFSET`/`INDIRECT`/`INDEX` nodes, evaluate
**only those nodes** through the interpreter to get the resulting
`BaseReferenceObject`, then call `.toUnitRange()` and feed those ranges
back to the dependency graph. This is the loop in
`_calculateListByFunctionRefNode` and `_calculateAddressFunction`.
Because the address function might in turn depend on other formulas,
those are pulled in and pre-calculated too — so dependency analysis and
calculation interleave for the dynamic-reference subset.

This is the part of Univer that we will _not_ attempt to replicate at
parity in v1; see §11.

---

## 4. Calculation / evaluator

### 4.1 Eager vs lazy

**Eager** along the dirty subset. After a mutation, the engine asks the
dependency generator to produce a run list of dirty trees, then walks the
run list in order, calling `Interpreter.execute(nodeData)` on each. The
interpreter is in
`packages/engine-formula/src/engine/interpreter/interpreter.ts` and is a
plain depth-first tree walker:

```text
execute(node):
  for child in node.children: execute(child)   # post-order
  if node is REFERENCE:  attach (refOffsetX, refOffsetY)
  if node is FUNCTION & async: await node.executeAsync()
  else: node.execute()
```

Each AST node carries its own value through `node.getValue()` —
intermediate values stay attached to the AST until the root produces the
final result. There is no separate evaluation stack.

### 4.2 Batching after a mutation

The pipeline is in
`packages/engine-formula/src/services/calculate-formula.service.ts`,
class `CalculateFormulaService`:

1. `execute(formulaDatasetConfig)` is called with the new dirty ranges,
   dirty defined names, dirty sheet renames, dirty other-formulas.
2. It acquires an async lock so concurrent `execute` calls serialise.
3. It runs `_executeStep()` once for normal formulas and then again with
   `isArrayFormulaState=true` for array formulas whose spilled output
   dirties further cells.
4. `_apply()` calls `_formulaDependencyGenerator.generate()` to get the
   ordered tree list, then iterates calling `interpreter.execute` (or
   `executeAsync` if any node in the AST is async).
5. Every `intervalCount` (default 500) iterations it yields with a
   `requestImmediateMacroTask` so the main thread stays responsive — this
   is Univer's cooperative-scheduling concession to the browser.
6. After execution it emits `_executionCompleteListener$.next(runtimeData)`
   carrying a per-cell map of new values; subscribers (the sheet UI
   layer) write them into the actual workbook.

### 4.3 Dirty tracking

Dirty input is _external_ — Univer does not auto-detect what changed.
Whoever mutates the workbook must supply `dirtyRanges`, `dirtyNameMap`,
`dirtyDefinedNameMap`, `dirtyUnitFeatureMap` to `execute()`. The
dependency generator translates dirty ranges into an R-tree query, which
yields the set of dependent tree-ids; the union of those (plus the
externally-named dirty trees) is the set the run-list builder operates on.

`FormulaDependencyTree` itself has `isDirty: boolean` and an `isCache:
boolean` flag — a tree built last time and unchanged keeps its parsed AST
and its R-tree entries; only its dirty bit toggles.

Why this is interesting: it means the engine cleanly supports
"partial recalc on a 500k-row sheet" — the dirty-range R-tree intersection
keeps the working set proportional to the affected range, not the sheet.

### 4.4 Volatile functions

`packages/engine-formula/src/basics/token-type.ts` exports:

```text
FORCED_RECALCULATION_FUNCTION_NAME = { 'RAND', 'RANDBETWEEN', 'NOW', 'TODAY' }
```

When the parser builds a `FunctionNode` whose token is in this set, it
calls `astNode.setForcedCalculateFunction()`. The dependency generator's
`_detectForcedRecalculationNode` walks the tree and, if any node carries
the flag, returns `true` from `_includeTree`, forcing the formula into
the dirty set every recalc regardless of range intersections.

This is a coarse approach (any RAND anywhere in a workbook re-runs every
time _something_ recalculates), but it is correct for Excel's volatility
semantics.

---

## 5. Async / iterative evaluation

### 5.1 Async functions

Yes. `BaseFunction` exposes `isAsync()` (default `false`), and async
functions override it. `FunctionNode` checks the executor's `isAsync()`
in its constructor and tags the AST node async; `Interpreter.checkAsyncNode`
walks an AST and reports whether **any** node is async, which decides
whether the calc loop awaits or runs synchronously per-formula:

```text
if interpreter.checkAsyncNode(rootNode):
  value = await interpreter.executeAsync(nodeData)
else:
  value = interpreter.execute(nodeData)
```

The async path matters for things like Univer's hyperlink/external-data
functions and for the `getExecutor(...)` API where third parties can
register custom async functions. In our scope we have no remote-data
functions, so we can stay synchronous and skip this branch.

### 5.2 Iterative calculation

Partially. `packages/engine-formula/src/config/config.ts` defines
`DEFAULT_CYCLE_REFERENCE_COUNT = 1`. The calc loop wraps the entire
recalc in `for (let i = 0; i < cycleReferenceCount; i++)`. Each pass:

- Runs `_executeStep()` once.
- Reads `_isCycleDependency`. If false, breaks early.
- If true and `i < cycleReferenceCount`, loops and re-runs.

This is the convergence loop Excel uses for circular references. Univer
defaults to 1 pass — i.e. circular refs are detected but **not**
iteratively converged unless the caller explicitly raises
`maxIteration`. The runtime exposes `enableCycleDependency` /
`disableCycleDependency`. There is no per-cell convergence threshold; the
loop just runs N times.

We will not implement iterative calc — see §11.

---

## 6. Function library

### 6.1 Where the functions live

Each Excel function is its own folder under
`packages/engine-formula/src/functions/<category>/<name>/`, with an
`index.ts` exporting a class extending `BaseFunction`. Categories:

- `array/` (2)
- `compatibility/` (~7 — old `BETADIST`-style functions)
- `cube/` (none yet, scaffolding only)
- `database/` (~12 `DSUM`, `DGET` family)
- `date/` (~27)
- `engineering/` (~57 — `BIN2HEX`, complex numbers, `BESSELI`, etc.)
- `financial/` (~54)
- `information/` (~23 — `IS*`, `TYPE`, `CELL`)
- `logical/` (~19)
- `lookup/` (~35)
- `math/` (~81)
- `meta/` (~6 — internal helpers, not Excel functions)
- `statistical/` (~107)
- `text/` (~47)
- `web/` (1 — `WEBSERVICE`)

That is roughly 470+ Excel functions implemented or stubbed. Actual
shipping count differs slightly because some are scaffolds.

### 6.2 Registration shape

Two layers:

1. **Executor registration** — each category folder has a
   `function-map.ts` exporting an array of `[ClassRef, NameEnum]`
   tuples (e.g. `[Sum, FUNCTION_NAMES_MATH.SUM]`) and a
   `function-names.ts` enum. A controller in `engine-formula/src/
controllers/` instantiates each class, passing the name to
   `BaseFunction`'s constructor, and calls
   `IFunctionService.registerExecutors(...instances)`.
2. **Description registration** — `IFunctionInfo`
   (`packages/engine-formula/src/basics/function.ts`) holds metadata for
   the formula bar (function name, type, abstract, parameter list with
   examples). `sheets-formula/src/services/function-list/` ships
   per-category description files used by `IDescriptionService`.

Lookup at evaluation time: `FunctionNodeFactory` consults
`IFunctionService.getExecutor(token)` to bind a `FunctionNode` to its
executor. Unknown function name → `ErrorNode.create(ErrorType.NAME)`
which propagates `#NAME?`.

### 6.3 BaseFunction shape

`packages/engine-formula/src/functions/base-function.ts`:

- Properties: `minParams`, `maxParams`, plus opt-in flags
  `needsExpandParams`, `needsReferenceObject`, `needsLocale`,
  `needsSheetsInfo`, `needsFormulaDataModel`, `needsSheetRowColumnCount`,
  `needsFilteredOutRows`.
- Hooks: `setRefInfo(unitId, subUnitId, row, col)`, `setSheetsInfo(...)`,
  etc.
- `calculate(...args: BaseValueObject[]): BaseValueObject` — the
  override point. Default returns `#VALUE!`.
- `isAsync()`, `isAddress()`, `isCustom()`, `isArgumentsIgnoreNumberPattern()`.

Concrete example, `Sum` in
`packages/engine-formula/src/functions/math/sum/index.ts`, conceptually:

```text
class Sum extends BaseFunction {
  minParams = 1; maxParams = 255;
  calculate(...variants):
    acc = NumberValueObject(0)
    for v in variants:
      if v.isString(): v = v.convertToNumberObjectValue()
      if v.isError(): return v
      if v.isArray(): v = v.sum()      # ArrayValueObject collapse
      acc = acc.plus(v)
    return acc
}
```

Two things to note:

- Argument coercion is _inside_ each function, not centralised. `Sum`
  decides that string `"42"` becomes a number; `Concat` would call
  `.toString()`. There is no schema-driven coercion table.
- Array inputs are handled by methods on `ArrayValueObject`
  (`engine/value-object/array-value-object.ts`), so functions like SUM
  can collapse a 2-D range with one call.

### 6.4 Argument coercion model

The value object hierarchy
(`packages/engine-formula/src/engine/value-object/`):

- `BaseValueObject` — abstract.
- `NumberValueObject`, `StringValueObject`, `BooleanValueObject`,
  `NullValueObject`, `ErrorValueObject` — primitives
  (`primitive-object.ts`).
- `ArrayValueObject` — 2-D matrix of `BaseValueObject`s.
- `LambdaValueObject`, `CubeValueObject` — for `LAMBDA(...)` and OLAP.

Coercion methods on `BaseValueObject`: `convertToNumberObjectValue`,
`convertToStringObjectValue`, `convertToBooleanObjectValue`. Functions
call these explicitly on each arg they want as a particular type, and
short-circuit on errors. There is no automatic schema-coercion layer.

### 6.5 Array formulas / dynamic arrays

Two paths:

- A function may _opt in_ to array semantics. `BaseFunction`'s
  `needsExpandParams` flag tells `FunctionNode` to broadcast scalar
  inputs against array inputs and call `calculate` per cell.
- Functions that natively return arrays (e.g. `SEQUENCE`, `UNIQUE`,
  `FILTER`, `RANDARRAY`) return an `ArrayValueObject`. The
  `CalculateFormulaService._executeStep` runs a _second_ pass with
  `isArrayFormulaState=true` — it discovers the spill range produced by
  any array-returning formula in pass one, marks those new cells dirty,
  and re-runs the dependency graph. This is how Univer keeps spilled
  output consistent with downstream formulas in a single recalc.

`#SPILL!` is raised when a dynamic array would overlap an occupied cell;
this is a property of the spill resolution pass, not of the function
itself.

---

## 7. Error model

### 7.1 Errors are first-class values

`ErrorValueObject extends BaseValueObject`. Every value-returning path
can return an error object, and arithmetic methods on `BaseValueObject`
short-circuit when called with an error operand. The error enum
(`basics/error-type.ts`) is the full Excel set:

| ErrorType     | Excel display   | When                                        |
| ------------- | --------------- | ------------------------------------------- |
| `DIV_BY_ZERO` | `#DIV/0!`       | `=1/0`, `MOD(1,0)`                          |
| `NAME`        | `#NAME?`        | unknown function, unparseable name, bad ref |
| `VALUE`       | `#VALUE!`       | wrong type to a function                    |
| `NUM`         | `#NUM!`         | numeric overflow / domain error             |
| `NA`          | `#N/A`          | lookup miss, `NA()`                         |
| `CYCLE`       | `#CYCLE!`       | unresolved circular reference               |
| `REF`         | `#REF!`         | reference to deleted/invalid cell           |
| `SPILL`       | `#SPILL!`       | dynamic array can't expand                  |
| `CALC`        | `#CALC!`        | empty array or calc-engine refusal          |
| `ERROR`       | `#ERROR!`       | catch-all parse failure                     |
| `CONNECT`     | `#GETTING_DATA` | external query in flight                    |
| `NULL`        | `#NULL!`        | space-intersection of disjoint ranges       |

### 7.2 Propagation

Three patterns coexist:

1. **Operator/arithmetic methods** — `BaseValueObject.plus`, `.minus`,
   `.compareBy`, etc. propagate errors immediately by checking
   `.isError()`.
2. **Function bodies** — call sites explicitly `if (v.isError()) return v;`
   inside `calculate`. Functions like `IFERROR`/`IFNA` use the same check
   to _catch_ the error.
3. **AST-level** — when the parser cannot build a valid node, it emits an
   `ErrorNode` carrying an `ErrorValueObject`; the interpreter's
   tree-walker yields it as the formula's result.

`ERROR_TYPE_SET` and `ERROR_TYPE_COUNT_ARRAY` enable the lexer to
recognise an error-literal token like `#REF!` typed by a user, so e.g.
`=IFERROR(#REF!, 0)` round-trips correctly.

### 7.3 Cached error objects

`ErrorValueObjectCache` (in `base-value-object.ts`) interns error objects
by enum so that repeated returns of `#VALUE!` reuse a single instance.
This keeps GC pressure down on workbooks with many error cells.

---

## 8. Formula text vs cached value vs serialized value

### 8.1 What's in a Cell

The persistent form is `IFormulaDataItem`
(`packages/engine-formula/src/basics/common.ts`):

```text
{ f?: string,      // formula text, e.g. "=SUM(A1:A10)"
  si?: string,     // shared formula id (for fill-down)
  x?: number,      // x offset from the canonical (si) cell
  y?: number   }   // y offset from the canonical (si) cell
```

The companion `IFormulaIdMap` records the canonical cell coordinate of
every `si`. So a fill-down stores the formula text _once_ with `si`,
and every dependent cell stores `{ si, x, y }` (no `f`). The lexer's
`moveFormulaRefOffset(formulaText, x, y)` reconstructs the per-cell
formula on demand.

The actual computed value lives in the workbook cell itself
(`ICellData.v`) outside the formula module — written back by the
`CalculateFormulaService.executionCompleteListener$` subscriber. So you
have two stores:

- Formula model: `f` (or `si+x+y`) per cell.
- Sheet model: the cached value `v`, the display string `p`, the
  number-format-applied string. None of these live inside the formula
  engine.

### 8.2 Persistence to OOXML

OOXML serialisation is done outside `engine-formula`. The pieces it needs
are exactly what's in `IFormulaDataItem`:

- `<f>` element on the cell carries the formula text (or `t="shared"
ref="A1:A10" si="0"` for the canonical cell of a shared formula, and
  `t="shared" si="0"` for dependents — matching `si+x+y` model).
- Cell `<v>` carries the cached value.
- Array formulas serialise as `<f t="array" ref="…">` plus the spill
  range cells.

The formula-data-model treats fill-down as _first-class state_, not as
syntactic sugar — which is why round-tripping shared formulas in OOXML is
straightforward.

---

## 9. R1C1 vs A1 conversion + relative/absolute reference adjustment

### 9.1 R1C1 vs A1

Univer stores reference text in **A1** form (with optional `$`
absolutes and `Sheet1!`/`[Book.xlsx]` prefixes). There is no R1C1
representation in the AST — R1C1 is purely a UI display mode. The lexer
regexes are A1-only.

That said, the _internal_ representation of an `IRange` is
`{startRow, startColumn, endRow, endColumn, startAbsoluteRefType,
endAbsoluteRefType}` — purely numeric and direction-agnostic, so an R1C1
display layer (which Univer does ship in the UI) is a one-pass formatter
over the same range objects.

### 9.2 Relative/absolute insertion adjustment

When a row is inserted, the workbook controller calls into
`packages/sheets-formula/src/services/formula-ref-range.service.ts`
(`FormulaRefRangeService`). This service:

1. Walks all formulas whose `rangeList` intersects (or sits below) the
   inserted row.
2. For each such reference, computes the new `IRange` by
   `moveRangeByOffset(range, offsetX, offsetY)`. The function honours
   `AbsoluteRefType` per side — absolute coordinates are not shifted.
3. Re-serialises the updated reference back into the formula text via
   the lexer's `serializeRangeToRefString` and replaces the corresponding
   token in the original formula string.

So the _formula text on disk_ is rewritten on insert/delete. The AST is
not mutated in place — it is re-parsed lazily next time the cell is
recalculated, and the dependency tree is rebuilt with the new ranges.

### 9.3 Copy/paste

Copy-paste uses the same `moveRangeByOffset` mechanism but with a
positive (target − source) offset, applied only to _relative_ portions
of references (because absolute `$A$1` is invariant). The fill-handle is
the `si+x+y` shared-formula model (§8.1) — Univer prefers shared formulas
for fill-down so the references stay implicit.

`AbsoluteRefType` (an `@univerjs/core` enum) has four values:
`NONE | ROW | COLUMN | ALL`, recording which side(s) of `A1` were
written with a `$`. The serializer reads that back when emitting the
text.

---

## 10. What we should KEEP for our 80%-scope engine

The pieces of Univer's design that earn their keep at our scope (~150
functions, single-workbook in-process calc, no plugin DI) are:

1. **Postfix-converted lexer tree** — combining tokenisation and
   bracket-grouping into one pass produces an artifact the AST builder
   can consume with a tiny stack-based RPN-to-tree algorithm. We will
   adopt this conceptually: a lexer that returns an infix tree of `(`-
   grouped chunks plus a Shunting-yard pass to flip each child list to
   postfix.
2. **Regex-classified leaves** — `A1`, `Sheet1!A1:B2`, `'My Sheet'!A:A`
   etc. live in a small catalogue of regexes. This is much shorter than
   a hand-rolled DFA and easier to audit.
3. **First-class error values** — `ErrorValueObject` short-circuiting
   inside arithmetic methods is the cleanest way to get correct
   propagation without try/catch sprinkled through every function.
4. **Forward graph on top of an R-tree** — for dirty-range queries
   over wide ranges (`SUM(A1:A100000)`), an R-tree is the right
   structure. We will use it (or a simple interval tree if a tree
   library is overkill in pass 1).
5. **`FORCED_RECALCULATION_FUNCTION_NAME` set** — the simplest possible
   volatile-function story. Mark the AST node, force the cell into the
   dirty set every recalc.
6. **`si + x + y` shared formula model** — saves a huge amount of
   memory on filled-down columns and matches OOXML's shared-formula
   serialisation directly.
7. **AbsoluteRefType per range side** — the four-state enum
   (NONE/ROW/COLUMN/ALL) is the only sane way to handle `$A$1`-style
   adjustment; we will adopt it.
8. **External dirty-range injection** — the engine should not snoop on
   the workbook for changes; mutators tell it what changed. This keeps
   the calc engine reusable from any layer (CLI, agent, tests).
9. **One folder per function with a `function-map.ts` per category** —
   easy to grep, easy to lazy-import, easy to delete. We adopt this
   layout 1:1.
10. **Iterative DFS run-list with state tags** — the deferred-append
    trick (mark `ADDED`, re-push under parents, append on revisit) is a
    nice way to get topological order without a separate Kahn-style
    in-degree map.
11. **Dependency tree pre-evaluation of address functions** — we will
    keep the _concept_ (compute `INDIRECT`/`OFFSET` ranges before
    treating their result as static dependencies), but limited to the
    subset of address functions we ship.
12. **Postfix `%` and `#`, prefix `@`/`-`/`+`** — having explicit
    `PrefixNode`/`SuffixNode` types in the AST keeps the operator table
    clean.

---

## 11. What we should DIFFER ON / SIMPLIFY

For our scope (in-process Node + browser, ~150 functions, no
collaborative editing, no remote connectors, no full Excel parity), we
deliberately drop:

1. **Plugin DI container** — Univer wires every service through
   `@Inject` and `createIdentifier`. We will use plain modules and
   constructor-injected dependencies; a function service is a
   `Map<string, FunctionImpl>` not a DI singleton.
2. **Async function path** — no `WEBSERVICE`, no remote data, no
   Univer-cloud functions. `Interpreter` stays synchronous; we drop
   `executeAsync` and `checkAsyncNode` entirely.
3. **Iterative calculation** — the
   `for (i = 0; i < cycleReferenceCount; i++)` loop and
   `_isCycleDependency` flag go away. Circular refs return `#CYCLE!`
   immediately.
4. **WASM acceleration** — Univer doesn't actually ship WASM in the
   formula engine, but where it would (e.g. heavy statistical
   functions), we explicitly stay in pure TS. Performance is a non-goal
   for the long tail.
5. **Cube functions, web functions, database functions** — we don't
   ship `CUBEVALUE`, `DGET`, `WEBSERVICE`. The whole `cube/`,
   `database/`, `web/` sub-trees go away.
6. **Engineering functions (mostly)** — `BIN2HEX`, `BESSELI`, complex
   numbers — we ship none of these in v1. Maybe revisit if needed.
7. **Lambda / LET / `MAP`/`REDUCE`/`MAKEARRAY`/`SCAN`/`BYROW`/`BYCOL`** —
   Univer supports user-defined `LAMBDA` and the modern functional set
   with their own `LambdaNode` / `LambdaParameterNode` and a runtime
   privacy-var environment. We do not need user-defined lambdas in v1;
   leaving them out drops two AST node types and the `_lambdaParameterHandler`
   entirely.
8. **`LET` rewrite** — `LET` expressions are rewritten into `LAMBDA` calls
   in `parser.ts._changeLetToLambda`. With no lambda support, no `LET`.
9. **Defined-name machinery** — Univer has a full `IDefinedNamesService`,
   plus a "dirty defined names" map plumbed through every recalc. v1
   can ship without defined names; if/when we add them, we expand them
   inline at parse time and skip the dirty-tracking layer (recalc all
   formulas referencing the name).
10. **Super tables (Excel structured references)** — Univer has a whole
    `super-table.service.ts`, `TableReferenceObject`, four regexes for
    `Table[…]` syntax. We treat structured references as out-of-scope in
    v1 (formulas using them will return `#REF!`).
11. **External workbook references** — the `[Book-1.xlsx]Sheet1!A1`
    syntax. Our scope is one workbook at a time; cross-book references
    return `#REF!`.
12. **Three caches: lexer, AST, sequence-node** — Univer keeps three
    LRU caches (size 2000 each). We start with a single AST cache keyed
    by `(unitId, sheetId, formulaText)` and revisit if profiling shows
    it matters.
13. **R-tree from `@univerjs/core`** — that R-tree is a dependency we
    don't want. v1 can use a `Map<rangeKey, Set<treeId>>` plus a
    flattened scan; if perf is an issue on workbooks with millions of
    cells we drop in `rbush` (MIT) or our own.
14. **FormulaDependencyTreeVirtual for shared formulas** — the
    distinction between a "real" and "virtual" tree exists because
    shared formulas don't get their own AST. We can take the simpler
    route: expand shared formulas into per-cell trees lazily, accepting
    the memory cost (sheets with 100k shared formulas are rare in our
    target use cases).
15. **Feature-formula and other-formula machinery** — Univer supports
    embedded calc for features like data validation, conditional
    formatting, pivot tables. For our package these formulas go through
    the same path as cell formulas.
16. **The cooperative `await requestImmediateMacroTask` every 500
    formulas** — we run in the agent runtime where freezing for a few
    seconds is fine; we can add this concession later if we run in a
    browser tab.
17. **`#CALC!`, `#GETTING_DATA`, `#CYCLE!`, `#SPILL!`** — we ship the
    other seven errors faithfully but treat these four as v1-only-emit-
    on-construction errors (no internal pathways produce them in our
    minimal engine).

---

## 12. Suggested architecture for `packages/xlsx/src/formula/`

### 12.1 Directory layout

```text
packages/xlsx/src/formula/
  README.md                  # short architecture note + non-goals
  index.ts                   # public surface: parse(), evaluate(), recalc()
  tokens.ts                  # operator/match/prefix/suffix enums + priority map
  errors.ts                  # ErrorType enum + ErrorValue class
  values.ts                  # NumberValue, StringValue, BoolValue,
                             # NullValue, ArrayValue, ErrorValue
  references.ts              # IRange, AbsoluteRefType, CellRef, RangeRef,
                             # serialize/parse helpers, moveByOffset
  lexer.ts                   # character scanner -> infix lexer tree
  shunting-yard.ts           # in-place rewrite of each lexer node's
                             # children to postfix
  ast.ts                     # AST node classes (Function, Operator,
                             # Reference, Value, Prefix, Suffix, Union,
                             # Root, Null, Error)
  parser.ts                  # RPN -> AST, factory dispatch, defined-name
                             # inlining (later)
  evaluator.ts               # post-order tree walker, single sync entry
  dependency-graph.ts        # forward graph + range-index for dirty queries
  recalc.ts                  # orchestration: take dirty ranges -> run list ->
                             # walk evaluator -> emit value events
  function-registry.ts       # Map<string, FunctionImpl>; volatile set;
                             # description metadata for UI hints
  functions/
    index.ts                 # registers all categories
    math.ts                  # SUM, AVERAGE, COUNT, ROUND, ABS, MOD, …
    logic.ts                 # IF, IFS, AND, OR, NOT, XOR, IFERROR, …
    lookup.ts                # VLOOKUP, HLOOKUP, INDEX, MATCH, OFFSET,
                             # INDIRECT, ROW, COLUMN, …
    text.ts                  # CONCAT, LEFT, RIGHT, MID, LEN, TRIM, …
    date.ts                  # TODAY, NOW, DATE, YEAR, MONTH, DATEDIF, …
    finance.ts               # PMT, PV, FV, RATE, NPER, NPV, IRR, SLN
    info.ts                  # ISBLANK, ISNUMBER, ISTEXT, ISERROR, TYPE, N
    array.ts                 # UNIQUE, SORT, FILTER, SEQUENCE (later)
    helpers/
      coerce.ts              # toNumber, toString, toBoolean coercers
      numeric.ts             # safe-divide, round-half-away, etc.
      date-serial.ts         # Excel epoch (1900-01-01 with leap-year bug)
      compare.ts             # excel-style comparison (text vs num vs bool)
  __tests__/                 # vitest specs per category
```

### 12.2 Design notes per file

- **`tokens.ts`** — single source of truth for operator priority. Match
  Univer's table including `^` at top priority. Don't duplicate the
  enum across files.
- **`errors.ts`** — `ErrorValue` subclass of `Value` with `kind:
ErrorType`. Export interned singletons per kind (`ErrorValue.VALUE`,
  `.NA`, etc.).
- **`values.ts`** — `Value` is an abstract class with `.kind`,
  `.toNumber()`, `.toString()`, `.toBoolean()`, `.isError()`,
  `.isArray()`. Use TypeScript discriminated unions where it helps;
  classes where polymorphism on `plus`/`minus` is cleaner. Pick one
  style and stick with it (see `typescript-exhaustive-switch.mdc`).
- **`references.ts`** — `IRange = {sheet, r0, c0, r1, c1, abs0, abs1}`.
  `serializeRange()` produces `Sheet1!$A$1:$B$2`-style text;
  `parseRange()` is the inverse and feeds the lexer for reference
  leaves.
- **`lexer.ts`** — single function `tokenize(text): LexerNode`. The
  `LexerNode` is `{ token: string, kind: NodeKind, children: (string |
LexerNode)[], start: number, end: number }`. Handles `()`, `{}`, `""`,
  `''` (sheet quoting), and `#ERROR!` literals.
- **`shunting-yard.ts`** — function `toPostfix(node: LexerNode): void`
  that mutates `node.children` in place. Recursion descends into nested
  function-call/parenthesised lexer nodes.
- **`ast.ts`** — one class per `NodeType`; each implements
  `evaluate(ctx: EvalContext): Value`. The context carries the sheet
  data, the current cell coords, the function registry. Children are
  evaluated post-order in `evaluate`.
- **`parser.ts`** — `parse(text: string): AstNode`. Calls `tokenize`
  then `toPostfix(root)`, then walks the lexer tree converting RPN to
  AST nodes. Unknown tokens become `ErrorNode(NAME)`.
- **`evaluator.ts`** — exposes `evaluate(node, ctx)`. Internally just
  delegates to `node.evaluate(ctx)`; the file mostly exists as a place
  to put orchestration helpers (timeouts, debug instrumentation).
- **`dependency-graph.ts`** — `class DepGraph` holds:
  - `Map<CellKey, FormulaNode>` — every cell that has a formula.
  - `RangeIndex` — interval-tree-ish structure mapping a range to the
    set of formula-cell keys whose `rangeList` intersects.
  - `parents: Map<CellKey, Set<CellKey>>` and `children: Map<CellKey,
Set<CellKey>>` for the forward graph.
  - `markDirty(ranges: IRange[]): Set<CellKey>` — bulk-search the
    range index plus volatile set, returns the dirty set.
  - `topoOrder(dirty): CellKey[]` — DFS with the `ADDED/SKIP` state
    trick from §3.4.
- **`recalc.ts`** — `class Recalc { run(dirty: IRange[]): RecalcResult }`.
  Pulls dirty cells from `DepGraph`, parses any not-yet-parsed formulas,
  walks `topoOrder` calling `evaluator.evaluate`, collects new cell
  values into a result map. Cycle detection sets `#CYCLE!` on every
  unresolved cell.
- **`function-registry.ts`** — `class FunctionRegistry { register(
name, impl, opts?: { volatile?, async? } ): void; get(name):
FunctionImpl | undefined }`. Single instance owned by the workbook;
  no DI.
- **Per-category function file** — each file imports from
  `function-registry.ts`'s register helper and exports nothing (or
  exports a `register(reg)` function the workbook calls at startup).
  Functions are plain functions over `Value[]`, not classes:

```text
export function SUM(args: Value[]): Value {
  let acc = 0
  for (const a of flatten(args)) {
    if (a.isError()) return a
    if (a.isString()) { const n = toNumber(a); if (n.isError()) return n; acc += n.value; continue }
    if (a.isNumber()) acc += a.value
  }
  return new NumberValue(acc)
}
```

Classes are heavier than we need; closures around state (e.g. SUMIF
with cached predicate) are easier than `BaseFunction` subclasses with
a dozen capability flags.

### 12.3 Public API

The package's `index.ts` exposes a tight surface:

- `parseFormula(text: string): AstNode | ErrorValue`
- `class Workbook` (probably from elsewhere) gets a `formula: Recalc`
  instance.
- `recalc.run(dirty: IRange[]): { values: Map<CellKey, Value>, errors:
Map<CellKey, ErrorValue> }`.
- `registerFunction(name, impl, opts)` — for tests / extensions.

No streaming/observable API in v1. Recalc returns a plain object;
callers update the workbook synchronously.

### 12.4 Conformance & test strategy

- A `__tests__/conformance/` folder with table-driven tests per
  function: input args (literals + small ranges) → expected output. We
  can seed a few hundred of these from public Excel test vectors
  (available as MS Office spec annexes) without copying anyone's code.
- Round-trip tests: `parseFormula(text); serialise(ast)` must equal
  `text` (modulo whitespace).
- Recalc tests: small workbook fixtures with known dependency chains;
  assert run order and final values.
- `prompt.md`'s 150-function priority list maps directly to the
  per-category files; we can track per-function status in a
  `coverage.md`.

---

## 13. Open questions for follow-up

These don't block the v1 architecture but should be revisited before we
ship:

- **Dynamic arrays as a recalc step** — should we adopt Univer's
  two-phase recalc (normal then array-formula spill) on day one, or
  defer until we ship `UNIQUE/SORT/FILTER/SEQUENCE`?
- **Persistence of cached value vs recompute on open** — Excel writes
  `<v>` for every formula cell; do we trust those values on open or
  always recompute? Recommend trust-then-validate (mark all formulas
  dirty in a background pass).
- **Range-index data structure** — start with a flat `Map`, switch to
  `rbush` when per-recalc dependency-search times exceed a target
  budget on a representative 50k-row workbook.
- **Number-format awareness in functions** — `TEXT(...)` and `DATEVALUE(...)`
  need a number-format engine. We have one in `@officeai/xlsx`'s
  serialiser; expose its parse/format primitives in `functions/helpers/`.
- **Sheet rename / delete invalidation** — Univer plumbs
  `dirtyNameMap`/`dirtyDefinedNameMap` through. We need an analogous
  signal in `recalc.run({ dirtyRanges, renamedSheets, deletedSheets })`.
- **Date math correctness** — Excel's "1900 is a leap year" bug is
  faithfully reproduced; pick a side (compatibility vs correctness) and
  document it in `functions/helpers/date-serial.ts`.

---

## Appendix A — Univer file map (for orientation)

Quick lookup table of where each concept lives upstream, useful when
checking "how did Univer handle X?":

| Concept                              | Univer file                                                           |
| ------------------------------------ | --------------------------------------------------------------------- |
| Operator precedence table            | `packages/engine-formula/src/basics/token.ts`                         |
| Error enum                           | `packages/engine-formula/src/basics/error-type.ts`                    |
| Volatile function set                | `packages/engine-formula/src/basics/token-type.ts`                    |
| Reference regex catalogue            | `packages/engine-formula/src/basics/regex.ts`                         |
| Cell-vs-range serialised shape       | `packages/engine-formula/src/basics/common.ts`                        |
| Lexer façade                         | `packages/engine-formula/src/engine/analysis/lexer.ts`                |
| Lexer tree builder + Shunting-yard   | `packages/engine-formula/src/engine/analysis/lexer-tree-builder.ts`   |
| Lexer node                           | `packages/engine-formula/src/engine/analysis/lexer-node.ts`           |
| AST builder                          | `packages/engine-formula/src/engine/analysis/parser.ts`               |
| AST node types                       | `packages/engine-formula/src/engine/ast-node/node-type.ts`            |
| Function AST node                    | `packages/engine-formula/src/engine/ast-node/function-node.ts`        |
| Reference AST node                   | `packages/engine-formula/src/engine/ast-node/reference-node.ts`       |
| Tree-walking interpreter             | `packages/engine-formula/src/engine/interpreter/interpreter.ts`       |
| Dependency tree                      | `packages/engine-formula/src/engine/dependency/dependency-tree.ts`    |
| Dependency generator + topo run-list | `packages/engine-formula/src/engine/dependency/formula-dependency.ts` |
| Range-index R-tree wrapper           | `packages/engine-formula/src/services/dependency-manager.service.ts`  |
| Recalc orchestrator                  | `packages/engine-formula/src/services/calculate-formula.service.ts`   |
| Iterative-calc setting               | `packages/engine-formula/src/config/config.ts`                        |
| Function registry                    | `packages/engine-formula/src/services/function.service.ts`            |
| BaseFunction                         | `packages/engine-formula/src/functions/base-function.ts`              |
| Per-category function map            | `packages/engine-formula/src/functions/<category>/function-map.ts`    |
| Value object hierarchy               | `packages/engine-formula/src/engine/value-object/`                    |
| Reference object hierarchy           | `packages/engine-formula/src/engine/reference-object/`                |
| Persistent formula state             | `packages/engine-formula/src/models/formula-data.model.ts`            |
| Insert/delete formula adjustment     | `packages/sheets-formula/src/services/formula-ref-range.service.ts`   |
| Function descriptions for UI         | `packages/sheets-formula/src/services/function-list/`                 |

---

## Appendix B — Pseudocode of the main pipelines

### B.1 Lexer (one pass)

```text
tokenize(text):
  root = new LexerNode("ROOT")
  cursor = root
  while not end:
    read char
    if in-string-quote: append to segment; on closing quote -> emit value-leaf
    elif char == '(' after identifier: cursor = cursor.openChild(FUNCTION, identifier)
    elif char == '(' alone:           cursor = cursor.openChild(GROUP)
    elif char == ')':                 cursor = cursor.parent
    elif char == ',' at function arg: cursor.beginNewParam()
    elif char in operator-set:        cursor.emit(operator)
    else:                             accumulate into segment
  return root
```

### B.2 Shunting-yard pass (per node)

```text
toPostfix(node):
  out=[]; ops=[]
  for child in node.children:
    if child is sub-LexerNode: toPostfix(child); out.append(child)
    elif child is operand:     out.append(child)
    elif child is operator:
      while ops not empty and prec(ops.top) <= prec(child): out.append(ops.pop())
      ops.push(child)
  while ops: out.append(ops.pop())
  node.children = out
```

### B.3 RPN to AST

```text
parse(node):
  stack=[]
  for child in node.children:
    if child is operator: pop two, attach as left+right, push operator-node
    else: push leaf-node (Value/Reference/Function/...)
  return stack.top
```

### B.4 Evaluator

```text
evaluate(node, ctx):
  for c in node.children: evaluate(c, ctx)
  return node.compute(ctx)   # specific to node type
```

### B.5 Recalc loop

```text
recalc(dirtyRanges):
  dirty = depGraph.markDirty(dirtyRanges) ∪ volatileCells
  order = depGraph.topoOrder(dirty)
  for cellKey in order:
    ast = parsedAstFor(cellKey)
    value = evaluate(ast, ctx.at(cellKey))
    sheet.setValue(cellKey, value)
```

That's the whole engine in five tiny pseudo-functions. Univer's
1500-line `formula-dependency.ts` is mostly bookkeeping for features we
explicitly don't ship (shared-formula virtual trees, address-function
pre-eval, feature-formula injection, cooperative scheduling). The small
core is what we should aim for.
