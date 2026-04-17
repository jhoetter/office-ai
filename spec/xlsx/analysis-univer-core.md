# Univer Core — Analysis (clean-room)

> Analyzed against upstream commit **`3d6f4182f2d2f0f9be155592d336b39eff06fdd7`** of [`dream-num/univer`](https://github.com/dream-num/univer) (Apache‑2.0). All references below are _file-path citations_ into a local clone at `/tmp/refs/univer`. **No source code was copied, lifted, or paraphrased into our codebase**; this document records architectural concepts only and is the basis for our own (independently-implemented) `@officeai/xlsx` design.

---

## 0. Repo shape (orientation)

The Univer monorepo is divided into ~60 `packages/*` workspaces. The relevant ones for spreadsheet core are:

- `packages/core` — DI container, command service, lifecycle, plugin registry, common shared utilities, **and the in-memory `Workbook` / `Worksheet` / `Range` / `Styles` model**. Despite its `core` name, the workbook data classes live here, not in `packages/sheets`.
- `packages/sheets` — sheet-aware _services_ and _controllers_: `SheetsSelectionsService`, `SheetInterceptorService`, `RefRangeService`, `SheetSkeletonService`, plus all **Commands & Mutations** that mutate worksheet data (`commands/commands/*`, `commands/mutations/*`, `commands/operations/*`).
- `packages/sheets-ui` — render bindings, DOM/canvas, keyboard. Out of scope for our headless build.
- `packages/sheets-formula` + `packages/engine-formula` — formula plumbing.
- `packages/sheets-conditional-formatting` — CF rules + view model.
- `packages/sheets-data-validation` (+ upstream `packages/data-validation`) — DV rule matrix.
- `packages/sheets-filter` — autofilter model.
- `packages/sheets-numfmt` — number-format application.

Univer's IoC container is the third-party library `@wendellhu/redi` (re-exported from `packages/core/src/common/di.ts`). Reactivity is `rxjs` `Subject` / `BehaviorSubject` everywhere.

There is **no XLSX import/export anywhere in the open-source repo** — the only "format" Univer speaks is its own JSON snapshot (`IWorkbookData`). This is significant for our design (see §7).

---

## 1. Workbook / Sheet / Cell / Range / Styles model

### 1.1 The "snapshot" / "model" duality

Univer keeps two parallel representations of every workbook:

1. A _plain-JSON snapshot_ — `IWorkbookData` — declared in `packages/core/src/sheets/typedef.ts`. This is the persistence/serialization shape and the input to the constructor. The same object is mutated in-place as commands run; calling `Workbook#save()` simply `Tools.deepClone`s it (`workbook.ts:132`).
2. A _runtime class wrapper_ — `Workbook`, `Worksheet`, `Range`, `Styles`, `RowManager`, `ColumnManager`, `SpanModel`, `SheetViewModel` — that holds the snapshot plus derived caches, observables, and methods. The class wrappers are constructed from the snapshot and **share live references with it**, so mutations on either side are visible on the other.

This is essentially "JSON snapshot is the source of truth; model classes are smart facades." There is no immutable / functional layer.

### 1.2 `Workbook` (`packages/core/src/sheets/workbook.ts`)

Key fields:

- `_snapshot: IWorkbookData` — the live JSON.
- `_unitId: string` — every Univer document (sheet/doc/slide) has a `unitId`. Workbooks live in `IUniverInstanceService` keyed by it.
- `_worksheets: Map<string, Worksheet>` — sheetId → worksheet wrapper. The map is rebuilt from `_snapshot.sheets` in `_parseWorksheetSnapshots()`.
- `_styles: Styles` — workbook-level shared style table (see §1.6).
- `_activeSheet$: BehaviorSubject<Nullable<Worksheet>>` — currently active sheet, exposed as RxJS observable.
- Lifecycle observables: `sheetCreated$`, `sheetDisposed$`, `name$`.
- `getRev() / incrementRev() / setRev()` — monotonically-increasing revision counter for collaboration.

`IWorkbookData` (`typedef.ts:29-83`) carries: `id`, `rev`, `name`, `appVersion`, `locale`, `styles` (the shared style dict), `sheetOrder: string[]`, `sheets: { [sheetId]: Partial<IWorksheetData> }`, optional `defaultStyle`, `resources` (plugin-owned blobs), `custom` (user metadata).

`Workbook` extends a generic base `UnitModel<TData, UNIVER_SHEET>` so docs/sheets/slides share lifecycle plumbing (see `packages/core/src/common/unit.ts`).

### 1.3 `Worksheet` (`packages/core/src/sheets/worksheet.ts`)

Holds:

- `_snapshot: IWorksheetData` — the per-sheet slice of JSON.
- `_cellData: ObjectMatrix<Nullable<ICellData>>` — sparse cell store (see §1.5).
- `_rowManager: RowManager`, `_columnManager: ColumnManager` — wrappers over `rowData` / `columnData` sparse maps.
- `_spanModel: SpanModel` — merged-cell index (built from `mergeData: IRange[]`).
- `_viewModel: SheetViewModel` — _interceptor-aware_ read façade (separate from raw cell access).
- `_isRowStylePrecedeColumnStyle: boolean` — precedence flag for composing inherited styles.

Important read-paths:

- `getCellRaw(row, col)` — bypasses interceptors and returns the raw `ICellData` from the matrix.
- `getCell(row, col)` — goes through `SheetViewModel`, which _fetches through interceptors_ registered by `SheetInterceptorService` (CF, NUMFMT, DV, hyperlinks, custom render…).
- `getComposedCellStyle(row, col, rowPriority?)` — composes default → row → column → cell → theme styles in either row-first or column-first order.

Interesting performance note (their own comment, `worksheet.ts:580-583`): `getCell` runs _all_ `CELL_CONTENT` interceptors and is explicitly flagged as a perf concern.

### 1.4 `Worksheet`'s sheet snapshot fields (`IWorksheetData`)

```
id, name, tabColor, hidden,
freeze: { xSplit, ySplit, startRow, startColumn },
rowCount, columnCount,
defaultColumnWidth, defaultRowHeight,
mergeData: IRange[],                                 // merged cells
cellData: IObjectMatrixPrimitiveType<ICellData>,     // sparse {row:{col:cell}}
rowData: IObjectArrayPrimitiveType<Partial<IRowData>>,
columnData: IObjectArrayPrimitiveType<Partial<IColumnData>>,
defaultStyle?, rowHeader, columnHeader,
showGridlines, gridlinesColor, rightToLeft,
custom?: CustomData
```

Row & column data are _sparse_ (`{ [index]: { h?, hd?, ah?, ia?, s?, custom? } }`). Width/height defaults live on the worksheet, only overridden rows/cols appear in the maps. Per-row and per-column styles are stored as a style id (string) or inline `IStyleData`.

### 1.5 Sparse cell representation — `ObjectMatrix`

Cells are stored as a _2-level sparse object_: `{ [rowIndex: number]: { [colIndex: number]: ICellData } }`. The wrapper class is `ObjectMatrix` (`packages/core/src/shared/object-matrix.ts`), which provides `getValue(r, c)`, `setValue`, `realDeleteValue`, `forValue` iteration, `getFragment(r1,r2,c1,c2)`, and bounds queries (`getRealRange`, `getStartEndScope`). Empty cells are simply absent keys.

This is the **only** representation — there is no 2D array fallback or column-major variant. Choosing nested objects keyed by string-coerced numbers means very cheap memory for sparse data (typical XLSX usage) but slow numeric iteration; the codebase explicitly notes "performance intensive" on iterators.

`ICellData` (`typedef.ts:239-283`) is intentionally minimal:

- `v` — origin scalar value (`string | number | boolean | null`).
- `t` — `CellValueType` (string / number / boolean / forceString).
- `s` — style: either a _string id_ into the workbook style table, or an _inline `IStyleData`_ object. Both are accepted everywhere.
- `f` — raw formula string (`=SUM(A1:B4)`).
- `si` — formula id (for shared/array formulas).
- `ref` — array-formula reference range.
- `xf` — Excel function prefix (`_xlfn.` etc) preserved verbatim — they keep this for round-trip.
- `p` — rich-text body (a `IDocumentData` document model, reusing the `@univerjs/core/docs` doc engine).
- `custom` — opaque user payload.

There's an enriched type `ICellDataForSheetInterceptor` (`typedef.ts:307-324`) that adds `interceptorStyle`, `markers`, `customRender[]`, `linkUrl`, `linkId`, `themeStyle`, `coverable`, `interceptorAutoHeight()`, etc. — these are _not_ persisted; they are produced on the fly when a renderer calls `getCell()`.

### 1.6 Styles — shared table with content-addressed dedup

`Styles` (`packages/core/src/sheets/styles.ts`) is a `Record<string, IStyleData>` plus an `LRUMap<string, string>` cache keyed by `JSON.stringify(style)`.

- `setValue(data)` → `JSON.stringify`, look up cache, then linear-scan `_styles` (`Tools.diffValue`) if cache misses, else allocate a 6-char random id and store. Returns the id.
- `addCustomStyle(id, data)` — keeps an existing id (used at load time so the author's existing IDs survive).
- `get(id)` returns the data; `getStyleByCell(cell)` understands both inline `cell.s = {…}` and id-based `cell.s = "abc123"`, plus folds in `interceptorStyle`.

So the style model is a **hybrid**: cells can either inline their style or reference the shared table; both are valid persisted forms. Composition (`composeStyles(default, col, row, themeStyle, cellStyle)`) handles cascading.

Caveat: the dedup `_getExistingStyleId` is `O(N)` per write — fine for small workbooks, painful at scale. Their LRU cache (capacity 100k) hides it for hot writes.

### 1.7 `Range` (`packages/core/src/sheets/range.ts`)

A thin façade around an `IRange` + `Worksheet` + a "deps" object that exposes `getStyles()`. Its only state is `{startRow, startColumn, endRow, endColumn}`. It provides Google-Sheets-style getters: `getValue()`, `getValues()`, `getBackgrounds()`, `getFontFamilies()`, `getA1Notation()`, etc. It does **not** mutate; mutation happens via Commands (§3).

`IRange` (`typedef.ts:481-487`) is the universal range struct everywhere in Univer:

```
{ startRow, startColumn, endRow, endColumn,
  rangeType?: RANGE_TYPE,             // NORMAL | ROW | COLUMN | ALL
  startAbsoluteRefType?, endAbsoluteRefType? }  // for $A$1 vs A1
```

`IRange` also optionally carries `unitId` and `sheetId` (via `IRangeLocation`); unscoped ranges are interpreted in the active sheet. Whole-row/column/all selections are encoded by a flag, not by setting columns to ±∞.

There are convenience aliases:

- `IGridRange = { sheetId, range }`,
- `IUnitRange = IGridRange & { unitId }`,
- `IUnitRangeName = { unitId, sheetName, range }` for cross-sheet references like `Sheet1!A1`.

### 1.8 Merged cells — `SpanModel`

`mergeData: IRange[]` on the snapshot is the source of truth. The runtime `SpanModel` (`packages/core/src/sheets/span-model.ts`, referenced from worksheet) builds an O(1) "is this cell inside a merge?" lookup. `Worksheet#getMergedCell(r,c)` returns the containing range or null; `getCellInfoInMergeData(r,c)` returns `{actualRow, actualColumn, isMerged, isMergedMainCell, startRow,…endColumn}` — the standard "main cell vs covered cell" duality.

When iterating ranges (`iterateByRow`/`iterateByColumn`), the iterator skips non-top-left cells of merges by default; only the top-left of a merged region carries a value.

### 1.9 Defined names

There is **no `DefinedName` type in `packages/core`** despite the user task mentioning it. Defined names are managed by the formula engine, in `packages/engine-formula/src/services/defined-names.service.ts`:

```
interface IDefinedNamesServiceParam {
  id: string;
  name: string;
  formulaOrRefString: string;     // e.g. "Sheet1!$A$1:$B$2" or "=SUM(...)"
  comment?: string;
  localSheetId?: string;          // workbook-scope vs sheet-scope
  hidden?: boolean;
  formulaOrRefStringWithPrefix?: string;  // Excel-compat _xlfn.* prefixed form
}
```

The service holds them per `unitId` and exposes Rx observables (`update$`, `focusRange$`). The data is also persisted into `IWorkbookData.resources` via the resource manager (see §2.4). Note the explicit Excel concession: `formulaOrRefStringWithPrefix` is kept around just for round-trip. This confirms Univer's I/O is designed to _route through their JSON_, with Excel-specific bits tucked into optional fields.

---

## 2. Plugin system

### 2.1 The `Plugin` base class

`packages/core/src/services/plugin/plugin.service.ts` defines the abstract `Plugin` (extends `Disposable`) with four lifecycle hooks:

- `onStarting()` — called when the plugin is instantiated and registered into the injector.
- `onReady()` — instances (workbook/doc) exist; controllers/services may now talk to them.
- `onRendered()` — first paint completed (only meaningful with UI plugins).
- `onSteady()` — all lazy work done.

Plus two _static_ fields used as identity:

```
static pluginName: string;
static packageName: string;          // from package.json
static version: string;              // from package.json
static type: UniverInstanceType;     // UNIVER_SHEET | UNIVER_DOC | UNIVER_SLIDE | UNIVER_UNKNOWN
```

A `@DependentOn(...PluginCtor[])` decorator stamps a `Symbol('DependentOn')` on the class so the registrar can do a topo sort. Example from `packages/sheets/src/plugin.ts:56`: `@DependentOn(UniverFormulaEnginePlugin)`.

### 2.2 Registration and lifecycle dispatch

`PluginService` (`plugin.service.ts:125`) holds:

- a `_pluginRegistry: Map<pluginName, {plugin: Ctor, options}>` of plugins waiting to be loaded,
- a `_pluginStore` of _instantiated_ plugins,
- `_seenPlugins` & `_loadedPlugins` for dedup,
- `_loadedPluginTypes: Set<UniverInstanceType>` to know which document-type runtimes are alive.

`registerPlugin(ctor, config)` does NOT instantiate immediately. It enqueues the plugin and, if a Univer instance of that type already exists, schedules a 4 ms macrotask (`_flushType`) to flush the queue. This lets users register plugins after `new Univer()` and have them activate when the matching `UniverSheet`/`UniverDoc` is created. Plugins typed `UNIVER_UNKNOWN` activate immediately and are global.

`_loadFromPlugins([…])` does a DFS topo sort honoring `DependentOnSymbol`, then calls `_initPlugin` (which uses `injector.createInstance(ctor, options)`). After all instantiation, `_pluginsRunLifecycle(plugins)` walks the _current_ lifecycle stage and any future stages by subscribing to `LifecycleService.lifecycle$`.

The lifecycle stages enum (`packages/core/src/services/lifecycle/lifecycle.ts`) is exactly:

```
Starting = 0, Ready = 1, Rendered = 2, Steady = 3
```

### 2.3 What a plugin actually does

A plugin's constructor typically:

1. Receives its config (typed) plus `@Inject(Injector)` plus `@IConfigService` etc.
2. Pushes its own config keys into `IConfigService` (for cross-plugin discovery).
3. Builds a `Dependency[]` array — each entry is a redi `[Identifier, {useClass | useValue | useFactory, deps}]` tuple — and passes it through `mergeOverrideWithDependencies(deps, this._config.override)`, then `registerDependencies(injector, deps)`.
4. Optionally calls `touchDependencies(injector, [[X], [Y]])` to _force eager instantiation_ of services that would otherwise be lazy.
5. Overrides `onStarting`/`onReady`/`onRendered` to touch additional dependencies _as their phase becomes valid_ (e.g. ranges-protection only at `Ready`).

`packages/sheets/src/plugin.ts` is the canonical example. It wires ~25 services and controllers — selections, ref-range, numfmt, permissions, sheet-interceptor, skeleton manager, range-themes, range-protection, defined names, etc. — and almost all of it is _just dependency wiring_; commands & mutations are registered separately by controllers.

### 2.4 Where commands/mutations actually get registered

Commands and mutations are **not** registered by the `Plugin` itself. The convention is:

1. The plugin loads a _controller_ (e.g. `BasicWorksheetController` in `packages/sheets/src/controllers/basic-worksheet.controller.ts`).
2. The controller is `@Inject(ICommandService)` dependent and, in its constructor, calls `commandService.registerCommand(SetRangeValuesCommand)`, `…(SetRangeValuesMutation)`, etc., adding the returned `IDisposable` to its own disposable bag.
3. When the controller is disposed (plugin unload), all its commands disappear too.

Resources (e.g. CF rules, DV rules, defined-names) are persisted via `IResourceManagerService` / `IResourceLoaderService`: a plugin calls `resourceManager.registerPluginResource({pluginName, businessName, hook})` with `toJson(unitId)` / `parseJson(json, unitId)` callbacks. At save time the manager iterates all hooks; at load it dispatches each resource blob back to the right plugin. This is how Univer carries plugin-owned data inside `IWorkbookData.resources` without `core` knowing about CF/DV/etc.

---

## 3. Command / Mutation / Operation pattern

This is the most opinionated piece of Univer's architecture and the part most worth borrowing.

### 3.1 The three command types

`packages/core/src/services/command/command.service.ts:38-56`:

- **COMMAND** — orchestration layer. It computes which mutations to dispatch based on params, pushes undo/redo into the stack, runs interceptors, calls `sequenceExecute`. Commands are the _only_ thing user code (UI, agent, REST) is supposed to call.
- **MUTATION** — atomic, _synchronous_, _deterministic_, **persisted-state-modifying** change. The smallest unit of conflict resolution for collaboration. A mutation handler is `(accessor, params) => boolean` and must be reproducible from `params` alone.
- **OPERATION** — like mutation, but the change is _not_ part of the persisted snapshot. Examples: scroll position, sidebar open/close, current selection. Operations are still command-bus-routed but never make it into the changeset.

This three-way split is the architectural backbone. Every action — `set-range-values`, `insert-row-col`, `set-style`, `clear-selection-content`, `set-defined-name`, etc. — is split into a command file (`commands/commands/set-range-values.command.ts`) and one or more mutation files (`commands/mutations/set-range-values.mutation.ts`).

### 3.2 The `ICommand` interface

```
interface ICommand<P, R = boolean> {
  id: string;                         // namespaced: 'sheet.command.set-range-values'
  type: CommandType;
  handler(accessor, params?, options?): Promise<R> | R;
}
```

Mutations narrow `type: CommandType.MUTATION` and require synchronous `handler`. Operations likewise.

`IMultiCommand` extends `ICommand` with `multi: true`, `priority: number`, and `preconditions(contextService): boolean`. Multiple implementations can register under the same id; at execution time the registry tries them in priority order until one returns truthy. This is how `Undo` and `Redo` get specialized for sheet vs doc vs editor contexts.

### 3.3 Dispatch surface — `ICommandService`

```
hasCommand(id): boolean
registerCommand(cmd): IDisposable
registerMultipleCommand(cmd): IDisposable
unregisterCommand(id)
executeCommand<P, R>(id, params?, options?): Promise<R>     // async
syncExecuteCommand<P, R>(id, params?, options?): R          // sync, throws if handler is async
beforeCommandExecuted(listener): IDisposable
onCommandExecuted(listener): IDisposable
onMutationExecutedForCollab(listener): IDisposable          // collab-only sink
```

`IExecutionOptions` carries `onlyLocal`, `fromCollab`, `fromChangeset`, `syncOnly`. The flags decide whether listeners fire and whether the execution is forwarded to remote peers.

`syncOnly: true` is a clever escape hatch: the mutation is _not_ run locally (already applied) but listeners fire so the collab client still sends it to the server.

A `_commandExecutionStack` tracks nested command/mutation calls; when a `MUTATION` runs, the service finds the most recent `COMMAND` on the stack and stamps `params.trigger = thatCommandId`. This lets mutation listeners attribute changes to the originating user intent — useful for telemetry, undo grouping, or ref-range follow-ups.

### 3.4 How a Command composes Mutations — concrete walk

Worth quoting at the structural level (no code copying):

`SetRangeValuesCommand.handler` (`packages/sheets/src/commands/commands/set-range-values.command.ts`):

1. Resolve target `{workbook, worksheet, unitId, subUnitId}` from params (or active selection).
2. Snapshot current selections (used later for undo restoration).
3. Normalize `value` (single cell / 2D array / `IObjectMatrixPrimitiveType`) into a single `cellValue` matrix.
4. Build `setRangeValuesMutationRedoParams = {unitId, subUnitId, cellValue}`.
5. Compute the inverse via `SetRangeValuesUndoMutationFactory(accessor, redoParams)` which reads current cells + styles and synthesizes the _exact_ mutation params that would restore them.
6. **Dry-run the redo**: `commandService.syncExecuteCommand(SetRangeValuesMutation.id, redoParams)`. If false, abort.
7. Ask `SheetInterceptorService.onCommandExecute({id, params})` for `{undos, redos, preUndos, preRedos}` from every registered interceptor (CF, NUMFMT, DV may all want to add follow-up mutations).
8. Ask `SheetInterceptorService.generateMutationsOfAutoHeight(...)` for row-height adjustments.
9. `sequenceExecute([...redos, ...autoHeightRedos], commandService)` to run the side-effect mutations.
10. Push to undo stack: `undoRedoService.pushUndoRedo({unitID, undoMutations: [setRangeValuesUndo, ...undos, ...autoHeightUndos, restoreSelections], redoMutations: [setRangeValuesRedo, ...redos, ...autoHeightRedos, followSelection]})`.

The mutation itself (`set-range-values.mutation.ts`) is _pure_: read the workbook by `unitId`, walk the matrix, deep-merge old & new cell data, write back. No undo computation, no broadcasting, no interceptors.

This split — **commands compute "what to do + how to undo + side-effects"; mutations just apply** — is the key invariant. It enables: undo/redo, OT-style collab replay, command logging, and command interception by _other_ features without those features touching the mutation handler.

### 3.5 Undo / redo stack

`packages/core/src/services/undoredo/undoredo.service.ts`:

- `IUndoRedoItem = { unitID, undoMutations: IMutationInfo[], redoMutations: IMutationInfo[], id? }`.
- Stacks are _per `unitId`_ and per "focus context" (sheet, formula-bar editor, normal editor — chosen by `_getFocusedUnitId()` reading `IContextService`).
- `STACK_CAPACITY = 20` (a fixed window — old items get spliced off the front).
- Pushing a new undo clears the redo stack.
- `UndoCommand` and `RedoCommand` are both registered as plain `COMMAND`s in the registry; their handler pulls the top item and calls `sequenceExecute(item.undoMutations, commandService)`.
- A `__tempBatchingUndoRedo(unitId)` returns a disposable that batches subsequent pushes into a single item until disposed (their workaround for find-and-replace).

Mutations are serializable by construction (the params are JSON-shaped), so a recorded `IMutationInfo[]` can be replayed verbatim on another peer to reproduce the change. That's how their collab plugin works (not in OSS, but the seam is here: `onMutationExecutedForCollab` listener and `IExecutionOptions.fromCollab` / `fromChangeset` flags).

### 3.6 Interceptor system (`SheetInterceptorService`)

A meta-feature on top of commands that lets _features_ hook into _other features'_ commands without coupling. `packages/sheets/src/services/sheet-interceptor/sheet-interceptor.service.ts`:

- `interceptCommand({priority, getMutations(info)})` — feature returns `{undos, redos, preUndos, preRedos}` to inject around the originating command's mutations.
- `interceptBeforeCommand({priority, performCheck(info)})` — async veto. All checks must resolve true.
- `intercept(INTERCEPTOR_POINT.CELL_CONTENT, {priority, effect, handler(value, ctx)})` — a _read-side_ interceptor: the worksheet's `SheetViewModel.getCell()` is composed of these in priority order. `effect` is a bitfield (`Style | Value`) so style-only and value-only readers don't pay for the others.
- `INTERCEPTOR_POINT.ROW_FILTERED` — read-side, used by autofilter to hide rows.
- `writeCellInterceptor` — `BEFORE_CELL_EDIT`, `AFTER_CELL_EDIT`, `VALIDATE_CELL` for editor flows.

This is how data validation, conditional formatting, number format, hyperlinks, and theme styles all _layer_ on top of raw cell data without modifying it: they intercept reads and contribute follow-up mutations to writes.

---

## 4. Dependency injection

Univer uses **`@wendellhu/redi`** (re-exported via `packages/core/src/common/di.ts`). It is a TypeScript decorator-based DI library similar in shape to Angular / VS Code's IoC.

### 4.1 Pattern

- **Token-based identifiers**: `createIdentifier<T>('univer.core.command-service')` mints an `IdentifierDecorator<T>`. The convention is to declare both `interface ICommandService { … }` and `const ICommandService = createIdentifier<ICommandService>(…)` with the same name; the decorator is used at injection sites (`@ICommandService private readonly _commandService`).
- **Constructor injection** with `@Inject(Class)` for concrete classes and `@IIdentifier` for token-typed services. Optional / many / lookup variants exist (`@Optional`, `Many`, `LookUp`, `Self`, `SkipSelf`).
- **Class registration**: `injector.add([Identifier, {useClass, deps?}])`, `[Identifier, {useValue}]`, or `[Identifier, {useFactory, deps}]`.
- **Resolution**: `injector.get(Identifier)`. Inside a command handler, the framework passes an `IAccessor` (a thin wrapper providing `accessor.get(Identifier)`) so handlers can do scoped lookups without grabbing the injector directly.
- **Eager/lazy**: services are lazy by default; `touchDependencies(injector, [[X],[Y]])` forces instantiation.

### 4.2 Injector hierarchy

`new Univer(config, parentInjector?)` (`packages/core/src/univer.ts:117`) creates a child injector when a parent is provided. Plugins all add into the same root injector — there's only one DI scope per Univer instance, no per-plugin sub-scopes by default.

### 4.3 Cross-plugin resolution

Because there's one root injector and tokens are global (`createIdentifier` strings must be unique), any plugin can `@IFoo` any other plugin's service. The convention is:

- `core` exposes pure-token services (`ICommandService`, `IUndoRedoService`, `IUniverInstanceService`, `IConfigService`, `IContextService`, `ILogService`, `IPermissionService`, `IResourceManagerService`, `IResourceLoaderService`).
- Each feature package declares its own tokens (`INumfmtService`, `IAutoFillService`, `IExclusiveRangeService`) and registers concrete classes against them.
- `mergeOverrideWithDependencies(deps, override)` lets the user replace any registered class via `IUniverConfig.override`, which is the official way to swap out, e.g., the local undo-redo service for a collaborative one.

---

## 5. Selection / range model

### 5.1 Selection at the data layer

`packages/sheets/src/services/selections/selection.service.ts` defines `SheetsSelectionsService` (per Univer instance) and delegates per-workbook state to `WorkbookSelectionModel` (`selection-data-model.ts`).

A selection is `ISelectionWithStyle` (`packages/sheets/src/basics/selection.ts:153`):

```
{
  range: IRange,                                  // rectangle / row / col / all
  primary: Nullable<ISelectionCell>,              // the highlighted "anchor" cell, with merge info
  style?: Partial<ISelectionStyle>                // colors/widths for rendering
}
```

`primary` is an `ISelectionCell` carrying `{actualRow, actualColumn, isMerged, isMergedMainCell, startRow, startColumn, endRow, endColumn}` so the caller knows how to highlight the anchor when it's inside a merged region.

Multi-selection is just `ISelectionWithStyle[]`. Cross-sheet selection is not a single object — `WorkbookSelectionModel` keeps a `Map<sheetId, ISelectionWithStyle[]>` so each sheet retains its own last selection. Reactivity is exposed through `selectionMoveStart$ / selectionMoving$ / selectionMoveEnd$ / selectionSet$ / selectionChanged$` (RxJS), with a `SelectionMoveType` enum (`MOVE_START | MOVING | MOVE_END | ONLY_SET`) so listeners can filter pointer-driven events from programmatic ones.

There is a separate `RefSelectionService` (extends `SheetsSelectionsService`) used while a formula is being authored, so reference picking doesn't pollute the main selection.

### 5.2 R1C1 vs A1

Internally, **everything is row/column integer indices** (0-based). `IRange` only ever speaks `{startRow, startColumn, endRow, endColumn}`. There is no R1C1 mode or A1 mode at the data layer.

A1 notation is a _string format_ used at the boundaries:

- `Range#getA1Notation()` produces `"B2:D4"` from indices via `Tools.numToWord(col+1)`.
- Formula reference parsing/serialization lives in `packages/engine-formula/src/engine/utils/reference.ts` (`handleRefStringInfo`, `serializeRange`). The formula engine accepts A1 strings (`"Sheet1!$A$1:$B$2"`), parses them into `IUnitRangeName`/`IRange`, then operates in indices.
- `AbsoluteRefType` (`NONE | ROW | COLUMN | ALL`) is stored on `IRange` so the absoluteness (`$A$1` vs `A1`) survives a round-trip through indices.

R1C1 is not natively modeled at all in core (consistent with Excel where R1C1 is a display preference, not a storage format).

---

## 6. Skeleton / data-model layering

Univer **does** keep a strict separation between data and rendering geometry, and they call this layer the "skeleton."

Two skeleton classes exist:

1. `SheetSkeleton` (`packages/core/src/sheets/sheet-skeleton.ts`) — the _headless geometry_ layer. Owns:
   - `_rowHeightAccumulation: number[]` and `_columnWidthAccumulation: number[]` (prefix sums for fast hit-testing via `searchArray`).
   - `_rowTotalHeight`, `_columnTotalWidth`, `_rowHeaderWidth`, `_columnHeaderHeight`.
   - Scroll & scale (`_scrollX/Y`, `_scaleX/Y`).
   - Methods: `getCellWithCoordByIndex(r, c)`, `getCellIndexByOffset(x, y, scaleX, scaleY, scrollXY)`, `expandRangeByMerge(range)`, `getOffsetByColumn(c)`, `getOffsetByRow(r)`, `_updateLayout()`.
   - A "gap" runtime concept (`ISheetGapConfig`) for visual row/column separators that don't affect data.
     It depends on a `Worksheet` (data) and `Styles` (shared styles) plus `ILocaleService`, `IContextService`, `IConfigService`, and `Injector`. **It computes everything needed to map between (row, col) and (x, y) without a renderer.**

2. `SpreadsheetSkeleton` in `packages/engine-render` (referenced from `packages/sheets/src/skeleton/skeleton.service.ts`) — the _rendering_ skeleton, extends or wraps `SheetSkeleton` and adds canvas-specific text shaping, font caches, image cache, etc. This is what `SheetSkeletonService` instantiates, keyed by `(unitId, subUnitId)`, lifecycle-tied to `Workbook#sheetCreated$/sheetDisposed$`.

For our headless build the lower `SheetSkeleton` (or rather, the _concept_ of it) is what matters: it shows that even a renderless XLSX core needs a layer for hit-testing, range expansion through merges, and width/height accumulation independent of the cell data. They co-locate it in `core` rather than `sheets-ui` precisely so it's reusable headless.

---

## 7. Sheet I/O — XLSX import/export

**There is none in the open-source repo.** Searching the entire monorepo for `xlsx`, `XLSX`, `SheetJS`, `jszip`, `fast-xml-parser`, `import.*xlsx`, `export.*xlsx` finds only:

- README banners describing Univer's _commercial_ import/export server.
- `docx` references in unrelated files.
- Hyperlink filename matchers like `*.xlsx`.

Univer's open-source data path is **`IWorkbookData` JSON only**: `new Workbook(workbookData)`, `workbook.save()`, `workbook.load(config)`. Round-trip with `.xlsx` is offered by their non-OSS server-side service `dream-num/server` and is opaque from this codebase.

What this _implies_ about their design:

- `IWorkbookData` is treated as the canonical storage format. They preserve Excel-isms (`xf`, `_xlfn.` formula prefixes via `formulaOrRefStringWithPrefix`, etc.) inside their JSON but do **not** keep the original OOXML parts.
- Plugin-owned data lives inside `IWorkbookData.resources: { [pluginName]: jsonString }`, populated by `IResourceLoaderService` at load and emitted by hooks at save. There is no concept of "preserve unknown OOXML parts": anything Univer doesn't model is lost on the JSON ↔ XLSX boundary.
- This is a major divergence point for our design (see §10): an OOXML-truth core that _parses XLSX directly into typed AST_ and re-serializes, preserving unknown XML, is a fundamentally different stance.

---

## 8. Conditional formatting, data validation, autofilter

All three follow the same architectural pattern: **a per-(`unitId`, `subUnitId`) model holding rules + a registered `INTERCEPTOR_POINT.CELL_CONTENT` interceptor that mutates the cell display on read**, plus a controller that listens to relevant mutations and calls `RuleMatrix.apply…` to translate range edits into rule edits.

### 8.1 Conditional formatting

`packages/sheets-conditional-formatting/src/models/conditional-formatting-rule-model.ts`:

- `_model: Map<unitId, Map<subUnitId, IConditionFormattingRule[]>>`. List ordering = priority.
- `addRule` unshifts (newest = highest priority).
- `_ruleChange$: Subject<{rule, oldRule?, unitId, subUnitId, type: 'delete'|'set'|'add'|'sort'}>` is the reactive surface.
- A separate `ConditionalFormattingViewModel` caches computed cell decorations.

Rules are persisted via the resource manager (one JSON blob per unit, namespaced by plugin name). Re-application after edits goes through `SheetInterceptorService.interceptCommand` for set-range-values etc., plus the read-side cell interceptor that overlays bg/cl/border on `getCell()`.

### 8.2 Data validation

`packages/sheets-data-validation/src/models/sheet-data-validation-model.ts`:

- `_ruleMatrixMap: Map<unitId, Map<subUnitId, RuleMatrix>>`. **`RuleMatrix`** is interesting — it stores a per-cell mapping from `(row, col)` to a `ruleId` so range-membership lookups are O(1). Re-application after row/col insertion is a `RuleMatrix` update, not a re-evaluation of every rule's range.
- Underlying rules live in the cross-format `DataValidationModel` from `@univerjs/data-validation` (shared with docs, in principle).
- `DataValidationCacheService` and `DataValidationFormulaService` cache validation results; `DataValidationCustomFormulaService` runs custom-formula validators.
- The status surface is `validStatusChange$: Subject<IValidStatusChange>`.

### 8.3 Autofilter

`packages/sheets-filter/src/models/filter-model.ts`:

- One `FilterModel` per `(unitId, subUnitId)`.
- Holds `_range: IRange | null` (the auto-filter ref), `_filterColumnByIndex: Map<number, FilterColumn>`, and `_alreadyFilteredOutRows: Set<number>` (cached predicate result).
- `filteredOutRows$: BehaviorSubject<Set<number>>` is the read-side surface; `INTERCEPTOR_POINT.ROW_FILTERED` reads from it.
- Per-column `FilterColumn` instances hold either `filters: IFilters` (value-list), `customFilters: ICustomFilters` (operator/value pairs), or `colorFilters: IColorFilters`. They serialize back to `IAutoFilter`.

For all three: range-edits (insert/delete row/col) are translated into rule-range-edits by listening for the relevant mutations via `RefRangeService` (`packages/sheets/src/services/ref-range/ref-range.service.ts`), which is the central hub for "I have this `IRange` registered; tell me how it transforms when row/col mutations fire." `RefRangeService` provides `registerRefRange(range, callback)` and emits `EffectRefRangeParams` describing how to adjust.

---

## 9. What we should KEEP

Concepts to import into our `@officeai/xlsx` design (we already have a docx-side `CommandBus` and `MutationStore` — these slot in directly):

1. **The Command / Mutation / Operation tri-split** — exactly what we're already doing for `@officeai/docx`. Confirms our architecture. Our naming aligns naturally: `Command` ≈ user-intent orchestrator, `Mutation` ≈ smallest persisted change, and we should add `Operation` for ephemeral state (selection, viewport, focus) so it shares the same dispatch surface but never enters the changeset/undo stack.
2. **Mutation handlers must be pure & sync.** Univer enforces this in `_syncExecute` (throws if a mutation handler returns a Promise). We should do the same — it's the precondition for collab, replay, and reliable undo.
3. **`IRange` as the universal range struct**, with an explicit `rangeType: NORMAL|ROW|COLUMN|ALL` flag (so whole-row/whole-column don't need sentinel ±∞ values) and `AbsoluteRefType` per endpoint to preserve `$A$1` semantics across index-based storage.
4. **Sparse cell storage as the first-class model.** `{[row]:{[col]:cell}}` (their `ObjectMatrix`) maps to OOXML's `<row r="…"><c r="…">` perfectly and is memory-cheap for the typical empty-cell distribution. We should use the same.
5. **Style table with content-addressed dedup + inline-or-id cell style.** Cells reference styles by id where possible, but inline styles are still legal — this matches OOXML's `cellXfs` indirection while letting agents write a cell without first allocating a style.
6. **Undo-redo as ordered list of `IMutationInfo[]`** (their `IUndoRedoItem`). Both undo and redo are _just_ mutation lists; the undo command sequence-executes them. Trivially serializable, replayable, collab-friendly.
7. **The factory-pattern for inverses** (`SetRangeValuesUndoMutationFactory`). Each command pairs a redo mutation with a synthesized undo mutation read from current state. We should adopt this idiom verbatim.
8. **Per-`unitId` undo/redo stacks** rather than one global stack — important once we support multiple open workbooks.
9. **Token-based DI with constructor injection.** We don't have to use `@wendellhu/redi`, but the _shape_ (token = `createIdentifier<T>('namespace.name')`, services accept `IAccessor`, eager-vs-lazy distinction) is the right shape.
10. **`SheetInterceptorService`-style read interceptors with `effect: Style | Value` bitfields**, so style-only readers don't pay the cost of value computation. Critical for headless contexts where an agent might only need styles for "explain this sheet" or only need values for formula evaluation.
11. **Plugin lifecycle stages** (`Starting / Ready / Rendered / Steady`) — even headless we want at least `Starting` (DI wiring) and `Ready` (workbook exists), and probably a `Steady` for lazy resource loading.
12. **`@DependentOn` topo sort on plugin registration** — we'll have feature plugins (formula, CF, DV, filter, hyperlink, table) and the order matters.
13. **`INTERCEPTOR_POINT.CELL_CONTENT` as the unified "read-side overlay" mechanism** — DV markers, CF backgrounds, hyperlink underlines, number-format display strings all belong here, not in the persisted cell.
14. **`RefRangeService` pattern** — centralized "this range moves when these mutations fire" registry, so every feature with a range (CF, DV, filter, named range, table, comment) plugs into one place rather than each feature listening to every row/col mutation.
15. **The two-skeleton split** (`SheetSkeleton` for headless geometry; rendering skeleton for canvas/DOM). We should have an analogous geometry layer in `@officeai/xlsx` so server-side code can hit-test, expand-by-merge, and compute print ranges without dragging in a renderer.
16. **Resource manager for plugin-owned data** — feature plugins persist their state into a namespaced `resources` blob on the workbook snapshot; core knows nothing about them. This is exactly how XLSX namespaces let parts coexist; we get the same property within our JSON (and our XLSX) by adopting it.

---

## 10. What we should DIFFER ON

Where Univer's choices conflict with our headless-first / agent-first / OOXML-truth stance:

1. **OOXML is the source of truth, not JSON.** Univer's `IWorkbookData` is the canonical persistence format and XLSX is a lossy import/export layer (in their commercial server, not even in OSS). We invert that: parse XLSX into a typed AST that preserves _everything_ — unknown elements, attributes, parts, namespaces — and serialize back so the bytes round-trip even when we don't model a feature. This is non-negotiable for an agent that may edit a 10-year-old auditor's workbook with custom XML, theme1.xml customizations, embedded objects, VBA, etc.
2. **No `IDocumentData` for rich text in cells.** Univer reuses their docs engine to model rich text (`ICellData.p: IDocumentData` is a full doc model). For us this is overkill: an XLSX cell's rich text is a flat `<r>` run sequence. Model it directly as `RichRun[]` with style references; don't drag a doc engine into the sheets package.
3. **Mutations should be _typed discriminated unions_, not loose `{id: string, params: unknown}` records.** Univer's `IMutationInfo<T>` is generic but the registry is keyed by string id with `unknown` params at the boundary. Our type system should make `Mutation` a TS union so a `switch (m.kind)` on the consumer side is exhaustive (we already enforce this in `@officeai/docx`).
4. **No `multi-command` priority chain** (their `IMultiCommand` for undo/redo per context). Headless-first means we don't have a "current focus"; the caller specifies the workbook and sheet explicitly. Replace `IMultiCommand` with explicit per-`unitId` dispatch.
5. **No reactive `BehaviorSubject` proliferation in core.** Univer wires `rxjs` everywhere for UI consumption. For us, prefer a thin event emitter / async iterator surface that's optional. UIs that want Rx can wrap; agents and servers shouldn't have to import it.
6. **Don't conflate "active sheet" with the data model.** Univer's `Workbook#getActiveSheet()` throws if no active sheet — fine for a UI, wrong for a headless agent who legitimately wants to read all sheets without "activating" any. Make active-sheet a property of an _editor session_, not the workbook.
7. **Style dedup must be O(log n) or hash-indexed**, not Univer's linear scan. With a 50k-row workbook authored by a CRM export, you'll allocate millions of style writes; their `LRUMap<JSON-string, id>` is a good cache, but the fallback `_getExistingStyleId` is `O(N×styles)`. Use a hash of normalized style → id (mirror OOXML's `cellXfs` index discipline).
8. **No `_xlfn.` / `xf` field smuggling.** Univer keeps Excel-specific bits as optional fields on `ICellData` (`xf`, `si`, `ref`) to survive the JSON↔XLSX trip. With OOXML-truth we don't need those: the formula AST and array-formula context are first-class, not stringly-typed.
9. **Plugin lifecycle should not fire `onRendered` / `onSteady`.** A headless build doesn't render. Reduce to `onStarting → onReady → onDisposing`. Feature plugins that _want_ a "lazy / steady" hook can opt in via an explicit timer or idle callback.
10. **`SheetSkeleton` has 1100+ lines and depends on `Injector`, `LocaleService`, `ThemeService`.** Way too much coupling for a geometry layer. Our equivalent should be a **pure function** — given `(rowData, colData, defaults, mergeData, rowGaps, colGaps)`, return `{rowAccum, colAccum, hitTest, expandByMerge, …}`. No DI, no theme, no locale.
11. **No `Operation` type for selection.** Selection in Univer is its own service that fires its own observables, but it's also routed through `SetSelectionsOperation` so it participates in undo. We should think harder: if the goal is "agent-driven", selection probably _isn't_ an `Operation` at all and shouldn't share the command bus. Editor-mode selection is a UI concern; agent edits address ranges directly via params.
12. **Defined names belong in `core`, not the formula engine.** Univer puts `IDefinedNamesService` in `engine-formula` (likely because the formula engine needs it for evaluation). For us: the workbook owns named ranges / named formulas as first-class entities (they live in OOXML's `definedNames` element under `workbook.xml`); the formula engine consumes them. This avoids forcing a "to use named ranges, install the formula plugin" coupling.
13. **No `__interceptViewModel` private callback hack.** Univer uses an underscored private method on `Worksheet` for `SheetInterceptorService` to wire itself in at construction time. With proper DI we instead have the worksheet _expose_ a `viewModel: SheetViewModel` property and let the interceptor service register against it via the standard surface — no underscored back-doors.

---

## 11. What we should IMPROVE on Univer's design

Gaps where an agent-native, OOXML-truth core can credibly do better:

1. **First-class machine-readable command schemas.** Univer's commands have TypeScript-typed params, but there's no JSON schema, no manifest, no introspection API ("what commands exist, what params, what does each do?"). For an LLM agent we should ship a generated `commands.json` describing every command, its params, examples, preconditions, side-effect mutations, and idempotence/inverse properties. `@officeai/docx` already does this via `commands/registry.ts` + a generated payload manifest — we extend the same pattern.
2. **Round-trip-preserving I/O.** Parse XLSX into a typed AST that retains: unknown XML elements (in a `__unknown` slot per node), namespace prefixes, attribute order where mandated, embedded binary parts (vba, drawings, charts), and content-types. Re-serialization should yield byte-identical-or-semantically-equivalent output for unmodified workbooks. This is the headline differentiator.
3. **Deterministic, hashable mutations.** Each mutation should serialize to a canonical JSON form with a content hash. Replay logs become git-like; collab becomes possible without the bespoke peer-discovery layer Univer needs in their commercial product.
4. **Pure-function style table.** Instead of `Styles.add(data) → random 6-char id`, use `styleHash(normalizedStyle) → id`. Two independent agents writing the same style produce the same id. Eliminates merge conflicts on the style table during collab.
5. **Versioned snapshot format.** Univer's `IWorkbookData.appVersion` is a string but there's no schema versioning, no migration runner. We should commit to an explicit `schemaVersion: number` plus a registered `Migration` chain.
6. **Cell-value coercion as a separate `valueType` system.** Univer has `CellValueType` (string|number|boolean|forceString) but coercion logic is sprinkled through `mergeCellData` / `getCellValue`. Make it a single typed module (`coerceCellValue(input, locale, numFmt) → {v, t}`) usable from commands, formula results, and import.
7. **Range-edit propagation via a generic `RangeIndex`.** Univer's `RefRangeService` is great but is sheets-only. Promote it to a generic `IndexedRanges<T>` that any feature (CF, DV, filter, named range, hyperlink, comment, table, drawing-anchor) registers ranges into and gets back transformed ranges on row/col mutations. Single index, queryable, batched.
8. **Skeleton/geometry as a pure function**, not a class. `computeLayout({rowData, colData, defaults, merges}) → Layout` (immutable struct). Eliminates the `dirty` flag, the `setScale/setScroll` methods, the locale/config dependencies, and makes layout safely shareable across workers.
9. **Native support for partial / streaming load.** Univer constructs a full `Workbook` synchronously from `IWorkbookData`. For large XLSX (gigabyte audit exports) we should support incremental sheet load (sheet metadata first, cell matrix on demand) and lazy materialization of `cellData[sheetId]`.
10. **Agent-native preconditions on commands.** Univer's `IBeforeCommandInterceptor.performCheck` is async-boolean. We should richer-type it: `precheck(cmd) → {ok: true} | {ok: false, reason, suggestedFix?}` so an agent gets back not just "no" but a structured explanation it can act on.
11. **Explicit "merge mode" in the type system.** Univer's `mergeCellData` walks keys with `overwriteCellPropertiesSet: Set<string>` to decide which fields replace vs deep-merge. Encode this in the type: `Mutation.kind === 'set-range-values'` carries `mode: 'replace' | 'patch' | 'clear'` so semantics are obvious to LLMs and reviewers.
12. **Drop the "interceptors mutate read results" pattern in favor of a layered view model.** Univer's CELL_CONTENT interceptor has every reader pay the cost of every interceptor. A layered model — base cell + overlay (CF result) + overlay (DV markers) + overlay (numfmt display) — composed lazily and cached per-(`row,col,layerSet`) is faster and more predictable.
13. **Treat formulas as ASTs in the cell, with the source string as a derived projection.** Univer stores `f: string` and parses on every formula-engine touch. A parsed AST in the cell (with the original source preserved verbatim for round-trip) avoids re-parsing and lets the agent reason structurally about formulas without owning a parser.
14. **Ship a deterministic test harness for mutations.** Each mutation should have a `generate(params) → {redo, undo}` contract that's property-tested: applying redo then undo to any starting state yields the starting state. Univer has unit tests but no enforced inverse property.

---

## Appendix — file-path index for reviewers

For each section above, the primary Univer files consulted (so a reviewer can verify the interpretation):

| Topic                                                  | Univer files                                                                                  |
| ------------------------------------------------------ | --------------------------------------------------------------------------------------------- |
| Workbook                                               | `packages/core/src/sheets/workbook.ts`                                                        |
| Worksheet                                              | `packages/core/src/sheets/worksheet.ts`                                                       |
| Cell / Range / IRange / IWorkbookData / IWorksheetData | `packages/core/src/sheets/typedef.ts`                                                         |
| Range façade                                           | `packages/core/src/sheets/range.ts`                                                           |
| Styles table                                           | `packages/core/src/sheets/styles.ts`                                                          |
| Sparse cell matrix                                     | `packages/core/src/shared/object-matrix.ts`                                                   |
| Merged cells                                           | `packages/core/src/sheets/span-model.ts`, `worksheet.ts:469-575`                              |
| View model / interceptor read                          | `packages/core/src/sheets/view-model.ts`                                                      |
| Skeleton (geometry)                                    | `packages/core/src/sheets/sheet-skeleton.ts`                                                  |
| DI re-exports                                          | `packages/core/src/common/di.ts` (re-exports `@wendellhu/redi`)                               |
| Univer entry                                           | `packages/core/src/univer.ts`                                                                 |
| Plugin base & service                                  | `packages/core/src/services/plugin/plugin.service.ts`                                         |
| Lifecycle                                              | `packages/core/src/services/lifecycle/lifecycle.ts`, `lifecycle.service.ts`                   |
| Command service                                        | `packages/core/src/services/command/command.service.ts`                                       |
| Undo/redo                                              | `packages/core/src/services/undoredo/undoredo.service.ts`                                     |
| Sheet plugin (canonical wiring)                        | `packages/sheets/src/plugin.ts`                                                               |
| Sheet interceptor service                              | `packages/sheets/src/services/sheet-interceptor/sheet-interceptor.service.ts`                 |
| Interceptor points                                     | `packages/sheets/src/services/sheet-interceptor/interceptor-const.ts`                         |
| Selection service                                      | `packages/sheets/src/services/selections/selection.service.ts`, `selection-data-model.ts`     |
| Selection types                                        | `packages/sheets/src/basics/selection.ts`                                                     |
| Ref-range propagation                                  | `packages/sheets/src/services/ref-range/ref-range.service.ts`                                 |
| Skeleton service (rendering)                           | `packages/sheets/src/skeleton/skeleton.service.ts`                                            |
| Example command                                        | `packages/sheets/src/commands/commands/set-range-values.command.ts`                           |
| Example mutation                                       | `packages/sheets/src/commands/mutations/set-range-values.mutation.ts`                         |
| Conditional formatting model                           | `packages/sheets-conditional-formatting/src/models/conditional-formatting-rule-model.ts`      |
| Data validation model                                  | `packages/sheets-data-validation/src/models/sheet-data-validation-model.ts`, `rule-matrix.ts` |
| Filter model                                           | `packages/sheets-filter/src/models/filter-model.ts`                                           |
| Defined names                                          | `packages/engine-formula/src/services/defined-names.service.ts`                               |

---

_End of analysis. No code was copied; all observations are derived from reading the upstream files at the commit pinned above._
