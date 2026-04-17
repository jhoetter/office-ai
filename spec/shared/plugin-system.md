# Shared Plugin System

> How features register themselves with the bus and the agent without
> hard-coding `if format === "docx"` branches into the core.

## Goal

Each format package (DOCX/XLSX/PPTX) registers a **format plugin** with the
core. The plugin advertises:

- which command types it handles
- which read operations it exposes (`getRange`, `search`)
- which export formats it supports

A consumer of `@officeai/core` needs to know **nothing** format-specific
beyond instantiating one plugin per format it cares about.

## Types

```typescript
export interface FormatPlugin<TSnapshot = unknown> {
  /** "docx" | "xlsx" | "pptx" — disjoint per plugin. */
  readonly format: "docx" | "xlsx" | "pptx";
  /** Handlers this plugin contributes to the command bus. */
  readonly handlers: ReadonlyArray<CommandHandler<unknown, TSnapshot>>;
  /** Parser: bytes -> initial snapshot. */
  parse(buffer: ArrayBuffer): Promise<TSnapshot>;
  /** Serializer: snapshot -> bytes. */
  serialize(snapshot: TSnapshot): Promise<ArrayBuffer>;
  /** Optional: project the snapshot to a plain-text or markdown digest for agents. */
  toMarkdown?(snapshot: TSnapshot): string;
}

export interface PluginRegistry {
  register<T>(plugin: FormatPlugin<T>): void;
  get(format: "docx" | "xlsx" | "pptx"): FormatPlugin<unknown>;
}
```

## Registration

```typescript
const registry = new PluginRegistry();
registry.register(docxPlugin);
// registry.register(xlsxPlugin); // deferred
// registry.register(pptxPlugin); // deferred
```

The `DocumentAgent` factory looks up the plugin by format and wires it into
a fresh `CommandBus`:

```typescript
const agent = createAgent("docx", await fs.readFile("doc.docx"), registry);
```

## Anti-rules

- Plugins **do not** import from each other. DOCX never imports from XLSX.
- The core **does not** import from any format package.
- A plugin's handlers receive `TSnapshot` typed to that plugin's snapshot
  shape; the registry erases this to `unknown` at storage time and re-types
  at retrieval time.
