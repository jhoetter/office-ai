# Shared OOXML Utilities

> The plumbing every format uses: zip container, XML parse/serialize,
> namespace handling, relationship graphs, byte-preservation.

## OoxmlContainer

```typescript
export interface OoxmlPart {
  /** Path inside the zip, e.g. "word/document.xml". */
  readonly path: string;
  /** Raw bytes, as loaded. Used for byte-identical re-emit. */
  readonly bytes: Uint8Array;
  /** Whether the part has been marked dirty by a command/serializer. */
  readonly dirty: boolean;
}

export class OoxmlContainer {
  /** Load a zip from an ArrayBuffer; index every part. */
  static load(buffer: ArrayBuffer): Promise<OoxmlContainer>;

  /** All parts, keyed by zip path. */
  readonly parts: ReadonlyMap<string, OoxmlPart>;

  /** Read a part as text; throws if missing. */
  readText(path: string): string;
  /** Read a part as raw bytes. */
  readBytes(path: string): Uint8Array;
  /** Replace a part. Marks it dirty. */
  writeText(path: string, content: string): void;
  writeBytes(path: string, content: Uint8Array): void;
  /** Add a new part (e.g. an inserted image). */
  addPart(path: string, content: Uint8Array, contentType: string): void;
  /** Remove a part. Marks the [Content_Types].xml dirty. */
  removePart(path: string): void;

  /** Re-emit the container as a zip ArrayBuffer. Untouched parts are byte-identical. */
  serialize(): Promise<ArrayBuffer>;

  /** SHA-256 of a part's current bytes. Used for snapshot.partHashes. */
  hash(path: string): string;
}
```

## XML utilities

We use [`fast-xml-parser`](https://github.com/NaturalIntelligence/fast-xml-parser)
(MIT) with these settings to preserve enough of the original document to
keep Word happy:

```typescript
const parserOptions = {
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  allowBooleanAttributes: true,
  parseTagValue: false, // never coerce "1" to 1; OOXML attrs are strings
  parseAttributeValue: false,
  preserveOrder: true, // CRITICAL: child order matters
  trimValues: false, // CRITICAL: whitespace can be significant in <w:t>
  unpairedTags: ["w:br", "w:tab", "w:cr"],
};
const builderOptions = {
  ...parserOptions,
  format: false, // do not pretty-print
  suppressEmptyNode: false, // keep <w:rPr/> when empty
  suppressBooleanAttributes: false,
};
```

```typescript
export function parseXml(xml: string): XmlTree;
export function serializeXml(tree: XmlTree, opts?: { xmlDeclaration?: string }): string;

/**
 * `XmlTree` is the array-of-objects shape produced by fast-xml-parser when
 * preserveOrder=true. We use it directly rather than projecting to a custom
 * tree type — it round-trips losslessly for any namespace that doesn't use
 * mixed content, which covers all of WordprocessingML.
 */
export type XmlTree = unknown;
```

## Namespace handling

OOXML parts declare namespaces on the root element. We preserve the original
declarations verbatim. When a serializer re-emits a part, it copies the
original root attributes (`xmlns:w`, `xmlns:r`, `mc:Ignorable`, etc.) before
rebuilding children.

## Relationship graph

```typescript
export interface Relationship {
  readonly id: string; // "rId7"
  readonly type: string; // "http://.../officeDocument/2006/relationships/hyperlink"
  readonly target: string; // "https://example.com" or "comments.xml"
  readonly targetMode?: "External" | "Internal";
}

export class RelationshipGraph {
  static loadFor(container: OoxmlContainer, partPath: string): RelationshipGraph;

  readonly relationships: ReadonlyArray<Relationship>;
  byId(id: string): Relationship | undefined;
  byType(type: string): ReadonlyArray<Relationship>;

  /** Allocate a fresh rId that doesn't collide with existing ones. */
  mintId(): string;
  add(rel: Omit<Relationship, "id"> & { id?: string }): Relationship;

  /** Re-emit to its part path. */
  writeBack(container: OoxmlContainer): void;
}
```

## Byte-preservation guarantee

The `OoxmlContainer.serialize()` contract:

1. For every part where `dirty === false`, the output bytes are
   **bit-identical** to the input bytes.
2. For every part where `dirty === true`, the output is the new bytes
   provided via `writeText`/`writeBytes`.
3. The zip archive itself uses **DEFLATE level matching the input** when
   possible; otherwise level 6. Note that Word does not check zip-level
   compression, so this is a best-effort property.
4. The `[Content_Types].xml` is **never edited** unless a part has been
   added or removed (or replaced with new content type).
5. The order of parts in the central directory follows the original order
   when possible, with new parts appended.

## Tests

The container ships with a self-test: load a fixture, serialize without any
edits, parse the output again, and assert every part is byte-identical to
the original.
