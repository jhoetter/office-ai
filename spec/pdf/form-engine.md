# PDF — Form Engine

> AcroForm read/write. XFA out of scope (AcroForm fallback only).
> Flatten semantics. Calc-order without JS execution.

Cross-references: typed `PdfFormField` in
[`document-model.md`](./document-model.md);
form-related agent commands and CLI invocations in
[`agent-commands.md`](./agent-commands.md) /
[`cli.md`](./cli.md);
edge cases in [`edge-cases.md`](./edge-cases.md).

## Scope

**In scope (P0):**

- AcroForm widget enumeration and value get/set.
- Validation: regex (`/V` regex constraints) + `/MaxLen`.
- Calc-order resolution **without JavaScript execution**.
- Flatten on save.
- Reset form.
- FDF/XFDF import/export.
- Signature-field detection + read-only display.
- Render the AcroForm fallback inside an XFA-bearing PDF; show a
  banner for pure-XFA forms.

**Out of scope:**

- XFA dynamic forms beyond the AcroForm fallback (Adobe is
  deprecating XFA).
- JavaScript actions in `/AA`, `/F`, document-level OpenAction, etc.
  These are sandboxed (no execution). Calc-order is resolved by
  static dependency analysis (see below).
- PKCS#12 cryptographic signing of signature fields. Visible
  signature image placement is in scope via the annotation pipeline;
  cryptographic signing is deferred.

## Field types

```typescript
export type PdfFormFieldType =
  | "text"      // /Tx
  | "checkbox"  // /Btn with /Ff bit 16 = checkbox
  | "radio"     // /Btn with /Ff bit 15 = radio
  | "choice"    // /Ch combo or list
  | "button"    // /Btn push-button
  | "signature" // /Sig
  | "unknown";
```

Mapping to PDF `/FT` and `/Ff` flag bits:

| Type        | `/FT` | `/Ff` flags                                    |
| ----------- | ----- | ---------------------------------------------- |
| `text`      | `/Tx` | `Multiline` (bit 13), `Password` (bit 14), `FileSelect` (bit 21), `DoNotSpellCheck` (bit 23), `RichText` (bit 26) |
| `checkbox`  | `/Btn`| `NoToggleToOff` (bit 15) for radios, `Radio` (bit 16) — checkbox = neither set |
| `radio`     | `/Btn`| `Radio` (bit 16) **set**, `RadiosInUnison` (bit 26) optional |
| `choice`    | `/Ch` | `Combo` (bit 18), `Edit` (bit 19), `MultiSelect` (bit 22), `DoNotSpellCheck` (bit 23) |
| `button`    | `/Btn`| `PushButton` (bit 17) |
| `signature` | `/Sig`| n/a |

`unknown` is a defensive fallback for malformed forms.

## Read path

`packages/pdf-forms` enumerates fields via the engine:

1. Walk `/AcroForm/Fields` recursively (terminal nodes are widgets;
   intermediate nodes are field-tree groupings carrying inheritable
   attributes).
2. For each terminal widget:
   - Resolve fully-qualified name `/T` (joined with parent `/T` via
     `.`).
   - Resolve `/V` (current value), `/DV` (default value).
   - Resolve `/Opt` (option list for `choice` and `radio`).
   - Resolve `/MaxLen`, `/Ff`.
   - Resolve `/P` (page reference) → `pageNumber`.
   - Resolve `/Rect`.
3. Project into `PdfFormField`.

Multi-widget fields (one field name, multiple widget annotations
across pages — common for "Sign here" repeated on every page) collapse
to **one** `PdfFormField` per name, with `pageNumber` + `rect` set
to the **first** widget. The set of all widgets is retained
internally for write-back.

## Write path

`pdf:fill-form` (typed command, see
[`agent-commands.md`](./agent-commands.md)) accepts a values object
keyed by fully-qualified field name:

```typescript
type FillFormPayload = {
  values: Record<string, string | boolean | string[]>;
  flatten?: boolean;
};
```

Write semantics per type:

- **`text`** — set `/V` to the value string. Re-emit appearance via
  `/DA` parsing → produce a fresh `/AP/N` content stream that draws
  the text using the same default appearance the field declared.
- **`checkbox`** — set `/V` to `/Yes` (true) or `/Off` (false), or
  the field's actual on-state name from `/AP/N` keys.
- **`radio`** — set the parent's `/V` to the on-state name of the
  selected child widget; siblings are implicitly off.
- **`choice`** — set `/V` to the selected option (or array for
  multi-select).
- **`button`** — push-buttons are not "filled"; the agent rejects
  attempts with `invalid-payload`.
- **`signature`** — read-only this session; the agent rejects writes
  with a clear "cryptographic signing deferred" error. A visible
  signature image at the field's `/Rect` can be added via
  `pdf:add-stamp` (an annotation, not a form fill).

Validation:

- **`/MaxLen`** — values longer than `/MaxLen` are rejected with
  `validation-failed`.
- **Regex** — fields with a regex constraint embedded in their `/V`
  formatter (Adobe-style format string) are validated. Regex syntax
  is the JavaScript subset.
- **Required** — empty string for a `/Required` field rejected with
  `validation-failed`.

## Calc-order resolution (without JS)

PDF's calculation order is declared in `/AcroForm/CO` (an array of
field references). For each field in `/CO`, `/AA/C` (calculate
action) optionally references a JS expression. **We do not execute
JS.** Instead:

- We **detect** dependent fields by parsing simple known patterns
  (`AFSimple_Calculate("SUM", ["a","b","c"])` style) — sum, product,
  min, max, average over a list of field names.
- Detected dependencies form a static DAG. After a fill, dependents
  recompute in topological order using their declared aggregation.
- Unrecognized JS calc actions are reported as
  `calc-order-skipped: <field name> <reason>` warnings on the
  `Mutation`. The fill still succeeds; the dependent fields keep
  their stale values. The `office-agent pdf inspect` output flags
  forms with unrecognized calc actions.

This handles 80% of real-world tax/admin forms (which use
`AFSimple_Calculate`) while never executing untrusted JS.

## Flatten on save

`flatten: true` (or `office-agent pdf flatten-form`):

1. For each field, render its current value into the page content
   stream at `/Rect`. Reuses the same AP-stream emit recipe as the
   write path, but writes directly into the page's content stream
   instead of into the field's `/AP/N`.
2. Remove the widget annotation from the page's `/Annots`.
3. Drop the field from `/AcroForm/Fields`.
4. If `/AcroForm/Fields` becomes empty, drop `/AcroForm` entirely
   from the catalog.

The output is visually identical but no longer fillable. This is the
canonical "lock the form" operation.

## Reset form

`pdf:reset-form` (or `office-agent pdf reset-form`):

For each field, set `/V` to `/DV` (or remove `/V` if no default).
Re-emit `/AP/N` from the default value. The form is restored to the
producer's initial state.

## FDF / XFDF import / export

```
office-agent pdf import-annotations --file f.pdf --annotations data.xfdf [--format xfdf|fdf|json] --out o.pdf
office-agent pdf export-annotations --file f.pdf [--format xfdf|fdf|json]
```

Same command as the annotation pipeline (FDF/XFDF carries both
annotations and form values, per Adobe's spec). Form-values entries
in an FDF/XFDF import populate the corresponding fields via the
`pdf:fill-form` handler internally.

## Signature fields

- Detected via `/FT /Sig`.
- `PdfFormField.type === "signature"` and `PdfFormField.readOnly === true`.
- The viewer displays a "Signature" badge with validity state from
  `/V/ByteRange` verification (delegated to `pdfjs-dist` /
  `@embedpdf/pdfium`'s built-in signature parser).
- The agent's `office-agent pdf list-signatures` outputs:

```json
[
  {
    "fieldName": "Sig1",
    "pageNumber": 1,
    "rect": [100, 200, 300, 250],
    "signedAt": "2026-04-19T18:32:00Z",
    "signedBy": "John Hoetter",
    "valid": true,
    "coversWholeDocument": true
  }
]
```

- Editing a signed PDF via incremental save preserves the signature.
  Editing via full re-serialize warns and breaks the signature
  intentionally — the `Mutation` carries
  `warning: "signature-broken"` so the UI can confirm.

## XFA fallback

XFA-bearing PDFs typically include an AcroForm fallback for non-Adobe
viewers. We:

- Render the AcroForm fallback if `/AcroForm/Fields` is non-empty.
  The user can fill, save, and round-trip.
- If only `/XFA` is present (pure-XFA), display a banner:
  **"This form was authored in XFA dynamic-forms format and requires
  Adobe Acrobat. The static layout is shown for reference only."**
  Field-level interaction is disabled.
- The `office-agent pdf list-form-fields` output flags pure-XFA forms
  with `"backing": "xfa-only"`.

## Round-trip guarantees

- Filled value via our writer → reopen in Adobe Acrobat → field shows
  the value with correct font / size / position.
- Filled in Adobe Acrobat → reopen in our viewer → value visible and
  editable.
- Round-trip a complex form (text + checkbox + radio + combo +
  multi-select listbox) byte-equality on every untouched object.
- A flattened form opens cleanly in Adobe / Preview / Chrome and is
  no longer reported as fillable.
