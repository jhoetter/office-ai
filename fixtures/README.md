# Fixture Matrix

`fixtures/MATRIX.json` is the canonical inventory for checked-in
DOCX/XLSX/PPTX/PDF fixtures. It records origin, license, complexity,
features, expected behavior and known risks for every real document file
under `fixtures/docx`, `fixtures/xlsx`, `fixtures/pptx` and
`fixtures/pdf`.

Run the local gate with:

```bash
pnpm fixtures:check
# or
make fixtures-check
```

The gate fails when:

- `fixtures/MATRIX.json` is missing or malformed.
- a fixture file is checked in but not indexed.
- an indexed file is missing, empty or under the wrong format root.
- a format has fewer than three simple and three complex fixtures.
- a format has no known-risk fixture that preserves opaque content or
  diagnoses unsupported structures.

Test code should select files from the matrix rather than hard-coding
ad hoc paths. The shared test helper is `tests/fixture-matrix.ts`; it is
used by roundtrip tests, MCP load-tool smoke tests and web open-file
smokes.

The per-format manifests remain useful for format-specific notes:

- DOCX: `fixtures/docx/MANIFEST.md`
- XLSX: `fixtures/xlsx/MANIFEST.md`
- PPTX: `fixtures/pptx/MANIFEST.md`
- PDF: `fixtures/pdf/README.md`
