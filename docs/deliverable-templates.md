# Deliverable templates

OfficeAI synthesis deliverables are template-driven at the MCP boundary.
Templates are data contracts: they declare a target format, slots they
can consume, required slots, and whether they are active or handoff-only.
They do not hard-code a Sonaloop dependency; any MCP client can supply
the same payload.

## Catalogue

Use:

```text
list_deliverable_templates
```

Active templates:

- `research_report` (`docx`): report with title, brand, summary,
  sections, findings, quotes, table summaries, assets and provenance.
- `executive_deck` (`pptx`): short deck built from PowerPoint layout
  placeholders; slides cover title, summary, key findings, signals and
  artifacts.
- `analysis_workbook` (`xlsx`): workbook with a `Synthesis` sheet and
  one row per payload item.

Handoff-only template:

- `pdf_review_handoff` (`pdf`): contract for annotating an existing PDF.
  Native PDF creation from synthesis text is intentionally deferred until
  PDF text-page authoring exists.

## Fallbacks

`create_deliverable_from_synthesis` accepts `template_id`. When the
requested template cannot produce one of the requested target formats,
OfficeAI falls back to the format default and emits
`template-fallback`.

Missing optional slots emit `template-slot-empty`. Extra pass-through
payload fields emit `template-payload-extra` so callers can see what was
preserved but not consumed by the selected template.

## Brand Data

Brand inputs are data:

```json
{
  "brand": {
    "name": "Care Ops",
    "accentColor": "2F6F73",
    "logoAssetId": "asset-logo"
  }
}
```

Templates may surface these values in generated documents and diagnostics.
They are not compiled assumptions; a different client can swap them per
call without code changes.
