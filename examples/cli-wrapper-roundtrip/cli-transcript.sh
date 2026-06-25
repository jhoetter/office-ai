#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
OUT="$ROOT/examples/_generated/cli-wrapper-roundtrip"
OA=(node "$ROOT/packages/agent/dist/cli.js" --deterministic-ids)

mkdir -p "$OUT"
export OFFICEAI_DATA_DIR="$OUT/data"

"${OA[@]}" docx read \
  --file "$ROOT/fixtures/docx/synthetic/03-headings-and-body.docx" \
  --format markdown > "$OUT/docx.md"

"${OA[@]}" xlsx read \
  --file "$ROOT/fixtures/xlsx/synthetic/01-single-sheet-numbers.xlsx" \
  --sheet Inventory \
  --range A1:D10 \
  --format json \
  --pretty > "$OUT/xlsx.json"

"${OA[@]}" pptx read \
  --file "$ROOT/fixtures/pptx/synthetic/03-title-and-content.pptx" \
  --format text > "$OUT/pptx.txt"

"${OA[@]}" pdf read-page \
  "$ROOT/fixtures/pdf/simple-text-1page.pdf" \
  --page 1 > "$OUT/pdf-page.json"

DOC_JSON="$("${OA[@]}" sessions import \
  --file "$ROOT/fixtures/docx/synthetic/03-headings-and-body.docx" \
  --json)"
DOC_ID="$(node -e 'const data=JSON.parse(process.argv[1]); console.log(data.document.documentId)' "$DOC_JSON")"

"${OA[@]}" sessions projection \
  --document-id "$DOC_ID" \
  --projection markdown > "$OUT/session-docx.md"

"${OA[@]}" sessions export \
  --document-id "$DOC_ID" \
  --out "$OUT/session-export.docx" \
  --json > "$OUT/session-export.json"

test -s "$OUT/session-export.docx"
