# `/api/convert` — Office document conversion

`POST /api/convert` shells out to **LibreOffice headless** (`soffice
--headless --convert-to <ext>`) to produce PDF or HTML from a DOCX,
XLSX, or PPTX upload.

## Request

`multipart/form-data` with these fields:

| Field       | Type    | Required | Notes                                                                                                          |
| ----------- | ------- | -------- | -------------------------------------------------------------------------------------------------------------- |
| `file`      | Blob    | yes      | Source document bytes.                                                                                         |
| `sourceExt` | string  | yes      | One of `docx`, `xlsx`, `pptx`.                                                                                 |
| `targetExt` | string  | yes      | One of `pdf`, `html`.                                                                                          |
| `filename`  | string  | no       | Base filename for the response (no extension).                                                                 |
| `pageRange` | string  | no       | PDF only. LibreOffice page-range expression (`"1-3,5,7"`). Used for DOCX page ranges and PPTX single/multi-slide PDFs. |

Limits: input ≤ 50 MB, conversion ≤ 60 s wall-clock. `pageRange` is
validated against `/^\d+(-\d+)?(,\d+(-\d+)?)*$/` (whitespace allowed)
before being relayed to the appropriate `*_pdf_Export` filter.

## Response

- `200 OK` with the converted bytes, `Content-Type` set to
  `application/pdf` or `text/html; charset=utf-8`, and a
  `Content-Disposition: attachment; filename="…"` header.
- `400` for invalid input, `413` if the upload exceeds the size cap,
  `500` for spawn / conversion failures, `504` on timeout. Errors
  are JSON: `{ "message": "…" }`.

## Operational dependency

The host running the Next.js server must have **LibreOffice**
installed and `soffice` on `PATH`.

- macOS: `brew install --cask libreoffice` (and ensure
  `/Applications/LibreOffice.app/Contents/MacOS/soffice` is reachable
  via `PATH`, or set `OFFICEAI_SOFFICE_BIN`).
- Debian/Ubuntu: `sudo apt-get install -y libreoffice-core
  libreoffice-writer libreoffice-calc libreoffice-impress`.
- Alpine: `apk add --no-cache libreoffice` (in a container, also
  install fonts: `ttf-dejavu ttf-liberation fontconfig`).
- Containers: include LibreOffice in the production image. Cold
  starts will be slower than subsequent calls because LibreOffice
  builds a user profile on first run; the route uses a per-request
  profile dir to keep parallel requests safe.

### Environment variables

- `OFFICEAI_SOFFICE_BIN` — override the binary name / absolute path
  used to invoke LibreOffice. Defaults to `soffice`.

## Why not pure JS?

Client-side PDF libraries can't fully reproduce DOCX/XLSX/PPTX
layout (kerning, styles, formulas, slide masters). LibreOffice gives
us the same fidelity as "Save as PDF" inside Office at a known
operational cost.
