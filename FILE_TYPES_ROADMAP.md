# File Types Roadmap

> Strategic brainstorm: which file types should `office-ai` support next, and _why_?
>
> **Context.** We already ship first-class, AI-native packages for **DOCX, XLSX, PPTX, and PDF**. The product is positioned as an **S3-backed, AI-native alternative to SharePoint** for SMEs. SharePoint's moat is not its editors — Microsoft's editors are objectively excellent. SharePoint's moat is _being the place where every business file lives_. To displace it (or even credibly sit beside it) we have to handle **the long tail of file types SMEs actually deal with every day**, not just the Office four.
>
> This document captures the candidate file types, grouped by strategic value, with the reasoning behind each. It is meant to drive prioritisation discussions, not to be a commitment.

---

## Guiding principles

Before the list, the lenses we use to rank a format:

1. **Frequency in the SME workday.** Does the average 5–200 person business touch this file weekly? Daily? Hourly?
2. **Pain with incumbents.** Is SharePoint / OneDrive / Google Drive / Dropbox _bad_ at this format today? Pain = wedge.
3. **AI leverage.** Does this format _unlock_ something that is structurally impossible without AI (transcription, OCR, semantic search, summarisation)? AI-shaped problems are where we can build a moat that Microsoft's 30-year editor lead doesn't matter.
4. **Adjacency to what we already have.** Does it pull users deeper into our existing DOCX/XLSX/PPTX/PDF surface area? E.g. opening an attached invoice from an `.eml` directly in our PDF editor.
5. **Build cost vs. payoff.** A `.txt` viewer is trivial; a `.dwg` viewer is a quarter of engineering. Cost matters.
6. **Trust / "everything just opens".** Some formats we support not because they're glamorous, but because _not_ supporting them makes the product feel like a toy. (`.zip`, `.html`, `.log` previews.)

---

## Tier 1 — Build next: high value, high frequency, AI-shaped

These are the formats where we believe there is a real wedge against SharePoint _today_, and where AI gives us an unfair advantage.

### 1. Email — `.eml`, `.msg`, `.mbox`, `.pst`

**Why this is #1.**

Email is the _real_ document management system of every SME on earth. Contracts, signed offers, NDAs, invoices, supplier confirmations, HR conversations, board updates — they all live as `.msg` files dragged into Outlook folders or forwarded into shared mailboxes. SharePoint is famously poor at email: `.msg` previews are inconsistent, attachment extraction is clunky, full-text search across body + attachments is unreliable, and threading is non-existent.

**What we'd build.**

- Parse `.eml` (RFC 5322) and `.msg` (Microsoft CFB / MAPI) into a normalised model: headers, body (text + html), attachments, inline images.
- Render a faithful viewer (header block, threaded conversation if `In-Reply-To` chains exist).
- Extract attachments and route them into the right editor: `.docx` → DOCX editor, `.pdf` → PDF editor, `.xlsx` → XLSX editor. **This is where the magic happens** — email becomes the entry point that pulls users into the rest of our suite.
- Index body + attachments for semantic search ("find the email where the supplier confirmed the new pricing").
- AI summary per thread, AI-extracted action items, AI-extracted entities (amounts, dates, parties).

**Why now.**

- Zero serious competition in the SME segment that combines `.msg` parsing + AI summarisation + S3 storage.
- Pulls value from every other package we already have.
- Migration story is easy: "drag your Outlook archive folder in."

---

### 2. Images — `.png`, `.jpg`/`.jpeg`, `.heic`/`.heif`, `.webp`, `.tiff`/`.tif`, `.gif`, `.svg`, `.bmp`

**Why.**

Receipts, ID scans, whiteboard photos, product shots, screenshots, logos, signatures, before/after photos for trades businesses, damage photos for insurance claims — SMEs _generate_ images constantly and have nowhere good to put them. SharePoint treats them as opaque blobs.

**HEIC specifically matters** because every iPhone since 2017 produces them by default, and the Microsoft / Windows ecosystem still trips over them regularly. Owning HEIC well is a small, concrete differentiator.

**What we'd build.**

- Universal viewer with format normalisation (HEIC/HEIF/TIFF → web-safe internally).
- OCR pass on every uploaded image, indexed for full-text search.
- AI captioning + tagging for semantic search ("find the photo of the whiteboard from the June offsite", "all receipts from Lufthansa").
- Lightweight editing: crop, rotate, redact, annotate. Nothing Photoshop-class — just enough that users don't have to leave.
- EXIF-aware: location, timestamp, camera. Useful for trades, real estate, insurance.

**Why now.** OCR + captioning is _the_ canonical AI-leverage workflow. We already have the AI plumbing from the existing packages.

---

### 3. Plain text family — `.txt`, `.md`, `.markdown`, `.rtf`, `.rst`, `.adoc`

**Why.**

Markdown is now the lingua franca of:

- AI-native workflows (system prompts, agent instructions, knowledge bases).
- Modern engineering / dev / consulting SMEs (READMEs, runbooks, ADRs).
- Meeting notes (Granola, Reflect, Obsidian, Notion exports, Cursor / Claude transcripts).

`.rtf` is the format every legacy Windows app exports to. `.txt` is universal. Cost to support is near-zero; perceived completeness gain is huge.

**What we'd build.**

- A high-quality Markdown editor with live preview, math, mermaid, tables. (We already have the frontend chops from the DOCX editor.)
- AI: rewrite, summarise, expand bullet → prose, prose → bullets, translate.
- Round-trip with DOCX (markdown ↔ docx is a well-trodden path; pandoc-style).

**Why now.** Trivial cost, high "this is my new Notion / my new Obsidian / my new Granola dump" potential. Pairs naturally with the email package (drop a `.md` summary of an email thread).

---

### 4. Tabular data files — `.csv`, `.tsv`, `.jsonl`, `.ndjson`, `.parquet`

**Why.**

These sit awkwardly between "spreadsheet" and "data". Every SaaS export, every Stripe report, every Shopify dump, every CRM extract is a CSV. SharePoint shows them as raw text. Excel mangles them on open (date coercion, scientific notation on long IDs, encoding bugs).

A great CSV experience is a **bridge into the XLSX product**: open a CSV, optionally promote to XLSX, edit, save back as either.

**What we'd build.**

- Streaming viewer (handles million-row CSVs without dying).
- Smart type inference with override (avoid the "Excel turned my SKU into a date" disaster).
- Filter / sort / search / pivot in the browser.
- "Open in XLSX editor" hand-off.
- Parquet for the more data-literate SMEs (small analytics teams, data-heavy startups).
- AI: "explain this column", "find anomalies", "summarise this dataset", "convert to chart".

**Why now.** Cheap to build on top of XLSX. Closes a real, daily, painful gap.

---

## Tier 2 — Build soon: high value, narrower or heavier

These are formats with strong demand but either narrower personas or higher build cost. They become Tier 1 the moment we get a beachhead customer who needs them.

### 5. Audio / video — `.mp3`, `.m4a`, `.wav`, `.ogg`, `.flac`, `.mp4`, `.mov`, `.webm`, `.mkv`

**Why.**

Voice memos, Zoom / Teams / Meet recordings, sales calls, training videos, customer support recordings, podcasts. SharePoint treats these as 200 MB blobs. The actual _value_ of these files is the _transcript_, the _summary_, the _searchable knowledge_ inside them — and that is impossible without AI.

This is potentially **a wedge product on its own**, the way Otter, Fireflies, Granola, and Fathom have built standalone businesses just on call recordings. We can offer it as a feature of the file system, not a separate $25/seat tool.

**What we'd build.**

- Transcription (speaker-diarised) on upload.
- Chapter detection, action item extraction, AI summary.
- Searchable transcripts with click-to-play timestamps.
- Lightweight trim / clip-export.
- Integration with the email package (auto-transcribe `.m4a` voice memos that arrive as attachments).

**Why now-ish.** Higher infra cost (GPU / transcription bills), but the differentiation is enormous. Likely a paid premium tier.

---

### 6. Calendar & contacts — `.ics`, `.ifb`, `.vcf`

**Why.**

Tiny formats, but extremely "businessy". Every meeting invite that lands in an inbox is an `.ics`. Every contact card shared from a phone is a `.vcf`. Rendering them properly makes the suite _feel_ like a real workspace, not a glorified file viewer.

**What we'd build.**

- Render `.ics` as a real calendar event card with RSVP state.
- Round-trip: edit and re-export.
- Bulk import `.vcf` to a contacts directory.
- AI: "schedule a follow-up two weeks after this", "summarise the agenda".

**Why now-ish.** Low effort, high "feels like a complete product" payoff.

---

### 7. Archives — `.zip`, `.7z`, `.rar`, `.tar`, `.tar.gz`, `.tgz`

**Why.**

"Client sent me a zip of 80 PDFs" is a weekly SME scenario. Today the workflow is: download, unzip locally, re-upload individual files, lose the structure. A browse-in-place archive viewer is a small piece of work with a large quality-of-life payoff.

**What we'd build.**

- In-browser tree view of archive contents without extraction.
- Stream individual files into the right editor (PDF → PDF editor, etc.).
- "Extract to folder" into the user's S3.
- Optional AI: "summarise the 80 PDFs in this zip".

**Why now-ish.** Real value, modest cost.

---

## Tier 3 — Domain-specific, but extremely sticky in their domain

These are not for everyone. But the moment a target SME _is_ in one of these domains, the format becomes a **non-negotiable**. We won't win architecture firms without `.dwg`. We won't win marketing agencies without `.psd`/`.ai`/`.fig`.

We pick these on a per-vertical basis when we have a beachhead.

### 8. CAD & engineering drawings — `.dwg`, `.dxf`, `.dwf`, `.stp`/`.step`, `.iges`, `.stl`, `.3mf`

- **Who:** architecture, construction, manufacturing, industrial design, prototyping.
- **Why:** SharePoint is brutal here — these files are heavy, version-sensitive, and reference-laden. A web viewer + version diff + AI bill-of-materials extraction would be a real product.
- **Cost:** high. Likely partner with an existing renderer (Autodesk Forge / equivalent) rather than build.

### 9. Design files — `.psd`, `.ai`, `.indd`, `.fig`, `.sketch`, `.xd`, `.afdesign`, `.afphoto`

- **Who:** marketing agencies, brand teams, in-house design.
- **Why:** these are the most-shared, least-previewable files in any agency Drive/SharePoint. Even just **rendering a faithful preview without Photoshop** is valuable. AI: "extract palette", "extract copy", "find all assets using this logo".
- **Cost:** medium-to-high (especially `.psd`/`.ai`).

### 10. Geospatial — `.kml`, `.kmz`, `.geojson`, `.gpx`, `.shp`, `.gpkg`

- **Who:** logistics, surveying, agritech, field services, real estate, environmental.
- **Why:** rendering on a map + simple analysis (area, distance, intersection) is tractable. Zero incumbent in the SME file-system space.
- **Cost:** medium (mapbox/maplibre + a parser).

### 11. Code & config — `.py`, `.ts`, `.js`, `.go`, `.rs`, `.java`, `.sql`, `.yaml`, `.toml`, `.json`, `.env`, `.ipynb`

- **Who:** small dev shops, technical consultancies, data teams.
- **Why:** syntax-highlighted preview + AI explain is table stakes for a technical SME.
- **Special case:** `.ipynb` (Jupyter) is a _huge_ unlock for anyone with a data team — render notebooks with outputs, run AI explanations on cells. Pairs naturally with `.csv`/`.parquet`.
- **Cost:** low for highlighting/preview, medium for executable notebooks.

### 12. eSignature & forms

- `.pdf` with AcroForms / XFA fields (we already touch PDF — extending into form-fill is small).
- DocuSign / Adobe Sign envelopes (`.docx` with content controls, signed PDFs with embedded certs).
- **Why:** signature workflows are the #1 reason SMEs pay for a separate SaaS (DocuSign, PandaDoc). Even a basic "sign here / send for signature" inside the PDF editor would prevent users from leaving.
- **Cost:** medium. Real legal/compliance care needed (eIDAS, ESIGN Act).

### 13. Accounting & banking exchange — `.qbo`, `.qfx`, `.ofx`, `.iif`, `.csv`-as-bank-export, `.xml` (UBL invoices, ZUGFeRD/Factur-X hybrid PDFs), `.edi`

- **Who:** every SME's bookkeeper.
- **Why:** the European e-invoicing mandates (ZUGFeRD/Factur-X, PEPPOL, X-Rechnung) are arriving country by country. A file-system that natively understands "this PDF is also a structured invoice" is genuinely novel.
- **Cost:** medium. Schema-heavy but well-documented.

---

## Tier 4 — Boring but trust-building: viewers only

We don't build editors here. We build _viewers_ so users never see a "cannot preview this file" message. The strategic value is **the absence of friction** — users trust that "office-ai opens everything".

- **`.html` / `.htm`** — sandboxed preview of saved web pages, exported reports.
- **`.xml`** — pretty-printed, schema-aware where possible (e.g. invoice schemas).
- **`.log` / `.out` / `.err`** — tail / search / AI-summarise. Surprisingly high value for technical SMEs.
- **`.epub` / `.mobi`** — internal training material, ebooks. Cheap.
- **`.dmg`, `.iso`, `.exe`, `.pkg`** — _don't try to open_, just store + label safely.
- **`.psd` thumbnail / `.ai` thumbnail** — even a thumbnail is better than nothing in the non-design-agency case.
- **Font files (`.ttf`, `.otf`, `.woff2`)** — preview the typeface. Cheap, charming.

---

## Anti-list — things we should _not_ chase (yet)

To be disciplined, here is what we explicitly de-prioritise:

- **Niche scientific / research formats** (`.fits`, `.mat`, `.nc`, `.h5`) — wrong audience for SME.
- **Proprietary game / 3D engine formats** — wrong audience.
- **Legacy Microsoft formats older than `.docx`/`.xlsx`/`.ppt`** (`.wpd`, `.pub`, `.mdb`) — long tail, low frequency, high parser pain. Convert-on-import only.
- **Encrypted containers we can't unlock without user keys** (`.kdbx`, `.gpg`, `.age`) — store, don't decrypt. Privacy story matters more than features here.
- **Anything that requires shipping a desktop runtime to render faithfully** — breaks the "browser-first, S3-backed" thesis.

---

## Recommended sequencing

Given our existing surface (DOCX / XLSX / PPTX / PDF) and our positioning (S3-backed, AI-native, SME-focused), the recommended order is:

1. **Email (`.msg` / `.eml`)** — biggest single wedge against SharePoint, pulls users into the existing editors via attachment hand-off. _This is the next package._
2. **Images with OCR + AI captioning** — broad, daily, AI-leveraged, cheap-ish.
3. **Markdown / plain text** — trivial cost, big completeness perception.
4. **CSV / TSV / JSONL** — natural extension of XLSX, closes a real pain.
5. **Audio / video with transcription** — likely a paid premium tier; a wedge product on its own.
6. **Archives + Calendar / Contacts** — quality-of-life, builds trust.
7. **Vertical bets** (CAD, design, geospatial, accounting/e-invoicing) — driven by the first paying beachhead in each vertical, not speculatively.
8. **Trust-building viewers** (HTML, XML, logs, EPUB, fonts) — bundled in throughout, never as a standalone milestone.

The unifying story across all of these: **office-ai is not a collection of editors. It is the place where every business file lives — and every file is alive, because every file is understood by AI.**
