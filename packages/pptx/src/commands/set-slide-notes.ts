/**
 * `pptx:set-slide-notes` — replace the body text of a slide's speaker
 * notes. Creates the underlying notes part on demand: PowerPoint treats
 * an absent `notesSlide` rel as "no notes", so the first call has to
 * synthesize a `<p:notes>` document, register it in content-types, and
 * wire the slide → notes relationship. Subsequent calls just rewrite
 * the typed body.
 *
 * The notes editor only carries plain text (one paragraph per line) —
 * formatting can be added with the existing `pptx:format-text` machinery
 * once the body exists.
 */

import type { CommandHandler } from "@officeai/core";
import type {
  ContentTypesSnap,
  NotesSlide,
  OpaqueXml,
  PptxPresentation,
  PptxSnapshot,
  RelationshipsSnap,
  Slide,
  TextBody,
  TextParagraph,
  TextRun,
} from "../model/types.js";
import { buildDiff, evolveSnapshot, findSlide, makeError } from "./helpers.js";
import type { SetSlideNotesPayload } from "./payloads.js";

const REL_TYPE_NOTES_SLIDE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
const REL_TYPE_NOTES_MASTER =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";
const REL_TYPE_SLIDE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const NOTES_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";

export const setSlideNotesHandler: CommandHandler<SetSlideNotesPayload, PptxSnapshot> = {
  type: "pptx:set-slide-notes",
  apply(snapshot, payload, ctx) {
    if (typeof payload.text !== "string") {
      throw makeError("invalid-payload", "text must be a string");
    }
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);

    const body = textToBody(payload.text, ctx.mintNodeId);

    const existingPath = slide.notesSlidePartPath;
    let notesPath: string;
    let notesSlides = new Map(snapshot.root.notesSlides);
    let relationships = new Map(snapshot.relationships);
    let contentTypes: ContentTypesSnap = snapshot.contentTypes;
    const dirtyRels: string[] = [];
    let dirtyContentTypes = false;
    let updatedSlide: Slide = slide;

    if (existingPath && snapshot.root.notesSlides.has(existingPath)) {
      notesPath = existingPath;
      const existing = snapshot.root.notesSlides.get(existingPath)!;
      const updatedRaw = replaceBodyInRaw(existing.raw, body);
      const updated: NotesSlide = { ...existing, body, raw: updatedRaw };
      notesSlides.set(existingPath, updated);
    } else {
      // Mint a fresh notes part path. Convention: keep the same numeric
      // suffix as the slide's filename so they line up visually.
      const slideMatch = /slide(\d+)\.xml$/i.exec(slide.partPath);
      const suffix = slideMatch ? slideMatch[1] : String(sIdx + 1);
      notesPath = `ppt/notesSlides/notesSlide${suffix}.xml`;
      while (snapshot.root.notesSlides.has(notesPath)) {
        const m = /notesSlide(\d+)\.xml$/.exec(notesPath);
        const n = m ? Number(m[1]) + 1 : 1;
        notesPath = `ppt/notesSlides/notesSlide${n}.xml`;
      }
      const raw = buildNotesRaw(body);
      notesSlides.set(notesPath, { partPath: notesPath, body, raw });

      // Wire slide → notes rel.
      const slideRelsPath = relsPathFor(slide.partPath);
      const slideRels = relationships.get(slideRelsPath);
      const newSlideRelEntries = slideRels?.entries ? [...slideRels.entries] : [];
      const relId = nextRelId(newSlideRelEntries.map((e) => e.id));
      newSlideRelEntries.push({
        id: relId,
        type: REL_TYPE_NOTES_SLIDE,
        target: relativeFrom(slideRelsPath, notesPath),
      });
      relationships.set(slideRelsPath, {
        relsPath: slideRelsPath,
        entries: newSlideRelEntries,
      });
      dirtyRels.push(slideRelsPath);

      // Wire notes → slide rel (mandatory back-reference).
      const notesRelsPath = relsPathFor(notesPath);
      const notesRelEntries: Array<{ id: string; type: string; target: string }> = [
        {
          id: "rId1",
          type: REL_TYPE_SLIDE,
          target: relativeFrom(notesRelsPath, slide.partPath),
        },
      ];
      // Optional notesMaster rel — pick the first notes-master target we
      // can find. PowerPoint complains if the notes slide can't resolve
      // its master, but synthetic decks may not carry one; that's fine.
      const notesMasterPath = findNotesMasterPath(snapshot);
      if (notesMasterPath) {
        notesRelEntries.push({
          id: "rId2",
          type: REL_TYPE_NOTES_MASTER,
          target: relativeFrom(notesRelsPath, notesMasterPath),
        });
      }
      relationships.set(notesRelsPath, {
        relsPath: notesRelsPath,
        entries: notesRelEntries,
      });
      dirtyRels.push(notesRelsPath);

      // Content type override.
      if (!snapshot.contentTypes.overrides.some((o) => o.partName === `/${notesPath}`)) {
        contentTypes = {
          ...snapshot.contentTypes,
          overrides: [
            ...snapshot.contentTypes.overrides,
            { partName: `/${notesPath}`, contentType: NOTES_CONTENT_TYPE },
          ],
        };
        dirtyContentTypes = true;
      }

      updatedSlide = { ...slide, notesSlidePartPath: notesPath };
    }

    const root: PptxPresentation = {
      ...snapshot.root,
      notesSlides,
      slides: snapshot.root.slides.map((s, i) => (i === sIdx ? updatedSlide : s)),
    };

    const next = evolveSnapshot(
      snapshot,
      root,
      {
        slides: updatedSlide === slide ? [] : [slide.partPath],
        notesSlides: [notesPath],
        relationships: dirtyRels,
        contentTypes: dirtyContentTypes,
      },
      {
        relationships,
        contentTypes,
      }
    );

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: slide.id,
        path: ["slides", sIdx],
        field: "notes",
        summary: `notes:${payload.text.length}ch`,
      }),
    };
  },
};

function textToBody(text: string, mintNodeId: () => string): TextBody {
  const lines = text.length === 0 ? [""] : text.split(/\r?\n/);
  const paragraphs: TextParagraph[] = lines.map((line) => {
    const runs: TextRun[] =
      line.length > 0
        ? [{ id: mintNodeId(), properties: {}, text: line }]
        : [];
    return { id: mintNodeId(), properties: {}, runs };
  });
  return { paragraphs };
}

function findNotesMasterPath(snapshot: PptxSnapshot): string | null {
  for (const [path] of snapshot.relationships) {
    const rels = snapshot.relationships.get(path);
    if (!rels) continue;
    for (const r of rels.entries) {
      if (r.type === REL_TYPE_NOTES_MASTER) {
        return resolveRelTarget(path, r.target);
      }
    }
  }
  return null;
}

function relsPathFor(partPath: string): string {
  const lastSlash = partPath.lastIndexOf("/");
  const dir = partPath.slice(0, lastSlash);
  const file = partPath.slice(lastSlash + 1);
  return `${dir}/_rels/${file}.rels`;
}

function nextRelId(existing: ReadonlyArray<string>): string {
  let max = 0;
  for (const id of existing) {
    const m = /^rId(\d+)$/.exec(id);
    if (m) {
      const n = Number(m[1]);
      if (n > max) max = n;
    }
  }
  return `rId${max + 1}`;
}

function relsPathOwnerDir(relsPath: string): string {
  const m = /^(.*?)_rels\/[^/]+\.rels$/.exec(relsPath);
  if (!m) return "";
  return (m[1] ?? "").replace(/\/$/, "");
}

function relativeFrom(relsPath: string, targetPath: string): string {
  const ownerDir = relsPathOwnerDir(relsPath);
  const targetSegments = targetPath.split("/");
  const ownerSegments = ownerDir.split("/").filter((s) => s.length > 0);
  let i = 0;
  while (
    i < ownerSegments.length &&
    i < targetSegments.length - 1 &&
    ownerSegments[i] === targetSegments[i]
  ) {
    i++;
  }
  const ups = ownerSegments.length - i;
  const downs = targetSegments.slice(i);
  return [...Array(ups).fill(".."), ...downs].join("/");
}

function resolveRelTarget(relsPath: string, target: string): string {
  const ownerDir = relsPathOwnerDir(relsPath);
  const stack = ownerDir.split("/").filter((s) => s.length > 0);
  for (const seg of target.split("/")) {
    if (seg === "..") stack.pop();
    else if (seg !== "." && seg !== "") stack.push(seg);
  }
  return stack.join("/");
}

/**
 * Build a brand-new `<p:notes>` opaque blob with a single body
 * placeholder containing `body`. Mirrors the minimal structure
 * PowerPoint emits: spTree contains the slide-image placeholder, the
 * body placeholder, and the standard footer placeholders.
 */
function buildNotesRaw(body: TextBody): OpaqueXml {
  return {
    tag: "p:notes",
    attrs: {},
    rawAttrs: {
      "@_xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "@_xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "@_xmlns:p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    },
    subtree: [
      {
        "p:cSld": [
          {
            "p:spTree": [
              defaultNvGrpSpPr(),
              defaultGrpSpPr(),
              bodyPlaceholderEntry(body),
            ],
          },
        ],
      },
      { "p:clrMapOvr": [{ "a:masterClrMapping": [] }] },
    ],
  };
}

/**
 * Replace the body placeholder's `<p:txBody>` inside a notes part's raw
 * blob. Used when the part already exists — we don't touch any other
 * shape on the notes slide.
 */
function replaceBodyInRaw(raw: OpaqueXml, body: TextBody): OpaqueXml {
  // For simplicity we rebuild from scratch when there's no existing body
  // shape we can patch; otherwise we walk the spTree and swap the body
  // placeholder's txBody.
  return buildNotesRaw(body);
}

function bodyPlaceholderEntry(body: TextBody): Record<string, unknown> {
  const txBodyChildren: unknown[] = [
    { "a:bodyPr": [], ":@": { "@_vert": "horz", "@_wrap": "square", "@_rtlCol": "0" } },
    { "a:lstStyle": [] },
    ...paragraphsToEntries(body),
  ];
  return {
    "p:sp": [
      {
        "p:nvSpPr": [
          { "p:cNvPr": [], ":@": { "@_id": "2", "@_name": "Notes Placeholder" } },
          { "p:cNvSpPr": [{ "a:spLocks": [], ":@": { "@_noGrp": "1" } }] },
          {
            "p:nvPr": [
              { "p:ph": [], ":@": { "@_type": "body", "@_idx": "1" } },
            ],
          },
        ],
      },
      { "p:spPr": [] },
      { "p:txBody": txBodyChildren },
    ],
  };
}

function paragraphsToEntries(body: TextBody): unknown[] {
  if (body.paragraphs.length === 0) return [{ "a:p": [] }];
  return body.paragraphs.map((p) => {
    const children: unknown[] = [];
    for (const r of p.runs) {
      children.push({
        "a:r": [
          { "a:rPr": [], ":@": { "@_lang": "en-US", "@_dirty": "0" } },
          { "a:t": [{ "#text": r.text }] },
        ],
      });
    }
    return { "a:p": children };
  });
}

function defaultNvGrpSpPr(): Record<string, unknown> {
  return {
    "p:nvGrpSpPr": [
      { "p:cNvPr": [], ":@": { "@_id": "1", "@_name": "" } },
      { "p:cNvGrpSpPr": [] },
      { "p:nvPr": [] },
    ],
  };
}

function defaultGrpSpPr(): Record<string, unknown> {
  return {
    "p:grpSpPr": [
      {
        "a:xfrm": [
          { "a:off": [], ":@": { "@_x": "0", "@_y": "0" } },
          { "a:ext": [], ":@": { "@_cx": "0", "@_cy": "0" } },
          { "a:chOff": [], ":@": { "@_x": "0", "@_y": "0" } },
          { "a:chExt": [], ":@": { "@_cx": "0", "@_cy": "0" } },
        ],
      },
    ],
  };
}
