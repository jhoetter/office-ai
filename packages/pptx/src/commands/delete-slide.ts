import type { CommandHandler } from "@officeai/core";
import type { PptxSnapshot } from "../model/types.js";
import { buildDiff, evolveSnapshot, findSlide } from "./helpers.js";
import type { DeleteSlidePayload } from "./payloads.js";

const PRES_RELS_PATH = "ppt/_rels/presentation.xml.rels";

export const deleteSlideHandler: CommandHandler<DeleteSlidePayload, PptxSnapshot> = {
  type: "pptx:delete-slide",
  apply(snapshot, payload) {
    const { slide, index } = findSlide(snapshot, payload.slideIndex);
    const slideRelsPath = relsPathFor(slide.partPath);

    const newSlides = [...snapshot.root.slides];
    newSlides.splice(index, 1);

    // presentation rels: drop the rel for this slide.
    const presRels = snapshot.relationships.get(PRES_RELS_PATH);
    const newRelationships = new Map(snapshot.relationships);
    if (presRels) {
      const filtered = presRels.entries.filter((e) => e.id !== slide.relId);
      newRelationships.set(PRES_RELS_PATH, { relsPath: PRES_RELS_PATH, entries: filtered });
    }
    newRelationships.delete(slideRelsPath);

    // content types: drop Override for this slide and (if any) notes-slide.
    const removedPartNames = new Set<string>([`/${slide.partPath}`]);
    const removedParts = new Set(snapshot.removedParts);
    removedParts.add(slide.partPath);
    removedParts.add(slideRelsPath);
    if (slide.notesSlidePartPath) {
      removedParts.add(slide.notesSlidePartPath);
      removedParts.add(relsPathFor(slide.notesSlidePartPath));
      removedPartNames.add(`/${slide.notesSlidePartPath}`);
    }
    const newOverrides = snapshot.contentTypes.overrides.filter(
      (o) => !removedPartNames.has(o.partName)
    );
    const newContentTypes = {
      ...snapshot.contentTypes,
      overrides: newOverrides,
    };

    const root = { ...snapshot.root, slides: newSlides };

    const next = evolveSnapshot(
      snapshot,
      root,
      {
        presentation: true,
        contentTypes: true,
        relationships: [PRES_RELS_PATH],
      },
      {
        relationships: newRelationships,
        contentTypes: newContentTypes,
        removedParts,
      }
    );

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-deleted",
        nodeId: slide.id,
        path: ["slides", index],
        summary: "slide",
      }),
    };
  },
};

function relsPathFor(partPath: string): string {
  const slash = partPath.lastIndexOf("/");
  const dir = slash >= 0 ? partPath.slice(0, slash) : "";
  const file = slash >= 0 ? partPath.slice(slash + 1) : partPath;
  return `${dir}${dir ? "/" : ""}_rels/${file}.rels`;
}
