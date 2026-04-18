import { CommandError, ooxml, sha256Hex, type CommandHandler, type DiffChange } from "@officeai/core";
import {
  EXTENSION_BY_CONTENT_TYPE,
  type ImageBlob,
  type SheetImage,
} from "../model/drawings.js";
import type { Sheet, XlsxSnapshot, XlsxWorkbook } from "../model/types.js";
import { buildDiff, evolveSnapshot, replaceSheet, type PartialDirtyFlags } from "./helpers.js";
import type {
  AddImagePayload,
  MoveImagePayload,
  RemoveImagePayload,
  ResizeImagePayload,
} from "./payloads.js";
import { resolveSheet } from "./validation.js";

/**
 * `xlsx:add-image` — splice a new image instance into a sheet.
 *
 * Pipeline:
 *   1. Resolve the sheet and validate the anchor + dimensions.
 *   2. Hash the bytes and dedupe against `workbook.images`. Identical
 *      bytes uploaded twice collapse to a single `xl/media/imageN.*`.
 *   3. If the bytes are new, mint a fresh `xl/media/imageN.*` slot
 *      and stash an {@link ImageBlob} on the workbook's image map.
 *   4. Append a {@link SheetImage} to `sheet.images` (later items
 *      overlay earlier ones — Excel's z-order semantics).
 *   5. Mint a `drawingPartPath` if the sheet didn't have one.
 *   6. Dirty `drawings`, `media`, `sheets` (so the worksheet XML
 *      gets the `<drawing>` ref injected), `sheetRels`, `rels`, and
 *      `contentTypes`.
 */
export const addImageHandler: CommandHandler<AddImagePayload, XlsxSnapshot> = {
  type: "xlsx:add-image",
  apply(snapshot, payload, ctx) {
    if (payload.bytes.byteLength === 0) {
      throw new CommandError("empty-image", "Image bytes must be non-empty");
    }
    if (payload.widthPx <= 0 || payload.heightPx <= 0) {
      throw new CommandError(
        "invalid-image-size",
        `Image size ${payload.widthPx}x${payload.heightPx} must be positive`
      );
    }
    if (payload.fromRow < 0 || payload.fromCol < 0) {
      throw new CommandError(
        "invalid-anchor",
        `Anchor (${payload.fromRow}, ${payload.fromCol}) must be non-negative`
      );
    }

    const sheet = resolveSheet(snapshot.root, payload.sheet);
    const hash = sha256Hex(payload.bytes);

    let mediaRef: string | undefined;
    let blob: ImageBlob | undefined;
    let isNewBlob = false;
    for (const existing of snapshot.root.images.values()) {
      if (existing.hash === hash && existing.contentType === payload.contentType) {
        mediaRef = existing.partPath;
        blob = existing;
        break;
      }
    }
    if (!mediaRef || !blob) {
      const ext = EXTENSION_BY_CONTENT_TYPE[payload.contentType];
      mediaRef = mintMediaPartPath(snapshot, ext);
      blob = {
        partPath: mediaRef,
        bytes: payload.bytes,
        contentType: payload.contentType,
        hash,
      };
      isNewBlob = true;
    }

    const newImage: SheetImage = {
      id: ctx.mintNodeId(),
      mediaRef,
      anchor: {
        fromRow: payload.fromRow,
        fromCol: payload.fromCol,
        fromOffsetXPx: payload.fromOffsetXPx ?? 0,
        fromOffsetYPx: payload.fromOffsetYPx ?? 0,
        widthPx: payload.widthPx,
        heightPx: payload.heightPx,
        editAs: "oneCell",
      },
      ...(payload.name ? { name: payload.name } : {}),
      ...(payload.altText ? { altText: payload.altText } : {}),
    };

    const drawingPartPath = sheet.drawingPartPath ?? mintDrawingPartPath(snapshot);
    const nextSheet: Sheet = {
      ...sheet,
      images: [...sheet.images, newImage],
      drawingPartPath,
    };

    const nextImages = new Map(snapshot.root.images);
    if (isNewBlob && blob) nextImages.set(blob.partPath, blob);
    const nextWorkbook: XlsxWorkbook = {
      ...replaceSheet(snapshot.root, nextSheet),
      images: nextImages,
    };

    const dirtyPatch: PartialDirtyFlags = {
      drawings: [sheet.partPath],
      sheets: [sheet.partPath],
      sheetRels: [ooxml.RelationshipGraph.relsPathFor(sheet.partPath)],
      contentTypes: true,
    };
    if (isNewBlob) dirtyPatch.media = [blob.partPath];

    const next = evolveSnapshot(snapshot, nextWorkbook, dirtyPatch);

    const changes: DiffChange[] = [
      {
        kind: "node-inserted",
        nodeId: nextSheet.id,
        path: ["sheets", sheet.index, "images", newImage.id],
        summary: `Added image to ${sheet.name}`,
        meta: {
          sheet: sheet.name,
          imageId: newImage.id,
          mediaRef,
          fromRow: payload.fromRow,
          fromCol: payload.fromCol,
          widthPx: payload.widthPx,
          heightPx: payload.heightPx,
        },
      },
    ];
    return { next, diff: buildDiff(snapshot.revision, next.revision, changes) };
  },
};

export const moveImageHandler: CommandHandler<MoveImagePayload, XlsxSnapshot> = {
  type: "xlsx:move-image",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    const idx = sheet.images.findIndex((i) => i.id === payload.imageId);
    if (idx === -1) {
      throw new CommandError(
        "image-not-found",
        `Image ${payload.imageId} not found on ${payload.sheet}`
      );
    }
    if (payload.fromRow < 0 || payload.fromCol < 0) {
      throw new CommandError(
        "invalid-anchor",
        `Anchor (${payload.fromRow}, ${payload.fromCol}) must be non-negative`
      );
    }
    const prev = sheet.images[idx]!;
    const nextImage: SheetImage = {
      ...prev,
      anchor: {
        ...prev.anchor,
        fromRow: payload.fromRow,
        fromCol: payload.fromCol,
        fromOffsetXPx: payload.fromOffsetXPx,
        fromOffsetYPx: payload.fromOffsetYPx,
      },
    };
    const images = sheet.images.slice();
    images[idx] = nextImage;
    const nextSheet: Sheet = { ...sheet, images };
    const nextWorkbook = replaceSheet(snapshot.root, nextSheet);
    const next = evolveSnapshot(snapshot, nextWorkbook, {
      drawings: [sheet.partPath],
      sheets: [sheet.partPath],
    });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, [
        {
          kind: "node-updated",
          nodeId: nextSheet.id,
          path: ["sheets", sheet.index, "images", nextImage.id],
          field: "anchor",
          summary: `Moved image on ${sheet.name}`,
          meta: {
            sheet: sheet.name,
            imageId: nextImage.id,
            fromRow: payload.fromRow,
            fromCol: payload.fromCol,
          },
        },
      ]),
    };
  },
};

export const resizeImageHandler: CommandHandler<ResizeImagePayload, XlsxSnapshot> = {
  type: "xlsx:resize-image",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    const idx = sheet.images.findIndex((i) => i.id === payload.imageId);
    if (idx === -1) {
      throw new CommandError(
        "image-not-found",
        `Image ${payload.imageId} not found on ${payload.sheet}`
      );
    }
    if (payload.widthPx <= 0 || payload.heightPx <= 0) {
      throw new CommandError(
        "invalid-image-size",
        `Image size ${payload.widthPx}x${payload.heightPx} must be positive`
      );
    }
    const prev = sheet.images[idx]!;
    const nextImage: SheetImage = {
      ...prev,
      anchor: {
        ...prev.anchor,
        widthPx: payload.widthPx,
        heightPx: payload.heightPx,
      },
    };
    const images = sheet.images.slice();
    images[idx] = nextImage;
    const nextSheet: Sheet = { ...sheet, images };
    const nextWorkbook = replaceSheet(snapshot.root, nextSheet);
    const next = evolveSnapshot(snapshot, nextWorkbook, {
      drawings: [sheet.partPath],
      sheets: [sheet.partPath],
    });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, [
        {
          kind: "node-updated",
          nodeId: nextSheet.id,
          path: ["sheets", sheet.index, "images", nextImage.id],
          field: "anchor",
          summary: `Resized image on ${sheet.name}`,
          meta: {
            sheet: sheet.name,
            imageId: nextImage.id,
            widthPx: payload.widthPx,
            heightPx: payload.heightPx,
          },
        },
      ]),
    };
  },
};

export const removeImageHandler: CommandHandler<RemoveImagePayload, XlsxSnapshot> = {
  type: "xlsx:remove-image",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    const idx = sheet.images.findIndex((i) => i.id === payload.imageId);
    if (idx === -1) {
      throw new CommandError(
        "image-not-found",
        `Image ${payload.imageId} not found on ${payload.sheet}`
      );
    }
    const removed = sheet.images[idx]!;
    const images = sheet.images.slice();
    images.splice(idx, 1);

    const nextSheet: Sheet = images.length === 0
      ? { ...sheet, images, drawingPartPath: undefined }
      : { ...sheet, images };

    // GC orphan media: a media blob is orphaned when no sheet (after
    // this removal) still references it. We diff at workbook scope to
    // catch images shared across sheets.
    const stillReferenced = new Set<string>();
    for (const s of snapshot.root.sheets) {
      const sheetImages = s.id === sheet.id ? images : s.images;
      for (const img of sheetImages) stillReferenced.add(img.mediaRef);
    }
    const removedMediaParts: string[] = [];
    const nextImages = new Map(snapshot.root.images);
    if (!stillReferenced.has(removed.mediaRef)) {
      nextImages.delete(removed.mediaRef);
      removedMediaParts.push(removed.mediaRef);
    }

    const nextWorkbook: XlsxWorkbook = {
      ...replaceSheet(snapshot.root, nextSheet),
      images: nextImages,
    };

    const dirtyPatch: PartialDirtyFlags = {
      drawings: [sheet.partPath],
      sheets: [sheet.partPath],
      sheetRels: [ooxml.RelationshipGraph.relsPathFor(sheet.partPath)],
      contentTypes: true,
    };
    if (removedMediaParts.length > 0) dirtyPatch.removedMediaParts = removedMediaParts;

    const next = evolveSnapshot(snapshot, nextWorkbook, dirtyPatch);
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, [
        {
          kind: "node-deleted",
          nodeId: nextSheet.id,
          path: ["sheets", sheet.index, "images", removed.id],
          summary: `Removed image from ${sheet.name}`,
          meta: {
            sheet: sheet.name,
            imageId: removed.id,
            mediaRef: removed.mediaRef,
            mediaCollected: removedMediaParts,
          },
        },
      ]),
    };
  },
};

function mintMediaPartPath(snapshot: XlsxSnapshot, ext: string): string {
  const taken = new Set<string>();
  for (const path of snapshot.container.parts.keys()) taken.add(path);
  for (const path of snapshot.root.images.keys()) taken.add(path);
  let i = 1;
  while (taken.has(`xl/media/image${i}.${ext}`)) i++;
  return `xl/media/image${i}.${ext}`;
}

function mintDrawingPartPath(snapshot: XlsxSnapshot): string {
  const taken = new Set<string>();
  for (const path of snapshot.container.parts.keys()) taken.add(path);
  for (const s of snapshot.root.sheets) {
    if (s.drawingPartPath) taken.add(s.drawingPartPath);
  }
  let i = 1;
  while (taken.has(`xl/drawings/drawing${i}.xml`)) i++;
  return `xl/drawings/drawing${i}.xml`;
}
