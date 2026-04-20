import type { CommandHandler } from "@officeai/core";
import type { PdfDocument, PdfSnapshot } from "../model/types.js";
import { buildDiff, evolvePdf } from "./helpers.js";
import type { SetMetadataPayload } from "./payloads.js";

export const setMetadataHandler: CommandHandler<SetMetadataPayload, PdfSnapshot> = {
  type: "pdf:set-metadata",
  apply(snapshot, payload) {
    const merged = { ...snapshot.root.metadata };
    for (const key of Object.keys(payload) as Array<keyof SetMetadataPayload>) {
      const value = payload[key];
      if (value === undefined) continue;
      (merged as Record<string, unknown>)[key] = value;
    }
    const root: PdfDocument = { ...snapshot.root, metadata: merged };
    const next = evolvePdf(snapshot, root);
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: "metadata",
        path: ["metadata"],
        field: "metadata",
        summary: "document metadata",
      }),
    };
  },
};
