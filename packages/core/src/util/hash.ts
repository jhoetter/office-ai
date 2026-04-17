import { sha256 } from "js-sha256";

/**
 * Synchronous SHA-256 (hex) for both Node and the browser.
 *
 * We use `js-sha256` (MIT) so the same code runs in `apps/web` (where
 * `node:crypto` isn't available to client bundles) and in tests / the
 * `office-agent` CLI.
 */
export function sha256Hex(bytes: Uint8Array | string): string {
  if (typeof bytes === "string") return sha256(bytes);
  return sha256(bytes);
}
