import { v4 as uuidv4 } from "uuid";
import type { NodeId } from "../types/document.js";

export interface IdMinter {
  (): NodeId;
}

export const defaultIdMinter: IdMinter = () => uuidv4();

/**
 * Deterministic minter for tests. Produces ids of the form `n-0`, `n-1`, ...
 */
export function deterministicIdMinter(prefix = "n"): IdMinter {
  let i = 0;
  return () => `${prefix}-${i++}`;
}
