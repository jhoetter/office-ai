import * as React from "react";
import type { PptxAgent } from "../../agent/agent.js";
import type { PptxSnapshot } from "../../model/types.js";

/**
 * Subscribe a React tree to the agent's snapshot stream. Returns the latest
 * approved snapshot.
 */
export function useAgentSnapshot(agent: PptxAgent): PptxSnapshot {
  const [snap, setSnap] = React.useState(() => agent.getSnapshot());
  React.useEffect(() => {
    setSnap(agent.getSnapshot());
    return agent.subscribe((s) => setSnap(s));
  }, [agent]);
  return snap;
}
