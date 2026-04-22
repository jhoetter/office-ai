import { describe, expect, it } from "vitest";
import { autoActivationSignature, resolveActiveTabId, visibleRibbonTabs } from "./ribbon-resolve";
import type { RibbonCatalogue, RibbonTab } from "./RibbonTypes";

interface Ctx {
  hasImage: boolean;
  inHeader: boolean;
}

const empty = () => null;

const catalogue: RibbonCatalogue<Ctx> = {
  defaultTabId: "start",
  tabs: [
    {
      id: "start",
      label: "Start",
      groups: [{ id: "g", label: "G", render: empty }],
    },
    {
      id: "insert",
      label: "Einfügen",
      groups: [{ id: "g", label: "G", render: empty }],
    },
    {
      id: "image-tools",
      label: "Bildtools",
      contextual: { accent: "image" },
      visible: (c) => c.hasImage,
      autoActivateWhen: (c) => c.hasImage,
      groups: [{ id: "g", label: "G", render: empty }],
    },
    {
      id: "hf-tools",
      label: "Kopf- und Fußzeile",
      contextual: { accent: "hf" },
      visible: (c) => c.inHeader,
      autoActivateWhen: (c) => c.inHeader,
      groups: [{ id: "g", label: "G", render: empty }],
    },
  ],
};

describe("ribbon-resolve", () => {
  it("hides contextual tabs whose visibility predicate is false", () => {
    const visible = visibleRibbonTabs(catalogue, { hasImage: false, inHeader: false });
    expect(visible.map((t) => t.id)).toEqual(["start", "insert"]);
  });

  it("shows contextual tabs whose visibility predicate is true", () => {
    const visible = visibleRibbonTabs(catalogue, { hasImage: true, inHeader: false });
    expect(visible.map((t) => t.id)).toEqual(["start", "insert", "image-tools"]);
  });

  it("falls back to the pinned tab when nothing auto-activates", () => {
    const visible = visibleRibbonTabs(catalogue, { hasImage: false, inHeader: false });
    expect(resolveActiveTabId(visible, { hasImage: false, inHeader: false }, "insert")).toBe("insert");
  });

  it("auto-activates a contextual tab when its trigger fires", () => {
    const visible = visibleRibbonTabs(catalogue, { hasImage: true, inHeader: false });
    expect(resolveActiveTabId(visible, { hasImage: true, inHeader: false }, "insert")).toBe("image-tools");
  });

  it("prefers the first matching contextual tab when multiple triggers fire", () => {
    const visible = visibleRibbonTabs(catalogue, { hasImage: true, inHeader: true });
    // image-tools is declared earlier than hf-tools in the catalogue
    // so it wins. Ordering matters: callers should put the more
    // specific contextual tab first if they have overlapping triggers.
    expect(resolveActiveTabId(visible, { hasImage: true, inHeader: true }, "start")).toBe("image-tools");
  });

  it("returns the first visible tab when the pinned id is no longer visible", () => {
    const visible = visibleRibbonTabs(catalogue, { hasImage: false, inHeader: false });
    expect(resolveActiveTabId(visible, { hasImage: false, inHeader: false }, "image-tools")).toBe("start");
  });

  it("returns empty string for an empty catalogue", () => {
    expect(
      resolveActiveTabId<Ctx>([] as ReadonlyArray<RibbonTab<Ctx>>, { hasImage: false, inHeader: false }, "x")
    ).toBe("");
  });

  it("respects the user's pinned tab when the auto-activation signature still matches", () => {
    // User has an image selected (image-tools auto-fires) and clicks
    // the persistent "insert" tab. As long as the same contextual
    // tab keeps firing, the click must win — Office's "I want to
    // navigate elsewhere even though there's a contextual tab open"
    // behaviour. The Ribbon captures the live signature at click
    // time and passes it as `suppressedAutoSignature`.
    const ctx = { hasImage: true, inHeader: false };
    const visible = visibleRibbonTabs(catalogue, ctx);
    const sig = autoActivationSignature(visible, ctx);
    expect(sig).toBe("image-tools");
    expect(resolveActiveTabId(visible, ctx, "insert", sig)).toBe("insert");
  });

  it("drops the user's override once the auto-activation signature changes", () => {
    // User overrode while only image-tools fired. When the user
    // clicks into a header (hf-tools also fires), the contextual
    // landscape changed and the override should expire so the new
    // contextual tab takes over (matches Word's behaviour when you
    // jump from a picture to the header).
    const ctxNew = { hasImage: true, inHeader: true };
    const visibleNew = visibleRibbonTabs(catalogue, ctxNew);
    const oldSig = "image-tools";
    expect(resolveActiveTabId(visibleNew, ctxNew, "insert", oldSig)).toBe("image-tools");
  });

  it("autoActivationSignature returns empty when no contextual tab fires", () => {
    const ctx = { hasImage: false, inHeader: false };
    const visible = visibleRibbonTabs(catalogue, ctx);
    expect(autoActivationSignature(visible, ctx)).toBe("");
  });

  it("autoActivationSignature joins multiple firing contextual tabs in catalogue order", () => {
    const ctx = { hasImage: true, inHeader: true };
    const visible = visibleRibbonTabs(catalogue, ctx);
    expect(autoActivationSignature(visible, ctx)).toBe("image-tools|hf-tools");
  });
});
