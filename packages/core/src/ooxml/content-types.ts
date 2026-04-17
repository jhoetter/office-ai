import { OoxmlContainer } from "./container.js";
import { asElement, filterChildren, makeElement, parseXml, serializeXml } from "./xml.js";

const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const CONTENT_TYPES_PATH = "[Content_Types].xml";

export interface DefaultEntry {
  readonly extension: string;
  readonly contentType: string;
}

export interface OverrideEntry {
  readonly partName: string; // "/word/comments.xml"
  readonly contentType: string;
}

/**
 * Lightweight reader/writer for [Content_Types].xml. We only edit it when
 * adding or removing a part with a content type not already registered.
 */
export class ContentTypes {
  constructor(
    public readonly defaults: DefaultEntry[],
    public readonly overrides: OverrideEntry[]
  ) {}

  static load(container: OoxmlContainer): ContentTypes {
    if (!container.has(CONTENT_TYPES_PATH)) {
      return new ContentTypes([], []);
    }
    const tree = parseXml(container.readText(CONTENT_TYPES_PATH));
    if (!Array.isArray(tree)) return new ContentTypes([], []);
    const root = (tree as unknown[]).map((n) => asElement(n)).find((e) => e?.tag === "Types");
    if (!root) return new ContentTypes([], []);
    const defaults: DefaultEntry[] = filterChildren(root.children, "Default").map((d) => ({
      extension: d.attrs.Extension ?? "",
      contentType: d.attrs.ContentType ?? "",
    }));
    const overrides: OverrideEntry[] = filterChildren(root.children, "Override").map((o) => ({
      partName: o.attrs.PartName ?? "",
      contentType: o.attrs.ContentType ?? "",
    }));
    return new ContentTypes(defaults, overrides);
  }

  hasOverride(partName: string): boolean {
    return this.overrides.some((o) => o.partName === partName);
  }

  hasDefault(extension: string): boolean {
    return this.defaults.some((d) => d.extension === extension);
  }

  addOverride(partName: string, contentType: string): void {
    if (!this.hasOverride(partName)) {
      this.overrides.push({ partName, contentType });
    }
  }

  addDefault(extension: string, contentType: string): void {
    if (!this.hasDefault(extension)) {
      this.defaults.push({ extension, contentType });
    }
  }

  removeOverride(partName: string): void {
    const idx = this.overrides.findIndex((o) => o.partName === partName);
    if (idx >= 0) this.overrides.splice(idx, 1);
  }

  writeBack(container: OoxmlContainer): void {
    const children: unknown[] = [];
    for (const d of this.defaults) {
      children.push(makeElement("Default", [], { Extension: d.extension, ContentType: d.contentType }));
    }
    for (const o of this.overrides) {
      children.push(makeElement("Override", [], { PartName: o.partName, ContentType: o.contentType }));
    }
    const tree = [makeElement("Types", children, { xmlns: CT_NS })];
    container.writeText(CONTENT_TYPES_PATH, serializeXml(tree));
  }
}
