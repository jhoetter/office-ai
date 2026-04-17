import { OoxmlContainer } from "./container.js";
import {
  asElement,
  filterChildren,
  findChild,
  getAttrs,
  makeElement,
  parseXml,
  serializeXml,
  type XmlElement,
} from "./xml.js";

export interface Relationship {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly targetMode?: "External" | "Internal";
}

const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

export class RelationshipGraph {
  private readonly _rels: Relationship[];

  constructor(
    /** The path to the relationships part itself, e.g. word/_rels/document.xml.rels. */
    public readonly relsPath: string,
    rels: ReadonlyArray<Relationship>,
  ) {
    this._rels = [...rels];
  }

  static relsPathFor(partPath: string): string {
    const slash = partPath.lastIndexOf("/");
    const dir = slash >= 0 ? partPath.slice(0, slash) : "";
    const name = slash >= 0 ? partPath.slice(slash + 1) : partPath;
    return `${dir}/_rels/${name}.rels`;
  }

  static loadFor(container: OoxmlContainer, partPath: string): RelationshipGraph {
    const relsPath = RelationshipGraph.relsPathFor(partPath);
    if (!container.has(relsPath)) {
      return new RelationshipGraph(relsPath, []);
    }
    const xml = container.readText(relsPath);
    const tree = parseXml(xml);
    if (!Array.isArray(tree)) return new RelationshipGraph(relsPath, []);
    const rootEntry = (tree as unknown[]).map((n) => asElement(n)).find((el) => el?.tag === "Relationships");
    if (!rootEntry) return new RelationshipGraph(relsPath, []);
    const rels: Relationship[] = filterChildren(rootEntry.children, "Relationship").map((el) =>
      relationshipFromElement(el),
    );
    return new RelationshipGraph(relsPath, rels);
  }

  get relationships(): ReadonlyArray<Relationship> {
    return this._rels;
  }

  byId(id: string): Relationship | undefined {
    return this._rels.find((r) => r.id === id);
  }

  byType(type: string): ReadonlyArray<Relationship> {
    return this._rels.filter((r) => r.type === type);
  }

  /** Allocate a fresh `rId` not present in this graph. */
  mintId(): string {
    const taken = new Set(this._rels.map((r) => r.id));
    let i = this._rels.length + 1;
    while (taken.has(`rId${i}`)) i++;
    return `rId${i}`;
  }

  add(rel: Omit<Relationship, "id"> & { id?: string }): Relationship {
    const id = rel.id ?? this.mintId();
    const r: Relationship = {
      id,
      type: rel.type,
      target: rel.target,
      ...(rel.targetMode ? { targetMode: rel.targetMode } : {}),
    };
    this._rels.push(r);
    return r;
  }

  remove(id: string): void {
    const idx = this._rels.findIndex((r) => r.id === id);
    if (idx >= 0) this._rels.splice(idx, 1);
  }

  writeBack(container: OoxmlContainer): void {
    const children = this._rels.map((r) =>
      makeElement("Relationship", [], {
        Id: r.id,
        Type: r.type,
        Target: r.target,
        ...(r.targetMode ? { TargetMode: r.targetMode } : {}),
      }),
    );
    const tree = [
      makeElement("Relationships", children, { xmlns: RELS_NS }),
    ];
    const xml = serializeXml(tree);
    container.writeText(this.relsPath, xml);
  }
}

function relationshipFromElement(el: XmlElement): Relationship {
  const attrs = el.attrs;
  return {
    id: attrs.Id ?? "",
    type: attrs.Type ?? "",
    target: attrs.Target ?? "",
    ...(attrs.TargetMode === "External" || attrs.TargetMode === "Internal"
      ? { targetMode: attrs.TargetMode as "External" | "Internal" }
      : {}),
  };
}

/** Re-export for module ergonomics. */
export { findChild, getAttrs };
