export { OoxmlContainer, OoxmlContainerError, type OoxmlPart } from "./container.js";

export {
  parseXml,
  serializeXml,
  asElement,
  isElementEntry,
  getTag,
  getChildren,
  getAttrs,
  getTextContent,
  findChild,
  filterChildren,
  makeElement,
  makeTextLeaf,
  xmlAttrPrefix,
  type XmlNode,
  type XmlElement,
  type AttrMap,
} from "./xml.js";

export { RelationshipGraph, type Relationship } from "./relationships.js";

export { ContentTypes, type DefaultEntry, type OverrideEntry } from "./content-types.js";

export {
  addEmbeddedPart,
  ensureDefaultContentType,
  relativeTarget,
  CT_DRAWINGML_CHART,
  CT_SPREADSHEETML_SHEET,
  REL_TYPE_CHART,
  REL_TYPE_IMAGE,
  REL_TYPE_OLE_OBJECT,
  REL_TYPE_PACKAGE,
  type AddPackagePartArgs,
  type AddPackagePartResult,
} from "./embeddings.js";

export {
  serializeChartXml,
  type ChartSpec,
  type ChartSeriesSpec,
  type ChartType as ChartSpecType,
  type ChartEmbeddingContext,
} from "./chart-spec.js";
