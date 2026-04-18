/**
 * Headless-safe renderer entry point: pure layout helpers + SVG factories.
 * No DOM/React imports; safe to call from Node and from the browser alike.
 *
 * The interactive React canvas + HTML overlay live in the sibling
 * `./react/` directory, which the package exports via `./renderer/react`.
 */

export * from "./layout/units.js";
export * from "./layout/slide.js";
export * from "./layout/shape.js";
export * from "./layout/color.js";
export * from "./svg/slide.js";
export * from "./svg/shapes.js";
export { escXml } from "./svg/escape.js";
