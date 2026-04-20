/**
 * @officeai/pdf-edit — pure page-level operations on PDF byte buffers.
 *
 * Each function is `(input: Uint8Array, opts) => Promise<Uint8Array>`
 * and produces a fresh PDF. Backed by pdf-lib (MIT). No DOM, no React.
 *
 * Spec: /spec/pdf/editing-pipeline.md
 */
export { rotatePages, type RotatePagesOptions } from "./rotate-pages.js";
export { reorderPages, type ReorderPagesOptions } from "./reorder-pages.js";
export { deletePages, type DeletePagesOptions } from "./delete-pages.js";
export { insertPages, type InsertPagesOptions } from "./insert-pages.js";
export { extractPages, type ExtractPagesOptions } from "./extract-pages.js";
export { mergePdfs, type MergePdfsOptions } from "./merge.js";
export { splitPdf, type SplitPdfOptions, type SplitPdfResult } from "./split.js";
export { setMetadata, type SetMetadataOptions } from "./set-metadata.js";
export { addPageNumbers, type AddPageNumbersOptions } from "./add-page-numbers.js";
export { addWatermark, type AddWatermarkOptions } from "./watermark.js";
export { cropPages, type CropPagesOptions } from "./crop.js";
