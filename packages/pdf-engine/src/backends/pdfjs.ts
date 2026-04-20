import { buildGlyphRuns } from "../text/build-glyph-runs.js";
import type {
  PdfEngine,
  PdfEngineAnnotationLite,
  PdfEngineDocument,
  PdfEngineFormFieldLite,
  PdfEngineGlyphRun,
  PdfEngineLoadOptions,
  PdfEngineMetadata,
  PdfEngineOutlineNode,
  PdfEnginePage,
  PdfEnginePageInfo,
  PdfEngineRenderOptions,
  PdfEngineTextContent,
  PdfEngineTextItem,
  PdfEngineViewport,
} from "../types.js";

/**
 * PDF.js backend (Apache 2.0). Uses the legacy build for environment
 * compatibility (works in both browser via bundler and Node ≥ 20 via
 * the .mjs entry).
 *
 * Worker-thread isolation is left to the consumer:
 *   - In the browser: register the worker via GlobalWorkerOptions.workerSrc.
 *   - In Node: pdf.js falls back to the main-thread implementation (slower
 *     but functional). For headless rendering CLI use, set the worker
 *     yourself if you need parallelism.
 */

type PdfjsModule = typeof import("pdfjs-dist/legacy/build/pdf.mjs");

let pdfjsModule: PdfjsModule | null = null;

const loadPdfjs = async (): Promise<PdfjsModule> => {
  if (pdfjsModule) return pdfjsModule;
  pdfjsModule = await import("pdfjs-dist/legacy/build/pdf.mjs");
  return pdfjsModule;
};

const ROTATION_VALUES: ReadonlyArray<0 | 90 | 180 | 270> = [0, 90, 180, 270];

const normalizeRotation = (raw: number): 0 | 90 | 180 | 270 => {
  const r = ((raw % 360) + 360) % 360;
  for (const v of ROTATION_VALUES) if (v === r) return v;
  return 0;
};

const annotationSubtype = (a: { subtype?: string }): string =>
  typeof a.subtype === "string" ? a.subtype : "Unknown";

const formFieldType = (annot: {
  fieldType?: string;
  multiLine?: boolean;
}): PdfEngineFormFieldLite["type"] => {
  switch (annot.fieldType) {
    case "Tx":
      return "text";
    case "Btn":
      return "checkbox";
    case "Ch":
      return "choice";
    case "Sig":
      return "signature";
    default:
      return "unknown";
  }
};

const buildPage = async (raw: import("pdfjs-dist").PDFPageProxy): Promise<PdfEnginePage> => {
  const viewport = raw.getViewport({ scale: 1 });
  const info: PdfEnginePageInfo = {
    pageNumber: raw.pageNumber,
    width: viewport.viewBox[2] - viewport.viewBox[0],
    height: viewport.viewBox[3] - viewport.viewBox[1],
    rotation: normalizeRotation(raw.rotate ?? 0),
  };

  const getTextContent = async (): Promise<PdfEngineTextContent> => {
    const text = await raw.getTextContent();
    const items: PdfEngineTextItem[] = [];
    let plain = "";
    for (const it of text.items) {
      if (typeof (it as { str?: unknown }).str !== "string") continue;
      const item = it as {
        str: string;
        transform: number[];
        width: number;
        height: number;
        fontName?: string;
        hasEOL?: boolean;
      };
      items.push({
        str: item.str,
        transform: [
          item.transform[0],
          item.transform[1],
          item.transform[2],
          item.transform[3],
          item.transform[4],
          item.transform[5],
        ] as const,
        width: item.width,
        height: item.height,
        ...(item.fontName !== undefined ? { fontName: item.fontName } : {}),
        ...(item.hasEOL !== undefined ? { hasEol: item.hasEOL } : {}),
      });
      plain += item.str;
      if (item.hasEOL) plain += "\n";
    }
    return { items, plain };
  };

  const getGlyphRuns = async (): Promise<ReadonlyArray<PdfEngineGlyphRun>> => {
    const { items } = await getTextContent();
    return buildGlyphRuns(items);
  };

  const getViewportApi = (opts: { scale: number; rotation?: 0 | 90 | 180 | 270 }): PdfEngineViewport => {
    const rotation = opts.rotation ?? 0;
    const vp = raw.getViewport({ scale: opts.scale, rotation });
    return {
      scale: opts.scale,
      rotation,
      width: vp.width,
      height: vp.height,
      raw: vp,
    };
  };

  // PDF.js v4 exposes `streamTextContent()` which yields a
  // `ReadableStream<TextContent>`; fall back to the eagerly-resolved
  // `getTextContent()` if the streaming variant isn't available
  // (older minor versions and the legacy build alias).
  const getTextContentSource = async (opts: { includeMarkedContent?: boolean } = {}): Promise<unknown> => {
    const includeMarkedContent = opts.includeMarkedContent ?? true;
    const params = {
      includeMarkedContent,
      // disableNormalization=false is the v4 default but we set it
      // explicitly to lock the documented Unicode normalisation
      // behaviour in.
      disableNormalization: false,
    };
    const streamer = (
      raw as unknown as {
        streamTextContent?: (opts: unknown) => ReadableStream;
      }
    ).streamTextContent;
    if (typeof streamer === "function") {
      return streamer.call(raw, params);
    }
    return raw.getTextContent(params as never);
  };

  const getAnnotations = async (): Promise<ReadonlyArray<PdfEngineAnnotationLite>> => {
    const raws = await raw.getAnnotations();
    return raws
      .filter((a) => Array.isArray((a as { rect?: unknown }).rect))
      .map((a, i) => {
        const annot = a as {
          id?: string;
          subtype?: string;
          rect: number[];
          contents?: string;
          contentsObj?: { str?: string };
          title?: string;
          url?: string;
          dest?: unknown;
        };
        const out: PdfEngineAnnotationLite = {
          id: annot.id ?? `annot-${i}`,
          subtype: annotationSubtype(annot),
          rect: [annot.rect[0], annot.rect[1], annot.rect[2], annot.rect[3]] as const,
          ...(annot.contents !== undefined
            ? { contents: annot.contents }
            : annot.contentsObj?.str !== undefined
              ? { contents: annot.contentsObj.str }
              : {}),
          ...(annot.title !== undefined ? { author: annot.title } : {}),
          ...(annot.url !== undefined ? { url: annot.url } : {}),
        };
        return out;
      });
  };

  const getFormFields = async (): Promise<ReadonlyArray<PdfEngineFormFieldLite>> => {
    const raws = await raw.getAnnotations({ intent: "any" });
    return raws
      .filter((a) => (a as { fieldType?: string }).fieldType !== undefined)
      .map((a, i) => {
        const annot = a as {
          id?: string;
          fieldName?: string;
          fieldValue?: unknown;
          fieldType?: string;
          options?: ReadonlyArray<{ exportValue?: string; displayValue?: string }>;
          readOnly?: boolean;
          required?: boolean;
          maxLen?: number;
          multiLine?: boolean;
          password?: boolean;
          rect: number[];
        };
        const opts: PdfEngineFormFieldLite = {
          id: annot.id ?? `field-${i}`,
          name: annot.fieldName ?? annot.id ?? `field-${i}`,
          type: formFieldType(annot),
          ...(annot.fieldValue !== undefined &&
          (typeof annot.fieldValue === "string" || typeof annot.fieldValue === "boolean")
            ? { value: annot.fieldValue }
            : {}),
          ...(annot.options !== undefined
            ? {
                options: annot.options
                  .map((o) => o.displayValue ?? o.exportValue ?? "")
                  .filter((v): v is string => typeof v === "string" && v.length > 0),
              }
            : {}),
          readOnly: Boolean(annot.readOnly),
          required: Boolean(annot.required),
          ...(annot.maxLen !== undefined ? { maxLength: annot.maxLen } : {}),
          ...(annot.multiLine !== undefined ? { multiline: annot.multiLine } : {}),
          ...(annot.password !== undefined ? { password: annot.password } : {}),
          pageNumber: raw.pageNumber,
          rect: [annot.rect[0], annot.rect[1], annot.rect[2], annot.rect[3]] as const,
        };
        return opts;
      });
  };

  const render = async (renderOpts: PdfEngineRenderOptions = {}): Promise<Uint8Array | undefined> => {
    const scale = renderOpts.dpi ? renderOpts.dpi / 72 : (renderOpts.scale ?? 1);
    const rotation = renderOpts.rotation ?? 0;
    const renderViewport = raw.getViewport({ scale, rotation });
    const canvas = renderOpts.canvas;
    if (!canvas) return undefined;
    const ctx = (canvas as HTMLCanvasElement).getContext("2d");
    if (!ctx) throw new Error("pdfjs backend: 2d canvas context unavailable");
    canvas.width = Math.ceil(renderViewport.width);
    canvas.height = Math.ceil(renderViewport.height);
    await raw.render({
      canvasContext: ctx as unknown as CanvasRenderingContext2D,
      viewport: renderViewport,
      canvas: canvas as HTMLCanvasElement,
    } as unknown as Parameters<typeof raw.render>[0]).promise;
    return undefined;
  };

  return {
    info,
    render,
    getTextContent,
    getGlyphRuns,
    getViewport: getViewportApi,
    getTextContentSource,
    getAnnotations,
    getFormFields,
    destroy: () => raw.cleanup(),
  };
};

const buildOutline = async (
  raw: import("pdfjs-dist").PDFDocumentProxy
): Promise<PdfEngineOutlineNode[] | null> => {
  const top = await raw.getOutline();
  if (!top) return null;

  const walk = async (
    nodes: ReadonlyArray<{
      title: string;
      dest?: unknown;
      url?: string;
      items?: ReadonlyArray<unknown>;
    }>
  ): Promise<PdfEngineOutlineNode[]> => {
    const out: PdfEngineOutlineNode[] = [];
    for (const n of nodes) {
      let pageNumber: number | undefined;
      if (n.dest != null) {
        try {
          const dest = typeof n.dest === "string" ? await raw.getDestination(n.dest) : n.dest;
          if (Array.isArray(dest) && dest[0]) {
            const idx = await raw.getPageIndex(dest[0] as never);
            pageNumber = idx + 1;
          }
        } catch {
          // ignore — outline entries can have unresolvable destinations.
        }
      }
      const node: PdfEngineOutlineNode = {
        title: n.title,
        ...(pageNumber !== undefined ? { pageNumber } : {}),
        ...(n.url !== undefined ? { uri: n.url } : {}),
        children: n.items ? await walk(n.items as never) : [],
      };
      out.push(node);
    }
    return out;
  };

  return walk(top as never);
};

/**
 * Normalise an `assetsBase` value into separate cmap / standard-font
 * URLs. PDF.js needs both with trailing slashes; consumers shouldn't
 * have to know that.
 */
const resolveAssetUrls = (
  assetsBase: string | undefined
): { cMapUrl?: string; cMapPacked?: boolean; standardFontDataUrl?: string } => {
  if (!assetsBase) return {};
  const base = assetsBase.endsWith("/") ? assetsBase : `${assetsBase}/`;
  return {
    cMapUrl: `${base}cmaps/`,
    cMapPacked: true,
    standardFontDataUrl: `${base}standard_fonts/`,
  };
};

export const pdfjsBackend: PdfEngine = {
  kind: "pdfjs",
  async load(buffer: Uint8Array, opts: PdfEngineLoadOptions = {}): Promise<PdfEngineDocument> {
    const pdfjs = await loadPdfjs();
    const assetUrls = resolveAssetUrls(opts.assetsBase);
    const loadingTask = pdfjs.getDocument({
      data: buffer,
      ...(opts.password ? { password: opts.password } : {}),
      isEvalSupported: false,
      // Fonts are typically on a CDN; consumers should configure this.
      disableFontFace: false,
      // Asset wiring: `cmaps/` is required for any PDF that uses a
      // non-standard CIDSystemInfo (CJK, custom encodings); without
      // it `getTextContent()` returns mojibake or empty strings for
      // entire languages. `standard_fonts/` is required for the 14
      // PDF base fonts when not embedded — without it pdfjs falls
      // back to a hard-coded sans-serif and glyph metrics drift.
      ...assetUrls,
      // Prefer system fonts when available — better visual fidelity
      // for unembedded standard fonts and noticeably less drift in
      // the text layer.
      useSystemFonts: true,
      verbosity: 0,
    });

    const doc = await loadingTask.promise;

    return {
      engine: "pdfjs",
      numPages: doc.numPages,
      async getPage(pageNumber: number): Promise<PdfEnginePage> {
        const page = await doc.getPage(pageNumber);
        return buildPage(page);
      },
      async getMetadata(): Promise<PdfEngineMetadata> {
        const md = await doc.getMetadata();
        const info = (md.info ?? {}) as Record<string, unknown>;
        const isStr = (v: unknown): v is string => typeof v === "string" && v.length > 0;
        const out: Record<string, unknown> = {};
        if (isStr(info.Title)) out.title = info.Title;
        if (isStr(info.Author)) out.author = info.Author;
        if (isStr(info.Subject)) out.subject = info.Subject;
        if (isStr(info.Keywords)) out.keywords = info.Keywords;
        if (isStr(info.Creator)) out.creator = info.Creator;
        if (isStr(info.Producer)) out.producer = info.Producer;
        if (isStr(info.CreationDate)) out.creationDate = info.CreationDate;
        if (isStr(info.ModDate)) out.modificationDate = info.ModDate;
        if (isStr(info.PDFFormatVersion)) out.pdfVersion = info.PDFFormatVersion;
        if (typeof info.IsLinearized === "boolean") out.linearized = info.IsLinearized;
        return out as PdfEngineMetadata;
      },
      getOutline: () => buildOutline(doc),
      async getAttachments(): Promise<ReadonlyArray<{ name: string; data: Uint8Array }>> {
        const att = await doc.getAttachments();
        if (!att) return [];
        return Object.entries(att as Record<string, { filename?: string; content?: Uint8Array }>).map(
          ([key, val]) => ({
            name: val.filename ?? key,
            data: val.content ?? new Uint8Array(0),
          })
        );
      },
      estimatedBytes: () => buffer.byteLength,
      async destroy(): Promise<void> {
        await doc.destroy();
        await loadingTask.destroy();
      },
    };
  },
};
