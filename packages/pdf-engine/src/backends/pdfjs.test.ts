import { describe, expect, it, vi } from "vitest";

interface ViewportCall {
  scale: number;
  rotation: number;
}

const viewportCalls: ViewportCall[] = [];

vi.mock("pdfjs-dist/legacy/build/pdf.mjs", () => {
  const fakePage = {
    pageNumber: 1,
    rotate: 0,
    getViewport: (opts: { scale?: number; rotation?: number }) => {
      viewportCalls.push({ scale: opts.scale ?? 1, rotation: opts.rotation ?? 0 });
      return { width: 100, height: 100, viewBox: [0, 0, 100, 100] };
    },
    render: () => ({ promise: Promise.resolve() }),
    getTextContent: () => Promise.resolve({ items: [], styles: {} }),
    getAnnotations: () => Promise.resolve([]),
    cleanup: () => undefined,
  };
  const fakeDoc = {
    numPages: 1,
    getPage: () => Promise.resolve(fakePage),
    getMetadata: () => Promise.resolve({ info: {}, metadata: null }),
    getOutline: () => Promise.resolve(null),
    getAttachments: () => Promise.resolve(null),
    destroy: () => Promise.resolve(),
  };
  return {
    getDocument: () => ({
      promise: Promise.resolve(fakeDoc),
      destroy: () => Promise.resolve(),
    }),
    GlobalWorkerOptions: {},
    version: "test",
  };
});

const { pdfjsBackend } = await import("./pdfjs.js");

interface StubCanvas {
  width: number;
  height: number;
  getContext: (kind: string) => CanvasRenderingContext2D | null;
}

function makeStubCanvas(): StubCanvas {
  const ctx = {} as CanvasRenderingContext2D;
  return {
    width: 0,
    height: 0,
    getContext: (kind: string) => (kind === "2d" ? ctx : null),
  };
}

/**
 * Verifies that PdfEngineRenderOptions.rotation is forwarded to PDF.js's
 * `getViewport({ rotation })`. The viewer relies on this to keep
 * canvas pixels and the un-rotated text layer in sync.
 */
describe("pdfjs backend rotation", () => {
  it("threads opts.rotation into getViewport and defaults to 0", async () => {
    viewportCalls.length = 0;
    const doc = await pdfjsBackend.load(new Uint8Array());
    const page = await doc.getPage(1);
    viewportCalls.length = 0; // ignore the construction-time getViewport().

    await page.render({ canvas: makeStubCanvas() as unknown as HTMLCanvasElement, scale: 2, rotation: 90 });
    await page.render({ canvas: makeStubCanvas() as unknown as HTMLCanvasElement, scale: 1 });

    expect(viewportCalls).toEqual([
      { scale: 2, rotation: 90 },
      { scale: 1, rotation: 0 },
    ]);
  });
});
