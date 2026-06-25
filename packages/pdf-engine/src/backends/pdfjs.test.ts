import { describe, expect, it, vi } from "vitest";

interface ViewportCall {
  scale: number;
  rotation: number;
}

const viewportCalls: ViewportCall[] = [];
const renderTasks: Array<{
  promise: Promise<void>;
  resolve: () => void;
  reject: (reason?: unknown) => void;
  cancel: ReturnType<typeof vi.fn>;
}> = [];

vi.mock("pdfjs-dist/legacy/build/pdf.mjs", () => {
  const fakePage = {
    pageNumber: 1,
    rotate: 0,
    getViewport: (opts: { scale?: number; rotation?: number }) => {
      viewportCalls.push({ scale: opts.scale ?? 1, rotation: opts.rotation ?? 0 });
      return { width: 100, height: 100, viewBox: [0, 0, 100, 100] };
    },
    render: () => {
      let resolve!: () => void;
      let reject!: (reason?: unknown) => void;
      const promise = new Promise<void>((res, rej) => {
        resolve = res;
        reject = rej;
      });
      const task = {
        promise,
        resolve,
        reject,
        cancel: vi.fn(() => reject(new Error("cancelled"))),
      };
      renderTasks.push(task);
      return task;
    },
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

async function waitForRenderTaskCount(count: number): Promise<void> {
  for (let i = 0; i < 10; i += 1) {
    if (renderTasks.length >= count) return;
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
  throw new Error(`Expected ${count} render task(s), saw ${renderTasks.length}`);
}

/**
 * Verifies that PdfEngineRenderOptions.rotation is forwarded to PDF.js's
 * `getViewport({ rotation })`. The viewer relies on this to keep
 * canvas pixels and the un-rotated text layer in sync.
 */
describe("pdfjs backend rotation", () => {
  it("threads opts.rotation into getViewport and defaults to 0", async () => {
    viewportCalls.length = 0;
    renderTasks.length = 0;
    const doc = await pdfjsBackend.load(new Uint8Array());
    const page = await doc.getPage(1);
    viewportCalls.length = 0; // ignore the construction-time getViewport().

    const first = page.render({
      canvas: makeStubCanvas() as unknown as HTMLCanvasElement,
      scale: 2,
      rotation: 90,
    });
    await waitForRenderTaskCount(1);
    renderTasks[0]?.resolve();
    await first;
    const second = page.render({ canvas: makeStubCanvas() as unknown as HTMLCanvasElement, scale: 1 });
    await waitForRenderTaskCount(2);
    renderTasks[1]?.resolve();
    await second;

    expect(viewportCalls).toEqual([
      { scale: 2, rotation: 90 },
      { scale: 1, rotation: 0 },
    ]);
  });

  it("cancels an active render before reusing the same canvas", async () => {
    renderTasks.length = 0;
    const doc = await pdfjsBackend.load(new Uint8Array());
    const page = await doc.getPage(1);
    const canvas = makeStubCanvas() as unknown as HTMLCanvasElement;

    const first = page.render({ canvas, scale: 1 });
    await waitForRenderTaskCount(1);
    const second = page.render({ canvas, scale: 2 });
    await expect(first).rejects.toThrow("cancelled");
    expect(renderTasks[0]?.cancel).toHaveBeenCalledTimes(1);

    await waitForRenderTaskCount(2);
    renderTasks[1]?.resolve();
    await second;
  });
});
