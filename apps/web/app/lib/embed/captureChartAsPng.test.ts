import { afterEach, describe, expect, it, vi } from "vitest";
import { captureChartAsPng } from "./captureChartAsPng";

/**
 * The vitest config targets the Node environment (no DOM), so these
 * tests stub out just the surface area `captureChartAsPng` touches:
 * `window`, `document.createElement("canvas")`, `XMLSerializer`,
 * `URL.createObjectURL`, and the global `Image` constructor. That
 * keeps the unit test cheap and deterministic — the real browser
 * round-trip is exercised by Playwright e2e once D4 lands.
 */

interface FakeCanvas {
  width: number;
  height: number;
  getContext: (kind: string) => { drawImage: () => void } | null;
  toDataURL: (mime: string) => string;
}

interface MutableGlobal {
  window?: unknown;
  document?: unknown;
  Image?: unknown;
  URL?: unknown;
  XMLSerializer?: unknown;
}

const g = globalThis as MutableGlobal;

function installBrowserStubs(opts: {
  pngBase64?: string;
  failImageLoad?: boolean;
}): { uninstall: () => void } {
  const originals: MutableGlobal = {
    window: g.window,
    document: g.document,
    Image: g.Image,
    URL: g.URL,
    XMLSerializer: g.XMLSerializer,
  };
  const pngBase64 = opts.pngBase64 ?? "iVBORw0KGgoAAAANSUhEUgFAKE";
  const fakeCanvas: FakeCanvas = {
    width: 0,
    height: 0,
    getContext: () => ({ drawImage: () => {} }),
    toDataURL: () => `data:image/png;base64,${pngBase64}`,
  };
  g.window = {};
  g.document = {
    createElement: (tag: string) => {
      if (tag === "canvas") return fakeCanvas;
      throw new Error(`unexpected createElement(${tag}) in test`);
    },
  };
  g.URL = {
    createObjectURL: () => "blob:fake",
    revokeObjectURL: () => {},
  };
  g.XMLSerializer = class {
    serializeToString(node: { outerHTML?: string }): string {
      return node.outerHTML ?? "<svg/>";
    }
  };
  g.Image = class {
    onload: (() => void) | null = null;
    onerror: (() => void) | null = null;
    set src(_v: string) {
      queueMicrotask(() => {
        if (opts.failImageLoad) this.onerror?.();
        else this.onload?.();
      });
    }
  };
  return {
    uninstall: () => {
      g.window = originals.window;
      g.document = originals.document;
      g.Image = originals.Image;
      g.URL = originals.URL;
      g.XMLSerializer = originals.XMLSerializer;
    },
  };
}

afterEach(() => {
  vi.restoreAllMocks();
});

describe("captureChartAsPng", () => {
  it("rasterises an <svg> source via an offscreen canvas and strips the data: prefix", async () => {
    const stubs = installBrowserStubs({ pngBase64: "TESTPAYLOAD" });
    try {
      const fakeSvg = {
        outerHTML: '<svg xmlns="http://www.w3.org/2000/svg" width="120" height="80"></svg>',
        clientWidth: 120,
        clientHeight: 80,
        getAttribute: (a: string) => (a === "width" ? "120" : "80"),
      } as unknown as SVGSVGElement;
      const out = await captureChartAsPng({ source: fakeSvg });
      expect(out.png).toBe("TESTPAYLOAD");
      expect(out.png.startsWith("data:")).toBe(false);
      expect(out.width).toBe(120);
      expect(out.height).toBe(80);
    } finally {
      stubs.uninstall();
    }
  });

  it("uses an HTMLCanvasElement source directly without going through Image()", async () => {
    const stubs = installBrowserStubs({ pngBase64: "FROMCANVAS" });
    try {
      const fakeCanvas = {
        width: 50,
        height: 25,
        toDataURL: () => "data:image/png;base64,FROMCANVAS",
      };
      const out = await captureChartAsPng({ source: fakeCanvas as unknown as HTMLCanvasElement });
      expect(out.png).toBe("FROMCANVAS");
      expect(out.width).toBe(50);
      expect(out.height).toBe(25);
    } finally {
      stubs.uninstall();
    }
  });

  it("rejects when no DOM is available", async () => {
    const original = g.window;
    g.window = undefined;
    try {
      await expect(
        captureChartAsPng({ source: {} as unknown as SVGSVGElement }),
      ).rejects.toThrow(/browser environment/);
    } finally {
      g.window = original;
    }
  });

  it("propagates Image() decode failures", async () => {
    const stubs = installBrowserStubs({ failImageLoad: true });
    try {
      const fakeSvg = {
        outerHTML: '<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"></svg>',
        clientWidth: 10,
        clientHeight: 10,
        getAttribute: () => "10",
      } as unknown as SVGSVGElement;
      await expect(captureChartAsPng({ source: fakeSvg })).rejects.toThrow(/decode SVG/);
    } finally {
      stubs.uninstall();
    }
  });
});
