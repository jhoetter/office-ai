import { describe, expect, it } from "vitest";
import { PDFDocument, StandardFonts } from "pdf-lib";
import { addAnnotations } from "./writer.js";
import { exportXfdf, importXfdf } from "./xfdf.js";

const buildPdf = async (n = 2): Promise<Uint8Array> => {
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  for (let i = 1; i <= n; i++) {
    const page = pdf.addPage([612, 792]);
    page.drawText(`Page ${i}`, { x: 50, y: 720, size: 18, font });
  }
  return pdf.save();
};

const annotsCount = async (buf: Uint8Array, pageIndex: number): Promise<number> => {
  const pdf = await PDFDocument.load(buf);
  const node = pdf.getPage(pageIndex).node;
  const annots = node.Annots();
  return annots ? annots.size() : 0;
};

describe("addAnnotations", () => {
  it("adds a highlight annotation onto the requested page", async () => {
    const out = await addAnnotations(await buildPdf(), {
      annotations: [
        {
          kind: "highlight",
          pageNumber: 1,
          rect: [50, 700, 200, 720],
          contents: "important",
          author: "Tester",
        },
      ],
    });
    expect(await annotsCount(out, 0)).toBe(1);
    expect(await annotsCount(out, 1)).toBe(0);
  });

  it("adds a sticky note annotation", async () => {
    const out = await addAnnotations(await buildPdf(), {
      annotations: [
        {
          kind: "sticky-note",
          pageNumber: 2,
          rect: [100, 600, 120, 620],
          contents: "todo: review",
        },
      ],
    });
    expect(await annotsCount(out, 1)).toBe(1);
  });

  it("adds a free-text annotation", async () => {
    const out = await addAnnotations(await buildPdf(), {
      annotations: [
        {
          kind: "free-text",
          pageNumber: 1,
          rect: [50, 600, 250, 640],
          contents: "Inline comment",
          fontSize: 14,
        },
      ],
    });
    expect(await annotsCount(out, 0)).toBe(1);
  });

  it("adds a link annotation pointing to a URL", async () => {
    const out = await addAnnotations(await buildPdf(), {
      annotations: [
        {
          kind: "link",
          pageNumber: 1,
          rect: [50, 500, 200, 520],
          url: "https://example.com",
        },
      ],
    });
    expect(await annotsCount(out, 0)).toBe(1);
  });

  it("adds a link annotation with a goto-page destination", async () => {
    const out = await addAnnotations(await buildPdf(3), {
      annotations: [
        {
          kind: "link",
          pageNumber: 1,
          rect: [50, 500, 200, 520],
          destPage: 3,
        },
      ],
    });
    expect(await annotsCount(out, 0)).toBe(1);
  });

  it("rejects annotations on out-of-range pages", async () => {
    await expect(
      addAnnotations(await buildPdf(), {
        annotations: [{ kind: "sticky-note", pageNumber: 99, rect: [0, 0, 10, 10], contents: "x" }],
      }),
    ).rejects.toThrow(/out of range/);
  });
});

describe("xfdf JSON form", () => {
  it("round-trips through export + import", () => {
    const annots = [
      {
        kind: "highlight" as const,
        pageNumber: 1,
        rect: [10, 20, 30, 40] as const,
        contents: "h",
      },
    ];
    const doc = exportXfdf(annots);
    expect(importXfdf(doc)).toEqual(annots);
  });
});
