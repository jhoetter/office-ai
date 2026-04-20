/**
 * GET /api/sample-files
 *
 * Lists the office-document samples bundled under
 * `apps/web/public/sample-files/`. The home page renders these as a
 * lightweight "file system" so users can jump straight into a real
 * document without first having to pick one from disk.
 *
 * The bytes themselves are served by Next's static handler under
 * `/sample-files/<name>`; this route only returns metadata so the
 * client can render the listing without parsing a directory itself.
 */

import { readdir, stat } from "node:fs/promises";
import path from "node:path";
import { NextResponse } from "next/server";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type Kind = "docx" | "xlsx" | "pptx" | "pdf";

const EXT_TO_KIND: Record<string, Kind> = {
  ".docx": "docx",
  ".xlsx": "xlsx",
  ".pptx": "pptx",
  ".pdf": "pdf",
};

interface SampleFileEntry {
  readonly name: string;
  readonly url: string;
  readonly kind: Kind;
  readonly size: number;
  readonly modifiedAt: string;
}

export async function GET(): Promise<NextResponse> {
  // `process.cwd()` for the Next dev/prod server is `apps/web`,
  // matching how `outputFileTracingRoot` is anchored. This keeps the
  // lookup the same in both `next dev` and a built standalone server.
  const dir = path.join(process.cwd(), "public", "sample-files");
  let entries: string[];
  try {
    entries = await readdir(dir);
  } catch (err) {
    const code = (err as NodeJS.ErrnoException)?.code;
    if (code === "ENOENT") {
      return NextResponse.json({ files: [] satisfies SampleFileEntry[] });
    }
    return NextResponse.json({ message: err instanceof Error ? err.message : String(err) }, { status: 500 });
  }

  const files: SampleFileEntry[] = [];

  const collect = async (subdir: string): Promise<void> => {
    const dirToWalk = subdir.length > 0 ? path.join(dir, subdir) : dir;
    let dirEntries: string[];
    try {
      dirEntries = await readdir(dirToWalk);
    } catch {
      return;
    }
    for (const name of dirEntries) {
      const ext = path.extname(name).toLowerCase();
      const kind = EXT_TO_KIND[ext];
      const full = path.join(dirToWalk, name);
      let info;
      try {
        info = await stat(full);
      } catch {
        continue;
      }
      if (!kind) {
        if (info.isDirectory() && subdir.length === 0) await collect(name);
        continue;
      }
      if (!info.isFile()) continue;
      const relativeName = subdir.length > 0 ? `${subdir}/${name}` : name;
      files.push({
        name,
        url: `/sample-files/${relativeName.split("/").map(encodeURIComponent).join("/")}`,
        kind,
        size: info.size,
        modifiedAt: info.mtime.toISOString(),
      });
    }
  };

  await collect("");

  files.sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: "base" }));
  return NextResponse.json({ files });
}
