/**
 * POST /api/convert
 *
 * Server-side office-document conversion via LibreOffice headless.
 *
 * Multipart form fields:
 *   - file       (Blob)   the source document bytes
 *   - sourceExt  (string) one of "docx" | "xlsx" | "pptx"
 *   - targetExt  (string) one of "pdf" | "html"
 *   - filename   (string, optional) base filename for the response
 *   - pageRange  (string, optional) PDF-only. LibreOffice page-range
 *                expression like "1-3,5,7" — passed straight through
 *                to the appropriate `*_pdf_Export` filter via
 *                `--convert-to`. Validated against a strict allowlist
 *                so we never pass user input into the shell.
 *
 * Returns the converted bytes as the response body with the correct
 * Content-Type and Content-Disposition. Errors are returned as JSON
 * with a `message` field and a 4xx/5xx status.
 *
 * Operational dependency: the host running this Next.js server must
 * have LibreOffice installed and `soffice` on PATH. See
 * `apps/web/app/api/convert/README.md`.
 */

import { spawn } from "node:child_process";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { NextResponse } from "next/server";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type SourceExt = "docx" | "xlsx" | "pptx";
type TargetExt = "pdf" | "html";

const SOURCE_EXTS: ReadonlySet<SourceExt> = new Set(["docx", "xlsx", "pptx"]);
const TARGET_EXTS: ReadonlySet<TargetExt> = new Set(["pdf", "html"]);

const TARGET_MIME: Record<TargetExt, string> = {
  pdf: "application/pdf",
  html: "text/html; charset=utf-8",
};

/** LibreOffice's PDF export filter is product-specific. */
const PDF_FILTER: Record<SourceExt, string> = {
  docx: "writer_pdf_Export",
  xlsx: "calc_pdf_Export",
  pptx: "impress_pdf_Export",
};

/**
 * Page-range expression accepted by LibreOffice's PDF export filter:
 * comma-separated list of single 1-based indices and `from-to` ranges,
 * with optional surrounding whitespace (e.g. `"1-3, 5, 7"`). We
 * validate strictly so the value can never inject shell metacharacters
 * even though we use `spawn` (no shell) — the JSON we build below also
 * winds up inside soffice's own option parser.
 */
const PAGE_RANGE_RE = /^\s*\d+(\s*-\s*\d+)?(\s*,\s*\d+(\s*-\s*\d+)?)*\s*$/;

/** Hard cap on input size to keep DoS surface narrow (50 MB). */
const MAX_INPUT_BYTES = 50 * 1024 * 1024;

/** Hard cap on conversion wall-clock time (60 seconds). */
const CONVERT_TIMEOUT_MS = 60_000;

function isSourceExt(value: string): value is SourceExt {
  return SOURCE_EXTS.has(value as SourceExt);
}

function isTargetExt(value: string): value is TargetExt {
  return TARGET_EXTS.has(value as TargetExt);
}

export async function POST(request: Request): Promise<Response> {
  let formData: FormData;
  try {
    formData = await request.formData();
  } catch {
    return NextResponse.json({ message: "Invalid multipart body." }, { status: 400 });
  }

  const file = formData.get("file");
  const sourceExtRaw = String(formData.get("sourceExt") ?? "").toLowerCase();
  const targetExtRaw = String(formData.get("targetExt") ?? "").toLowerCase();
  const filenameInput = String(formData.get("filename") ?? "document");
  const pageRangeRaw = formData.get("pageRange");

  if (!(file instanceof Blob)) {
    return NextResponse.json({ message: "Missing `file` field." }, { status: 400 });
  }
  if (!isSourceExt(sourceExtRaw)) {
    return NextResponse.json({ message: `Unsupported sourceExt "${sourceExtRaw}".` }, { status: 400 });
  }
  if (!isTargetExt(targetExtRaw)) {
    return NextResponse.json({ message: `Unsupported targetExt "${targetExtRaw}".` }, { status: 400 });
  }
  if (file.size > MAX_INPUT_BYTES) {
    return NextResponse.json({ message: `Input exceeds ${MAX_INPUT_BYTES} bytes.` }, { status: 413 });
  }

  const sourceExt = sourceExtRaw;
  const targetExt = targetExtRaw;
  const baseFilename = sanitizeFilenameBase(filenameInput);

  let pageRange: string | undefined;
  if (typeof pageRangeRaw === "string" && pageRangeRaw.length > 0) {
    if (targetExt !== "pdf") {
      return NextResponse.json(
        { message: "`pageRange` is only supported when targetExt is 'pdf'." },
        { status: 400 }
      );
    }
    if (!PAGE_RANGE_RE.test(pageRangeRaw)) {
      return NextResponse.json({ message: `Invalid pageRange "${pageRangeRaw}".` }, { status: 400 });
    }
    pageRange = pageRangeRaw.replace(/\s+/g, "");
  }

  let workdir: string | null = null;
  try {
    workdir = await mkdtemp(path.join(tmpdir(), "officeai-convert-"));
    const inputName = `input.${sourceExt}`;
    const inputPath = path.join(workdir, inputName);
    const bytes = new Uint8Array(await file.arrayBuffer());
    await writeFile(inputPath, bytes);

    await runSoffice({
      cwd: workdir,
      inputPath,
      outdir: workdir,
      sourceExt,
      targetExt,
      pageRange,
    });

    const outputPath = path.join(workdir, `input.${targetExt}`);
    const outputBytes = await readFile(outputPath);

    const outName = `${baseFilename}.${targetExt}`;
    return new Response(new Uint8Array(outputBytes), {
      status: 200,
      headers: {
        "Content-Type": TARGET_MIME[targetExt],
        "Content-Disposition": `attachment; filename="${encodeFilenameForHeader(outName)}"`,
        "Cache-Control": "no-store",
      },
    });
  } catch (err) {
    const message =
      err instanceof ConvertError
        ? err.message
        : "Conversion failed. Is LibreOffice installed on the server?";
    const status = err instanceof ConvertError ? err.status : 500;
    console.error("[api/convert] failed:", err);
    return NextResponse.json({ message }, { status });
  } finally {
    if (workdir) {
      await rm(workdir, { recursive: true, force: true }).catch(() => {});
    }
  }
}

class ConvertError extends Error {
  readonly status: number;
  constructor(message: string, status: number) {
    super(message);
    this.status = status;
  }
}

function runSoffice(args: {
  readonly cwd: string;
  readonly inputPath: string;
  readonly outdir: string;
  readonly sourceExt: SourceExt;
  readonly targetExt: TargetExt;
  readonly pageRange?: string;
}): Promise<void> {
  // `soffice` insists on a unique user profile per invocation when
  // multiple conversions can overlap; without it, parallel calls die
  // with a "another instance is running" error.
  const profileDir = path.join(args.cwd, "lo-profile");
  const sofficeBin = process.env.OFFICEAI_SOFFICE_BIN || "soffice";
  const convertTo = buildConvertToArg(args.sourceExt, args.targetExt, args.pageRange);
  const cliArgs = [
    "--headless",
    "--norestore",
    "--nologo",
    "--nofirststartwizard",
    `-env:UserInstallation=file://${profileDir}`,
    "--convert-to",
    convertTo,
    "--outdir",
    args.outdir,
    args.inputPath,
  ];

  return new Promise<void>((resolve, reject) => {
    const child = spawn(sofficeBin, cliArgs, {
      cwd: args.cwd,
      stdio: ["ignore", "pipe", "pipe"],
    });

    let stderr = "";
    child.stderr.on("data", (chunk: Buffer) => {
      stderr += chunk.toString("utf8");
      // Cap so a runaway process doesn't blow up memory.
      if (stderr.length > 64 * 1024) stderr = stderr.slice(-64 * 1024);
    });
    // Drain stdout to avoid back-pressure.
    child.stdout.on("data", () => {});

    const timer = setTimeout(() => {
      child.kill("SIGKILL");
      reject(new ConvertError("Conversion timed out.", 504));
    }, CONVERT_TIMEOUT_MS);

    child.on("error", (err) => {
      clearTimeout(timer);
      // ENOENT — soffice not installed.
      const msg =
        (err as NodeJS.ErrnoException).code === "ENOENT"
          ? "LibreOffice (`soffice`) is not installed on the server."
          : `Failed to spawn LibreOffice: ${err.message}`;
      reject(new ConvertError(msg, 500));
    });
    child.on("close", (code) => {
      clearTimeout(timer);
      if (code === 0) {
        resolve();
        return;
      }
      reject(
        new ConvertError(`LibreOffice exited with code ${code}.${stderr ? ` ${stderr.trim()}` : ""}`, 500)
      );
    });
  });
}

/**
 * Build the `--convert-to` argument. For plain conversions this is
 * just the target extension (e.g. `"pdf"`). When the caller supplied
 * extra filter options we encode them in LibreOffice's documented
 * `target:filter:json-options` form, e.g.
 * `pdf:writer_pdf_Export:{"PageRange":{"type":"string","value":"1-3"}}`.
 */
function buildConvertToArg(
  sourceExt: SourceExt,
  targetExt: TargetExt,
  pageRange: string | undefined
): string {
  if (targetExt !== "pdf" || !pageRange) return targetExt;
  const filter = PDF_FILTER[sourceExt];
  const options = {
    PageRange: { type: "string", value: pageRange },
  };
  return `pdf:${filter}:${JSON.stringify(options)}`;
}

function sanitizeFilenameBase(name: string): string {
  // Strip path components and any trailing extension; keep the stem
  // safe for both URLs and Content-Disposition.
  const stem = name.replace(/\\/g, "/").split("/").pop() ?? "document";
  const dot = stem.lastIndexOf(".");
  const withoutExt = dot > 0 ? stem.slice(0, dot) : stem;
  const cleaned = withoutExt.replace(/[\u0000-\u001f"\\/]/g, "_").trim();
  return cleaned.length > 0 ? cleaned : "document";
}

function encodeFilenameForHeader(name: string): string {
  // Conservative ASCII-only fallback; non-ASCII characters are
  // escaped so the header stays valid even without RFC 5987 support
  // on the client side.
  return name.replace(/[^\x20-\x7e]/g, (ch) => encodeURIComponent(ch));
}
