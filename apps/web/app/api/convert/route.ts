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

  if (!(file instanceof Blob)) {
    return NextResponse.json({ message: "Missing `file` field." }, { status: 400 });
  }
  if (!isSourceExt(sourceExtRaw)) {
    return NextResponse.json(
      { message: `Unsupported sourceExt "${sourceExtRaw}".` },
      { status: 400 }
    );
  }
  if (!isTargetExt(targetExtRaw)) {
    return NextResponse.json(
      { message: `Unsupported targetExt "${targetExtRaw}".` },
      { status: 400 }
    );
  }
  if (file.size > MAX_INPUT_BYTES) {
    return NextResponse.json(
      { message: `Input exceeds ${MAX_INPUT_BYTES} bytes.` },
      { status: 413 }
    );
  }

  const sourceExt = sourceExtRaw;
  const targetExt = targetExtRaw;
  const baseFilename = sanitizeFilenameBase(filenameInput);

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
      targetExt,
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
  readonly targetExt: TargetExt;
}): Promise<void> {
  // `soffice` insists on a unique user profile per invocation when
  // multiple conversions can overlap; without it, parallel calls die
  // with a "another instance is running" error.
  const profileDir = path.join(args.cwd, "lo-profile");
  const sofficeBin = process.env.OFFICEAI_SOFFICE_BIN || "soffice";
  const cliArgs = [
    "--headless",
    "--norestore",
    "--nologo",
    "--nofirststartwizard",
    `-env:UserInstallation=file://${profileDir}`,
    "--convert-to",
    args.targetExt,
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
        new ConvertError(
          `LibreOffice exited with code ${code}.${stderr ? ` ${stderr.trim()}` : ""}`,
          500
        )
      );
    });
  });
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
