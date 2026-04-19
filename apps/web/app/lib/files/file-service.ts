/**
 * Cross-product file service.
 *
 * Provides three things:
 *   1. `openFilePicker(accept)` — wraps File System Access API where
 *      available, falls back to a hidden `<input type="file">`.
 *   2. `saveFile(handle, blob, fallbackName)` — writes back to the
 *      original on-disk file when we have a handle, otherwise falls
 *      back to a download.
 *   3. `downloadBlob(blob, name)` — pure download helper used for
 *      Export and as the no-handle Save fallback.
 *
 * No persistence (no IndexedDB, no recent files) — the user opens
 * files via Open or drag-drop today and from S3 later. This service
 * stays the single integration point so swapping storage doesn't
 * require touching every editor.
 */

export interface OpenedFile {
  readonly name: string;
  readonly bytes: Uint8Array;
  /** Present when the browser supports the File System Access API
   * and the user picked the file via that path. We hold onto it so
   * a subsequent `Save` can write the bytes back in place. */
  readonly handle?: FileSystemFileHandle;
}

interface PickerType {
  readonly description: string;
  /** Map of MIME → string[] of extensions. */
  readonly accept: Record<string, string[]>;
}

interface ShowOpenFilePickerOptions {
  readonly multiple?: boolean;
  readonly types?: PickerType[];
}

/** Subset of the File System Access API we actually use. Typed
 * locally to avoid lib.dom additions that may not be present. */
interface FileSystemAccessWindow {
  showOpenFilePicker?: (options?: ShowOpenFilePickerOptions) => Promise<FileSystemFileHandle[]>;
}

function fsa(): FileSystemAccessWindow | null {
  if (typeof window === "undefined") return null;
  return window as unknown as FileSystemAccessWindow;
}

/**
 * Open a file using the File System Access API where possible.
 * Returns `null` if the user cancelled. Throws on actual errors.
 *
 * `accept` is a map of human description → array of file extensions
 * (each with leading dot, e.g. `.docx`).
 */
export async function openFilePickerWithFsa(
  description: string,
  mimeToExt: Record<string, string[]>
): Promise<OpenedFile | null> {
  const w = fsa();
  if (!w?.showOpenFilePicker) return null;
  try {
    const [handle] = await w.showOpenFilePicker({
      multiple: false,
      types: [{ description, accept: mimeToExt }],
    });
    if (!handle) return null;
    const file = await handle.getFile();
    const bytes = new Uint8Array(await file.arrayBuffer());
    return { name: file.name, bytes, handle };
  } catch (err) {
    // AbortError fires when the user dismisses the picker.
    if (err instanceof DOMException && err.name === "AbortError") return null;
    throw err;
  }
}

/**
 * Trigger a hidden `<input type="file">` for browsers without File
 * System Access. The accept string is the standard `<input accept>`
 * format ("comma-separated MIMEs and extensions").
 */
export function pickFileLegacy(accept: string): Promise<OpenedFile | null> {
  return new Promise((resolve, reject) => {
    if (typeof document === "undefined") {
      resolve(null);
      return;
    }
    const input = document.createElement("input");
    input.type = "file";
    input.accept = accept;
    input.style.position = "fixed";
    input.style.left = "-9999px";
    document.body.appendChild(input);
    let settled = false;
    const cleanup = () => {
      input.removeEventListener("change", onChange);
      input.removeEventListener("cancel", onCancel);
      input.remove();
    };
    const onChange = async () => {
      if (settled) return;
      settled = true;
      try {
        const file = input.files?.[0] ?? null;
        if (!file) {
          cleanup();
          resolve(null);
          return;
        }
        const bytes = new Uint8Array(await file.arrayBuffer());
        cleanup();
        resolve({ name: file.name, bytes });
      } catch (err) {
        cleanup();
        reject(err);
      }
    };
    const onCancel = () => {
      if (settled) return;
      settled = true;
      cleanup();
      resolve(null);
    };
    input.addEventListener("change", onChange);
    input.addEventListener("cancel", onCancel);
    input.click();
  });
}

/**
 * Combined open: try FSA first, fall back to legacy. The `accept`
 * string is the legacy `<input accept>` value; `mimeToExt` is the
 * structured FSA descriptor. Both must agree.
 */
export async function openFile(opts: {
  description: string;
  mimeToExt: Record<string, string[]>;
  accept: string;
}): Promise<OpenedFile | null> {
  const fsaResult = await openFilePickerWithFsa(opts.description, opts.mimeToExt);
  if (fsaResult) return fsaResult;
  // FSA might be unsupported (no `showOpenFilePicker`). Note: if FSA
  // is supported but the user cancelled, openFilePickerWithFsa
  // returns `null` and we'd fall through to a second native picker —
  // disambiguate by checking the API surface.
  if (fsa()?.showOpenFilePicker) return null;
  return pickFileLegacy(opts.accept);
}

/**
 * Save bytes back to the original file (FSA writable stream) when we
 * have a handle, otherwise fall back to a download. Returns true if
 * we wrote in place.
 */
export async function saveFile(
  bytes: Uint8Array,
  fallbackName: string,
  mime: string,
  handle?: FileSystemFileHandle
): Promise<boolean> {
  if (handle && "createWritable" in handle) {
    try {
      const writable = await handle.createWritable();
      // The FSA writable accepts ArrayBuffer / typed arrays / Blob /
      // string. Hand it the underlying buffer so we don't pay for an
      // extra Blob allocation.
      await writable.write(new Blob([bytes as BlobPart], { type: mime }));
      await writable.close();
      return true;
    } catch (err) {
      // Permission revoked or storage error — fall through.
      console.warn("[file-service] FSA save failed, falling back to download", err);
    }
  }
  downloadBlob(new Blob([bytes as BlobPart], { type: mime }), fallbackName);
  return false;
}

/** Plain download helper. Used for Export and the no-handle Save
 * fallback. */
export function downloadBlob(blob: Blob, filename: string): void {
  if (typeof document === "undefined") return;
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  // Allow time for the download to start before revoking.
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/** Per-product accept descriptors. Centralised so we never disagree
 * about which extensions/MIMEs are valid. */
export const PRODUCT_FILE_TYPES: Record<
  "docx" | "xlsx" | "pptx",
  {
    description: string;
    mimeToExt: Record<string, string[]>;
    accept: string;
    primaryMime: string;
    primaryExt: string;
  }
> = {
  docx: {
    description: "Word document",
    mimeToExt: {
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document": [".docx"],
    },
    accept: ".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    primaryMime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    primaryExt: "docx",
  },
  xlsx: {
    description: "Excel workbook",
    mimeToExt: {
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
    },
    accept: ".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    primaryMime: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    primaryExt: "xlsx",
  },
  pptx: {
    description: "PowerPoint presentation",
    mimeToExt: {
      "application/vnd.openxmlformats-officedocument.presentationml.presentation": [".pptx"],
    },
    accept: ".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation",
    primaryMime: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    primaryExt: "pptx",
  },
};

/** True when the active environment supports File System Access. */
export function supportsFsa(): boolean {
  return Boolean(fsa()?.showOpenFilePicker);
}
