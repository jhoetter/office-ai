export interface ImageViewerSource {
  readonly filename: string;
  readonly mediaType: string;
  readonly extension: string;
  readonly browserRenderable: boolean;
  readonly displayBytes: Uint8Array;
  readonly diagnostics: ReadonlyArray<string>;
  readonly normalizedPreviewSvg?: string;
}

const DIRECT_IMAGE_TYPES: Readonly<Record<string, string>> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  webp: "image/webp",
  gif: "image/gif",
  svg: "image/svg+xml",
  bmp: "image/bmp",
};

const NORMALIZED_IMAGE_TYPES: Readonly<Record<string, string>> = {
  tif: "image/tiff",
  tiff: "image/tiff",
  heic: "image/heic",
  heif: "image/heif",
};

export async function normalizeImageForViewer(
  bytes: Uint8Array,
  filename: string
): Promise<ImageViewerSource> {
  const extension = extensionFromName(filename);
  const directMediaType = DIRECT_IMAGE_TYPES[extension];
  if (directMediaType) {
    return {
      filename,
      mediaType: directMediaType,
      extension,
      browserRenderable: true,
      displayBytes: bytes,
      diagnostics: [],
    };
  }

  const mediaType = NORMALIZED_IMAGE_TYPES[extension] ?? "application/octet-stream";
  const label = extension ? extension.toUpperCase() : "IMAGE";
  const preview = normalizedPreviewSvg(label, bytes.byteLength);
  return {
    filename,
    mediaType,
    extension,
    browserRenderable: false,
    displayBytes: new TextEncoder().encode(preview),
    normalizedPreviewSvg: preview,
    diagnostics: [
      `${label} is preserved as the working artifact; this browser build exposes a normalized metadata preview.`,
    ],
  };
}

export function extensionFromName(filename: string): string {
  const match = /\.([a-z0-9]+)$/i.exec(filename.trim());
  return match?.[1]?.toLowerCase() ?? "";
}

function normalizedPreviewSvg(label: string, bytes: number): string {
  const escapedLabel = escapeXml(label);
  const escapedBytes = escapeXml(formatBytes(bytes));
  return `<svg xmlns="http://www.w3.org/2000/svg" width="960" height="540" viewBox="0 0 960 540" role="img" aria-label="${escapedLabel} preview">
  <rect width="960" height="540" fill="#f7f8fb"/>
  <rect x="32" y="32" width="896" height="476" rx="8" fill="#ffffff" stroke="#d8dde8"/>
  <text x="480" y="238" text-anchor="middle" font-family="Arial, sans-serif" font-size="56" font-weight="700" fill="#263044">${escapedLabel}</text>
  <text x="480" y="292" text-anchor="middle" font-family="Arial, sans-serif" font-size="20" fill="#657089">Normalized preview</text>
  <text x="480" y="326" text-anchor="middle" font-family="Arial, sans-serif" font-size="16" fill="#8791a5">${escapedBytes}</text>
</svg>`;
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} bytes`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function escapeXml(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
