import type { WebOfficeFormat } from "@/lib/sessions/web-sessions";

export interface EmailHeader {
  readonly name: string;
  readonly value: string;
}

export interface EmailAttachment {
  readonly id: string;
  readonly filename: string;
  readonly contentType: string;
  readonly size: number;
  readonly bytes: Uint8Array;
  readonly routeFormat: WebOfficeFormat | null;
}

export interface EmailMessage {
  readonly sourceKind: "eml" | "msg";
  readonly subject: string;
  readonly from: string;
  readonly to: string;
  readonly date: string;
  readonly headers: ReadonlyArray<EmailHeader>;
  readonly bodyText: string;
  readonly bodyHtml: string;
  readonly attachments: ReadonlyArray<EmailAttachment>;
  readonly diagnostics: ReadonlyArray<string>;
}

interface ParsedPart {
  readonly headers: ReadonlyArray<EmailHeader>;
  readonly body: string;
}

export function parseEmailMessage(bytes: Uint8Array, filename: string): EmailMessage {
  const lower = filename.toLowerCase();
  if (lower.endsWith(".msg")) return parseMsgFallback(bytes);
  return parseEml(bytes);
}

function parseEml(bytes: Uint8Array): EmailMessage {
  const raw = decodeText(bytes);
  const root = parsePart(raw);
  const rootHeaders = root.headers;
  const rootContentType = headerValue(rootHeaders, "content-type");
  const boundary = parameterValue(rootContentType, "boundary");
  const diagnostics: string[] = [];
  let bodyText = "";
  let bodyHtml = "";
  const attachments: EmailAttachment[] = [];

  if (boundary) {
    const parts = splitMultipart(root.body, boundary).map(parsePart);
    parts.forEach((part, index) => {
      const contentType = headerValue(part.headers, "content-type") || "text/plain";
      const disposition = headerValue(part.headers, "content-disposition");
      const filename = parameterValue(disposition, "filename") || parameterValue(contentType, "name") || "";
      const decoded = decodePartBody(part);
      const isAttachment = /attachment/i.test(disposition) || filename.length > 0;
      if (isAttachment) {
        const attachmentName = filename || `attachment-${index + 1}`;
        attachments.push({
          id: `att_${index + 1}`,
          filename: attachmentName,
          contentType: contentType.split(";")[0]?.trim() || "application/octet-stream",
          size: decoded.byteLength,
          bytes: decoded,
          routeFormat: routeFormatForName(attachmentName),
        });
        return;
      }
      const decodedText = decodeText(decoded);
      if (/text\/html/i.test(contentType) && !bodyHtml) bodyHtml = decodedText;
      if (/text\/plain/i.test(contentType) && !bodyText) bodyText = decodedText;
    });
  } else {
    const decoded = decodePartBody(root);
    const contentType = rootContentType || "text/plain";
    if (/text\/html/i.test(contentType)) bodyHtml = decodeText(decoded);
    else bodyText = decodeText(decoded);
  }

  if (!bodyText && bodyHtml) bodyText = stripHtml(bodyHtml);
  if (!bodyText && !bodyHtml) diagnostics.push("No readable message body was found.");

  return {
    sourceKind: "eml",
    subject: decodeHeader(headerValue(rootHeaders, "subject")) || "(no subject)",
    from: decodeHeader(headerValue(rootHeaders, "from")),
    to: decodeHeader(headerValue(rootHeaders, "to")),
    date: decodeHeader(headerValue(rootHeaders, "date")),
    headers: rootHeaders,
    bodyText,
    bodyHtml,
    attachments,
    diagnostics,
  };
}

function parseMsgFallback(bytes: Uint8Array): EmailMessage {
  const utf8 = decodeText(bytes).replace(/\u0000/g, "");
  const utf16 = new TextDecoder("utf-16le", { fatal: false }).decode(bytes).replace(/\u0000/g, "");
  const text = utf8.length > utf16.length ? utf8 : utf16;
  const headerLines = text
    .split(/\r?\n/)
    .filter((line) => /^(subject|from|to|date):/i.test(line))
    .slice(0, 24);
  const headers = parseHeaders(headerLines.join("\r\n"));
  const bodyText = text
    .split(/\r?\n/)
    .filter((line) => line.trim().length > 0 && !/^(subject|from|to|date):/i.test(line))
    .slice(0, 80)
    .join("\n");
  return {
    sourceKind: "msg",
    subject: decodeHeader(headerValue(headers, "subject")) || "(MSG message)",
    from: decodeHeader(headerValue(headers, "from")),
    to: decodeHeader(headerValue(headers, "to")),
    date: decodeHeader(headerValue(headers, "date")),
    headers,
    bodyText,
    bodyHtml: "",
    attachments: [],
    diagnostics: [
      "MSG support uses a dependency-free metadata/text fallback and preserves the source artifact.",
    ],
  };
}

function parsePart(raw: string): ParsedPart {
  const [headerText, ...bodyParts] = raw.split(/\r?\n\r?\n/);
  return {
    headers: parseHeaders(headerText ?? ""),
    body: bodyParts.join("\n\n"),
  };
}

function parseHeaders(raw: string): ReadonlyArray<EmailHeader> {
  const unfolded = raw.replace(/\r?\n[ \t]+/g, " ");
  return unfolded
    .split(/\r?\n/)
    .map((line) => {
      const index = line.indexOf(":");
      if (index < 0) return null;
      return {
        name: line.slice(0, index).trim(),
        value: line.slice(index + 1).trim(),
      };
    })
    .filter((header): header is EmailHeader => Boolean(header?.name));
}

function splitMultipart(body: string, boundary: string): ReadonlyArray<string> {
  const marker = `--${boundary}`;
  return body
    .split(marker)
    .map((part) => part.replace(/^\r?\n/, "").replace(/\r?\n$/, ""))
    .filter((part) => part.trim().length > 0 && part.trim() !== "--");
}

function decodePartBody(part: ParsedPart): Uint8Array {
  const transfer = headerValue(part.headers, "content-transfer-encoding").toLowerCase();
  if (transfer === "base64") return base64ToBytes(part.body.replace(/\s+/g, ""));
  if (transfer === "quoted-printable") return encodeText(decodeQuotedPrintable(part.body));
  return encodeText(part.body.replace(/\r?\n$/, ""));
}

function headerValue(headers: ReadonlyArray<EmailHeader>, name: string): string {
  return headers.find((header) => header.name.toLowerCase() === name.toLowerCase())?.value ?? "";
}

function parameterValue(value: string, parameter: string): string {
  const re = new RegExp(`${parameter}\\*?=(?:"([^"]*)"|([^;]+))`, "i");
  const match = re.exec(value);
  return decodeHeader((match?.[1] ?? match?.[2] ?? "").trim());
}

function decodeHeader(value: string): string {
  return value.replace(/=\?([^?]+)\?([bq])\?([^?]+)\?=/gi, (_match, charset, encoding, encoded) => {
    const bytes =
      String(encoding).toLowerCase() === "b"
        ? base64ToBytes(String(encoded))
        : encodeText(
            String(encoded)
              .replace(/_/g, " ")
              .replace(/=([0-9a-f]{2})/gi, (_, hex) => String.fromCharCode(Number.parseInt(hex, 16)))
          );
    try {
      return new TextDecoder(String(charset), { fatal: false }).decode(bytes);
    } catch {
      return decodeText(bytes);
    }
  });
}

function decodeQuotedPrintable(value: string): string {
  return value
    .replace(/=\r?\n/g, "")
    .replace(/=([0-9a-f]{2})/gi, (_match, hex) => String.fromCharCode(Number.parseInt(hex, 16)));
}

function routeFormatForName(name: string): WebOfficeFormat | null {
  const lower = name.toLowerCase();
  if (lower.endsWith(".docx")) return "docx";
  if (lower.endsWith(".xlsx")) return "xlsx";
  if (lower.endsWith(".pptx")) return "pptx";
  if (lower.endsWith(".pdf")) return "pdf";
  if (lower.endsWith(".eml") || lower.endsWith(".msg")) return "email";
  if (/\.(png|jpe?g|webp|gif|svg|bmp|tiff?|heic|heif)$/i.test(lower)) return "image";
  return null;
}

function stripHtml(html: string): string {
  return html
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|div|li|tr|h[1-6])>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .trim();
}

function decodeText(bytes: Uint8Array): string {
  return new TextDecoder("utf-8", { fatal: false }).decode(bytes);
}

function encodeText(value: string): Uint8Array {
  return new TextEncoder().encode(value);
}

function base64ToBytes(value: string): Uint8Array {
  if (typeof atob === "function") {
    const binary = atob(value);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
    return bytes;
  }
  const maybeBuffer = (globalThis as unknown as { Buffer?: typeof Buffer }).Buffer;
  if (maybeBuffer) return new Uint8Array(maybeBuffer.from(value, "base64"));
  throw new Error("No base64 decoder is available.");
}
