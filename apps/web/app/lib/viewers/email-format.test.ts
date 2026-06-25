import { describe, expect, it } from "vitest";
import { parseEmailMessage } from "./email-format";

describe("parseEmailMessage", () => {
  it("normalizes EML headers, body and attachments", () => {
    const eml = [
      "From: Ada <ada@example.com>",
      "To: Team <team@example.com>",
      "Subject: Report",
      'Content-Type: multipart/mixed; boundary="b1"',
      "",
      "--b1",
      "Content-Type: text/plain; charset=utf-8",
      "",
      "Body text",
      "--b1",
      'Content-Type: application/pdf; name="report.pdf"',
      'Content-Disposition: attachment; filename="report.pdf"',
      "Content-Transfer-Encoding: base64",
      "",
      "JVBERi0x",
      "--b1--",
    ].join("\r\n");

    const parsed = parseEmailMessage(new TextEncoder().encode(eml), "message.eml");
    expect(parsed).toMatchObject({
      sourceKind: "eml",
      subject: "Report",
      from: "Ada <ada@example.com>",
      to: "Team <team@example.com>",
      bodyText: "Body text",
    });
    expect(parsed.attachments).toHaveLength(1);
    expect(parsed.attachments[0]).toMatchObject({
      filename: "report.pdf",
      contentType: "application/pdf",
      routeFormat: "pdf",
    });
    expect(new TextDecoder().decode(parsed.attachments[0]?.bytes)).toBe("%PDF-1");
  });

  it("extracts dependency-free MSG fallback text", () => {
    const parsed = parseEmailMessage(
      new TextEncoder().encode("Subject: MSG subject\nFrom: Ops\n\nMSG body"),
      "message.msg"
    );
    expect(parsed).toMatchObject({
      sourceKind: "msg",
      subject: "MSG subject",
      from: "Ops",
      bodyText: "MSG body",
    });
    expect(parsed.diagnostics[0]).toContain("MSG support");
  });
});
