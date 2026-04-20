import { readFile } from "node:fs/promises";
import { resolve } from "node:path";

export const FIXTURE_DIR = resolve(__dirname, "../../../fixtures/pdf");

export async function loadFixture(name: string): Promise<Uint8Array> {
  const buf = await readFile(resolve(FIXTURE_DIR, name));
  return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
}

/** Quick magic-byte sanity check — every PDF file starts with `%PDF-`. */
export function isPdfBytes(bytes: Uint8Array): boolean {
  return (
    bytes.length > 5 &&
    bytes[0] === 0x25 && // %
    bytes[1] === 0x50 && // P
    bytes[2] === 0x44 && // D
    bytes[3] === 0x46 && // F
    bytes[4] === 0x2d // -
  );
}
