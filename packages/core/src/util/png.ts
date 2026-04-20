/**
 * Minimal pure-JS PNG encoder. Used to author OLE-spreadsheet preview
 * images and other tiny raster previews where pulling in `sharp` /
 * `node-canvas` would be overkill (and break edge / browser bundles
 * that import this module transitively).
 *
 * Only supports 8-bit RGBA. Pixel layout matches the input order: row
 * 0 first, top-left to bottom-right, four bytes per pixel `[r, g, b, a]`.
 *
 * The encoder uses filter type 0 (None) on every scanline because the
 * previews we generate are small and simple enough that compression
 * gains from filtering are negligible.
 *
 * The IDAT zlib stream is built from uncompressed deflate blocks (BTYPE=00)
 * so we don't need `node:zlib` and the module can be loaded in browser /
 * edge bundles transitively.
 */
const PNG_SIGNATURE = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

const CRC_TABLE: Uint32Array = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(bytes: Uint8Array): number {
  let c = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) {
    const idx = (c ^ (bytes[i] ?? 0)) & 0xff;
    c = (CRC_TABLE[idx] ?? 0) ^ (c >>> 8);
  }
  return (c ^ 0xffffffff) >>> 0;
}

function adler32(bytes: Uint8Array): number {
  let a = 1;
  let b = 0;
  const MOD = 65521;
  for (let i = 0; i < bytes.length; i++) {
    a = (a + (bytes[i] ?? 0)) % MOD;
    b = (b + a) % MOD;
  }
  return ((b << 16) | a) >>> 0;
}

function zlibStored(data: Uint8Array): Uint8Array {
  const MAX_BLOCK = 0xffff;
  const blockCount = Math.max(1, Math.ceil(data.length / MAX_BLOCK));
  const out = new Uint8Array(2 + data.length + blockCount * 5 + 4);
  let p = 0;
  out[p++] = 0x78;
  out[p++] = 0x01;
  for (let off = 0; off < data.length || off === 0; off += MAX_BLOCK) {
    const remaining = data.length - off;
    const len = Math.min(MAX_BLOCK, remaining);
    const isLast = off + len >= data.length;
    out[p++] = isLast ? 0x01 : 0x00;
    out[p++] = len & 0xff;
    out[p++] = (len >>> 8) & 0xff;
    const nlen = ~len & 0xffff;
    out[p++] = nlen & 0xff;
    out[p++] = (nlen >>> 8) & 0xff;
    if (len > 0) {
      out.set(data.subarray(off, off + len), p);
      p += len;
    }
    if (isLast) break;
  }
  const adler = adler32(data);
  out[p++] = (adler >>> 24) & 0xff;
  out[p++] = (adler >>> 16) & 0xff;
  out[p++] = (adler >>> 8) & 0xff;
  out[p++] = adler & 0xff;
  return out.subarray(0, p);
}

function chunk(type: string, data: Uint8Array): Uint8Array {
  const length = data.length;
  const buf = new Uint8Array(8 + length + 4);
  const dv = new DataView(buf.buffer);
  dv.setUint32(0, length, false);
  buf[4] = type.charCodeAt(0);
  buf[5] = type.charCodeAt(1);
  buf[6] = type.charCodeAt(2);
  buf[7] = type.charCodeAt(3);
  buf.set(data, 8);
  const crcBytes = buf.subarray(4, 8 + length);
  dv.setUint32(8 + length, crc32(crcBytes), false);
  return buf;
}

export interface PngEncodeArgs {
  /** Width in pixels. */
  readonly width: number;
  /** Height in pixels. */
  readonly height: number;
  /** RGBA pixels, row-major, length = width * height * 4. */
  readonly rgba: Uint8Array;
}

/**
 * Encode an 8-bit RGBA pixel grid into a valid PNG file (returned as
 * raw bytes). Throws if the buffer length is wrong.
 */
export function encodePng(args: PngEncodeArgs): Uint8Array {
  const { width, height, rgba } = args;
  if (rgba.length !== width * height * 4) {
    throw new Error(`encodePng: expected ${width * height * 4} bytes, got ${rgba.length}`);
  }

  const ihdr = new Uint8Array(13);
  const ihdrView = new DataView(ihdr.buffer);
  ihdrView.setUint32(0, width, false);
  ihdrView.setUint32(4, height, false);
  ihdr[8] = 8;
  ihdr[9] = 6;
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;

  const stride = width * 4;
  const filtered = new Uint8Array((stride + 1) * height);
  for (let y = 0; y < height; y++) {
    filtered[y * (stride + 1)] = 0;
    filtered.set(rgba.subarray(y * stride, y * stride + stride), y * (stride + 1) + 1);
  }

  const idat = zlibStored(filtered);

  const ihdrChunk = chunk("IHDR", ihdr);
  const idatChunk = chunk("IDAT", idat);
  const iendChunk = chunk("IEND", new Uint8Array(0));

  const out = new Uint8Array(PNG_SIGNATURE.length + ihdrChunk.length + idatChunk.length + iendChunk.length);
  let offset = 0;
  out.set(PNG_SIGNATURE, offset); offset += PNG_SIGNATURE.length;
  out.set(ihdrChunk, offset); offset += ihdrChunk.length;
  out.set(idatChunk, offset); offset += idatChunk.length;
  out.set(iendChunk, offset);
  return out;
}
