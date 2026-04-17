#!/usr/bin/env node
/**
 * OOXML XSD fetcher — populates `vendor/ooxml-xsd/` with the ECMA-376
 * 5th-edition (December 2016) **Transitional** WordprocessingML / DrawingML /
 * SpreadsheetML / VML schemas, plus the W3C `xml.xsd` they depend on.
 *
 * We pull the canonical ECMA bundle (ECMA-376 Part 4, "Transitional Migration
 * Features"), then unpack the inner `OfficeOpenXML-XMLSchema-Transitional.zip`
 * into `vendor/ooxml-xsd/`. Transitional (not Strict) is the schema set Word /
 * LibreOffice / Google actually emit in real-world `.docx` exports today, so
 * it's the right target for our serializer-validation gate (W9 / Theme D4).
 *
 * Sources (pinned):
 *   1. https://www.ecma-international.org/wp-content/uploads/ECMA-376-4_5th_edition_december_2016.zip
 *      SHA256 bd25da1109f73762356596918bf5ff8b74a1331642dba5f1c1d1dfc6bed34ecd  (~8.4 MB)
 *      → contains OfficeOpenXML-XMLSchema-Transitional.zip → 26 .xsd files.
 *   2. https://www.w3.org/2001/xml.xsd
 *      SHA256 61960fb3131e38022caad5360e2f33a3382578ab3c80cd58bd74320ede61b20c  (~9 KB)
 *      → the W3C XML namespace schema. The OOXML XSDs `<xsd:import>` it
 *        without a `schemaLocation`, expecting the validator to resolve it
 *        from a built-in catalog. xmllint has no such catalog by default,
 *        so we ship the xml.xsd locally and patch the imports below.
 *
 * Post-process: every `<xsd:import namespace="…/XML/1998/namespace"/>` is
 * rewritten to `<xsd:import namespace="…/XML/1998/namespace" schemaLocation=
 * "xml.xsd"/>`. Without this fixup `xmllint` cannot compile the schema graph
 * (you get "QName '{…}space' does not resolve to an attribute declaration").
 *
 * License: ECMA-376 (5th Ed.) is published by Ecma International and is also
 * available under Microsoft's "Open Specification Promise" — both are
 * MIT-compatible for ship-time usage as build-time validation assets. The
 * W3C xml.xsd is "W3C Software Notice and Document License", also
 * MIT-compatible. We never redistribute these files (they are fetched on
 * demand and gitignored under `vendor/ooxml-xsd/.gitignore`).
 *
 * Idempotent: skips the network round-trip when `vendor/ooxml-xsd/wml.xsd`
 * already exists. Hard-fails if the SHA-256 of any downloaded file does not
 * match the pinned hash above (supply-chain guard).
 *
 * Run via `make xsd-fetch` or `node scripts/fetch-ooxml-xsd.mjs`.
 */

import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const ECMA_URL =
  "https://www.ecma-international.org/wp-content/uploads/ECMA-376-4_5th_edition_december_2016.zip";
const ECMA_SHA256 = "bd25da1109f73762356596918bf5ff8b74a1331642dba5f1c1d1dfc6bed34ecd";
const INNER_ZIP_NAME = "OfficeOpenXML-XMLSchema-Transitional.zip";

const W3C_XML_XSD_URL = "https://www.w3.org/2001/xml.xsd";
const W3C_XML_XSD_SHA256 = "61960fb3131e38022caad5360e2f33a3382578ab3c80cd58bd74320ede61b20c";

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, "..");
const vendorDir = resolve(root, "vendor/ooxml-xsd");
const cacheDir = join(vendorDir, ".cache");
const sentinel = join(vendorDir, "wml.xsd");

function log(msg) {
  console.log(`xsd-fetch: ${msg}`);
}

function fail(msg) {
  console.error(`xsd-fetch: ${msg}`);
  process.exit(1);
}

async function downloadTo(url, dest) {
  log(`fetching ${url}`);
  const res = await fetch(url, { redirect: "follow" });
  if (!res.ok) fail(`HTTP ${res.status} fetching ${url}`);
  const buf = Buffer.from(await res.arrayBuffer());
  writeFileSync(dest, buf);
  return buf;
}

function sha256(buf) {
  return createHash("sha256").update(buf).digest("hex");
}

function unzipSync(zipPath, outDir, members) {
  // We shell out to the system `unzip` (universally available on macOS and
  // Linux CI). Adding a JS unzip dependency just for build-time XSD fetching
  // would be overkill — the same pattern that we use for `xmllint` in the
  // validate script.
  mkdirSync(outDir, { recursive: true });
  const args = ["-o", "-q", zipPath, ...(members ?? []), "-d", outDir];
  const r = spawnSync("unzip", args, { encoding: "utf8" });
  if (r.error) fail(`unzip not available on PATH: ${r.error.message}`);
  if (r.status !== 0) {
    fail(`unzip failed (exit ${r.status}): ${r.stderr || r.stdout || "(no output)"}`);
  }
}

async function fetchVerifiedFile({ url, dest, expectedSha256, label }) {
  let buf;
  if (existsSync(dest)) {
    log(`reusing cached ${label} at ${dest}`);
    buf = readFileSync(dest);
  } else {
    buf = await downloadTo(url, dest);
  }
  const observed = sha256(buf);
  if (observed !== expectedSha256) {
    rmSync(dest, { force: true });
    fail(
      `SHA-256 mismatch for ${label}\n  expected: ${expectedSha256}\n  observed: ${observed}\nDeleted the corrupt cache; please re-run.`
    );
  }
  log(`${label} SHA-256 OK (${observed.slice(0, 12)}…)`);
  return buf;
}

function patchXmlNsImports(xsdDir) {
  // The ECMA XSDs import the XML 1998 namespace without a schemaLocation,
  // which xmllint cannot resolve. Rewrite each `<xsd:import>` so it points
  // at the local xml.xsd we just placed alongside.
  const NEEDLE_RE = /<xsd:import\s+namespace="http:\/\/www\.w3\.org\/XML\/1998\/namespace"\s*\/>/g;
  const REPLACEMENT =
    '<xsd:import namespace="http://www.w3.org/XML/1998/namespace" schemaLocation="xml.xsd"/>';
  let touched = 0;
  for (const f of readdirSync(xsdDir)) {
    if (!f.endsWith(".xsd") || f === "xml.xsd") continue;
    const path = join(xsdDir, f);
    const before = readFileSync(path, "utf8");
    if (!NEEDLE_RE.test(before)) continue;
    NEEDLE_RE.lastIndex = 0;
    const after = before.replace(NEEDLE_RE, REPLACEMENT);
    if (after !== before) {
      writeFileSync(path, after, "utf8");
      touched++;
    }
  }
  log(`patched xml-ns import in ${touched} schema(s)`);
}

async function main() {
  if (existsSync(sentinel)) {
    log(`vendor/ooxml-xsd/ already populated (found wml.xsd) — skipping fetch.`);
    return 0;
  }

  mkdirSync(vendorDir, { recursive: true });
  mkdirSync(cacheDir, { recursive: true });

  // 1. ECMA-376 Part 4 (Transitional) bundle.
  const outerZip = join(cacheDir, "ECMA-376-4.zip");
  await fetchVerifiedFile({
    url: ECMA_URL,
    dest: outerZip,
    expectedSha256: ECMA_SHA256,
    label: "ECMA-376 part 4 zip",
  });

  // Stage 1: pull the inner Transitional XSD zip out of the ECMA bundle.
  unzipSync(outerZip, cacheDir, [INNER_ZIP_NAME]);
  const innerZip = join(cacheDir, INNER_ZIP_NAME);
  if (!existsSync(innerZip)) {
    fail(`inner zip ${INNER_ZIP_NAME} not present after extracting ${outerZip}`);
  }

  // Stage 2: unpack the .xsd files into vendor/ooxml-xsd/ (flat layout — the
  // schemas reference each other via plain `<xs:import schemaLocation="…"/>`
  // entries that resolve relative to the current file).
  unzipSync(innerZip, vendorDir);

  // 2. W3C xml.xsd (separate fetch — needed for xs:import resolution).
  const xmlXsdDest = join(vendorDir, "xml.xsd");
  if (!existsSync(xmlXsdDest)) {
    await fetchVerifiedFile({
      url: W3C_XML_XSD_URL,
      dest: xmlXsdDest,
      expectedSha256: W3C_XML_XSD_SHA256,
      label: "W3C xml.xsd",
    });
  } else {
    log(`reusing existing xml.xsd at ${xmlXsdDest}`);
  }

  // 3. Patch the OOXML xsd imports to point at the local xml.xsd.
  patchXmlNsImports(vendorDir);

  const xsds = readdirSync(vendorDir).filter((f) => f.endsWith(".xsd"));
  log(`extracted ${xsds.length} xsd file(s) to ${vendorDir}`);
  if (!existsSync(sentinel)) {
    fail(`expected wml.xsd in vendor/ooxml-xsd/ after extract; got: ${xsds.join(", ")}`);
  }
  log(`✅ vendor/ooxml-xsd/ is ready.`);
  return 0;
}

const code = await main();
process.exit(code);
