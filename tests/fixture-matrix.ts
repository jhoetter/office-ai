import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

export type FixtureFormat = "docx" | "xlsx" | "pptx" | "pdf";
export type FixtureComplexity = "simple" | "complex";
export type FixtureOrigin = "synthetic" | "generated-real-shape" | "collected-real" | "public-domain";
export type FixtureExpectedBehavior =
  | "import"
  | "projection"
  | "noop-roundtrip"
  | "mutation-roundtrip"
  | "preserve-opaque"
  | "diagnose-unsupported"
  | "performance-smoke";

export interface FixtureMatrixEntry {
  readonly id: string;
  readonly format: FixtureFormat;
  readonly path: string;
  readonly origin: FixtureOrigin;
  readonly complexity: FixtureComplexity;
  readonly license: string;
  readonly source: string;
  readonly features: readonly string[];
  readonly expectedBehaviors: readonly FixtureExpectedBehavior[];
  readonly knownRisks: readonly string[];
}

export interface FixtureMatrix {
  readonly version: 1;
  readonly updated: string;
  readonly policy: Record<string, string>;
  readonly fixtures: readonly FixtureMatrixEntry[];
}

export interface FixtureMatrixFilter {
  readonly format?: FixtureFormat;
  readonly origin?: FixtureOrigin;
  readonly complexity?: FixtureComplexity;
  readonly id?: string;
  readonly feature?: string;
  readonly expectedBehavior?: FixtureExpectedBehavior;
}

export const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
export const MATRIX_PATH = resolve(REPO_ROOT, "fixtures/MATRIX.json");

let cached: FixtureMatrix | null = null;

export function loadFixtureMatrix(): FixtureMatrix {
  if (!cached) {
    cached = JSON.parse(readFileSync(MATRIX_PATH, "utf8")) as FixtureMatrix;
  }
  return cached;
}

export function matrixFixtures(filter: FixtureMatrixFilter = {}): FixtureMatrixEntry[] {
  return loadFixtureMatrix()
    .fixtures.filter((fixture) => {
      if (filter.id && fixture.id !== filter.id) return false;
      if (filter.format && fixture.format !== filter.format) return false;
      if (filter.origin && fixture.origin !== filter.origin) return false;
      if (filter.complexity && fixture.complexity !== filter.complexity) return false;
      if (filter.feature && !fixture.features.includes(filter.feature)) return false;
      if (filter.expectedBehavior && !fixture.expectedBehaviors.includes(filter.expectedBehavior))
        return false;
      return true;
    })
    .sort((a, b) => a.path.localeCompare(b.path));
}

export function requiredMatrixFixture(
  format: FixtureFormat,
  filter: Omit<FixtureMatrixFilter, "format"> = {}
): FixtureMatrixEntry {
  const [fixture] = matrixFixtures({ ...filter, format });
  if (!fixture) {
    throw new Error(`No fixture matrix entry matched ${format} ${JSON.stringify(filter)}`);
  }
  return fixture;
}

export function fixturePath(fixture: FixtureMatrixEntry): string {
  return resolve(REPO_ROOT, fixture.path);
}
