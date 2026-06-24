import type { NextConfig } from "next";
import path from "path";

const nextConfig: NextConfig = {
  transpilePackages: ["@officeai/agent", "@officeai/ui", "@officeai/design-tokens"],
  devIndicators: false,
  outputFileTracingRoot: path.resolve(__dirname, "../.."),
  // ESLint runs as a dedicated step in `make verify` (`make lint-web`) so a
  // lint regression surfaces in seconds. Re-running it inside `next build`
  // would only delay the same failure to the slow build stage AND cascade
  // into every CI job that calls `pnpm build` directly (libreoffice
  // roundtrip, perf, OOXML XSD, web e2e). The lint gate stays authoritative
  // at the verify stage; build is for build issues only.
  eslint: { ignoreDuringBuilds: true },
};

export default nextConfig;
