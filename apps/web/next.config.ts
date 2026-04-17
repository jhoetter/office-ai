import type { NextConfig } from "next";
import path from "path";

const nextConfig: NextConfig = {
  transpilePackages: ["@officeai/ui", "@officeai/design-tokens"],
  devIndicators: false,
  outputFileTracingRoot: path.resolve(__dirname, "../.."),
};

export default nextConfig;
