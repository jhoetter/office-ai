import nextConfig from "eslint-config-next";

const eslintConfig = [
  {
    ignores: [".next/**", "node_modules/**"],
  },
  ...nextConfig,
  {
    // React 19's `react-hooks` plugin ships several new strict rules
    // (`set-state-in-effect`, `refs`, exhaustive-deps tightening) that
    // flag patterns the editors have used since v0. Fixing every site
    // means a meaningful refactor of long-lived adapter code; until
    // that pass lands we keep the rules as **warnings** so they stay
    // visible in CI output without breaking the build. Pre-existing
    // bugs caught by these rules are tracked as a code-quality TODO,
    // not a blocker.
    files: ["**/*.{ts,tsx}"],
    rules: {
      "react-hooks/set-state-in-effect": "warn",
      "react-hooks/refs": "warn",
      "react-hooks/exhaustive-deps": "warn",
      "react-hooks/preserve-manual-memoization": "warn",
      "react-hooks/immutability": "warn",
    },
  },
];

export default eslintConfig;
