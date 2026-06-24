/**
 * Replacement for `next/link` used when the editor bundle runs inside a
 * non-Next host (e.g. embedding host' Vite SPA). The real `next/link` enables
 * client-side navigation in Next's app router, which is irrelevant
 * outside Next — a plain `<a>` is the safe lowest-common-denominator
 * and keeps `EditorTopBar` rendering correctly without dragging the
 * Next runtime into the bundle.
 *
 * Aliased via the esbuild build (see `build.mjs`); the apps/web Next
 * build is unaffected because it never goes through this bundler.
 */
import * as React from "react";

type AnyLinkProps = React.AnchorHTMLAttributes<HTMLAnchorElement> & {
  href: string;
  prefetch?: boolean | null;
  replace?: boolean;
  scroll?: boolean;
  shallow?: boolean;
  passHref?: boolean;
  legacyBehavior?: boolean;
  locale?: string | false;
};

const Link = React.forwardRef<HTMLAnchorElement, AnyLinkProps>(function Link(
  {
    href,
    prefetch: _prefetch,
    replace: _replace,
    scroll: _scroll,
    shallow: _shallow,
    passHref: _passHref,
    legacyBehavior: _legacyBehavior,
    locale: _locale,
    children,
    ...rest
  },
  ref
) {
  return (
    <a ref={ref} href={typeof href === "string" ? href : "#"} {...rest}>
      {children}
    </a>
  );
});

export default Link;
