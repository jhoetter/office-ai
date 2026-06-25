"use client";

import { ThemeProvider as NextThemesProvider } from "next-themes";
import { useEffect } from "react";
import { useTheme } from "next-themes";

function SonaloopThemeAttributeSync(): React.ReactNode {
  const { resolvedTheme } = useTheme();

  useEffect(() => {
    const root = document.documentElement;
    if (resolvedTheme === "dark" || resolvedTheme === "light") {
      root.dataset.theme = resolvedTheme;
      return;
    }
    delete root.dataset.theme;
  }, [resolvedTheme]);

  return null;
}

export function ThemeProvider({ children }: { children: React.ReactNode }) {
  return (
    <NextThemesProvider attribute="class" defaultTheme="system" enableSystem>
      <SonaloopThemeAttributeSync />
      {children}
    </NextThemesProvider>
  );
}
