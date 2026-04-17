"use client";

import { useState, useEffect } from "react";
import { useTheme } from "next-themes";
import { Monitor, Moon, Sun } from "lucide-react";
import { cn } from "../lib/cn";

interface ThemeToggleProps {
  className?: string;
  /** Compact mode shows only the active icon as a cycling button */
  compact?: boolean;
}

const options = [
  { value: "light" as const, icon: Sun, label: "Light" },
  { value: "system" as const, icon: Monitor, label: "System" },
  { value: "dark" as const, icon: Moon, label: "Dark" },
];

export function ThemeToggle({ className, compact }: ThemeToggleProps) {
  const { theme, setTheme } = useTheme();
  const [mounted, setMounted] = useState(false);

  useEffect(() => setMounted(true), []);

  if (!mounted) {
    return <div className={cn("h-7", compact ? "w-7" : "w-[104px]", className)} />;
  }

  if (compact) {
    const currentIdx = options.findIndex((o) => o.value === theme);
    const current = options[currentIdx >= 0 ? currentIdx : 1];
    const next = options[(options.indexOf(current) + 1) % options.length];
    const Icon = current.icon;
    return (
      <button
        onClick={() => setTheme(next.value)}
        className={cn(
          "flex h-7 w-7 items-center justify-center rounded-md text-secondary hover:bg-hover hover:text-foreground transition-colors duration-150",
          className
        )}
        title={`Theme: ${current.label}. Click for ${next.label}`}
      >
        <Icon size={14} />
      </button>
    );
  }

  return (
    <div className={cn("inline-flex items-center rounded-md bg-hover p-0.5 gap-0.5", className)}>
      {options.map(({ value, icon: Icon, label }) => (
        <button
          key={value}
          onClick={() => setTheme(value)}
          className={cn(
            "flex h-6 items-center justify-center rounded px-2 transition-colors duration-150",
            theme === value
              ? "bg-background text-foreground shadow-sm"
              : "text-secondary hover:text-foreground"
          )}
          title={label}
        >
          <Icon size={13} />
        </button>
      ))}
    </div>
  );
}

export type { ThemeToggleProps };
