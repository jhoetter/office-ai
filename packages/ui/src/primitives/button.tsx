import { cva, type VariantProps } from "class-variance-authority";
import { forwardRef, type ButtonHTMLAttributes } from "react";
import { cn } from "../lib/cn";

const buttonVariants = cva(
  "inline-flex items-center justify-center gap-2 whitespace-nowrap font-medium transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--accent)]/40 focus-visible:ring-offset-1 disabled:pointer-events-none disabled:opacity-50",
  {
    variants: {
      variant: {
        primary: "bg-foreground text-background hover:opacity-85 active:opacity-100",
        secondary: "bg-hover text-foreground hover:bg-divider active:bg-hover",
        ghost: "text-foreground hover:bg-hover active:bg-divider",
        destructive: "bg-error/10 text-error hover:bg-error/20 active:bg-error/10",
        link: "text-foreground underline-offset-4 hover:underline",
        accent: "bg-[var(--accent)] text-white hover:opacity-85 active:opacity-100",
      },
      size: {
        sm: "h-7 rounded-md px-2.5 text-xs",
        md: "h-8 rounded-md px-3 text-sm",
        lg: "h-10 rounded-md px-5 text-sm",
        icon: "h-8 w-8 rounded-md",
      },
    },
    defaultVariants: {
      variant: "primary",
      size: "md",
    },
  },
);

export interface ButtonProps
  extends ButtonHTMLAttributes<HTMLButtonElement>,
    VariantProps<typeof buttonVariants> {}

const Button = forwardRef<HTMLButtonElement, ButtonProps>(
  ({ className, variant, size, ...props }, ref) => {
    return (
      <button
        className={cn(buttonVariants({ variant, size, className }))}
        ref={ref}
        {...props}
      />
    );
  },
);
Button.displayName = "Button";

export { Button, buttonVariants };
