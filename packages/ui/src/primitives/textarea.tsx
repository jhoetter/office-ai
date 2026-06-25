import { forwardRef, type TextareaHTMLAttributes } from "react";
import { cn } from "../lib/cn";

export interface TextareaProps extends TextareaHTMLAttributes<HTMLTextAreaElement> {
  label?: string;
  error?: string;
}

const Textarea = forwardRef<HTMLTextAreaElement, TextareaProps>(
  ({ className, label, error, id, ...props }, ref) => {
    return (
      <div className={cn("sl-field flex flex-col gap-1.5", error && "sl-field--invalid")}>
        {label && (
          <label htmlFor={id} className="sl-field__label text-sm font-medium text-foreground">
            {label}
          </label>
        )}
        <textarea
          id={id}
          className={cn(
            "sl-textarea flex min-h-[80px] w-full rounded-md border border-divider bg-surface px-3 py-2 text-sm text-foreground transition-colors duration-150 resize-none",
            "placeholder:text-tertiary",
            "hover:border-tertiary",
            "focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--accent)]/40 focus-visible:border-[var(--accent)]",
            "disabled:cursor-not-allowed disabled:opacity-50",
            error && "border-error focus-visible:ring-error/40",
            className
          )}
          ref={ref}
          {...props}
        />
        {error && <p className="sl-field__error text-xs text-error">{error}</p>}
      </div>
    );
  }
);
Textarea.displayName = "Textarea";

export { Textarea };
