import Link from "next/link";
import { Button, ThemeToggle } from "@officeai/ui";

export default function HomePage() {
  return (
    <main className="mx-auto flex min-h-full max-w-content flex-col px-6 py-16">
      <header className="flex items-center justify-between">
        <div className="flex items-center gap-2">
          <div className="h-6 w-6 rounded-md bg-[var(--office-blue)]" aria-hidden />
          <span className="font-semibold tracking-tight">officeAI</span>
        </div>
        <ThemeToggle />
      </header>

      <section className="mt-24 flex flex-1 flex-col items-start gap-6">
        <span className="rounded-full bg-[var(--accent-light)] px-2.5 py-0.5 text-xs font-medium text-[var(--accent)]">
          scaffold
        </span>
        <h1 className="text-4xl font-semibold tracking-tight text-foreground sm:text-5xl">
          A clean starting point.
        </h1>
        <p className="max-w-prose text-base text-secondary">
          officeAI is a minimal product scaffold sharing the design language and
          tech stack of our other products. The hello-world below is a tiny
          in-memory text editor that exercises the design system, theme toggle,
          and the FastAPI backend.
        </p>
        <div className="flex items-center gap-2">
          <Link href="/editor">
            <Button variant="accent" size="lg">
              Open the editor
            </Button>
          </Link>
          <Link href="https://localhost:8000/docs" target="_blank">
            <Button variant="ghost" size="lg">
              API docs
            </Button>
          </Link>
        </div>
      </section>
    </main>
  );
}
