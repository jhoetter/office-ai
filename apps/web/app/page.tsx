import Link from "next/link";
import { ArrowRight, Sparkles } from "lucide-react";
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
        <span className="inline-flex items-center gap-1.5 rounded-full bg-[var(--ai-violet-light)] px-2.5 py-0.5 text-xs font-medium text-[var(--ai-violet)]">
          <Sparkles size={12} />
          AI-native office editors
        </span>
        <h1 className="text-4xl font-semibold tracking-tight text-foreground sm:text-5xl">
          Headless-first office editors with a real command bus.
        </h1>
        <p className="max-w-prose text-base text-secondary">
          Word-, Excel- and PowerPoint-compatible editors built around an OOXML-faithful core. Every change
          — human or AI — flows through the same typed command bus, so agents and humans collaborate without
          two parallel mutation paths. DOCX, XLSX and PPTX editors: all live.
        </p>
        <div className="flex flex-wrap items-center gap-2">
          <Link href="/editor">
            <Button variant="accent" size="lg">
              Open the DOCX editor
              <ArrowRight size={14} />
            </Button>
          </Link>
          <Link href="/xlsx-editor">
            <Button variant="secondary" size="lg">
              Open the XLSX editor
              <ArrowRight size={14} />
            </Button>
          </Link>
          <Link href="/pptx-editor">
            <Button variant="secondary" size="lg">
              Open the PPTX editor
              <ArrowRight size={14} />
            </Button>
          </Link>
          <Link href="https://github.com" target="_blank">
            <Button variant="ghost" size="lg">
              Read the specs
            </Button>
          </Link>
        </div>

        <ul className="mt-12 grid w-full max-w-3xl grid-cols-1 gap-4 text-sm text-secondary sm:grid-cols-3">
          <li className="rounded-lg border border-divider bg-surface p-4">
            <div className="text-foreground">OOXML round-trip</div>
            <p className="mt-1 text-xs">
              Untouched parts are preserved byte-for-byte; edited parts re-serialize cleanly.
            </p>
          </li>
          <li className="rounded-lg border border-divider bg-surface p-4">
            <div className="text-foreground">Command bus</div>
            <p className="mt-1 text-xs">
              Every mutation is a typed `Command`. Agents stage `pending` changes for review.
            </p>
          </li>
          <li className="rounded-lg border border-divider bg-surface p-4">
            <div className="text-foreground">CLI + headless API</div>
            <p className="mt-1 text-xs">
              <code className="font-mono">office-agent</code> wraps the same headless agent for server-side AI
              workflows.
            </p>
          </li>
        </ul>
      </section>
    </main>
  );
}
