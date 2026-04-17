"use client";

import Link from "next/link";
import { useEffect, useState } from "react";
import { ArrowLeft, Check, Loader2 } from "lucide-react";
import {
  Button,
  Card,
  CardContent,
  CardHeader,
  CardTitle,
  Input,
  Textarea,
  ThemeToggle,
  cn,
} from "@officeai/ui";
import { api, ApiError, type Document } from "@/lib/api";

type SaveState = "idle" | "saving" | "saved" | "error";

export default function EditorPage() {
  const [docId, setDocId] = useState<string | null>(null);
  const [title, setTitle] = useState("Untitled");
  const [content, setContent] = useState("");
  const [saveState, setSaveState] = useState<SaveState>("idle");
  const [error, setError] = useState<string | null>(null);
  const [loaded, setLoaded] = useState(false);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const docs = await api.listDocuments();
        if (cancelled) return;
        const latest = docs[0];
        if (latest) {
          setDocId(latest.id);
          setTitle(latest.title);
          setContent(latest.content);
        }
      } catch (e) {
        if (e instanceof ApiError) setError(`Backend unreachable (${e.status}).`);
        else setError("Backend unreachable.");
      } finally {
        if (!cancelled) setLoaded(true);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  async function handleSave() {
    setSaveState("saving");
    setError(null);
    try {
      let saved: Document;
      if (docId) {
        saved = await api.updateDocument(docId, { title, content });
      } else {
        saved = await api.saveDocument({ title, content });
        setDocId(saved.id);
      }
      setSaveState("saved");
      setTimeout(() => setSaveState((s) => (s === "saved" ? "idle" : s)), 1500);
    } catch (e) {
      setSaveState("error");
      setError(e instanceof Error ? e.message : "Save failed");
    }
  }

  return (
    <main className="mx-auto flex min-h-full max-w-prose flex-col px-6 py-10">
      <header className="flex items-center justify-between">
        <Link
          href="/"
          className="inline-flex items-center gap-1.5 text-sm text-secondary transition-colors hover:text-foreground"
        >
          <ArrowLeft size={14} />
          Back
        </Link>
        <ThemeToggle />
      </header>

      <Card className="mt-8 animate-fade-in">
        <CardHeader>
          <CardTitle>Editor</CardTitle>
          <p className="text-sm text-secondary">
            A minimal text editor wired to the FastAPI backend. Documents are
            stored in-memory for now — the goal is to verify the full stack.
          </p>
        </CardHeader>
        <CardContent className="flex flex-col gap-4">
          <Input
            id="doc-title"
            label="Title"
            value={title}
            onChange={(e) => setTitle(e.target.value)}
            placeholder="Give it a title"
            disabled={!loaded}
          />
          <Textarea
            id="doc-content"
            label="Content"
            value={content}
            onChange={(e) => setContent(e.target.value)}
            placeholder="Start typing…"
            rows={14}
            className="min-h-[280px] font-mono"
            disabled={!loaded}
          />
          {error && (
            <p className="rounded-md border border-[var(--error-border)] bg-[var(--error-bg)] px-3 py-2 text-xs text-[var(--error-text)]">
              {error}
            </p>
          )}
          <div className="flex items-center justify-between">
            <span className="text-xs text-secondary">
              {docId ? `id: ${docId.slice(0, 8)}…` : "unsaved"}
            </span>
            <Button
              variant="accent"
              size="md"
              onClick={handleSave}
              disabled={!loaded || saveState === "saving"}
              className={cn(saveState === "saved" && "bg-[var(--success)]")}
            >
              {saveState === "saving" ? (
                <>
                  <Loader2 size={14} className="animate-spin" />
                  Saving…
                </>
              ) : saveState === "saved" ? (
                <>
                  <Check size={14} />
                  Saved
                </>
              ) : (
                "Save"
              )}
            </Button>
          </div>
        </CardContent>
      </Card>
    </main>
  );
}
