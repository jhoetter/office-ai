"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Loader2 } from "lucide-react";
import { cn } from "@officeai/ui";
import { PptxAgent } from "@officeai/pptx/agent";
import { SlideCanvas, SlidesSidebar } from "@officeai/pptx/renderer/react";
import type { Mutation } from "@officeai/core";
import type { TextShape } from "@officeai/pptx";
import { buildSamplePptx } from "@/lib/sample-pptx";
import { dispatchToLlmPptx } from "@/lib/llm-client-pptx";
import { PptxToolbar } from "./PptxToolbar";
import { PptxAgentPanel, type PptxAgentDispatch } from "./PptxAgentPanel";

interface ToastMessage {
  id: number;
  kind: "info" | "warn" | "error";
  text: string;
}

export function PptxEditor(): React.ReactNode {
  const [agent, setAgent] = useState<PptxAgent | null>(null);
  const agentRef = useRef<PptxAgent | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const [ready, setReady] = useState(false);
  const [docName, setDocName] = useState("welcome.pptx");
  const [activeIndex, setActiveIndex] = useState(0);
  const [pending, setPending] = useState<Mutation[]>([]);
  const [tick, setTick] = useState(0);
  const [toasts, setToasts] = useState<ToastMessage[]>([]);
  const toastId = useRef(0);

  const pushToast = useCallback((kind: ToastMessage["kind"], text: string) => {
    const id = ++toastId.current;
    setToasts((prev) => [...prev, { id, kind, text }]);
    setTimeout(() => setToasts((prev) => prev.filter((t) => t.id !== id)), 4000);
  }, []);

  const mountAgent = useCallback(async (buf: ArrayBuffer) => {
    const next = await PptxAgent.fromBuffer(buf);
    agentRef.current = next;
    setAgent(next);
    setActiveIndex(0);
    setPending([...next.getPendingMutations()]);
    setReady(true);
    setTick((t) => t + 1);
    next.subscribe(() => {
      setPending([...next.getPendingMutations()]);
      setTick((t) => t + 1);
    });
  }, []);

  useEffect(() => {
    let cancelled = false;
    void (async () => {
      try {
        const buf = await buildSamplePptx();
        if (!cancelled) await mountAgent(buf);
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [mountAgent, pushToast]);

  const onError = useCallback(
    (err: unknown) => pushToast("error", err instanceof Error ? err.message : String(err)),
    [pushToast]
  );

  const handleFile = useCallback(
    async (file: File) => {
      const buf = await file.arrayBuffer();
      setDocName(file.name);
      try {
        await mountAgent(buf);
        pushToast("info", `Opened ${file.name}`);
      } catch (err) {
        onError(err);
      }
    },
    [mountAgent, onError, pushToast]
  );

  const handleExport = useCallback(async () => {
    const a = agentRef.current;
    if (!a) return;
    try {
      const buf = await a.exportFile();
      const blob = new Blob([buf], {
        type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = docName;
      link.click();
      URL.revokeObjectURL(url);
      pushToast("info", `Exported ${docName}`);
    } catch (err) {
      onError(err);
    }
  }, [docName, onError, pushToast]);

  const snap = agent?.getSnapshot() ?? null;
  void tick;
  const slides = snap?.root.slides ?? [];
  const slideSize = snap?.root.slideSize ?? { cxEmu: 9144000, cyEmu: 6858000 };

  const addSlide = useCallback(async () => {
    const a = agentRef.current;
    if (!a) return;
    try {
      await a.applyCommand({
        type: "pptx:add-slide",
        payload: {},
        source: "human",
      });
      setActiveIndex(a.getSnapshot().root.slides.length - 1);
    } catch (err) {
      onError(err);
    }
  }, [onError]);

  const deleteSlide = useCallback(async () => {
    const a = agentRef.current;
    if (!a) return;
    if (a.getSnapshot().root.slides.length <= 1) {
      pushToast("warn", "Cannot delete the last slide.");
      return;
    }
    try {
      await a.applyCommand({
        type: "pptx:delete-slide",
        payload: { slideIndex: activeIndex },
        source: "human",
      });
      setActiveIndex((i) => Math.max(0, i - 1));
    } catch (err) {
      onError(err);
    }
  }, [activeIndex, onError, pushToast]);

  const duplicateSlide = useCallback(async () => {
    const a = agentRef.current;
    if (!a) return;
    try {
      await a.applyCommand({
        type: "pptx:duplicate-slide",
        payload: { slideIndex: activeIndex },
        source: "human",
      });
      setActiveIndex(activeIndex + 1);
    } catch (err) {
      onError(err);
    }
  }, [activeIndex, onError]);

  const addTextBox = useCallback(async () => {
    const a = agentRef.current;
    if (!a) return;
    try {
      await a.applyCommand({
        type: "pptx:add-text-box",
        payload: {
          slideIndex: activeIndex,
          text: "New text box",
          x: 1000000,
          y: 1000000,
          width: 4000000,
          height: 800000,
        },
        source: "human",
      });
    } catch (err) {
      onError(err);
    }
  }, [activeIndex, onError]);

  const toggleMark = useCallback(
    async (mark: "bold" | "italic" | "underline") => {
      const a = agentRef.current;
      if (!a) return;
      const slide = a.getSnapshot().root.slides[activeIndex];
      if (!slide) return;
      const ts = slide.shapes.find((s): s is TextShape => s.kind === "text");
      if (!ts || ts.txBody.paragraphs.length === 0) {
        pushToast("info", "No text shape on this slide.");
        return;
      }
      const p = ts.txBody.paragraphs[0];
      const flatLen = p.runs.reduce((acc, r) => acc + (r.isLineBreak ? 0 : r.text.length), 0);
      try {
        await a.applyCommand({
          type: "pptx:format-text",
          payload: {
            slideIndex: activeIndex,
            shapeId: ts.id,
            range: { paragraph: 0, start: 0, end: flatLen },
            format: { [mark]: true },
          },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError, pushToast]
  );

  // Routes prompts through the shared `/api/llm` bridge with `format: "pptx"`.
  // When no `OPENAI_API_KEY` is configured server-side the helper transparently
  // falls back to the in-process intent parser so the editor stays usable
  // without any env vars.
  const promptDispatch: PptxAgentDispatch = useCallback(
    async (text: string) => {
      const a = agentRef.current;
      if (!a) return;
      const result = await dispatchToLlmPptx(text, a, activeIndex);
      if (result.note) pushToast("warn", result.note);
      if (result.commands.length === 0) {
        pushToast("info", result.rationale || "No commands proposed.");
        return;
      }
      await a.applyCommands(result.commands);
      pushToast("info", `Applied ${result.commands.length} command(s).`);
    },
    [activeIndex, pushToast]
  );

  // Empty media URL map for now (no embedded images in the sample).
  const mediaUrls = useMemo(() => new Map<string, string>(), []);

  return (
    <div className="pptx-editor flex h-full min-h-0 flex-col gap-3">
      <PptxToolbar
        disabled={!ready}
        slideCount={slides.length}
        onOpenFile={() => fileInputRef.current?.click()}
        onExport={() => void handleExport()}
        onAddSlide={() => void addSlide()}
        onDeleteSlide={() => void deleteSlide()}
        onDuplicateSlide={() => void duplicateSlide()}
        onAddTextBox={() => void addTextBox()}
        onToggleBold={() => void toggleMark("bold")}
        onToggleItalic={() => void toggleMark("italic")}
        onToggleUnderline={() => void toggleMark("underline")}
      />
      <input
        ref={fileInputRef}
        type="file"
        accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation"
        className="hidden"
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) void handleFile(f);
          e.target.value = "";
        }}
      />
      <div className="flex min-h-0 flex-1 gap-3">
        <aside
          data-testid="pptx-sidebar"
          className="hidden w-[200px] shrink-0 overflow-y-auto rounded-md border border-divider bg-surface md:block"
        >
          {snap ? (
            <SlidesSidebar
              slides={slides}
              slideSize={slideSize}
              mediaUrls={mediaUrls}
              activeIndex={activeIndex}
              onSelect={setActiveIndex}
              thumbnailWidth={170}
            />
          ) : null}
        </aside>
        <section className="relative flex min-h-0 flex-1 items-center justify-center rounded-md border border-divider bg-background p-4">
          {agent ? (
            <div className="w-full max-w-[1100px]">
              <SlideCanvas
                agent={agent}
                slideIndex={activeIndex}
                mediaUrls={mediaUrls}
                onError={onError}
              />
            </div>
          ) : null}
          {!ready ? (
            <div className="absolute inset-0 flex items-center justify-center text-sm text-secondary">
              <Loader2 className="mr-2 animate-spin" size={14} />
              Loading…
            </div>
          ) : null}
        </section>
        <aside
          data-testid="pptx-agent-panel"
          className="hidden w-[280px] shrink-0 flex-col gap-4 border-l border-divider pl-3 lg:flex"
        >
          <PptxAgentPanel
            ready={ready}
            pending={pending}
            onApprove={(id) => agentRef.current?.approveMutation(id)}
            onReject={(id) => agentRef.current?.rejectMutation(id)}
            onApproveAll={() => pending.forEach((m) => agentRef.current?.approveMutation(m.id))}
            onRejectAll={() => pending.forEach((m) => agentRef.current?.rejectMutation(m.id))}
            dispatch={promptDispatch}
          />
        </aside>
      </div>
      <div className="pointer-events-none fixed bottom-6 left-1/2 z-50 flex -translate-x-1/2 flex-col items-center gap-2">
        {toasts.map((t) => (
          <div
            key={t.id}
            role="status"
            className={cn(
              "pointer-events-auto rounded-md border px-3 py-1.5 text-xs shadow-sm",
              t.kind === "info" && "border-divider bg-surface text-foreground",
              t.kind === "warn" && "border-[var(--warning)] bg-[var(--warning)]/10 text-[var(--warning)]",
              t.kind === "error" &&
                "border-[var(--error-border)] bg-[var(--error-bg)] text-[var(--error-text)]"
            )}
          >
            {t.text}
          </div>
        ))}
      </div>
    </div>
  );
}
