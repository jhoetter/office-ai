"use client";

/**
 * Word-style "Restrict Editing" dialog.
 *
 * Mirrors the right-rail panel in Microsoft Word. Lets the user pick
 * an allowed-edits mode (`readOnly` / `comments` / `trackedChanges` /
 * `forms` or `none` to clear), enforce / un-enforce the protection,
 * optionally restrict formatting changes, and (optionally) seal it
 * with a password.
 *
 * Submission dispatches `docx:set-protection`. The password is hashed
 * client-side (SHA-512 with a random salt + 100 000 spins) so the
 * resulting `<w:documentProtection>` element matches what Word writes
 * when the user types a password into its native dialog. When the
 * user removes protection (`mode === "none"`) we dispatch with
 * `enabled: false` to drop the element entirely.
 */

import { useEffect, useRef, useState } from "react";
import type { ReactNode } from "react";
import { X } from "@officeai/ui/sonaloop-icons";
import { useFocusTrap } from "@officeai/ui";

export type ProtectionEdit = "none" | "readOnly" | "comments" | "trackedChanges" | "forms";

export interface ProtectDocumentSubmit {
  readonly enabled: boolean;
  readonly edit?: ProtectionEdit;
  readonly enforce?: boolean;
  readonly formatting?: boolean;
  readonly algorithmName?: string;
  readonly hashValue?: string;
  readonly saltValue?: string;
  readonly spinCount?: number;
}

export interface ProtectDocumentDialogProps {
  readonly open: boolean;
  readonly current: {
    readonly enabled: boolean;
    readonly edit?: ProtectionEdit;
    readonly enforce?: boolean;
    readonly formatting?: boolean;
  };
  readonly onClose: () => void;
  readonly onSubmit: (payload: ProtectDocumentSubmit) => void;
}

export function ProtectDocumentDialog(props: ProtectDocumentDialogProps): ReactNode {
  const { open, current, onClose, onSubmit } = props;
  const [edit, setEdit] = useState<ProtectionEdit>(current.edit ?? "readOnly");
  const [enforce, setEnforce] = useState<boolean>(current.enforce ?? true);
  const [formatting, setFormatting] = useState<boolean>(current.formatting ?? false);
  const [password, setPassword] = useState<string>("");
  const [working, setWorking] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const panelRef = useRef<HTMLDivElement>(null);
  useFocusTrap(panelRef, { enabled: open, onEscape: onClose });

  useEffect(() => {
    if (!open) return;
    setEdit(current.edit ?? "readOnly");
    setEnforce(current.enforce ?? true);
    setFormatting(current.formatting ?? false);
    setPassword("");
    setError(null);
  }, [open, current.edit, current.enforce, current.formatting]);

  if (!open) return null;

  const handleApply = async (): Promise<void> => {
    setError(null);
    if (edit === "none") {
      onSubmit({ enabled: false });
      onClose();
      return;
    }
    let hashFields: Pick<ProtectDocumentSubmit, "algorithmName" | "hashValue" | "saltValue" | "spinCount"> =
      {};
    if (password.length > 0) {
      try {
        setWorking(true);
        hashFields = await hashPassword(password);
      } catch (err) {
        setError(err instanceof Error ? err.message : String(err));
        setWorking(false);
        return;
      }
      setWorking(false);
    }
    onSubmit({
      enabled: true,
      edit,
      enforce,
      formatting,
      ...hashFields,
    });
    onClose();
  };

  const handleRemove = (): void => {
    onSubmit({ enabled: false });
    onClose();
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="docx-protect-title"
      className="fixed inset-0 z-[60] flex items-center justify-center bg-black/40 px-4 py-6"
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div
        ref={panelRef}
        tabIndex={-1}
        className="flex w-full max-w-md flex-col rounded-lg border border-divider bg-surface shadow-2xl outline-none"
        onClick={(e) => e.stopPropagation()}
      >
        <header className="flex items-center justify-between gap-3 border-b border-divider px-5 py-3">
          <h2 id="docx-protect-title" className="text-base font-semibold">
            Restrict editing
          </h2>
          <button
            type="button"
            aria-label="Close"
            onClick={onClose}
            className="rounded p-1 text-secondary transition-colors hover:bg-hover hover:text-default"
          >
            <X size={16} />
          </button>
        </header>

        <div className="flex flex-col gap-4 px-5 py-4 text-sm">
          <fieldset className="flex flex-col gap-1">
            <legend className="text-xs font-medium uppercase tracking-wide text-secondary">
              Editing restrictions
            </legend>
            <Radio
              label="No restrictions"
              checked={edit === "none"}
              onChange={() => setEdit("none")}
              testId="docx-protect-none"
            />
            <Radio
              label="Tracked changes"
              checked={edit === "trackedChanges"}
              onChange={() => setEdit("trackedChanges")}
              testId="docx-protect-tracked"
            />
            <Radio
              label="Comments"
              checked={edit === "comments"}
              onChange={() => setEdit("comments")}
              testId="docx-protect-comments"
            />
            <Radio
              label="Filling in forms"
              checked={edit === "forms"}
              onChange={() => setEdit("forms")}
              testId="docx-protect-forms"
            />
            <Radio
              label="No changes (read only)"
              checked={edit === "readOnly"}
              onChange={() => setEdit("readOnly")}
              testId="docx-protect-readonly"
            />
          </fieldset>

          {edit !== "none" ? (
            <>
              <fieldset className="flex flex-col gap-1.5">
                <legend className="text-xs font-medium uppercase tracking-wide text-secondary">
                  Options
                </legend>
                <Check
                  label="Enforce protection"
                  checked={enforce}
                  onChange={setEnforce}
                  testId="docx-protect-enforce"
                />
                <Check
                  label="Restrict formatting changes"
                  checked={formatting}
                  onChange={setFormatting}
                  testId="docx-protect-formatting"
                />
              </fieldset>

              <label className="flex flex-col gap-1">
                <span className="text-xs text-secondary">Password (optional)</span>
                <input
                  type="password"
                  value={password}
                  onChange={(e) => setPassword(e.currentTarget.value)}
                  data-testid="docx-protect-password"
                  className="rounded border border-divider bg-background px-2 py-1"
                />
              </label>
            </>
          ) : null}

          {error ? <p className="text-xs text-red-500">{error}</p> : null}
        </div>

        <footer className="flex items-center justify-between gap-2 border-t border-divider px-5 py-3">
          {current.enabled ? (
            <button
              type="button"
              onClick={handleRemove}
              data-testid="docx-protect-remove"
              className="rounded text-xs text-red-500 underline hover:text-red-600"
            >
              Stop protection
            </button>
          ) : (
            <span />
          )}
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={onClose}
              className="rounded border border-divider bg-background px-3 py-1.5 text-sm hover:bg-hover"
            >
              Cancel
            </button>
            <button
              type="button"
              onClick={() => void handleApply()}
              disabled={working}
              data-testid="docx-protect-apply"
              className="rounded bg-accent px-3 py-1.5 text-sm font-medium text-white hover:opacity-90 disabled:opacity-40"
            >
              {working ? "Working…" : "OK"}
            </button>
          </div>
        </footer>
      </div>
    </div>
  );
}

function Radio(props: { label: string; checked: boolean; onChange: () => void; testId: string }): ReactNode {
  const { label, checked, onChange, testId } = props;
  return (
    <label className="flex items-center gap-2">
      <input type="radio" checked={checked} onChange={onChange} data-testid={testId} />
      <span>{label}</span>
    </label>
  );
}

function Check(props: {
  label: string;
  checked: boolean;
  onChange: (v: boolean) => void;
  testId: string;
}): ReactNode {
  const { label, checked, onChange, testId } = props;
  return (
    <label className="flex items-center gap-2">
      <input
        type="checkbox"
        checked={checked}
        onChange={(e) => onChange(e.currentTarget.checked)}
        data-testid={testId}
      />
      <span>{label}</span>
    </label>
  );
}

/**
 * Hash a Word-style password using SHA-512 + 100 000 spin counts.
 * Mirrors the algorithm Word writes when sealing a document — using
 * the same spin count and salt size means a future "Unprotect"
 * dialog (if we ever build one) could verify the original password.
 *
 * Falls back to throwing when `crypto.subtle` is unavailable; the
 * editor surfaces the toast.
 */
async function hashPassword(password: string): Promise<{
  algorithmName: string;
  hashValue: string;
  saltValue: string;
  spinCount: number;
}> {
  const subtle = typeof crypto !== "undefined" && crypto.subtle ? crypto.subtle : null;
  if (!subtle) throw new Error("Web Crypto API unavailable in this browser.");
  const saltBytes = new Uint8Array(16);
  crypto.getRandomValues(saltBytes);
  const utf16le = utf16LE(password);
  let hash = await subtle.digest("SHA-512", toArrayBuffer(concat(saltBytes, utf16le)));
  for (let i = 0; i < 100_000; i++) {
    const iter = new Uint8Array(4);
    new DataView(iter.buffer).setUint32(0, i, true);
    hash = await subtle.digest("SHA-512", toArrayBuffer(concat(new Uint8Array(hash), iter)));
  }
  return {
    algorithmName: "SHA-512",
    hashValue: bytesToBase64(new Uint8Array(hash)),
    saltValue: bytesToBase64(saltBytes),
    spinCount: 100_000,
  };
}

function utf16LE(text: string): Uint8Array {
  const out = new Uint8Array(text.length * 2);
  const view = new DataView(out.buffer);
  for (let i = 0; i < text.length; i++) {
    view.setUint16(i * 2, text.charCodeAt(i), true);
  }
  return out;
}

function concat(a: Uint8Array, b: Uint8Array): Uint8Array {
  const out = new Uint8Array(a.length + b.length);
  out.set(a, 0);
  out.set(b, a.length);
  return out;
}

/**
 * Force a Uint8Array onto a fresh ArrayBuffer. `crypto.subtle.digest`
 * is typed as `BufferSource` whose `ArrayBufferView<ArrayBuffer>`
 * variant rejects the SharedArrayBuffer-compatible default Uint8Array
 * type — copying the bytes onto a plain ArrayBuffer keeps TypeScript
 * happy without any runtime cost worth measuring.
 */
function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  const out = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(out).set(bytes);
  return out;
}

function bytesToBase64(bytes: Uint8Array): string {
  let bin = "";
  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i] ?? 0);
  if (typeof btoa === "function") return btoa(bin);
  return Buffer.from(bin, "binary").toString("base64");
}
