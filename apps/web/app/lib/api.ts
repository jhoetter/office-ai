/**
 * Minimal API client for the officeAI backend.
 *
 * Reads the base URL from `NEXT_PUBLIC_API_BASE_URL` and falls back to
 * the local dev server. Kept intentionally tiny — replace with a generated
 * client (e.g. from openapi-typescript) when the surface grows.
 */

export const API_BASE_URL =
  process.env.NEXT_PUBLIC_API_BASE_URL ?? "http://localhost:8000";

export class ApiError extends Error {
  constructor(
    public status: number,
    message: string,
  ) {
    super(message);
    this.name = "ApiError";
  }
}

async function request<T>(path: string, init?: RequestInit): Promise<T> {
  const res = await fetch(`${API_BASE_URL}${path}`, {
    ...init,
    headers: {
      "Content-Type": "application/json",
      ...(init?.headers ?? {}),
    },
    cache: "no-store",
  });
  if (!res.ok) {
    const text = await res.text().catch(() => res.statusText);
    throw new ApiError(res.status, text || res.statusText);
  }
  if (res.status === 204) return undefined as T;
  return (await res.json()) as T;
}

export interface Document {
  id: string;
  title: string;
  content: string;
  updated_at: string;
}

export const api = {
  health: () => request<{ status: string; service: string }>("/health"),
  listDocuments: () => request<Document[]>("/api/v1/documents"),
  getDocument: (id: string) => request<Document>(`/api/v1/documents/${id}`),
  saveDocument: (input: { title: string; content: string }) =>
    request<Document>("/api/v1/documents", {
      method: "POST",
      body: JSON.stringify(input),
    }),
  updateDocument: (id: string, input: { title: string; content: string }) =>
    request<Document>(`/api/v1/documents/${id}`, {
      method: "PUT",
      body: JSON.stringify(input),
    }),
};
