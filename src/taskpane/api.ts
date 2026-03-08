/**
 * Backend API client. Base URL is set at build time via API_BASE_URL env var
 * (default: http://localhost:5000).
 */
declare const __API_BASE_URL__: string;
const API_BASE = typeof __API_BASE_URL__ !== "undefined" ? __API_BASE_URL__ : "http://localhost:5000";

function url(path: string): string {
  const base = API_BASE.replace(/\/$/, "");
  const p = path.startsWith("/") ? path : `/${path}`;
  return `${base}${p}`;
}

/** GET /health – validate backend availability. */
export async function getHealth(): Promise<{ ok: boolean; status: number; error?: string }> {
  try {
    const res = await fetch(url("/health"), { method: "GET" });
    return { ok: res.ok, status: res.status };
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return { ok: false, status: 0, error: message };
  }
}

/** GET /document – inspect backend source material (for debugging). */
export async function getDocument(): Promise<{ ok: boolean; status: number; body?: unknown; error?: string }> {
  try {
    const res = await fetch(url("/document"), { method: "GET" });
    if (!res.ok) return { ok: false, status: res.status };
    const contentType = res.headers.get("content-type") ?? "";
    let body: unknown;
    if (contentType.includes("application/json")) {
      body = await res.json().catch(() => ({}));
    } else {
      body = await res.text();
    }
    return { ok: true, status: res.status, body };
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return { ok: false, status: 0, error: message };
  }
}

export interface AnalyzeResponse {
  source_id: string;
  citation_text: string;
  confidence: number;
  url: string;
}

export interface AnalyzeRequest {
  text: string;
  document_id?: string;
  user_id?: string;
}

/** POST /analyze – analyze selected text and get citation. */
export async function analyze(
  request: AnalyzeRequest
): Promise<{ ok: boolean; status: number; data?: AnalyzeResponse; error?: string }> {
  try {
    const res = await fetch(url("/analyze"), {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        text: request.text,
        ...(request.document_id != null && { document_id: request.document_id }),
        ...(request.user_id != null && { user_id: request.user_id }),
      }),
    });
    const data = res.ok ? await res.json() : undefined;
    return { ok: res.ok, status: res.status, data };
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return { ok: false, status: 0, error: message };
  }
}
