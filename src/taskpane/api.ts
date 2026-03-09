/**
 * Backend API client. Base URL is set at build time via API_BASE_URL env var
 * (default: Cloud Run citation service).
 */
const API_BASE = "/api";


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

const DEFAULT_DOCUMENT_ID = "trustops-handbook-v1";
const DEFAULT_USER_ID = "candidate_1";

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
        document_id: request.document_id ?? DEFAULT_DOCUMENT_ID,
        user_id: request.user_id ?? DEFAULT_USER_ID,
      }),
    });
    let data: AnalyzeResponse | undefined;
    let errorMessage: string | undefined;
    const contentType = res.headers.get("content-type") ?? "";
    if (contentType.includes("application/json")) {
      const body = await res.json().catch(() => ({}));
      if (res.ok) {
        data = body;
      } else if (typeof body?.error === "string") {
        errorMessage = body.error;
      }
    }
    if (!res.ok && !errorMessage) {
      errorMessage = `HTTP ${res.status}`;
    }
    return {
      ok: res.ok,
      status: res.status,
      data,
      error: errorMessage,
    };
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return { ok: false, status: 0, error: message };
  }
}
