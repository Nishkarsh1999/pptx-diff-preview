import type { CompareResponse } from "./types";

const BASE = import.meta.env.VITE_API_URL ?? "";

export async function compare(
  oldFile: File,
  newFile: File,
  enrich: boolean
): Promise<CompareResponse> {
  const fd = new FormData();
  fd.append("old_file", oldFile);
  fd.append("new_file", newFile);
  fd.append("enrich", String(enrich));

  const res = await fetch(`${BASE}/api/compare`, { method: "POST", body: fd });
  if (!res.ok) {
    let detail = `Request failed (${res.status})`;
    try {
      const j = await res.json();
      if (j?.detail) detail = typeof j.detail === "string" ? j.detail : JSON.stringify(j.detail);
    } catch {
      /* ignore */
    }
    throw new Error(detail);
  }
  return res.json();
}
