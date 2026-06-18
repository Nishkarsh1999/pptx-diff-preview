import { useMemo, useState } from "react";
import { compare } from "./api";
import type { CompareResponse, SlideResult, SlideStatus } from "./types";

type Filter = "all" | "changed" | "unchanged" | "added" | "removed";
type Tab = "sidebyside" | "changes" | "notes";

const STATUS_LABEL: Record<SlideStatus, string> = {
  changed: "Changed",
  unchanged: "No change",
  added: "Added",
  removed: "Removed",
};

function mapping(s: SlideResult): string {
  const o = s.old_slide ?? "—";
  const n = s.new_slide ?? "—";
  return `${o} → ${n}`;
}

export default function SlideDiffViewer() {
  const [oldFile, setOldFile] = useState<File | null>(null);
  const [newFile, setNewFile] = useState<File | null>(null);
  const [enrich, setEnrich] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [result, setResult] = useState<CompareResponse | null>(null);

  const [selected, setSelected] = useState(0);
  const [filter, setFilter] = useState<Filter>("all");
  const [tab, setTab] = useState<Tab>("sidebyside");
  const [showRaw, setShowRaw] = useState(false);

  async function runCompare() {
    if (!oldFile || !newFile) return;
    setLoading(true);
    setError(null);
    setResult(null);
    try {
      const data = await compare(oldFile, newFile, enrich);
      setResult(data);
      const firstChanged = data.slides.findIndex((s) => s.status !== "unchanged");
      setSelected(firstChanged >= 0 ? firstChanged : 0);
      setFilter("all");
      setTab("sidebyside");
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setLoading(false);
    }
  }

  function downloadJson() {
    if (!result) return;
    // Strip the heavy base64 images from the downloadable JSON.
    const slim = {
      ...result,
      slides: result.slides.map(({ old_image, new_image, ...rest }) => rest),
    };
    const blob = new Blob([JSON.stringify(slim, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "diff-result.json";
    a.click();
    URL.revokeObjectURL(url);
  }

  const filtered = useMemo(() => {
    if (!result) return [];
    return result.slides
      .map((s, i) => ({ s, i }))
      .filter(({ s }) => filter === "all" || s.status === filter);
  }, [result, filter]);

  const current = result?.slides[selected];

  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">pptx-diff-preview</div>
        <div className="tagline">Text-only slide comparison · deletions in red · insertions in green</div>
      </header>

      <section className="uploader">
        <div className="file-inputs">
          <label className="file-input">
            <span className="file-label">Old deck (.pptx)</span>
            <input type="file" accept=".pptx" onChange={(e) => setOldFile(e.target.files?.[0] ?? null)} />
            <span className="file-name">{oldFile?.name ?? "No file selected"}</span>
          </label>
          <label className="file-input">
            <span className="file-label">New deck (.pptx)</span>
            <input type="file" accept=".pptx" onChange={(e) => setNewFile(e.target.files?.[0] ?? null)} />
            <span className="file-name">{newFile?.name ?? "No file selected"}</span>
          </label>
        </div>
        <div className="controls">
          <label className="enrich-toggle">
            <input type="checkbox" checked={enrich} onChange={(e) => setEnrich(e.target.checked)} />
            AI summaries
          </label>
          <button className="primary" onClick={runCompare} disabled={!oldFile || !newFile || loading}>
            {loading ? "Comparing…" : "Compare"}
          </button>
        </div>
      </section>

      {error && <div className="error">⚠ {error}</div>}

      {loading && <div className="placeholder">Rendering slides and computing the diff…</div>}

      {!loading && !result && !error && (
        <div className="placeholder">
          Upload an old and a new <code>.pptx</code>, then press Compare. Only text differences are
          detected; the slides are rendered so changed words can be colored in place.
        </div>
      )}

      {result && (
        <>
          <div className="summary-bar">
            <Stat label="Old slides" value={result.summary.total_old} />
            <Stat label="New slides" value={result.summary.total_new} />
            <Stat label="Changed" value={result.summary.changed} tone="changed" />
            <Stat label="No change" value={result.summary.unchanged} />
            <Stat label="Added" value={result.summary.added} tone="added" />
            <Stat label="Removed" value={result.summary.removed} tone="removed" />
            <button className="ghost" onClick={downloadJson}>Download JSON</button>
          </div>

          {result.warnings.length > 0 && (
            <details className="warnings">
              <summary>{result.warnings.length} note(s) about this comparison</summary>
              <ul>{result.warnings.map((w, i) => <li key={i}>{w}</li>)}</ul>
            </details>
          )}

          <div className="results">
            <aside className="sidebar">
              <div className="filters">
                {(["all", "changed", "unchanged", "added", "removed"] as Filter[]).map((f) => (
                  <button
                    key={f}
                    className={`filter ${filter === f ? "active" : ""}`}
                    onClick={() => setFilter(f)}
                  >
                    {f === "all" ? "All" : STATUS_LABEL[f as SlideStatus]}
                  </button>
                ))}
              </div>
              <ul className="slide-list">
                {filtered.map(({ s, i }) => (
                  <li
                    key={i}
                    className={`slide-item ${selected === i ? "selected" : ""}`}
                    onClick={() => { setSelected(i); setTab("sidebyside"); }}
                  >
                    <div className="slide-item-top">
                      <span className="slide-map">{mapping(s)}</span>
                      <span className={`badge ${s.status}`}>{STATUS_LABEL[s.status]}</span>
                    </div>
                    <div className="slide-item-summary">{s.summary}</div>
                    <div className="slide-item-meta">
                      <span>sim {(s.text_similarity * 100).toFixed(0)}%</span>
                      {s.possibly_reordered && <span className="reorder">reordered</span>}
                      {s.has_notes_changes && <span className="notes-flag">notes</span>}
                    </div>
                  </li>
                ))}
                {filtered.length === 0 && <li className="empty">No slides match this filter.</li>}
              </ul>
            </aside>

            <main className="detail">
              {current ? (
                <>
                  <div className="detail-head">
                    <h2>Slide {mapping(current)}</h2>
                    <span className={`badge ${current.status}`}>{STATUS_LABEL[current.status]}</span>
                    {current.ai_summary && <span className="ai-pill">AI summary</span>}
                  </div>
                  <p className="detail-summary">{current.summary}</p>

                  <div className="tabs">
                    <button className={tab === "sidebyside" ? "active" : ""} onClick={() => setTab("sidebyside")}>
                      Side by side
                    </button>
                    <button className={tab === "changes" ? "active" : ""} onClick={() => setTab("changes")}>
                      Text changes ({current.changes.length})
                    </button>
                    <button className={tab === "notes" ? "active" : ""} onClick={() => setTab("notes")}>
                      Notes {current.has_notes_changes ? "•" : ""}
                    </button>
                  </div>

                  {tab === "sidebyside" && (
                    <div className="sidebyside">
                      <figure>
                        <figcaption>Old (deletions in red)</figcaption>
                        {current.old_image
                          ? <img src={current.old_image} alt={`old slide ${current.old_slide}`} />
                          : <div className="no-image">No old slide</div>}
                      </figure>
                      <figure>
                        <figcaption>New (insertions in green)</figcaption>
                        {current.new_image
                          ? <img src={current.new_image} alt={`new slide ${current.new_slide}`} />
                          : <div className="no-image">No new slide</div>}
                      </figure>
                    </div>
                  )}

                  {tab === "changes" && (
                    <div className="changes">
                      {current.changes.length === 0 && <p className="muted">No text changes on this slide.</p>}
                      {current.changes.map((c, i) => (
                        <div key={i} className="change-row">
                          <span className={`chip ${c.type}`}>{c.type}</span>
                          {c.old && <span className="old-text">{c.old}</span>}
                          {c.old && c.new && <span className="arrow">→</span>}
                          {c.new && <span className="new-text">{c.new}</span>}
                        </div>
                      ))}
                      {current.change_tags.length > 0 && (
                        <div className="tags">
                          {current.change_tags.map((t, i) => (
                            <span key={i} className="tag" title={t.detail}>{t.type}</span>
                          ))}
                        </div>
                      )}
                      <button className="link" onClick={() => setShowRaw((v) => !v)}>
                        {showRaw ? "Hide" : "Show"} raw change data
                      </button>
                      {showRaw && <pre className="raw">{JSON.stringify(current.changes, null, 2)}</pre>}
                    </div>
                  )}

                  {tab === "notes" && (
                    <div className="notes">
                      {!current.has_notes_changes && <p className="muted">Speaker notes are unchanged.</p>}
                      <div className="notes-cols">
                        <div>
                          <h4>Old notes</h4>
                          <pre>{current.old_notes || "(none)"}</pre>
                        </div>
                        <div>
                          <h4>New notes</h4>
                          <pre>{current.new_notes || "(none)"}</pre>
                        </div>
                      </div>
                    </div>
                  )}
                </>
              ) : (
                <div className="placeholder">Select a slide.</div>
              )}
            </main>
          </div>
        </>
      )}
    </div>
  );
}

function Stat({ label, value, tone }: { label: string; value: number; tone?: string }) {
  return (
    <div className={`stat ${tone ?? ""}`}>
      <div className="stat-value">{value}</div>
      <div className="stat-label">{label}</div>
    </div>
  );
}
