# pptx-diff-preview

Compare two PowerPoint decks and see the **text differences** rendered directly
on the slide images: deletions in red on the old deck, insertions in green on the
new deck.

![side by side: old slide with FY25/8%/1200 in red, new slide with FY26/12%/1450 in green]

It is built for the common review question — *"what actually changed in this deck
between versions?"* — and answers it precisely at the word level, without you
having to eyeball two decks side by side.

---

## What it does

- Pairs slides between the two decks by content similarity (not by position), so
  inserting or reordering a slide does **not** make everything after it look
  changed.
- Detects text changes deterministically with a word-level diff, then recolors
  **only the changed words**, in place, preserving each run's own formatting
  (bold, size, color).
- Renders each marked-up deck to per-slide PNGs so you get a true visual preview.
- Lists structured before/after changes per slide and diffs speaker notes.
- Optionally adds a short AI-written summary per slide — **advisory only**; it
  never decides what is colored, and the app works fully with no API key.

## What it does not do

It is **text-only by design**. It does not detect image swaps, recolored shapes,
moved/resized objects, chart data changes, or animation changes. See
[Known limitations](#known-limitations).

---

## Architecture

```
┌──────────────┐   two .pptx    ┌───────────────────────── backend (FastAPI) ─────────────────────────┐
│  frontend    │  ───────────▶  │  1. flatten each slide → ordered word-token stream                   │
│  (React/Vite)│                │     (text frames + table cells + groups, in document order)          │
│              │                │  2. match slides   → Hungarian assignment over title/body similarity │
│  side-by-side│  ◀───────────  │  3. diff tokens    → difflib.SequenceMatcher  (THE source of truth)  │
│  preview +   │   JSON +       │  4. recolor runs   → split a run only at a change boundary, deep-copy │
│  change list │   base64 PNGs  │     rPr so formatting survives; red = deleted, green = inserted       │
└──────────────┘                │  5. render         → LibreOffice headless → PDF → pdftoppm → PNG      │
                                │  6. summarize      → deterministic; optional advisory LLM            │
                                └──────────────────────────────────────────────────────────────────────┘
```

**Detection is deterministic and text-only.** `difflib.SequenceMatcher` over the
word-token stream is the *only* authority for what gets colored. Rendering exists
purely to produce the preview image; it never influences detection. The LLM layer
(if enabled) only writes prose summaries of changes the diff already found.

### How the coloring stays correct

A slide is flattened into a single ordered stream of word tokens. Paragraph
boundaries are marked with a sentinel so a diff never silently merges the last
word of one paragraph with the first word of the next — that is what stops a
paragraph split/merge from being reported as a text change.

When a run contains a mix of changed and unchanged characters, it is split at the
boundary into multiple runs, each a deep copy of the original run's properties
(`rPr`), and only the changed segment's color is overridden. Unchanged runs are
never touched. This is why a bold word next to an edited word keeps its bold.

---

## Running with Docker (recommended)

```bash
cp .env.example .env        # optional: edit settings / add an API key
docker compose up --build
```

- Frontend: <http://localhost:8080>
- Backend API: <http://localhost:8000> (`/health`, `/api/compare`)

The frontend container serves the built SPA via nginx and proxies `/api` to the
backend container, so there is nothing else to configure.

---

## Running locally without Docker

The backend needs **LibreOffice** (`soffice` on `PATH`) and **poppler**
(`pdftoppm`). On Debian/Ubuntu (incl. WSL2):

```bash
sudo apt-get install -y libreoffice-impress poppler-utils fonts-liberation
```

### Backend (PowerShell)

```powershell
cd backend
python -m venv env
.\env\Scripts\Activate.ps1
pip install -r requirements.txt
uvicorn app.main:app --reload --port 8000
```

### Frontend (PowerShell, second terminal)

```powershell
cd frontend
npm install
npm run dev
```

Vite serves on <http://localhost:5173> and proxies `/api` to `http://localhost:8000`
during development (see `vite.config.ts`). To point the frontend at a different
backend, set `VITE_API_URL` before building, e.g. `VITE_API_URL=https://api.example.com`.

> Note: LibreOffice runs inside WSL2; `soffice` must be on the PATH of the shell
> that launches uvicorn. Running the backend in Docker avoids this entirely.

---

## API

### `POST /api/compare`

Multipart form:

| field      | type            | notes                                  |
|------------|-----------------|----------------------------------------|
| `old_file` | file (`.pptx`)  | required                               |
| `new_file` | file (`.pptx`)  | required                               |
| `enrich`   | bool            | optional; request advisory AI summaries |

Response (abridged):

```jsonc
{
  "summary": { "total_old": 2, "total_new": 3, "changed": 2, "added": 1, "removed": 0, "unchanged": 0 },
  "slides": [
    {
      "old_slide": 1, "new_slide": 1,
      "status": "changed",
      "possibly_reordered": false,
      "has_changes": true,
      "has_notes_changes": false,
      "text_similarity": 0.941,
      "changes": [
        { "type": "changed", "old": "FY25", "new": "FY26" },
        { "type": "changed", "old": "8%",   "new": "12%" }
      ],
      "summary": "“FY25” → “FY26”; “8%” → “12%”; “1200” → “1450”",
      "ai_summary": false,
      "old_image": "data:image/png;base64,…",   // old slide, deletions in red
      "new_image": "data:image/png;base64,…"    // new slide, insertions in green
    }
  ],
  "warnings": ["Comparison is text-only; …"]
}
```

Images are returned inline as base64 data URIs. This is stateless and safe for
multi-replica deployments with no shared storage. For very large decks this makes
the JSON heavy; for that case, upload the PNGs to blob storage in
`app/main.py` and return URLs instead of data URIs.

Errors: `400` invalid/empty/non-pptx upload, `413` over the size limit, `422`
unreadable PPTX content. A rendering failure does **not** fail the request — the
text diff is still returned, with a message in `warnings` and `*_image` set to
`null`.

---

## Configuration

All optional; sensible defaults shown.

| Variable                     | Default            | Meaning                                                        |
|------------------------------|--------------------|----------------------------------------------------------------|
| `MAX_UPLOAD_MB`              | `100`              | Per-file upload limit.                                         |
| `LIBREOFFICE_TIMEOUT_SECONDS`| `120`              | Per-render timeout for LibreOffice and pdftoppm.              |
| `RENDER_DPI`                 | `150`              | Preview image resolution.                                      |
| `MATCH_THRESHOLD`            | `0.55`             | Min similarity (0–1) to pair two slides; below ⇒ added/removed.|
| `ENABLE_AI_SUMMARY`          | `false`            | Turn on advisory LLM summaries.                                |
| `AI_PROVIDER`                | `anthropic`        | `anthropic` or `openai`.                                       |
| `AI_MODEL`                   | `claude-sonnet-4-6`| Model id for the chosen provider.                              |
| `ANTHROPIC_API_KEY`          | —                  | Required only if provider is anthropic and AI is on.           |
| `OPENAI_API_KEY`             | —                  | Required only if provider is openai and AI is on.              |

---

## Tests

```bash
cd backend
pip install pytest
python -m pytest
```

The suite builds real `.pptx` objects in memory and inspects the colored runs
directly, so it validates the diff/match/recolor logic **without** needing
LibreOffice. It covers word edits, paragraph split/merge (no false change),
promotion of text into a table, table cell additions, within-word changes
(`FY25`→`FY26`), formatting preservation on a changed paragraph, whitespace-only
suppression, the slide matcher (including the mid-deck-insert no-cascade case),
and added/removed classification. The single render test is skipped automatically
when `soffice` is absent.

---

## Known limitations

- **Text only.** Image swaps, recolored/moved/resized shapes, chart visuals, and
  animations are not compared. Only text in text frames, table cells, and grouped
  shapes is diffed.
- **Fonts and rendering.** Previews are produced by LibreOffice. If a deck uses a
  font that is not installed, LibreOffice substitutes one, which can shift where
  colored words sit. Both decks are substituted the same way, so the old/new
  comparison stays internally consistent, but the preview will not be pixel-identical
  to PowerPoint. Add fonts to the backend image to improve fidelity.
- **Text inside images** cannot be recolored — it is not text in the file.
- **Field runs** (slide numbers, dates via `a:fld`) are read for diffing but not
  recolored, since their cached text is fragile to rewrite.
- **Merged table cells** are de-duplicated and read once; complex nested tables
  are flattened in document order.
- **Heavily mixed-format runs**: a changed paragraph keeps each run's formatting,
  but if a *single* run mixes properties via run-level XML tricks the split is at
  the run granularity python-pptx exposes.
- **Hidden slides.** Image mapping assumes LibreOffice emits one PNG per
  python-pptx slide, in order. If a deck contains hidden slides and LibreOffice's
  export enumeration differs from python-pptx's, the image-to-slide mapping can
  drift. Verify against a deck with a hidden slide before relying on previews in
  production.
- **AI summaries are advisory.** They never affect detection or coloring and fall
  back to a deterministic summary on any error or when disabled.

## Project layout

```
pptx-diff-preview/
├─ backend/
│  ├─ app/
│  │  ├─ slide_diff.py   # flatten · diff · recolor-in-place · Hungarian match · render
│  │  ├─ enrich.py       # deterministic + optional advisory AI summary
│  │  ├─ config.py       # env-driven settings
│  │  └─ main.py         # FastAPI: /health, /api/compare
│  ├─ tests/test_slide_diff.py
│  ├─ requirements.txt
│  └─ Dockerfile
├─ frontend/
│  ├─ src/               # React + TypeScript viewer
│  ├─ package.json
│  ├─ vite.config.ts
│  ├─ Dockerfile
│  └─ nginx.conf
├─ docker-compose.yml
└─ .env.example
```
