"""
FastAPI application for pptx-diff-preview.

POST /api/compare  (multipart: old_file, new_file, enrich?) -> per-slide diff JSON
GET  /health

Images are returned inline as base64 data URIs. This is stateless and safe for
multi-replica deployments (Azure Container Apps) with no shared storage. For very
large decks, swap _data_uri for an upload to blob storage and return URLs instead
(see README).
"""

from __future__ import annotations

import base64
import shutil
import tempfile
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware

from .config import settings
from .enrich import summarize
from .slide_diff import RenderError, build_marked_decks, pptx_to_images

app = FastAPI(title="pptx-diff-preview", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],          # tighten for production
    allow_methods=["*"],
    allow_headers=["*"],
)

ZIP_MAGIC = b"PK\x03\x04"
MAX_BYTES = settings.max_upload_mb * 1024 * 1024


@app.get("/health")
def health():
    return {"status": "ok"}


def _validate(name: str, data: bytes):
    if not name.lower().endswith(".pptx"):
        raise HTTPException(400, detail=f"{name!r} is not a .pptx file")
    if len(data) == 0:
        raise HTTPException(400, detail=f"{name!r} is empty")
    if len(data) > MAX_BYTES:
        raise HTTPException(413, detail=f"{name!r} exceeds {settings.max_upload_mb} MB limit")
    if not data.startswith(ZIP_MAGIC):
        raise HTTPException(400, detail=f"{name!r} is not a valid PPTX (bad file signature)")


def _data_uri(png_path: Path | None) -> str | None:
    if png_path is None:
        return None
    return "data:image/png;base64," + base64.b64encode(png_path.read_bytes()).decode("ascii")


@app.post("/api/compare")
async def compare(
    old_file: UploadFile = File(...),
    new_file: UploadFile = File(...),
    enrich: bool = Form(False),
):
    old_bytes = await old_file.read()
    new_bytes = await new_file.read()
    _validate(old_file.filename or "old.pptx", old_bytes)
    _validate(new_file.filename or "new.pptx", new_bytes)

    tmp = Path(tempfile.mkdtemp(prefix="pptxdiff_"))
    warnings: list[str] = [
        "Comparison is text-only; image swaps, recolors, moved shapes, and chart "
        "visuals are not detected.",
        "Rendering uses LibreOffice; if a deck's fonts are not installed, text is "
        "substituted, which can shift where colored words land. Both decks are "
        "substituted the same way, so the old/new comparison stays internally "
        "consistent.",
    ]
    try:
        old_path = tmp / "old.pptx"
        new_path = tmp / "new.pptx"
        old_path.write_bytes(old_bytes)
        new_path.write_bytes(new_bytes)

        # 1) Text detection + in-place coloring (cannot fail on rendering).
        try:
            old_prs, new_prs, entries = build_marked_decks(
                old_path, new_path, threshold=settings.match_threshold
            )
        except Exception as e:
            raise HTTPException(422, detail=f"Could not read PPTX content: {e}")

        # 2) Render the marked decks (graceful: warn instead of crashing).
        old_imgs: list[Path] = []
        new_imgs: list[Path] = []
        work = tmp / "work"
        work.mkdir(parents=True, exist_ok=True)
        try:
            old_marked = work / "old_marked.pptx"
            new_marked = work / "new_marked.pptx"
            old_prs.save(str(old_marked))
            new_prs.save(str(new_marked))
            old_imgs = pptx_to_images(old_marked, work / "old",
                                      dpi=settings.render_dpi, timeout=settings.libreoffice_timeout)
            new_imgs = pptx_to_images(new_marked, work / "new",
                                      dpi=settings.render_dpi, timeout=settings.libreoffice_timeout)
        except RenderError as e:
            warnings.append(f"Slide rendering failed; previews are unavailable. Detail: {e}")

        # 3) Assemble response (summary always present; AI is advisory + optional).
        slides = []
        for e in entries:
            payload = e.public_dict()
            oi, ni = e.old_index0, e.new_index0
            payload["old_image"] = _data_uri(old_imgs[oi] if (oi is not None and oi < len(old_imgs)) else None)
            payload["new_image"] = _data_uri(new_imgs[ni] if (ni is not None and ni < len(new_imgs)) else None)

            if enrich:
                s = summarize(payload)
                payload["summary"] = s["summary"]
                payload["change_tags"] = s["tags"]
                payload["ai_summary"] = s["ai"]
            else:
                from .enrich import deterministic_summary
                payload["summary"] = deterministic_summary(payload)
                payload["change_tags"] = []
                payload["ai_summary"] = False

            slides.append(payload)

        counts = {
            "total_old": sum(1 for e in entries if e.old_slide is not None),
            "total_new": sum(1 for e in entries if e.new_slide is not None),
            "changed": sum(1 for e in entries if e.status == "changed"),
            "unchanged": sum(1 for e in entries if e.status == "unchanged"),
            "added": sum(1 for e in entries if e.status == "added"),
            "removed": sum(1 for e in entries if e.status == "removed"),
        }
        return {"summary": counts, "slides": slides, "warnings": warnings}
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
