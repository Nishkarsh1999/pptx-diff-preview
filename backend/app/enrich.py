"""
Advisory summary layer.

HARD RULE: this never decides coloring. difflib is the sole authority for the
red/green spans (see slide_diff.py). This module only turns the already-detected,
structured change spans into a short human-readable sentence.

- If AI is disabled or no key is configured -> deterministic_summary (always works).
- If AI is enabled and configured -> ai_summary, which falls back to the
  deterministic summary on ANY error. AI failure never fails the request.
"""

from __future__ import annotations

import json

from .config import settings


def _fmt_change(c: dict) -> str:
    t = c.get("type")
    if t == "changed":
        return f'“{c["old"]}” → “{c["new"]}”'
    if t == "added":
        return f'added “{c["new"]}”'
    if t == "removed":
        return f'removed “{c["old"]}”'
    return ""


def deterministic_summary(entry: dict) -> str:
    """Build a summary purely from the mechanical diff. No model involved."""
    status = entry.get("status")
    if status == "added":
        return "This slide was added."
    if status == "removed":
        return "This slide was removed."

    changes = entry.get("changes", [])
    notes_changed = entry.get("has_notes_changes", False)

    if not changes:
        if notes_changed:
            return "Only the speaker notes changed."
        return "No text changes."

    shown = [s for s in (_fmt_change(c) for c in changes[:5]) if s]
    summary = "; ".join(shown)
    extra = len(changes) - len(shown)
    if extra > 0:
        summary += f"; and {extra} more change{'s' if extra != 1 else ''}"
    if notes_changed:
        summary += " (speaker notes also changed)"
    return summary


_SYSTEM = (
    "You are a precise slide comparison assistant. You receive deterministic, "
    "word-level text-diff data for one slide. Summarize ONLY the changes present "
    "in that data. Do not invent changes, do not claim a change the data does not "
    "show, and do not produce coordinates or bounding boxes. Return STRICT JSON, "
    "no prose, no code fences, shaped exactly as:\n"
    '{"summary": "<one concise sentence>", '
    '"changes": [{"type": "moved|reworded|substantive|cosmetic", "detail": "<short>"}]}'
)


def _ai_summary_anthropic(entry: dict) -> dict:
    from anthropic import Anthropic

    client = Anthropic(api_key=settings.anthropic_api_key)
    user = json.dumps({
        "status": entry.get("status"),
        "changes": entry.get("changes", []),
        "has_notes_changes": entry.get("has_notes_changes", False),
        "old_text": entry.get("old_text", "")[:4000],
        "new_text": entry.get("new_text", "")[:4000],
    }, ensure_ascii=False)
    msg = client.messages.create(
        model=settings.ai_model,
        max_tokens=400,
        system=_SYSTEM,
        messages=[{"role": "user", "content": user}],
    )
    raw = "".join(b.text for b in msg.content if getattr(b, "type", None) == "text")
    return _parse(raw)


def _ai_summary_openai(entry: dict) -> dict:
    from openai import OpenAI

    client = OpenAI(api_key=settings.openai_api_key)
    user = json.dumps({
        "status": entry.get("status"),
        "changes": entry.get("changes", []),
        "has_notes_changes": entry.get("has_notes_changes", False),
        "old_text": entry.get("old_text", "")[:4000],
        "new_text": entry.get("new_text", "")[:4000],
    }, ensure_ascii=False)
    resp = client.chat.completions.create(
        model=settings.ai_model,
        max_tokens=400,
        response_format={"type": "json_object"},
        messages=[{"role": "system", "content": _SYSTEM},
                  {"role": "user", "content": user}],
    )
    return _parse(resp.choices[0].message.content or "")


def _parse(raw: str) -> dict:
    cleaned = raw.strip().removeprefix("```json").removeprefix("```").removesuffix("```").strip()
    data = json.loads(cleaned)
    return {
        "summary": str(data.get("summary", "")).strip(),
        "changes": data.get("changes", []),
    }


def summarize(entry: dict) -> dict:
    """Return {"summary": str, "tags": list, "ai": bool}.
    Always succeeds; falls back to the deterministic summary on any AI error."""
    fallback = {"summary": deterministic_summary(entry), "tags": [], "ai": False}

    if not settings.ai_ready or not entry.get("has_changes"):
        return fallback

    try:
        if settings.ai_provider == "anthropic":
            out = _ai_summary_anthropic(entry)
        elif settings.ai_provider == "openai":
            out = _ai_summary_openai(entry)
        else:
            return fallback
        if not out.get("summary"):
            return fallback
        return {"summary": out["summary"], "tags": out.get("changes", []), "ai": True}
    except Exception:
        # Never let an AI failure break the job.
        return fallback
