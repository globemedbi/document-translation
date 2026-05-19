from __future__ import annotations

"""
Solution — TEXT mode
====================
For PDFs with embedded, selectable text (e.g. Word exports, digital forms).

Pipeline:
  1. Extract every text span via PyMuPDF (exact position, font size, colour).
  2. Group spans into logical lines, deduplicate, batch-translate via Claude.
  3. For each page: redact all changed spans at once, then re-insert translations.

Much faster than VISION/CODING — Claude is only called for translation, not
for extraction or reconstruction.
"""

import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import NamedTuple

import fitz
from loguru import logger

from src.config import get_client, settings
from src.llm.schema.text_translation import TextTranslationResponse
from src.llm.structured import structured_call

# Minimum characters for a span to be worth translating (skip bullets, numbers, etc.)
_MIN_SPAN_LEN = 2
# How many strings to send per translation batch
_BATCH_SIZE = 40


class _Span(NamedTuple):
    page_idx: int
    bbox: tuple[float, float, float, float]   # x0, y0, x1, y1
    text: str
    font_size: float
    color: tuple[float, float, float]          # r, g, b  (0..1)


class TextLayerTranslator:
    """
    Translates a text-based PDF by extracting all text spans,
    translating them with Claude, then redacting originals and
    inserting translated text in-place.
    """

    def __init__(self) -> None:
        self.client = get_client()
        self.target_language = settings.target_language
        logger.info(
            f"TextLayerTranslator initialised "
            f"(provider={settings.llm_provider}, model={settings.model}, lang={self.target_language})"
        )

    def translate(self, pdf_path: Path, output_path: Path) -> Path:
        logger.info("=== TEXT LAYER TRANSLATION START ===")
        logger.info(f"  source : {pdf_path}")
        logger.info(f"  target : {self.target_language}")
        logger.info(f"  output : {output_path}")

        doc = fitz.open(str(pdf_path))
        spans = _extract_spans(doc)
        logger.info(f"  extracted {len(spans)} text spans from {doc.page_count} page(s)")

        # Translate all unique, non-trivial strings
        unique = _unique_texts(spans)
        logger.info(f"  translating {len(unique)} unique strings …")
        translations = self._batch_translate(unique)

        # Apply changes page by page
        for page_idx in range(doc.page_count):
            page_spans = [s for s in spans if s.page_idx == page_idx]
            if not page_spans:
                continue
            _apply_page(doc[page_idx], page_spans, translations)
            logger.info(f"  page {page_idx + 1}: overlay applied")

        output_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(output_path), garbage=4, deflate=True)
        doc.close()

        logger.info(f"=== TEXT LAYER DONE → {output_path} ===")
        return output_path

    # ── Translation ──────────────────────────────────────────────────────────

    def _batch_translate(self, texts: list[str]) -> dict[str, str]:
        """Return {original: translation} for every text in the list."""
        batches = [texts[i : i + _BATCH_SIZE] for i in range(0, len(texts), _BATCH_SIZE)]
        result: dict[str, str] = {}
        with ThreadPoolExecutor(max_workers=min(len(batches), 5)) as pool:
            futures = {pool.submit(self._call_translate, batch): idx for idx, batch in enumerate(batches)}
            for future in as_completed(futures):
                idx = futures[future]
                translated = future.result()
                result.update(translated)
                logger.debug(f"  batch {idx + 1}/{len(batches)}: {len(translated)} items translated")
        return result

    def _call_translate(self, texts: list[str]) -> dict[str, str]:
        """Send one batch to Claude, return {original: translated}."""
        numbered = "\n".join(f"{j+1}. {t}" for j, t in enumerate(texts))
        prompt = (
            f"Translate each numbered item to {self.target_language}.\n"
            "Rules:\n"
            "- Preserve formatting (ALL CAPS → ALL CAPS, Title Case → Title Case).\n"
            "- Keep proper nouns, dates, numbers, and codes unchanged.\n"
            "- Do not add explanations.\n\n"
            f"Texts:\n{numbered}\n\n"
            "Fill the 'translations' list with one entry per item: "
            "'original' = the exact original string, 'translated' = its translation."
        )
        try:
            result = structured_call(
                self.client,
                settings.model,
                messages=[{"role": "user", "content": prompt}],
                schema=TextTranslationResponse,
                max_tokens=4096,
                reasoning_effort="low",
            )
            return {item.original: item.translated for item in result.translations}
        except Exception as exc:
            logger.warning(f"Translation batch failed ({exc}); using originals")
            return {}


# ── Helpers ───────────────────────────────────────────────────────────────────

def _extract_spans(doc: fitz.Document) -> list[_Span]:
    """Extract all text spans with position and style info."""
    spans: list[_Span] = []
    for page_idx in range(doc.page_count):
        page = doc[page_idx]
        try:
            blocks = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]
        except Exception:
            continue
        for block in blocks:
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    text = span.get("text", "")
                    if len(text.strip()) < _MIN_SPAN_LEN:
                        continue
                    color_int = span.get("color", 0)
                    r = ((color_int >> 16) & 0xFF) / 255.0
                    g = ((color_int >> 8) & 0xFF) / 255.0
                    b = (color_int & 0xFF) / 255.0
                    spans.append(_Span(
                        page_idx=page_idx,
                        bbox=(
                            span["bbox"][0], span["bbox"][1],
                            span["bbox"][2], span["bbox"][3],
                        ),
                        text=text,
                        font_size=float(span.get("size", 10.0)),
                        color=(r, g, b),
                    ))
    return spans


# Spans larger than this are decorative/structural (section dividers, watermarks)
# and should not be translated or repositioned
_MAX_TRANSLATE_FONT_SIZE = 20.0


def _unique_texts(spans: list[_Span]) -> list[str]:
    """Return deduplicated list of non-trivial strings worth translating."""
    seen: set[str] = set()
    out: list[str] = []
    for s in spans:
        if s.font_size > _MAX_TRANSLATE_FONT_SIZE:
            continue
        t = s.text.strip()
        if t and t not in seen and re.search(r"[a-zA-Z؀-ۿ]", t):
            seen.add(t)
            out.append(t)
    return out


def _apply_page(
    page: fitz.Page,
    spans: list[_Span],
    translations: dict[str, str],
) -> None:
    """Redact original spans and insert translations for a single page."""
    to_replace: list[tuple[_Span, str]] = []

    for span in spans:
        if span.font_size > _MAX_TRANSLATE_FONT_SIZE:
            continue
        translated = translations.get(span.text.strip())
        if not translated or translated == span.text.strip():
            continue
        to_replace.append((span, translated))

    if not to_replace:
        return

    # Add all redaction annotations first
    for span, _ in to_replace:
        rect = fitz.Rect(*span.bbox)
        page.add_redact_annot(rect, fill=(1, 1, 1))

    # Apply all redactions in one pass (more efficient, avoids position drift)
    page.apply_redactions()

    # Re-insert all translations
    for span, translated in to_replace:
        rect = fitz.Rect(*span.bbox)
        box_h = rect.height
        # Keep font size, cap slightly below box height
        font_size = max(5.0, min(span.font_size, box_h * 0.90))
        # Baseline: bottom of box, adjusted for ascender
        baseline = fitz.Point(rect.x0, rect.y1 - box_h * 0.15)
        try:
            page.insert_text(
                baseline,
                translated,
                fontname="helv",
                fontsize=font_size,
                color=span.color,
            )
        except Exception as exc:
            logger.debug(f"  insert_text skipped ({exc}): {translated[:30]!r}")
