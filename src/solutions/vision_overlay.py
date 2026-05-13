from __future__ import annotations

"""
Solution 1 — VISION mode (v3)
==============================
Uses PyMuPDF's exact text extraction to find precise bounding boxes, font sizes,
and text directions for each original field. Falls back to Claude's estimated
bboxes only when the text cannot be found in the PDF.

Overlay steps for each field:
  1. Search for the original text with page.search_for() to get the exact quad.
  2. Build a font-info lookup via page.get_text("dict") for font size detection.
  3. Draw a white filled rectangle/quad precisely over the original text.
  4. Write the translated text at the same position, same size, same direction.
"""

import math
from pathlib import Path

import fitz  # PyMuPDF
from loguru import logger

from src.models import BBox, TranslatedDocument, TranslatedField

# ── Visual constants ──────────────────────────────────────────────────────────
_FONT           = "helv"
_BASE_FONT_SIZE = 8.5
_MIN_FONT_SIZE  = 5.0
_TEXT_COLOR     = (0.0, 0.0, 0.0)    # black — maximum readability over original content
_MASK_COLOR     = (1.0, 1.0, 1.0)    # white
_MASK_PAD       = 1.0                 # tight padding around exact bbox
# Fallback-bbox guard: skip if bbox covers more than this fraction of the page
# (too-large bboxes from Claude's estimates cause big white erasure patches)
_MAX_FALLBACK_W_FRACTION = 0.5
_MAX_FALLBACK_H_FRACTION = 0.15
# Field types where we skip key overlay — structural elements, not data entries
_SKIP_KEY_TYPES = {"header", "signature"}


class VisionOverlay:
    """
    Produces a translated PDF by overlaying translated text onto the original.
    Uses PyMuPDF text extraction for exact bounding boxes when possible.
    """

    def render(
        self,
        translated_doc: TranslatedDocument,
        original_pdf: Path,
        output_path: Path,
    ) -> Path:
        logger.info("=== VISION OVERLAY START ===")
        logger.info(f"  source : {original_pdf}")
        logger.info(f"  output : {output_path}")
        output_path.parent.mkdir(parents=True, exist_ok=True)

        try:
            doc = fitz.open(str(original_pdf))
        except Exception as exc:
            logger.error(f"Cannot open PDF: {exc}")
            raise

        for tpage in translated_doc.pages:
            page_idx = tpage.page_number - 1
            if page_idx >= doc.page_count:
                logger.warning(
                    f"Page {tpage.page_number} exceeds PDF page count "
                    f"({doc.page_count}) — skipping"
                )
                continue

            page = doc[page_idx]
            pw = page.rect.width
            ph = page.rect.height
            logger.info(
                f"  page {tpage.page_number}: {len(tpage.translated_fields)} fields "
                f"({pw:.0f}×{ph:.0f} pt)"
            )

            # Build exact text lookup for this page
            text_info = _build_text_info(page)

            for tf in tpage.translated_fields:
                self._overlay_field(page, tf, pw, ph, text_info)

        try:
            doc.save(str(output_path), garbage=4, deflate=True)
        except Exception as exc:
            logger.error(f"Failed to save output PDF: {exc}")
            raise
        finally:
            doc.close()

        logger.info(f"=== VISION OVERLAY DONE → {output_path} ===")
        return output_path

    # ── Field dispatcher ──────────────────────────────────────────────────────

    def _overlay_field(
        self,
        page: fitz.Page,
        tf: TranslatedField,
        pw: float,
        ph: float,
        text_info: dict,
    ) -> None:
        field = tf.original

        # Overlay key/label — skip structural types (headers, signatures) and
        # unchanged translations, to avoid erasing original form structure
        if (
            field.key_bbox
            and tf.translated_key.strip()
            and tf.translated_key.strip() != field.key.strip()
            and field.field_type not in _SKIP_KEY_TYPES
        ):
            self._overlay_text(
                page, field.key, tf.translated_key,
                field.key_bbox, pw, ph, text_info,
            )

        # Overlay value — skip empty, redacted, or unchanged
        _SKIP = {"", "[REDACTED]", "[BLANK]", "[redacted]", "[blank]"}
        if (
            field.value_bbox
            and tf.translated_value.strip()
            and tf.translated_value not in _SKIP
            and field.value.strip() not in _SKIP
            and tf.translated_value.strip() != field.value.strip()
        ):
            self._overlay_text(
                page, field.value, tf.translated_value,
                field.value_bbox, pw, ph, text_info,
            )

    def _overlay_text(
        self,
        page: fitz.Page,
        original: str,
        translated: str,
        fallback_bbox: BBox,
        pw: float,
        ph: float,
        text_info: dict,
    ) -> None:
        """
        Try exact PyMuPDF search first (text-based PDFs).
        For scanned pages, fall back to Claude's estimated bbox.
        The fallback is guarded against oversized bboxes that would erase too much.
        """
        if not translated.strip():
            return

        original_stripped = original.strip()

        # ── Try exact search ─────────────────────────────────────────────────
        if original_stripped:
            try:
                quads = page.search_for(original_stripped, quads=True)
            except Exception:
                quads = []

            if quads:
                span_info = text_info.get(original_stripped, {})
                font_size = span_info.get("size", _BASE_FONT_SIZE)
                for quad in quads:
                    self._draw_quad_overlay(page, quad, translated, font_size)
                logger.debug(f"  exact match: {original_stripped[:40]!r} → {translated[:40]!r}")
                return

        # ── Fallback: Claude's estimated bbox ─────────────────────────────────
        # Skip if the bbox is implausibly large (Claude sometimes estimates huge areas)
        if (
            fallback_bbox.w > _MAX_FALLBACK_W_FRACTION
            or fallback_bbox.h > _MAX_FALLBACK_H_FRACTION
        ):
            logger.debug(
                f"  skipping oversized fallback bbox "
                f"({fallback_bbox.w:.2f}×{fallback_bbox.h:.2f}): {translated[:30]!r}"
            )
            return

        span_info = text_info.get(original_stripped, {})
        font_size = span_info.get("size", 0)
        self._draw_bbox_overlay(page, fallback_bbox, translated, pw, ph, font_size)
        logger.debug(f"  fallback bbox: {translated[:40]!r}")

    # ── Exact-quad overlay (from page.search_for) ─────────────────────────────

    def _draw_quad_overlay(
        self,
        page: fitz.Page,
        quad: fitz.Quad,
        text: str,
        detected_font_size: float,
    ) -> None:
        if not text.strip():
            return

        rect = quad.rect
        if rect.width < 2 or rect.height < 2:
            return

        # Determine rotation angle from quad corners
        ul = quad.ul
        ur = quad.ur
        angle_deg = math.degrees(math.atan2(ur.y - ul.y, ur.x - ul.x))

        # Height of the quad (perpendicular to text direction)
        side_h = math.hypot(quad.ul.x - quad.ll.x, quad.ul.y - quad.ll.y)
        font_size = max(
            _MIN_FONT_SIZE,
            min(
                detected_font_size if detected_font_size > _MIN_FONT_SIZE else _BASE_FONT_SIZE,
                side_h * 0.85,
            ),
        )

        # White mask — use exact rect padded
        mask = fitz.Rect(
            rect.x0 - _MASK_PAD,
            rect.y0 - _MASK_PAD,
            rect.x1 + _MASK_PAD,
            rect.y1 + _MASK_PAD,
        ).intersect(page.rect)

        try:
            page.draw_rect(mask, color=_MASK_COLOR, fill=_MASK_COLOR, width=0)
        except Exception as exc:
            logger.debug(f"  draw_rect failed: {exc}")
            return

        # Text insertion point: baseline just inside the top of the rect
        baseline = fitz.Point(rect.x0, rect.y0 + font_size * 0.9)

        try:
            if abs(angle_deg) < 1.5:
                # Horizontal text
                lines = _wrap_text(text, font_size, rect.width)
                line_h = font_size * 1.2
                for i, line in enumerate(lines):
                    y = rect.y0 + font_size * 0.9 + i * line_h
                    if y > rect.y1 + _MASK_PAD:
                        break
                    page.insert_text(
                        fitz.Point(rect.x0, y),
                        line,
                        fontname=_FONT,
                        fontsize=font_size,
                        color=_TEXT_COLOR,
                    )
            else:
                # Rotated text — morph rotates the glyph around baseline
                rot = fitz.Matrix(angle_deg)
                page.insert_text(
                    baseline,
                    text,
                    fontname=_FONT,
                    fontsize=font_size,
                    color=_TEXT_COLOR,
                    morph=(baseline, rot),
                )
        except Exception as exc:
            logger.warning(f"  insert_text failed ({text[:30]!r}): {exc}")

    # ── Fallback bbox overlay (Claude's estimated bbox) ───────────────────────

    def _draw_bbox_overlay(
        self,
        page: fitz.Page,
        bbox: BBox,
        text: str,
        pw: float,
        ph: float,
        detected_font_size: float,
    ) -> None:
        if not text.strip():
            return

        x0 = bbox.x * pw
        y0 = bbox.y * ph
        x1 = (bbox.x + bbox.w) * pw
        y1 = (bbox.y + bbox.h) * ph

        mask = fitz.Rect(
            x0 - _MASK_PAD, y0 - _MASK_PAD,
            x1 + _MASK_PAD, y1 + _MASK_PAD,
        ).intersect(page.rect)

        if mask.width < 4 or mask.height < 4:
            return

        box_h = y1 - y0
        box_w = x1 - x0
        # If no font detected, estimate from box height; cap at 11pt to avoid giant text
        inferred = min(box_h * 0.75, 11.0)
        font_size = max(
            _MIN_FONT_SIZE,
            min(
                detected_font_size if detected_font_size > _MIN_FONT_SIZE else inferred,
                box_h * 0.85,
            ),
        )

        try:
            page.draw_rect(mask, color=_MASK_COLOR, fill=_MASK_COLOR, width=0)
            lines = _wrap_text(text, font_size, box_w)
            line_h = font_size * 1.2
            for i, line in enumerate(lines):
                y = y0 + font_size * 0.9 + i * line_h
                if y > y1 + _MASK_PAD:
                    break
                page.insert_text(
                    fitz.Point(x0, y),
                    line,
                    fontname=_FONT,
                    fontsize=font_size,
                    color=_TEXT_COLOR,
                )
        except Exception as exc:
            logger.warning(f"  Failed to overlay {text[:40]!r}: {exc}")


# ── Page text info builder ─────────────────────────────────────────────────────

def _build_text_info(page: fitz.Page) -> dict:
    """
    Build a lookup: original_text → {"size": float, "font": str}
    using PyMuPDF's span-level text extraction.
    """
    lookup: dict = {}
    try:
        blocks = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]
    except Exception:
        return lookup

    for block in blocks:
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span.get("text", "").strip()
                if text:
                    lookup[text] = {
                        "size": span.get("size", _BASE_FONT_SIZE),
                        "font": span.get("font", "Helvetica"),
                    }
    return lookup


# ── Text wrapping ─────────────────────────────────────────────────────────────

def _wrap_text(text: str, font_size: float, box_w: float) -> list[str]:
    avg_char_w = font_size * 0.52
    max_chars = max(4, int(box_w / avg_char_w))

    words = text.split()
    lines: list[str] = []
    current = ""
    for word in words:
        candidate = (current + " " + word).strip()
        if len(candidate) <= max_chars:
            current = candidate
        else:
            if current:
                lines.append(current)
            while len(word) > max_chars:
                lines.append(word[:max_chars])
                word = word[max_chars:]
            current = word
    if current:
        lines.append(current)
    return lines or [text[:max_chars]]
