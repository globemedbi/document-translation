from __future__ import annotations

import base64
import io
from pathlib import Path

import fitz  # PyMuPDF
from loguru import logger
from PIL import Image

# 150 DPI gives a good balance of quality vs. token cost for Claude Vision
_RENDER_DPI = 150
_SCALE = _RENDER_DPI / 72  # 72 is PDF's native DPI


def pdf_to_images(pdf_path: Path) -> list[tuple[int, Image.Image]]:
    """Return (1-based page_number, PIL Image) for every page of the PDF."""
    logger.info(f"pdf_to_images: opening {pdf_path.name}")
    pages: list[tuple[int, Image.Image]] = []

    try:
        doc = fitz.open(str(pdf_path))
    except Exception as exc:
        logger.error(f"Cannot open PDF {pdf_path}: {exc}")
        raise

    logger.info(f"pdf_to_images: {doc.page_count} pages found")
    for idx in range(doc.page_count):
        page = doc[idx]
        mat = fitz.Matrix(_SCALE, _SCALE)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        pages.append((idx + 1, img))
        logger.debug(f"  page {idx + 1}: {pix.width}×{pix.height} px")

    doc.close()
    logger.info(f"pdf_to_images: done ({len(pages)} pages)")
    return pages


def save_page_images(pdf_path: Path, output_dir: Path) -> list[Path]:
    """Render every PDF page as PNG and return the saved paths."""
    output_dir.mkdir(parents=True, exist_ok=True)
    saved: list[Path] = []
    for page_num, img in pdf_to_images(pdf_path):
        dest = output_dir / f"page_{page_num:03d}.png"
        img.save(dest, format="PNG")
        saved.append(dest)
        logger.debug(f"  saved page image: {dest}")
    logger.info(f"save_page_images: {len(saved)} images → {output_dir}")
    return saved


def image_to_base64(img: Image.Image, fmt: str = "PNG") -> str:
    buf = io.BytesIO()
    img.save(buf, format=fmt)
    return base64.b64encode(buf.getvalue()).decode("utf-8")


def get_page_count(pdf_path: Path) -> int:
    doc = fitz.open(str(pdf_path))
    count = doc.page_count
    doc.close()
    return count


def is_text_pdf(pdf_path: Path, min_chars_per_page: int = 200) -> bool:
    """
    Return True if the PDF has sufficient embedded text for TEXT mode.

    Heuristic: average character count per page must exceed min_chars_per_page.
    PDFs exported from Word/InDesign typically have 500-3000 chars/page; scanned
    images have near-zero. The default threshold (200) is conservative enough to
    avoid false-positives on forms that have a little bit of embedded text but are
    mostly images.
    """
    try:
        doc = fitz.open(str(pdf_path))
        total = sum(len(doc[i].get_text().strip()) for i in range(doc.page_count))
        pages = doc.page_count
        doc.close()
        avg = total / max(pages, 1)
        logger.debug(f"is_text_pdf: {pdf_path.name} — {total} chars, avg {avg:.0f}/page")
        return avg >= min_chars_per_page
    except Exception as exc:
        logger.warning(f"is_text_pdf check failed: {exc}")
        return False
