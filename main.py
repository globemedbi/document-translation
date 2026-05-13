#!/usr/bin/env python3
"""
PDF Form Translator — entry point.

Usage:
    uv run python main.py <pdf_file>
    uv run python main.py <pdf_file> --mode CODING --language French
    uv run python main.py <pdf_file> --mode AUTO   # auto-detects TEXT vs VISION
    uv run python main.py <pdf_file> --output my_output.pdf
"""
from __future__ import annotations

import sys
from pathlib import Path

import click
from loguru import logger
from rich.console import Console
from rich.panel import Panel

from src.config import SolutionMode, settings

console = Console()


def _setup_logging(output_dir: Path) -> None:
    logger.remove()
    logger.add(
        sys.stderr,
        format=(
            "<green>{time:HH:mm:ss}</green> | "
            "<level>{level: <8}</level> | "
            "<cyan>{name}</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan>"
            " — <level>{message}</level>"
        ),
        level=settings.log_level,
        colorize=True,
    )
    output_dir.mkdir(parents=True, exist_ok=True)
    log_file = output_dir / "translation.log"
    logger.add(
        str(log_file),
        format="{time:YYYY-MM-DD HH:mm:ss} | {level: <8} | {name}:{function}:{line} — {message}",
        level="DEBUG",
        rotation="10 MB",
        retention="7 days",
        encoding="utf-8",
    )


@click.command()
@click.argument("pdf_path", type=click.Path(exists=True, path_type=Path))
@click.option("--output", "-o", type=click.Path(path_type=Path), default=None,
              help="Output file path.")
@click.option("--mode", "-m",
              type=click.Choice(["VISION", "CODING", "TEXT", "AUTO"], case_sensitive=False),
              default=None,
              help="Reconstruction mode (AUTO detects TEXT vs VISION automatically).")
@click.option("--language", "-l", default=None,
              help="Target language (e.g. 'French', 'Arabic', 'Spanish').")
def main(pdf_path: Path, output: Path | None, mode: str | None, language: str | None) -> None:
    """Translate a PDF form to a target language and reconstruct it."""

    if mode:
        settings.solution_mode = SolutionMode(mode.upper())
    if language:
        settings.target_language = language.strip()

    _setup_logging(settings.output_dir)
    stem = pdf_path.stem.replace(" ", "_")

    console.print(Panel(
        f"[bold]PDF Form Translator[/bold]\n\n"
        f"  Input   : {pdf_path}\n"
        f"  Mode    : [cyan]{settings.solution_mode.value}[/cyan]\n"
        f"  Language: [cyan]{settings.target_language}[/cyan]\n"
        f"  Model   : [cyan]{settings.anthropic_model}[/cyan]",
        title="Configuration",
        style="bold white",
    ))

    # ── Resolve AUTO ──────────────────────────────────────────────────────────
    if settings.solution_mode == SolutionMode.AUTO:
        from src.pdf_utils import is_text_pdf
        detected = SolutionMode.TEXT if is_text_pdf(pdf_path) else SolutionMode.VISION
        logger.info(f"AUTO: selected {detected.value}")
        console.print(f"[cyan]AUTO[/cyan] → [bold]{detected.value}[/bold]")
        settings.solution_mode = detected

    # ── TEXT mode — direct text-layer replacement (no Vision extraction) ──────
    if settings.solution_mode == SolutionMode.TEXT:
        if output is None:
            d = settings.output_dir / "text"
            d.mkdir(parents=True, exist_ok=True)
            output = d / f"{stem}_translated.pdf"

        console.rule("[bold]TEXT — Direct Translation[/bold]")
        logger.info("TEXT mode: translating embedded text spans …")

        from src.solutions.text_layer import TextLayerTranslator
        try:
            result_path = TextLayerTranslator().translate(pdf_path, output)
        except Exception as exc:
            logger.error(f"Text layer failed: {exc}")
            console.print(f"[red bold]✗ Text layer failed:[/red bold] {exc}")
            sys.exit(1)

        console.print(f"[green]✓[/green] Translated PDF → [bold]{result_path}[/bold]")

    # ── VISION / CODING — need extraction + translation first ────────────────
    else:
        # Step 1: Extract
        console.rule("[bold]Step 1 — Extract[/bold]")
        logger.info("Step 1: Extracting form data from PDF …")
        from src.extractor import Extractor
        try:
            extraction = Extractor().extract_document(pdf_path)
        except Exception as exc:
            logger.error(f"Extraction failed: {exc}")
            console.print(f"[red bold]✗ Extraction failed:[/red bold] {exc}")
            sys.exit(1)
        total_fields = sum(len(p.fields) for p in extraction.pages)
        console.print(
            f"[green]✓[/green] Extracted [bold]{total_fields}[/bold] fields "
            f"from [bold]{len(extraction.pages)}[/bold] pages "
            f"(lang: {extraction.source_language})"
        )

        # Step 2: Translate
        console.rule("[bold]Step 2 — Translate[/bold]")
        logger.info(f"Step 2: Translating to {settings.target_language} …")
        from src.translator import Translator
        try:
            translated_doc = Translator().translate_document(extraction)
        except Exception as exc:
            logger.error(f"Translation failed: {exc}")
            console.print(f"[red bold]✗ Translation failed:[/red bold] {exc}")
            sys.exit(1)
        console.print(
            f"[green]✓[/green] Translation to [bold]{settings.target_language}[/bold] complete"
        )

        # Step 3: Reconstruct
        console.rule(f"[bold]Step 3 — Reconstruct ({settings.solution_mode.value})[/bold]")

        if settings.solution_mode == SolutionMode.VISION:
            if output is None:
                d = settings.output_dir / "vision"
                d.mkdir(parents=True, exist_ok=True)
                output = d / f"{stem}_translated.pdf"

            from src.solutions.vision_overlay import VisionOverlay
            try:
                result_path = VisionOverlay().render(translated_doc, pdf_path, output)
            except Exception as exc:
                logger.error(f"Vision overlay failed: {exc}")
                console.print(f"[red bold]✗ Vision overlay failed:[/red bold] {exc}")
                sys.exit(1)
            console.print(f"[green]✓[/green] PDF saved → [bold]{result_path}[/bold]")

        else:  # CODING
            if output is None:
                d = settings.output_dir / "coding"
                d.mkdir(parents=True, exist_ok=True)
                output = d / f"{stem}_translated.docx"

            from src.pdf_utils import save_page_images
            from src.solutions.coding_agent import PAOCodingAgent
            try:
                images_dir = settings.output_dir / "page_images"
                image_paths = save_page_images(pdf_path, images_dir)
                result_path = PAOCodingAgent().run(
                    translated_doc=translated_doc,
                    page_image_paths=image_paths,
                    output_docx=output,
                )
            except Exception as exc:
                logger.error(f"Coding agent failed: {exc}")
                console.print(f"[red bold]✗ Coding agent failed:[/red bold] {exc}")
                sys.exit(1)

            if result_path.exists():
                console.print(
                    f"[green]✓[/green] {result_path.suffix.upper().lstrip('.')} saved "
                    f"→ [bold]{result_path}[/bold]"
                )

    console.print(Panel(
        f"[bold green]Done![/bold green]\n\n"
        f"  Output: {output}\n"
        f"  Logs  : {settings.output_dir / 'translation.log'}",
        title="Complete",
        style="green",
    ))


if __name__ == "__main__":
    main()
