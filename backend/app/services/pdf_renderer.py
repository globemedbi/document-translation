from __future__ import annotations

from pathlib import Path
from typing import Any

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    HRFlowable,
    KeepTogether,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from backend.app.models.schemas import DocumentField, DocumentTable, ExtractedDocument

# ---------------------------------------------------------------------------
# Colour palette
# ---------------------------------------------------------------------------
C_NAVY        = colors.HexColor("#1E3A8A")
C_NAVY_MID    = colors.HexColor("#3B5FCE")
C_NAVY_LIGHT  = colors.HexColor("#DBEAFE")
C_TEXT        = colors.HexColor("#111827")
C_SLATE       = colors.HexColor("#475569")
C_MUTED       = colors.HexColor("#94A3B8")
C_BG_ALT      = colors.HexColor("#F8FAFC")
C_BORDER      = colors.HexColor("#CBD5E1")
C_WHITE       = colors.white

# Severity palette  (bg, left-bar, text)
_SEV: dict[str, tuple[colors.Color, colors.Color, colors.Color]] = {
    "critical": (colors.HexColor("#FEF2F2"), colors.HexColor("#EF4444"), colors.HexColor("#991B1B")),
    "warning":  (colors.HexColor("#FFFBEB"), colors.HexColor("#F59E0B"), colors.HexColor("#92400E")),
    "info":     (colors.HexColor("#F0F9FF"), colors.HexColor("#38BDF8"), colors.HexColor("#075985")),
}

# Confidence colours
_CONF_COLOR: dict[str, colors.Color] = {
    "high":   colors.HexColor("#166534"),
    "medium": colors.HexColor("#92400E"),
    "low":    colors.HexColor("#991B1B"),
}

# Page geometry
PAGE_W, PAGE_H = A4
MARGIN_H = 1.6 * cm
MARGIN_V = 1.5 * cm
BODY_W   = PAGE_W - 2 * MARGIN_H   # ≈ 17.8 cm


class PdfRenderer:
    # ------------------------------------------------------------------
    # Column widths for the fields table (Field / English / Original / Conf)
    # ------------------------------------------------------------------
    _FIELD_COLS = [3.6 * cm, 7.0 * cm, 4.8 * cm, 2.0 * cm]   # sum = 17.4 cm

    def __init__(self) -> None:
        base = getSampleStyleSheet()

        def _add(name: str, **kw: Any) -> ParagraphStyle:
            style = ParagraphStyle(name=name, **kw)
            base.add(style)
            return style

        self.s_doc_title = _add(
            "DocTitle",
            fontName="Helvetica-Bold",
            fontSize=18,
            leading=22,
            textColor=C_NAVY,
            spaceAfter=4,
        )
        self.s_meta = _add(
            "Meta",
            fontName="Helvetica",
            fontSize=9,
            leading=13,
            textColor=C_SLATE,
            spaceAfter=2,
        )
        self.s_section = _add(
            "SectionTitle",
            fontName="Helvetica-Bold",
            fontSize=11,
            leading=14,
            textColor=C_WHITE,
            backColor=C_NAVY,
            leftIndent=6,
            rightIndent=6,
            spaceBefore=14,
            spaceAfter=6,
            borderPad=4,
        )
        self.s_table_title = _add(
            "TableTitle",
            fontName="Helvetica-Bold",
            fontSize=9,
            leading=12,
            textColor=C_NAVY_MID,
            spaceBefore=8,
            spaceAfter=3,
        )
        self.s_summary = _add(
            "Summary",
            fontName="Helvetica-Oblique",
            fontSize=9,
            leading=13,
            textColor=C_SLATE,
            spaceAfter=5,
        )
        self.s_body = _add(
            "Body",
            fontName="Helvetica",
            fontSize=9,
            leading=12,
            textColor=C_TEXT,
        )
        self.s_cell_head = _add(
            "CellHead",
            fontName="Helvetica-Bold",
            fontSize=8.5,
            leading=11,
            textColor=C_TEXT,
        )
        self.s_cell = _add(
            "Cell",
            fontName="Helvetica",
            fontSize=8.5,
            leading=11,
            textColor=C_TEXT,
            wordWrap="CJK",
        )
        self.s_cell_muted = _add(
            "CellMuted",
            fontName="Helvetica",
            fontSize=8,
            leading=11,
            textColor=C_SLATE,
            wordWrap="CJK",
        )

    # ------------------------------------------------------------------
    # Public entry point
    # ------------------------------------------------------------------
    def render(self, document: ExtractedDocument, output_path: Path) -> Path:
        print(f"[pdf_renderer] Rendering PDF — sections={len(document.sections)}, warnings={len(document.warnings)}")
        try:
            output_path.parent.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            print(f"[pdf_renderer] ERROR: Cannot create output directory {output_path.parent} — {exc}")
            raise

        doc = SimpleDocTemplate(
            str(output_path),
            pagesize=A4,
            leftMargin=MARGIN_H,
            rightMargin=MARGIN_H,
            topMargin=MARGIN_V,
            bottomMargin=MARGIN_V,
        )

        story: list[Any] = []
        self._build_header(story, document)
        self._build_warnings(story, document)
        self._build_sections(story, document)

        try:
            doc.build(story)
        except Exception as exc:
            print(f"[pdf_renderer] ERROR: ReportLab failed to build PDF — {exc}")
            raise

        print(f"[pdf_renderer] PDF written to {output_path}")
        return output_path

    # ------------------------------------------------------------------
    # Header
    # ------------------------------------------------------------------
    def _build_header(self, story: list[Any], document: ExtractedDocument) -> None:
        story.append(Paragraph(document.title, self.s_doc_title))
        story.append(Paragraph(
            f"<b>Source language:</b> {document.source_language_guess or 'Unknown'} &nbsp;|&nbsp; "
            f"<b>Target language:</b> {document.target_language}",
            self.s_meta,
        ))
        story.append(HRFlowable(width="100%", thickness=1.5, color=C_NAVY, spaceAfter=8))

    # ------------------------------------------------------------------
    # Warnings block
    # ------------------------------------------------------------------
    def _build_warnings(self, story: list[Any], document: ExtractedDocument) -> None:
        if not document.warnings:
            return

        story.append(Paragraph("Review Notes", self.s_section))

        for warning in document.warnings:
            sev = warning.severity
            bg, bar, txt_color = _SEV.get(sev, _SEV["info"])
            page_label = f"Page {warning.page_number}" if warning.page_number else "General"

            label_style = ParagraphStyle(
                f"WarnLabel_{sev}",
                parent=self.s_body,
                fontName="Helvetica-Bold",
                fontSize=8,
                textColor=txt_color,
            )
            text_style = ParagraphStyle(
                f"WarnText_{sev}",
                parent=self.s_body,
                fontSize=8.5,
                leading=12,
                textColor=C_TEXT,
                wordWrap="CJK",
            )

            row = Table(
                [[
                    Paragraph(f"{page_label.upper()} · {sev.upper()}", label_style),
                    Paragraph(warning.message, text_style),
                ]],
                colWidths=[2.4 * cm, BODY_W - 2.4 * cm - 0.4 * cm],
                hAlign="LEFT",
            )
            row.setStyle(TableStyle([
                ("BACKGROUND",   (0, 0), (-1, -1), bg),
                ("LINEAFTER",    (0, 0), (0,  0),  1.5, bar),
                ("LEFTPADDING",  (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING",   (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING",(0, 0), (-1, -1), 5),
                ("VALIGN",       (0, 0), (-1, -1), "TOP"),
            ]))
            story.append(row)
            story.append(Spacer(1, 2))

        story.append(Spacer(1, 6))

    # ------------------------------------------------------------------
    # Sections
    # ------------------------------------------------------------------
    def _build_sections(self, story: list[Any], document: ExtractedDocument) -> None:
        for section in document.sections:
            block: list[Any] = []
            block.append(Paragraph(section.title, self.s_section))

            if section.page_numbers:
                pages_str = ", ".join(str(p) for p in section.page_numbers)
                block.append(Paragraph(f"Pages: {pages_str}", self.s_meta))

            if section.summary:
                block.append(Paragraph(section.summary, self.s_summary))

            if section.fields:
                block.append(self._fields_table(section.fields))
                block.append(Spacer(1, 4))

            for table in section.tables:
                block.append(Paragraph(table.title, self.s_table_title))
                block.append(self._data_table(table))
                block.append(Spacer(1, 4))

            story.append(KeepTogether(block[:4]))   # keep header + first element together
            for item in block[4:]:
                story.append(item)

    # ------------------------------------------------------------------
    # Fields table  (Field | English Value | Original | Confidence)
    # ------------------------------------------------------------------
    def _fields_table(self, fields: list[DocumentField]) -> Table:
        header = [
            Paragraph("Field",          self.s_cell_head),
            Paragraph("English Value",  self.s_cell_head),
            Paragraph("Original",       self.s_cell_head),
            Paragraph("Conf.",          self.s_cell_head),
        ]
        rows = [header]
        for f in fields:
            conf_color = _CONF_COLOR.get(f.confidence, C_TEXT)
            conf_style = ParagraphStyle(
                f"Conf_{f.confidence}",
                parent=self.s_cell,
                textColor=conf_color,
                fontName="Helvetica-Bold",
                fontSize=8,
            )
            rows.append([
                Paragraph(f.label_en or "",          self.s_cell),
                Paragraph(f.value_en or "",          self.s_cell),
                Paragraph(f.value_original or "—",   self.s_cell_muted),
                Paragraph(f.confidence,              conf_style),
            ])

        return self._make_table(rows, self._FIELD_COLS)

    # ------------------------------------------------------------------
    # Generic data table (arbitrary columns)
    # ------------------------------------------------------------------
    def _data_table(self, table: DocumentTable) -> Table:
        col_widths = self._even_widths(len(table.columns))
        header = [Paragraph(c, self.s_cell_head) for c in table.columns]
        rows = [header]
        for row in table.rows:
            cells = list(row.cells)
            # pad or trim to match column count
            while len(cells) < len(table.columns):
                cells.append("")
            cells = cells[:len(table.columns)]
            rows.append([Paragraph(str(c), self.s_cell) for c in cells])

        return self._make_table(rows, col_widths)

    # ------------------------------------------------------------------
    # Shared table builder
    # ------------------------------------------------------------------
    def _make_table(self, rows: list[list[Any]], col_widths: list[float]) -> Table:
        tbl = Table(rows, colWidths=col_widths, repeatRows=1, hAlign="LEFT")
        n = len(rows)
        tbl.setStyle(TableStyle([
            # Header row
            ("BACKGROUND",    (0, 0), (-1, 0),      colors.HexColor("#E2E8F0")),
            ("LINEBELOW",     (0, 0), (-1, 0),      1.0, C_NAVY),
            # Alternating body rows
            ("ROWBACKGROUNDS",(0, 1), (-1, n - 1),  [C_WHITE, C_BG_ALT]),
            # Grid
            ("GRID",          (0, 0), (-1, -1),     0.35, C_BORDER),
            # Padding
            ("LEFTPADDING",   (0, 0), (-1, -1),     6),
            ("RIGHTPADDING",  (0, 0), (-1, -1),     6),
            ("TOPPADDING",    (0, 0), (-1, -1),     5),
            ("BOTTOMPADDING", (0, 0), (-1, -1),     5),
            ("VALIGN",        (0, 0), (-1, -1),     "TOP"),
        ]))
        return tbl

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    @staticmethod
    def _even_widths(n: int) -> list[float]:
        available = BODY_W - 0.4 * cm   # small breathing room
        if n <= 0:
            return []
        return [available / n] * n
