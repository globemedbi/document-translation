from __future__ import annotations

from pydantic import BaseModel, Field


class TableSpec(BaseModel):
    name: str = Field(description="Short identifier, e.g. 'patient_info', 'medical_table'")
    rows: int
    cols: int
    col_widths_pct: list[int] = Field(description="Column widths as percentages, must sum to ~100")
    has_header_row: bool
    header_bg_color: str = Field(description="Hex without #, e.g. '1F4E79', or 'none'")
    header_text_color: str = Field(description="Hex without #, e.g. 'FFFFFF', or '000000'")
    row_bg_color: str = Field(description="Hex for data rows background, or 'none'")
    border_color: str = Field(description="Hex for borders, or 'none'")
    fields: list[str] = Field(description="Field keys that belong in this table, in order")
    notes: str = Field(default="", description="Any special layout notes for this table")


class ColorSpec(BaseModel):
    section: str = Field(description="Section name, e.g. 'page_header', 'section_a', 'footer'")
    bg_color: str = Field(description="Hex without #, or 'none'")
    text_color: str = Field(description="Hex without #")
    border_color: str = Field(default="none")


class LayoutPlan(BaseModel):
    header: str = Field(description="Description of the page header: content, logo position, colors, fields")
    footer: str = Field(description="Description of the footer, or 'none'")
    layout: str = Field(description="Overall page structure: number of sections, columns, margins, reading order (LTR/RTL)")
    colors: list[ColorSpec] = Field(description="One entry per distinct colored section or element")
    tables: list[TableSpec] = Field(description="All tables on the page, in top-to-bottom order")
    images: str = Field(description="Logo or image placements with approximate position, or 'none'")
    content: str = Field(description="Non-table content: standalone paragraphs, signature lines, checkboxes outside tables, dividers")
