from __future__ import annotations

import json
import os
from typing import Any

import requests
import streamlit as st

BACKEND_URL = os.getenv("BACKEND_URL", "http://localhost:8000")
TRANSLATE_ENDPOINT = f"{BACKEND_URL}/api/v1/translate"

st.set_page_config(page_title="PDF Translator", page_icon="📄", layout="wide")

st.title("PDF Translator")
st.caption("Upload a scanned or digital medical/claim PDF and receive a clean English PDF.")

uploaded_file = st.file_uploader("Upload a PDF", type=["pdf"])


def render_sections(document: dict[str, Any]) -> None:
    st.subheader(document.get("title", "Translated Document"))
    st.write(f"**Source language guess:** {document.get('source_language_guess', 'Unknown')}")
    warnings = document.get("warnings", [])
    if warnings:
        with st.expander("Warnings / Review notes", expanded=True):
            for warning in warnings:
                page = warning.get("page_number")
                prefix = f"Page {page}: " if page else ""
                st.warning(f"{prefix}{warning.get('message', '')}")

    for section in document.get("sections", []):
        st.markdown(f"### {section.get('title', 'Section')}")
        page_numbers = section.get("page_numbers", [])
        if page_numbers:
            st.caption("Pages: " + ", ".join(str(page) for page in page_numbers))
        if section.get("summary"):
            st.write(section["summary"])

        fields = section.get("fields", [])
        if fields:
            st.markdown("**Fields**")
            field_rows = []
            for field in fields:
                field_rows.append(
                    {
                        "Field": field.get("label_en", ""),
                        "English Value": field.get("value_en", ""),
                        "Original": field.get("value_original", ""),
                        "Confidence": field.get("confidence", ""),
                    }
                )
            st.dataframe(field_rows, use_container_width=True)

        tables = section.get("tables", [])
        for table in tables:
            st.markdown(f"**{table.get('title', 'Table')}**")
            columns = table.get("columns", [])
            rows = table.get("rows", [])
            if columns and rows:
                table_data = []
                for row in rows:
                    cells = row.get("cells", [])
                    mapped = {
                        columns[index]: cells[index] if index < len(cells) else ""
                        for index in range(len(columns))
                    }
                    table_data.append(mapped)
                st.dataframe(table_data, use_container_width=True)


if uploaded_file and st.button("Translate document", type="primary"):
    with st.spinner("Uploading and translating PDF..."):
        files = {"file": (uploaded_file.name, uploaded_file.getvalue(), "application/pdf")}
        print("here")
        response = requests.post(TRANSLATE_ENDPOINT, files=files, timeout=300)

    if response.ok:
        result = response.json()
        st.success("Translation complete.")

        left, right = st.columns([2, 1])
        with left:
            render_sections(result["extracted_document"])
        with right:
            st.metric("Pages", result["pages"])
            st.metric("Model used", result["model_used"])
            st.metric("Fallback used", "Yes" if result["fallback_used"] else "No")
            st.link_button("Download PDF", result["pdf_download_url"])
            st.link_button("Download JSON", result["json_download_url"])
            with st.expander("Raw JSON"):
                st.code(json.dumps(result, ensure_ascii=False, indent=2), language="json")
    else:
        st.error(f"Backend error: {response.text}")
