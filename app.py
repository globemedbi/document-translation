"""
PDF Form Translator — Streamlit frontend.

Sends the PDF to the FastAPI backend and shows a spinner while waiting.

Start backend: uv run uvicorn api:app --port 8000
Start this:    uv run streamlit run app.py
"""

from __future__ import annotations

import os
from pathlib import Path

import httpx
import streamlit as st

st.set_page_config(page_title="PDF Form Translator", page_icon="📄", layout="centered")

LANGUAGES = [
    "French",
    "English",
    "Arabic",
    "Spanish",
    "German",
    "Italian",
    "Portuguese",
    "Dutch",
    "Turkish",
    "Persian (Farsi)",
    "Urdu",
    "Hindi",
    "Russian",
    "Chinese (Simplified)",
    "Japanese",
]

MODES = {
    "AUTO": "Auto-detect — TEXT for digital PDFs, VISION for scanned",
    "TEXT": "Direct replacement — fastest, requires selectable text",
    "VISION": "Overlay on original — fast, works with any PDF",
    "CODING": "Recreate document — slowest, best for complex forms",
}


def _backend_ok(url: str) -> bool:
    try:
        return httpx.get(f"{url.rstrip('/')}/health", timeout=2.0).status_code == 200
    except Exception:
        return False


def main() -> None:
    st.title("📄 PDF Form Translator")

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.header("⚙️ Settings")

        backend = st.text_input("Backend URL", value=os.getenv("BACKEND_URL", "http://localhost:8000"))

        language = st.selectbox("Target Language", LANGUAGES)
        custom = st.text_input("Or custom language", placeholder="e.g. Swedish …")
        if custom.strip():
            language = custom.strip()

        st.divider()
        mode = st.radio("Mode", list(MODES.keys()), captions=list(MODES.values()))

        st.divider()
        if _backend_ok(backend):
            st.success("Backend connected ✅")
        else:
            st.error("Backend offline — run:\n```\nuv run uvicorn api:app --port 8000\n```")

    # ── Upload ────────────────────────────────────────────────────────────────
    uploaded = st.file_uploader("Upload a PDF", type=["pdf"], label_visibility="collapsed")
    if not uploaded:
        st.info("Upload a PDF to get started.")
        return

    st.success(f"📎 **{uploaded.name}** — {len(uploaded.getvalue()) // 1024} KB")

    _, col, _ = st.columns([2, 1, 2])
    with col:
        go = st.button("🚀 Translate", type="primary", use_container_width=True)

    if not go:
        return

    if not _backend_ok(backend):
        st.error("Backend is not reachable. Start it first.")
        return

    # ── Call API ──────────────────────────────────────────────────────────────
    with st.spinner(f"Translating to {language} ({mode} mode) — this may take a few minutes …"):
        try:
            resp = httpx.post(
                f"{backend.rstrip('/')}/translate",
                data={"language": language, "mode": mode},
                files={"file": (uploaded.name, uploaded.getvalue(), "application/pdf")},
                timeout=1800.0,  # up to 30 min (high reasoning effort needs more time)
            )
        except httpx.TimeoutException:
            st.error("Request timed out. Please try again.")
            return
        except Exception as exc:
            st.error(f"Request failed: {exc}")
            return

    if resp.status_code != 200:
        st.error(f"Translation failed (HTTP {resp.status_code}):\n\n{resp.text[-1000:]}")
        return

    # ── Download ──────────────────────────────────────────────────────────────
    st.success("✅ Translation complete!")

    cd = resp.headers.get("content-disposition", "")
    if 'filename="' in cd:
        fname = cd.split('filename="')[1].rstrip('"')
    else:
        stem = Path(uploaded.name).stem
        ext = ".pdf" if "application/pdf" in resp.headers.get("content-type", "") else ".docx"
        fname = f"{stem}_translated{ext}"

    st.download_button(
        label=f"⬇️ Download {fname}",
        data=resp.content,
        file_name=fname,
        mime=resp.headers.get("content-type", "application/octet-stream"),
        use_container_width=True,
    )


if __name__ == "__main__":
    main()
