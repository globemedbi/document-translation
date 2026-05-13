FROM python:3.12-slim

# System dependencies
# libgl1 + libglib2.0-0  — required by PyMuPDF for rendering
# fonts-liberation        — Latin fonts for PDF text rendering
# fonts-noto              — broad Unicode coverage (Arabic, CJK, etc.)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libgl1 \
    libglib2.0-0 \
    fonts-liberation \
    fonts-noto \
    && rm -rf /var/lib/apt/lists/*

# Install uv (official binary — fastest Python package manager)
COPY --from=ghcr.io/astral-sh/uv:latest /uv /uvx /usr/local/bin/

WORKDIR /app

# ── Dependency layer (cached unless pyproject.toml or uv.lock changes) ────────
COPY pyproject.toml uv.lock ./
RUN uv sync --frozen --no-dev --no-install-project

# ── Application source ────────────────────────────────────────────────────────
COPY . .
RUN uv sync --frozen --no-dev

# Create runtime directories (overridden by volume mounts in production)
RUN mkdir -p output workspace

# Ports: 8000 = FastAPI, 8501 = Streamlit
EXPOSE 8000 8501
