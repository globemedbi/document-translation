FROM python:3.12-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
    libgl1 \
    libglib2.0-0 \
    fonts-liberation \
    fonts-noto \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
ENV PYTHONPATH=/app

COPY pyproject.toml ./
RUN pip install --no-cache-dir .

COPY . .
RUN mkdir -p output workspace

EXPOSE 8000 8501
