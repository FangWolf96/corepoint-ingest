# -------- Build stage: install deps ----------
FROM python:3.11-slim AS build

# Avoid bytecode and pip progress bars
ENV PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1

# System deps for pandas/openpyxl/xlsxwriter if needed
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN python -m venv /opt/venv && /opt/venv/bin/pip install --upgrade pip && \
    /opt/venv/bin/pip install -r requirements.txt

# -------- Runtime stage: tiny image ----------
FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PATH="/opt/venv/bin:$PATH"

# Add a non-root user
RUN useradd -ms /bin/bash appuser
WORKDIR /app

# Copy venv from build stage
COPY --from=build /opt/venv /opt/venv

# Copy app code
COPY app.py /app/app.py

# Healthcheck so orchestrators know it’s alive
HEALTHCHECK --interval=30s --timeout=5s --retries=3 CMD wget -qO- http://127.0.0.1:8000/ || exit 1

# Listen on 8000 inside container
EXPOSE 8000

USER appuser

# Gunicorn with 2 workers is fine here
CMD ["gunicorn", "-w", "2", "-b", "0.0.0.0:8000", "app:app"]
