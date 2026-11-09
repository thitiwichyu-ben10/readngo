FROM python:3.11-slim-bookworm

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PLAYWRIGHT_BROWSERS_PATH=/ms-playwright \
    FILE_STORAGE_DIR=/tmp/generated_files

WORKDIR /app

# Thai and Latin fonts are installed in the image; Playwright installs Chromium OS libraries.
RUN apt-get update && apt-get install -y --no-install-recommends \
    ca-certificates fonts-noto-core fonts-noto-extra fonts-thai-tlwg \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt \
    && playwright install --with-deps chromium

COPY . .
RUN mkdir -p /tmp/generated_files

EXPOSE 10000
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT:-10000}"]
