FROM python:3.11-slim

# Install build dependencies for libgourou, OCR, and related tools
RUN apt-get update && apt-get install -y --no-install-recommends \
    git cmake make g++ \
    libpugixml-dev libzip-dev libssl-dev libcurl4-openssl-dev \
    tesseract-ocr \
    tesseract-ocr-osd \
    tesseract-ocr-eng \
    tesseract-ocr-chi-tra \
    tesseract-ocr-chi-sim \
    tesseract-ocr-jpn \
    tesseract-ocr-kor \
    ghostscript \
    unpaper \
    pngquant \
    curl \
    && rm -rf /var/lib/apt/lists/*

# libmupdf-dev intentionally NOT installed — PyMuPDF bundles its own
# MuPDF and the system package can cause import conflicts / segfaults.

# Ensure tesseract can always find its trained models.
# python:3.11-slim is based on Debian Bookworm which ships tesseract 5.x.
# If the path doesn't exist at build time (unlikely), tesseract falls back
# to its compiled-in default, so this is a safety net, not a hard requirement.
ENV TESSDATA_PREFIX=/usr/share/tesseract-ocr/5/tessdata/

WORKDIR /app

# Clone and build libgourou
RUN git clone --recurse-submodules https://forge.soutade.fr/soutade/libgourou.git /app/libgourou \
    && cd /app/libgourou \
    && make BUILD_UTILS=1 BUILD_STATIC=1 BUILD_SHARED=0 \
    && ls -la /app/libgourou/utils/acsmdownloader

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy app code
COPY app.py converter.py ./
COPY templates/ templates/

# Create data directories
RUN mkdir -p uploads output covers

EXPOSE 8080

HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD curl -f http://localhost:8080/login || exit 1

CMD ["gunicorn", "app:app", "--bind", "0.0.0.0:8080", "--threads", "4", "--timeout", "600"]
