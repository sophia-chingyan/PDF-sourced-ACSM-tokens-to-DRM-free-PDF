FROM python:3.11-slim

# Install build dependencies for libgourou and PyMuPDF
RUN apt-get update && apt-get install -y --no-install-recommends \
    git cmake make g++ \
    libpugixml-dev libzip-dev libssl-dev libcurl4-openssl-dev \
    libmupdf-dev \
    && rm -rf /var/lib/apt/lists/*

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

CMD ["gunicorn", "app:app", "--bind", "0.0.0.0:8080", "--threads", "4", "--timeout", "300"]
