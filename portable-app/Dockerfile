# Report Card Generator - Docker Image
# Includes Python 3.11 + LibreOffice for PDF generation

FROM python:3.11-slim-bookworm

# Install LibreOffice and dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-writer \
    fonts-liberation \
    fonts-dejavu \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy requirements first (for Docker layer caching)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY app/ ./app/
COPY config/ ./config/
COPY input_files/ ./input_files/

# Create output directory
RUN mkdir -p /app/output

# Set environment variables
ENV FLASK_APP=app/main.py
ENV PYTHONUNBUFFERED=1
ENV PORT=8080

# Expose Flask port
EXPOSE 8080

# Run the Flask app (uses PORT env variable)
CMD ["python", "app/main.py"]

