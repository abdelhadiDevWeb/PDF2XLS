FROM python:3.10-slim-bookworm

# Tesseract for scanned PDFs; libgl1/libglib2.0-0 are needed by OpenCV (camelot, img2table).
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        tesseract-ocr tesseract-ocr-eng tesseract-ocr-fra libgl1 libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    OCR_LANG=fra+eng \
    PORT=7860 \
    WEB_WORKERS=2

RUN useradd --create-home --uid 1000 app
WORKDIR /home/app/pdf2xls

COPY requirements.txt .
RUN pip install -r requirements.txt

COPY --chown=app:app app.py converter.py ./
COPY --chown=app:app templates ./templates

USER app
EXPOSE 7860

# Hosts such as Render and Cloud Run inject $PORT; OCR of large scans can take minutes.
CMD gunicorn --bind 0.0.0.0:${PORT} --workers ${WEB_WORKERS} --threads 4 --timeout 300 app:app
