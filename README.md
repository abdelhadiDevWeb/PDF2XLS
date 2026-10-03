---
title: PDF to Excel
sdk: docker
app_port: 7860
---

# PDF → Excel

Upload a PDF, get the table as an `.xlsx` file (one sheet, only the table, values as written in the PDF).

- Text PDFs: ruled-row detection → Camelot (lattice) → pdfplumber → Camelot (stream)
- Scanned PDFs: Tesseract OCR + img2table
- Tables continued over several pages are joined; by default only the main table is kept.

## Run locally (Windows)

```powershell
python -m venv .venv
.\.venv\Scripts\pip install -r requirements-dev.txt
.\.venv\Scripts\python app.py          # http://127.0.0.1:5000
.\.venv\Scripts\python -m unittest discover -s tests -v
```

Scanned PDFs need Tesseract: `winget install UB-Mannheim.TesseractOCR`.

## Configuration (environment variables)

| Variable       | Default              | Purpose                                              |
|----------------|----------------------|------------------------------------------------------|
| `APP_PASSWORD` | *(empty = no login)* | Ask for a password (any username) before using the app |
| `OCR_LANG`     | `eng` (`fra+eng` in Docker) | Tesseract languages for scanned PDFs           |
| `PORT`         | `7860` in Docker     | Port the server listens on (set by most hosts)        |
| `WEB_WORKERS`  | `2`                  | Parallel conversions (use 1 on 512 MB plans)          |

Uploaded PDFs and results are deleted after 1 hour.

## Deploy

The app ships as a Docker image (`Dockerfile`), which bundles Tesseract OCR, so any host that
builds Dockerfiles can run it.

### Hugging Face Spaces (free, enough RAM for OCR)

1. Create a Space at <https://huggingface.co/new-space> → SDK **Docker** → *Blank*.
2. Push this folder to it:
   ```powershell
   git remote add space https://huggingface.co/spaces/<your-user>/<space-name>
   git push space main          # password = a Hugging Face access token with write access
   ```
3. Optional: Space → *Settings* → *Variables and secrets* → add secret `APP_PASSWORD`.

If the push is rejected because of binary files, remove the sample PDFs from git
(`git rm -r --cached samples` then commit) and push again.

### Render

1. Push this folder to a GitHub repository.
2. <https://dashboard.render.com> → **New** → **Blueprint** → pick the repository
   (`render.yaml` configures everything) and enter an `APP_PASSWORD` when asked.

The free plan sleeps when idle (first request takes ~1 minute) and has little memory; choose a
paid plan for large scanned PDFs.

### Google Cloud Run

```bash
gcloud run deploy pdf2xls --source . --region europe-west1 --allow-unauthenticated \
  --memory 2Gi --timeout 300 --set-env-vars APP_PASSWORD=change-me
```

### Any server with Docker (VPS)

```bash
docker build -t pdf2xls .
docker run -d --restart unless-stopped -p 80:7860 -e APP_PASSWORD=change-me pdf2xls
```
