FROM python:3.13-slim

# Tesseract is an operating-system package, not just a Python dependency.
# Installing it in the production image lets UPG OCR image-only bank PDFs
# while retaining word coordinates for its existing geometry parser.
RUN apt-get update && apt-get install -y --no-install-recommends \
    tesseract-ocr \
    tesseract-ocr-eng \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt
COPY . ./

ENV PYTHONDONTWRITEBYTECODE=1
CMD ["python", "-B", "run_current.py"]
