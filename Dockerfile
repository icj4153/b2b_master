FROM mcr.microsoft.com/playwright/python:v1.51.0-jammy

WORKDIR /app

ENV PYTHONUNBUFFERED=1 \
    TZ=KST-9 \
    B2B_DOWNLOAD_DIR=/app/b2b_downloads \
    B2B_OUTPUT_DIR=/app/output

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt \
    && python -m playwright install chromium

COPY . .

CMD ["python", "scheduler.py"]
