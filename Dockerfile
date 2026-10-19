# Toll Operations Desk on a server (Railway or any Docker host). See DEPLOY_RAILWAY.md.
FROM python:3.12-slim

# TZ: the report is dated "today" in India, not in the server's UTC.
# DATA_DIR: the persistent disk (a Railway volume mounted at /data) holding history, the saved portal sign-in and .env.
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    TZ=Asia/Kolkata \
    HOSTED=1 \
    DATA_DIR=/data \
    PLAYWRIGHT_BROWSERS_PATH=/ms-playwright

WORKDIR /app

COPY requirements.txt .
RUN apt-get update \
 && apt-get install -y --no-install-recommends tzdata \
 && pip install --no-cache-dir -r requirements.txt \
 && playwright install --with-deps chromium-headless-shell \
 && rm -rf /var/lib/apt/lists/*

COPY . .

# One worker only: the program keeps a single signed-in browser that every request shares. Threads serve the
# dashboard and the long report run at the same time.
CMD ["sh", "-c", "exec gunicorn --bind 0.0.0.0:${PORT:-8080} --workers 1 --threads 16 --timeout 900 --access-logfile - app:app"]
