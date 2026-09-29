FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY tapestry_scraper.py .

# Downloads land in /export; the reusable login lives in /cache.
# Mount both from the host so they survive the container.
VOLUME [ "/export", "/cache" ]

ENV TAPESTRY_SESSION_CACHE=/cache/session.json \
    PYTHONUNBUFFERED=1

ENTRYPOINT [ "python", "/app/tapestry_scraper.py", "--output", "/export" ]
