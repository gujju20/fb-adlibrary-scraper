FROM python:3.12-slim-bookworm
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1 PLAYWRIGHT_BROWSERS_PATH=/ms-playwright HOST=0.0.0.0 PORT=8000 DATA_DIR=/app/data
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt && python -m playwright install --with-deps chromium && rm -rf /var/lib/apt/lists/*
COPY meta_ads_export.py download_france_assets.py asset_zip.py jobs.py web_app.py country_names.json ./
COPY web ./web
RUN useradd --create-home studio && mkdir -p /app/data && chown -R studio:studio /app /ms-playwright
USER studio
EXPOSE 8000
CMD ["python", "web_app.py"]
