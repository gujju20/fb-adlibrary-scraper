# Ad Library Studio

A Python web app for Facebook Ad Library research: paste a search link, run Chromium headlessly, download an Excel workbook, and collect available ad images/videos in a ZIP. The original desktop UI and command-line tools remain available.

## Run locally on Windows

1. Install Python 3.12 or newer.
2. Double-click `setup_web.bat` once (installs dependencies and Chromium).
3. Open **http://127.0.0.1:8000** in your browser.
4. On later runs, double-click `start_web.bat`.

Or run:

```powershell
python -m pip install -r requirements.txt
python -m playwright install chromium
python web_app.py
```

Linux needs browser OS dependencies: `python -m playwright install --with-deps chromium`.

## Use the app

1. Paste a Facebook Ad Library URL containing `country=XX` or `country=ALL`.
2. Leave the ad limit at **0** to load all available matching results, or set a smaller test limit.
3. Click **Start export**. The browser runs on the server, not on the visitor's PC.
4. Download the **Excel workbook** when scraping finishes.
5. Click **Download all assets**. The server downloads the exported media and prepares a ZIP, then the browser starts a normal file download to the user's computer. If automatic downloading is blocked, use **Download prepared ZIP again**.

The ZIP includes `media/<library_id>_<number>.<extension>`, `download_report.csv`, and `summary.json`. Downloads stream to server disk and ZIP files stream to the browser; the whole archive is not held in browser memory.

The app collects Asset URL values from every exported country ad sheet and the matching Versions rows, deduplicating by Library ID + URL. Actual media bytes determine the file extension. Failed, expired, missing, oversized and cancelled downloads are reported. Previews are not substituted for missing video URLs. A ZIP with no available media still contains the report.

**All countries** uses Meta's `country=ALL` search as supplied. It does not launch hundreds of separate country searches or claim to infer each ad's country. Single-country and worldwide searches use the same Excel and ZIP flow.

The original scraper exports **active ads**. Even if the link includes all statuses, only active cards are exported. Available ad details, versions and transparency depend on what Meta renders. A successful run means the loaded results were processed; it does not guarantee Meta exposed every matching ad or every original video URL.

## GitHub and live deployment

Push this project to a GitHub repository. **GitHub Pages cannot run this Python/Chromium backend.** Deploy the repository to a Docker-capable server or hosting provider. A public URL has not been provisioned by this repository alone.

### Docker / your server

1. Copy `.env.example` to `.env` and set a long, random `APP_PASSWORD`.
2. Run `docker compose up --build -d`.
3. Open port 8000 through your HTTPS reverse proxy/domain.
4. Sign in with username **admin** and the `APP_PASSWORD` you configured.

Compose uses a persistent data volume, one browser worker, process init and a 60-second shutdown grace period. Terminate HTTPS at your hosting provider or reverse proxy. Do not expose unprotected port 8000 directly to the public Internet. Only trust forwarded headers from your actual proxy (`FORWARDED_ALLOW_IPS` for Uvicorn if needed).

### Render from GitHub

`render.yaml` supplies a Docker web service and a persistent disk. In Render, connect your GitHub repository and create a Blueprint from this file, then set `APP_PASSWORD`. The template selects a **paid Standard service and 10 GB disk**; review the provider's current price before deploying. Do not select Static Site. Keep one instance. Alternatively create a Docker Web Service manually, set `/health` as the health check, mount persistent storage at `/app/data`, and configure the environment variables below.

The container builds Chromium and its system dependencies using the installed Playwright version. Browser-heavy unlimited searches need sufficient RAM and disk. The included GitHub Actions workflow runs unit tests, builds the Docker image and checks Chromium launch. This does not deploy automatically or run real Facebook searches in CI.

Official references: [Playwright Docker](https://playwright.dev/python/docs/docker), [Render Docker](https://render.com/docs/docker), [Render persistent disks](https://render.com/docs/disks).

## Runtime settings

| Variable | Default | Purpose |
| --- | --- | --- |
| `HOST` | `127.0.0.1` locally; `0.0.0.0` in Docker | Bind address |
| `PORT` | `8000` | HTTP port; hosting providers can override |
| `APP_PASSWORD` | Empty locally | Shared HTTP Basic password; username is `admin` |
| `DATA_DIR` | `./data` | Job metadata, workbooks and ZIP files |
| `JOB_CONCURRENCY` | `1` | Concurrent browser/media workers |
| `MAX_PENDING_JOBS` | `10` | Global active/queued job limit; two active jobs per browser session |
| `JOB_TIMEOUT_SECONDS` | `21600` | Six-hour worker time limit |
| `RETENTION_HOURS` | `24` | Terminal jobs expire; removed at startup or before a new export |
| `MAX_ASSET_BYTES` | `536870912` | Maximum bytes per media file (512 MiB) |
| `MAX_ZIP_BYTES` | `4294967296` | Maximum media bytes per ZIP (4 GiB, plus report/ZIP overhead) |

Run **one web server process / one container instance per data directory**. The queue is held by that process; job metadata is persisted as JSON. Browser cookies isolate users' jobs. Refreshing the page reconnects to the same job. Clearing cookies loses access to that session's exports. This is a shared-password research tool, not a multi-account SaaS.

Server restarts mark unfinished jobs interrupted and retain saved files; they do not automatically resume scraping. Start another export after interruption. Do not deploy while important jobs are running. `Stop & save` waits for the current network action; a worker that does not stop within 45 seconds is terminated, leaving only already-saved files.

Progress follows worker events. Discovery is indeterminate because the total is unknown. Detail and asset phases show measured counts and estimates after three samples; estimates become unavailable during long stalls. Grouped ads, slow servers and large videos can change the estimate. The final ZIP still needs time to transfer to the user's PC.

Only HTTPS media URLs on Meta-owned CDN domains are downloaded. Redirects and resolved addresses are validated; private/internal addresses are rejected. If an exported original asset uses another host, its failure is recorded for review.

## Verification

```powershell
python -m pip install httpx
python -m unittest -v test_export test_web
```

For an optional real end-to-end browser test with one worldwide ad and its media:

```powershell
python verify_live.py
```

This makes real Facebook/CDN requests and writes XLSX, ZIP and desktop/mobile screenshots under `exports/web-verification/`. Do not use it as an unattended production health check. `/health` only checks that the HTTP app responds.

## Files

- `web_app.py`, `jobs.py`: web API, isolated sessions and background jobs.
- `web/`: responsive UI with progress, logs and downloads.
- `meta_ads_export.py`: existing scraper with structured events and cancellation.
- `asset_zip.py`: media ZIP generation using validation from `download_france_assets.py`.
- `download_france_assets.py`: original standalone asset downloader.
- `scraper_ui.py`, `start_ui.bat`: original desktop tool.

Generated data, exports, local runtime files and secrets are ignored by Git and Docker. Keep `country_names.json` with the project.
