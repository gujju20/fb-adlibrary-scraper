"""Download a country's Asset URL media, named by Library ID.

Usage:
    python download_france_assets.py --country Germany
    python download_france_assets.py --sheet "France Ads" --xlsx topinfoz_ads.xlsx
"""

import argparse
import csv
import io
import json
import re
import time
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode, urlsplit
from urllib.request import Request, urlopen


SHEET_ID = "1NcQkxgxfjN_gQhP5aInp0Wmec20j4OP7lXPeFgm4S-U"
HEADERS = {"User-Agent": "Mozilla/5.0"}


def normalize(value):
    return re.sub(r"[\s_]+", " ", str(value or "").strip()).casefold()


def configure_destination(args):
    if args.sheet is not None:
        args.sheet = args.sheet.strip()
        if not args.sheet:
            raise ValueError("Sheet name cannot be empty.")
        country = re.sub(r"\s+Ads$", "", args.sheet, flags=re.I).strip()
    else:
        country = (args.country if args.country is not None else "France").strip()
        if not country:
            raise ValueError("Country name cannot be empty.")
        names_path = Path(__file__).with_name("country_names.json")
        names = json.loads(names_path.read_text(encoding="utf-8-sig")) if names_path.exists() else {}
        country = next((name for code, name in names.items()
                        if country.casefold() in {code.casefold(), name.casefold()}), country.title())
        # Match the exporter, including Excel's 31-character sheet name limit.
        args.sheet = re.sub(r"[\\/*?:\[\]]", "-", country)[:27] + " Ads"
    if args.folder is None:
        safe_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "-", country).strip(" .").lower()
        if not safe_name:
            raise ValueError("Country/sheet name must contain a usable folder name.")
        args.folder = Path(f"{safe_name} asset")


def read_rows(args):
    if args.xlsx:
        from openpyxl import load_workbook

        book = load_workbook(args.xlsx, read_only=True, data_only=True)
        try:
            title = next((name for name in book.sheetnames
                          if name.casefold() == args.sheet.casefold()), None)
            if title is None:
                raise ValueError(f"Sheet '{args.sheet}' not found. Available: {', '.join(book.sheetnames)}")
            rows = list(book[title].values)
        finally:
            book.close()
    else:
        url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?" + urlencode(
            {"tqx": "out:csv", "sheet": args.sheet}
        )
        try:
            with urlopen(Request(url, headers=HEADERS), timeout=args.timeout) as response:
                content_type = response.headers.get("Content-Type", "").lower()
                body = response.read().decode("utf-8-sig")
            if "html" in content_type or body.lstrip().startswith("<"):
                raise ValueError("Google returned a login/error page instead of CSV.")
        except (HTTPError, URLError, ValueError) as exc:
            raise ValueError(
                "Cannot read the Google Sheet anonymously. Download it as Excel and run "
                "with --xlsx path/to/file.xlsx, or enable link viewing if appropriate. "
                f"Details: {exc}"
            ) from exc
        rows = list(csv.reader(io.StringIO(body)))

    if not rows:
        raise ValueError(f"The '{args.sheet}' sheet is empty.")
    columns = [normalize(value) for value in rows[0]]
    for required in ("library id", "asset url"):
        if required not in columns:
            raise ValueError(f"Missing column: {required}. Found: {columns}")
    return [dict(zip(columns, row)) for row in rows[1:]]


def media_extension(data):
    # Check the actual bytes so login/error HTML is never saved as media.
    if data.startswith(b"\xff\xd8\xff"):
        return ".jpg"
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return ".png"
    if data.startswith((b"GIF87a", b"GIF89a")):
        return ".gif"
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return ".webp"
    if data[4:8] == b"ftyp":
        brands = data[8:64]
        if b"avif" in brands or b"avis" in brands:
            return ".avif"
        if any(brand in brands for brand in (b"heic", b"heix", b"mif1")):
            return ".heic"
        return ".mp4"
    if data.startswith(b"\x1aE\xdf\xa3"):
        return ".webm"
    raise ValueError("URL did not return a supported image/video (possibly expired or a web page).")


def download(url, folder, stem, timeout, retries):
    if urlsplit(url).scheme not in {"http", "https"}:
        raise ValueError("Asset URL must start with http:// or https://")
    temporary = folder / f"{stem}.part"
    for attempt in range(1, retries + 1):
        try:
            with urlopen(Request(url, headers=HEADERS), timeout=timeout) as response:
                first = response.read(8192)
                extension = media_extension(first)
                target = folder / f"{stem}{extension}"
                expected = response.headers.get("Content-Length")
                written = len(first)
                with temporary.open("wb") as output:
                    output.write(first)
                    while chunk := response.read(1024 * 1024):
                        output.write(chunk)
                        written += len(chunk)
                if expected and written != int(expected):
                    raise OSError(f"Incomplete download: {written}/{expected} bytes")
            temporary.replace(target)
            return target
        except (OSError, ValueError) as exc:
            temporary.unlink(missing_ok=True)
            if attempt == retries or isinstance(exc, ValueError):
                raise
            print(f"  Retry {attempt}/{retries}: {exc}", flush=True)
            time.sleep(min(attempt * 2, 10))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    selection = parser.add_mutually_exclusive_group()
    selection.add_argument("--country", help="Country name or code, e.g. Germany or DE (default: France)")
    selection.add_argument("--sheet", help="Exact sheet name, e.g. 'Germany Ads'")
    parser.add_argument("--folder", type=Path, help="Override the default '<country> asset' folder")
    parser.add_argument("--xlsx", type=Path, help="Read an Excel download instead of Google Sheets")
    parser.add_argument("--timeout", type=int, default=60)
    parser.add_argument("--retries", type=int, default=3)
    args = parser.parse_args()
    if args.timeout < 1 or args.retries < 1:
        parser.error("Timeout and retries must be positive.")
    try:
        configure_destination(args)
        print(f"Sheet: {args.sheet} | Folder: {args.folder}", flush=True)
        rows = read_rows(args)
    except (OSError, ValueError, KeyError) as exc:
        parser.exit(1, f"Error: {exc}\n")

    args.folder.mkdir(parents=True, exist_ok=True)
    counts = {}
    seen = set()
    downloaded = failed = skipped = 0
    report = args.folder / "download_report.csv"
    with report.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(["Library ID", "Asset URL", "File", "Status", "Error"])
        for row in rows:
            library_id = str(row.get("library id") or "").strip()
            value = str(row.get("asset url") or "").strip()
            if not library_id and not value:
                continue
            if not re.fullmatch(r"[0-9]+", library_id) or not value:
                skipped += 1
                writer.writerow([library_id, value, "", "Skipped", "Missing URL or invalid Library ID"])
                handle.flush()
                continue
            for url in re.split(r"\s*\|\s*|[\r\n]+", value):
                url = url.strip()
                if not url or (library_id, url) in seen:
                    continue
                seen.add((library_id, url))
                counts[library_id] = counts.get(library_id, 0) + 1
                number = counts[library_id]
                stem = library_id if number == 1 else f"{library_id}_{number}"
                try:
                    target = download(url, args.folder, stem, args.timeout, args.retries)
                    downloaded += 1
                    writer.writerow([library_id, url, target.name, "Downloaded", ""])
                    print(f"Downloaded: {target.name}", flush=True)
                except (OSError, ValueError) as exc:
                    failed += 1
                    writer.writerow([library_id, url, "", "Failed", str(exc)])
                    print(f"Failed: {stem}: {exc}", flush=True)
                handle.flush()
    print(f"Done: {downloaded} downloaded, {failed} failed, {skipped} skipped. "
          f"Folder: {args.folder.resolve()}")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
