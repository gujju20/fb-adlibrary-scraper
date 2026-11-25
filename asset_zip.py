"""Build a ZIP of exported media, with a report for every asset/missing URL."""
import argparse
import csv
import io
import ipaddress
import json
import os
from pathlib import Path
import re
import socket
import tempfile
import time
from urllib.parse import urlsplit
from urllib.request import Request, build_opener, HTTPRedirectHandler, ProxyHandler
from zipfile import ZipFile, ZIP_STORED

from openpyxl import load_workbook
from download_france_assets import HEADERS, media_extension, normalize

ALLOWED_DOMAINS = ('fbcdn.net', 'fbsbx.com', 'facebook.com', 'cdninstagram.com')


class AssetCancelled(Exception):
    pass


def check_cancel(flag):
    if flag and Path(flag).exists():
        raise AssetCancelled('Stopped by user')


def validate_media_url(url):
    parts = urlsplit(url)
    host = (parts.hostname or '').lower()
    if (parts.scheme != 'https' or parts.username or parts.password
            or parts.port not in (None, 443)
            or not any(host == domain or host.endswith('.' + domain) for domain in ALLOWED_DOMAINS)):
        raise ValueError('Only HTTPS Meta media hosts are supported')
    addresses = socket.getaddrinfo(host, 443, type=socket.SOCK_STREAM)
    if not addresses or any(not ipaddress.ip_address(item[4][0]).is_global for item in addresses):
        raise ValueError('Media host must resolve to a public address')


class SafeRedirect(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        validate_media_url(newurl)
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def collect_assets(xlsx, country='ALL'):
    book = load_workbook(xlsx, read_only=True, data_only=True)
    assets, missing, seen, ids = [], [], set(), set()
    try:
        titles = [title for title in book.sheetnames if title.endswith(' Ads') and title != 'Repeated Ads']
        if country.upper() != 'ALL':
            wanted = re.sub(r'[\\/*?:\[\]]', '-', country)[:27] + ' Ads'
            titles = [title for title in titles if title.casefold() == wanted.casefold()]
        if not titles:
            raise ValueError('No matching country ad sheets in workbook')

        def rows(title):
            values = iter(book[title].values)
            fields = [normalize(cell) for cell in next(values, ())]
            for row in values:
                yield dict(zip(fields, row))

        def collect(row, source):
            library_id = str(row.get('library id') or '').strip()
            if not re.fullmatch(r'\d+', library_id):
                return
            value = str(row.get('asset url') or '').strip()
            version = str(row.get('version') or '')
            if not value:
                missing.append(dict(library_id=library_id, source=source, version=version,
                                    url='', status='Missing', error='No downloadable asset URL was exported'))
            for url in re.split(r'\s*\|\s*|[\r\n]+', value):
                url = url.strip()
                if url and (library_id, url) not in seen:
                    seen.add((library_id, url))
                    assets.append(dict(library_id=library_id, source=source, version=version, url=url))

        for title in titles:
            for row in rows(title):
                ids.add(str(row.get('library id') or '').strip())
                collect(row, title)
        if 'Versions' in book.sheetnames:
            for row in rows('Versions'):
                if str(row.get('library id') or '').strip() in ids:
                    collect(row, 'Versions')
    finally:
        book.close()
    return assets, missing


def download_media(url, path, cancel_file=None, max_bytes=512 * 1024 * 1024, timeout=30, retries=3):
    opener = build_opener(ProxyHandler({}), SafeRedirect())
    for attempt in range(retries):
        check_cancel(cancel_file)
        try:
            validate_media_url(url)
            with opener.open(Request(url, headers=HEADERS), timeout=timeout) as response:
                first = response.read(8192)
                extension = media_extension(first)
                expected = response.headers.get('Content-Length')
                if expected and int(expected) > max_bytes:
                    raise ValueError('Asset exceeds configured file size limit')
                written = 0
                with path.open('wb') as handle:
                    chunk = first
                    while chunk:
                        check_cancel(cancel_file)
                        written += len(chunk)
                        if written > max_bytes:
                            raise ValueError('Asset exceeds configured file size limit')
                        handle.write(chunk)
                        chunk = response.read(1024 * 1024)
                if expected and written != int(expected):
                    raise OSError('Incomplete media response')
                return extension, written
        except (OSError, ValueError):
            path.unlink(missing_ok=True)
            if attempt + 1 == retries:
                raise
            for _ in range((attempt + 1) * 5):
                check_cancel(cancel_file)
                time.sleep(.2)
    raise RuntimeError('Download did not finish')


def build_zip(xlsx, output, country='ALL', cancel_file=None, emit=lambda **data: None,
              downloader=download_media, max_total=None, max_asset=None):
    max_total = max_total or int(os.environ.get('MAX_ZIP_BYTES', 4 * 1024**3))
    max_asset = max_asset or int(os.environ.get('MAX_ASSET_BYTES', 512 * 1024**2))
    assets, missing = collect_assets(xlsx, country)
    output = Path(output)
    temporary = output.with_suffix('.part')
    downloaded = failed = completed = used = 0
    cancelled = False
    counts = {}
    started = time.monotonic()
    report = io.StringIO(newline='')
    writer = csv.writer(report)
    writer.writerow(['Library ID', 'Source', 'Version', 'Asset URL', 'File', 'Status', 'Error'])
    def report_row(row, filename, status, error=''):
        # Untrusted sheet text must not become an Excel formula in the CSV report.
        cells = [row['library_id'], row['source'], row['version'], row['url'], filename, status, error]
        writer.writerow(["'" + str(cell) if str(cell).startswith(('=', '+', '-', '@')) else cell for cell in cells])
    for row in missing:
        report_row(row, '', row['status'], row['error'])
    try:
        with tempfile.TemporaryDirectory(dir=output.parent) as folder, ZipFile(temporary, 'w', compression=ZIP_STORED, allowZip64=True) as archive:
            emit(phase='assets', completed=0, total=len(assets), downloaded=0, failed=0, elapsed=0)
            for row in assets:
                try:
                    check_cancel(cancel_file)
                    if used >= max_total:
                        raise ValueError('ZIP media size limit reached; increase MAX_ZIP_BYTES on the server')
                    path = Path(folder) / 'asset.part'
                    extension, size = downloader(row['url'], path, cancel_file=cancel_file,
                                                 max_bytes=min(max_asset, max_total - used))
                    counts[row['library_id']] = counts.get(row['library_id'], 0) + 1
                    name = f"media/{row['library_id']}_{counts[row['library_id']]:03d}{extension}"
                    archive.write(path, name)
                    path.unlink(missing_ok=True)
                    used += size
                    downloaded += 1
                    report_row(row, name, 'Downloaded')
                    emit(phase='log', message=f'Downloaded {name}')
                except AssetCancelled:
                    cancelled = True
                    for remaining in assets[completed:]:
                        report_row(remaining, '', 'Cancelled', 'Stopped before download finished')
                    break
                except (OSError, ValueError) as exc:
                    failed += 1
                    report_row(row, '', 'Failed', str(exc))
                    emit(phase='log', message=f"Asset for {row['library_id']} failed: {exc}")
                completed += 1
                emit(phase='assets', completed=completed, total=len(assets), downloaded=downloaded,
                     failed=failed, elapsed=time.monotonic() - started)
            summary = dict(downloaded=downloaded, failed=failed, missing=len(missing), cancelled=cancelled,
                           total=len(assets), completed=completed, bytes=used)
            archive.writestr('download_report.csv', '\ufeff' + report.getvalue())
            archive.writestr('summary.json', json.dumps(summary, indent=2))
        temporary.replace(output)
        emit(phase='finished', success=not cancelled and not failed and not missing,
             message='ZIP saved; see download_report.csv for missing or failed assets', **summary)
        return summary
    except BaseException:
        temporary.unlink(missing_ok=True)
        raise


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--xlsx', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--country', default='ALL')
    parser.add_argument('--cancel-file', type=Path)
    args = parser.parse_args()
    def emit(**data):
        print('@event ' + json.dumps(data), flush=True)
    try:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        build_zip(args.xlsx, args.output, args.country, args.cancel_file, emit)
    except Exception as exc:
        emit(phase='error', message=str(exc))
        return 1
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
