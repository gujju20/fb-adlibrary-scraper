import argparse
import json
import re
import shutil
import time
from zipfile import ZipFile
from collections import Counter
from datetime import date, datetime
from pathlib import Path
from urllib.parse import parse_qs, urlencode, urlsplit

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
from playwright.sync_api import Error, TimeoutError, sync_playwright


AD_ID = re.compile(r"^Library ID:\s*(\d+)$")
VERSION_COUNT = re.compile(r"^(\d+) of (\d+)$")
MORE_RESULTS = re.compile(r"^(See|Load) more$", re.I)
CARD_SELECTOR = "div.x1plvlek"

# Change this link for the next country, or pass --url on the command line.
LIBRARY_URL = (
    "https://www.facebook.com/ads/library/?active_status=active&ad_type=all"
    "&country=DE&is_targeted_country=false&media_type=all&q=topinfoz.com"
    "&search_type=keyword_unordered&sort_data[direction]=desc"
    "&sort_data[mode]=total_impressions"
)
COUNTRY_LIMITS = {"DE": 260, "FR": 270}
EVENTS = False
CANCEL_FILE = None


class Cancelled(Exception):
    pass


def check_cancel():
    if CANCEL_FILE and CANCEL_FILE.exists():
        raise Cancelled("Stopped by user")


def event(phase, **values):
    if EVENTS:
        print("@event " + json.dumps(dict(phase=phase, **values)), flush=True)


def source_settings(url, query=None):
    parsed = urlsplit(url.replace("\\", ""))
    if (parsed.scheme != "https" or parsed.hostname not in {"facebook.com", "www.facebook.com"}
            or parsed.path.rstrip("/") != "/ads/library" or parsed.username or parsed.port):
        raise ValueError("Use a Facebook Ad Library URL containing country=XX.")
    params = parse_qs(parsed.query)
    code = params.get("country", [""])[0].upper()
    names_file = Path(__file__).with_name("country_names.json")
    if names_file.exists():
        names = json.loads(names_file.read_text(encoding="utf-8-sig"))
    else:
        raise ValueError("Missing country_names.json. Keep it beside meta_ads_export.py.")
    names["ALL"] = "All countries"
    if code not in names:
        raise ValueError(f"Unknown country code: {code}")
    params["country"] = [code]
    if query:
        params["q"] = [query]
    return parsed._replace(path="/ads/library/", query=urlencode(params, doseq=True), fragment="").geturl(), code, names[code]


def sheet_records(sheet):
    rows = sheet.iter_rows(values_only=True)
    fields = [str(value or "").strip().lower().replace(" ", "_") for value in next(rows, ())]
    for values in rows:
        row = dict(zip(fields, (value if value is not None else "" for value in values)))
        if row.get("library_id"):
            row["library_id"] = str(row["library_id"])
            yield row


class Archive:
    def __init__(self, path, country, limit):
        self.book = load_workbook(path) if path.exists() else Workbook()
        if not path.exists():
            self.book.remove(self.book.active)
        if "Ads" in self.book.sheetnames:
            if "France Ads" in self.book.sheetnames:
                raise ValueError("Both Ads and France Ads exist; cannot safely migrate automatically.")
            self.book["Ads"].title = "France Ads"
        self.country = country
        self.limit = limit
        self.countries = {}
        self.memberships = {}
        self.owners = {}
        self.titles = {}
        if "Country Index" in self.book.sheetnames:
            for name, title in self.book["Country Index"].iter_rows(min_row=2, max_col=2, values_only=True):
                if name and title:
                    self.titles[name] = title
        names_by_title = {title: name for name, title in self.titles.items()}
        for sheet in self.book:
            if sheet.title.endswith(" Ads") and sheet.title != "Repeated Ads":
                name = names_by_title.get(sheet.title, sheet.title[:-4])
                self.titles[name] = sheet.title
                self.countries[name] = {row["library_id"]: row for row in sheet_records(sheet)}
                for library_id in self.countries[name]:
                    self.owners.setdefault(library_id, name)
                    self.memberships.setdefault(library_id, set()).add(name)
        if "Repeated Ads" in self.book.sheetnames:
            for row in sheet_records(self.book["Repeated Ads"]):
                self.memberships.setdefault(row["library_id"], set()).update(
                    name.strip() for name in str(row.get("countries", "")).split(",") if name.strip()
                )
        self.existing_ids = set(self.owners)
        self.countries.setdefault(country, {})
        if country not in self.titles:
            title = re.sub(r"[\\/*?:\[\]]", "-", country)[:27] + " Ads"
            if title in self.titles.values():
                raise ValueError(f"Country sheet name collision: {title}")
            self.titles[country] = title
        self.seen = set()
        self.attempted = set()
        self.versions = {}
        self.locations = {}
        if "Versions" in self.book.sheetnames:
            for row in sheet_records(self.book["Versions"]):
                self.versions[row["library_id"], int(row["version"] or 1)] = row
        if "Locations" in self.book.sheetnames:
            for row in sheet_records(self.book["Locations"]):
                key = tuple(row.get(field, "") for field in (
                    "library_id", "region", "location_type", "location", "source_location_type"))
                self.locations[key] = row

    def register(self, ad):
        library_id = ad["library_id"]
        if library_id not in self.seen:
            if self.limit and len(self.seen) >= self.limit:
                return None
            self.seen.add(library_id)
        self.memberships.setdefault(library_id, set()).add(self.country)
        if library_id in self.existing_ids:
            saved = self.countries[self.owners[library_id]][library_id]
            self.countries[self.country].setdefault(library_id, dict(saved))
            # Include the saved ad in this country, but never scrape it again.
            return None
        self.owners[library_id] = self.country
        return self.countries[self.country].setdefault(library_id, ad)

AD_FIELDS = """
library_id facebook_page_name facebook_page_id active_status started_running_on
active_days primary_text headline description asset_url preview_asset_url
destination_url facebook_page_url version_count versions_exported
locations_exported details_status transparency_status
""".split()

VERSION_FIELDS = """
library_id version version_count primary_text headline description asset_url
preview_asset_url destination_url player_source asset_status
""".split()

LOCATION_FIELDS = """
library_id facebook_page_name region location_type location source_location_type
""".split()

READ_CREATIVE = """
root => {
    const links = [...root.querySelectorAll('a[href]')];
    const cta = links.find(a =>
        a.hostname === 'l.facebook.com' && a.pathname === '/l.php'
    ) || links.find(a => a.querySelector('div._4ik4._4ik5'));

    const fields = [...(cta?.querySelectorAll('div._4ik4._4ik5') || [])];
    const page = links.find(a =>
        /(^|[.])facebook[.]com$/.test(a.hostname)
        && a.hostname !== 'l.facebook.com'
        && !a.pathname.startsWith('/ads/')
        && (a.querySelector('img') || /^[/][0-9]+[/]?$/.test(a.pathname))
    );

    const primaryText = [...root.querySelectorAll('[role="button"] ._4ik4')]
        .filter(node => node.getClientRects().length && !node.closest('a'))
        .map(node => node.querySelector('[style*="white-space: pre-wrap"]'))
        .find(node => node?.textContent.trim());

    const video = root.querySelector('video');
    const images = [...root.querySelectorAll('img')]
        .filter(img => !page?.contains(img) && img.getBoundingClientRect().width >= 100)
        .map(img => img.currentSrc || img.src)
        .filter(Boolean);
    const playerSource = video?.getAttribute('src') || video?.currentSrc
        || video?.querySelector('source')?.src || '';

    return {
        text: root.innerText,
        facebook_page_name: page?.innerText.trim() || page?.querySelector('img')?.alt || '',
        facebook_page_url: page?.href || '',
        destination_url: cta?.href || '',
        primary_text: primaryText?.textContent.trim() || '',
        headline: fields[1]?.innerText.trim() || '',
        description: fields[2]?.innerText.trim() || '',
        player_source: /^https?:/.test(playerSource) ? playerSource : '',
        preview_asset_url: video?.poster || images[0] || '',
        image_urls: video ? '' : [...new Set(images)].join(' | ')
    };
}
"""

READ_TABLES = """
root => [...root.querySelectorAll('table')]
    .filter(table => table.getClientRects().length)
    .map(table => [...table.rows].map(row =>
        [...row.cells].map(cell => cell.innerText.trim())
    ))
"""

NEW_RESULTS = r"""
known => {
    const oldIds = new Set(known);
    return [...document.body.innerText.matchAll(/Library ID:\s*(\d+)/g)]
        .some(match => !oldIds.has(match[1]));
}
"""


def ad_cards(scope):
    return scope.locator(CARD_SELECTOR).filter(has_text=re.compile(r"Library ID:\s*\d+"))


def read_cards(scope):
    return ad_cards(scope).evaluate_all(f"cards => cards.map({READ_CREATIVE})")


def visible_ids(page):
    return {
        match[1]
        for label in page.get_by_text(AD_ID).all_text_contents()
        if (match := AD_ID.fullmatch(label.strip()))
    }


def load_results(page, timeout, retries=20, expected_ads=0, max_ads=0, snapshot=None):
    known = visible_ids(page)
    if snapshot:
        snapshot()
    stalled = 0
    end_checks = 0

    while True:
        check_cancel()
        event("loading", loaded=len(known))
        if max_ads and len(known) >= max_ads:
            return f"Reached the {max_ads}-ad limit; exporting at most {max_ads} unique ads."
        buttons = page.get_by_role("button", name=MORE_RESULTS).filter(visible=True)
        action_finished = False
        try:
            if buttons.count() and buttons.last.is_enabled():
                buttons.last.click(timeout=timeout)
            else:
                page.keyboard.press("End")
            action_finished = True
            page.wait_for_function(
                NEW_RESULTS,
                arg=list(known),
                timeout=timeout,
            )
        except TimeoutError as exc:
            print(f"Waiting for more ads ({len(known)} IDs loaded): "
                  f"{str(exc).splitlines()[0]}", flush=True)

        current = visible_ids(page)
        if known - current:
            raise Error("Meta replaced earlier cards; full loading cannot be confirmed.")
        if not current - known:
            stalled += 1
            # A failed click does not count as confirmation of the end.
            at_end = action_finished and not buttons.count()
            end_checks = end_checks + 1 if at_end else 0
            if end_checks >= 3 and len(known) >= expected_ads:
                return (f"Reached the end: {len(known)} Library IDs loaded; "
                        "no load button or new IDs in three consecutive waits.")
            if stalled >= retries:
                raise Error(
                    f"Loading stalled at {len(known)} Library IDs after {retries} retries. "
                    "Full loading was not confirmed; increase --load-timeout or --load-retries."
                )
            print(f"Retrying loading {stalled}/{retries} "
                  f"({len(known)} IDs loaded)", flush=True)
            page.mouse.wheel(0, 1500)
            continue

        known = current
        if snapshot:
            snapshot()
        stalled = 0
        end_checks = 0
        print(f"Loaded {len(known)} Library IDs", flush=True)


def clean_creative(raw):
    row = dict(raw)
    destination = urlsplit(raw["destination_url"])
    if destination.hostname == "l.facebook.com":
        row["destination_url"] = parse_qs(destination.query).get("u", [""])[0]

    facebook_page = urlsplit(raw["facebook_page_url"])
    path_id = facebook_page.path.strip("/").split("/")[-1]
    row["facebook_page_id"] = (
        path_id if path_id.isdecimal()
        else parse_qs(facebook_page.query).get("id", [""])[0]
    )
    row["asset_url"] = row["player_source"] or row["image_urls"]
    return row


def make_ad(raw):
    text = raw["text"]
    match = re.search(r"^Library ID:\s*(\d+)$", text, re.M)
    if not match or not re.search(r"^Active$", text, re.M):
        return None

    row = clean_creative(raw)
    row.update(
        library_id=match[1],
        active_status="Active",
        version_count="",
        details_status="Listing only",
        transparency_status="Unverified",
    )

    started = re.search(r"Started running on ([^\n]+)", text)
    if started:
        raw_date = started[1].strip()
        row["started_running_on"] = raw_date
        for pattern in ("%d %b %Y", "%d %B %Y", "%b %d, %Y", "%B %d, %Y"):
            try:
                start_date = datetime.strptime(raw_date, pattern).date()
            except ValueError:
                continue
            row["started_running_on"] = start_date.isoformat()
            row["active_days"] = (date.today() - start_date).days
            break
    return row


def close_dialog(dialog):
    if dialog.count() and dialog.is_visible():
        dialog.get_by_role("button", name="Close", exact=True).first.click()


def version_row(ad, creative, index, total, status):
    return dict(
        creative,
        library_id=ad["library_id"],
        version=index,
        version_count=total,
        asset_status=status,
    )


def thumbnail_images(panel, counter):
    depth = counter.first.evaluate("""
    element => {
        for (let node = element.parentElement, depth = 1;
             node && depth <= 8; node = node.parentElement, depth++) {
            if (node.querySelectorAll('img[alt=""]').length >= 2) return depth;
        }
        return 0;
    }
    """)
    if depth:
        ancestor = counter.first.locator("xpath=" + "/".join([".."] * depth))
        return ancestor.locator('img[alt=""]')
    return panel.locator('img[alt=""]')


def read_versions(panel, ad, versions, timeout):
    counter = panel.get_by_text(VERSION_COUNT)
    total = 1
    if counter.count():
        match = VERSION_COUNT.fullmatch(counter.first.inner_text().strip())
        if match:
            total = int(match[2])
    elif panel.get_by_text("This ad has multiple versions", exact=True).count():
        total = ""

    ad["version_count"] = total
    first_player = ""
    for index in range(1, (total or 1) + 1):
        check_cancel()
        if index > 1:
            thumbnails = thumbnail_images(panel, counter)
            if thumbnails.count() < index:
                next_button = panel.locator('[aria-label="Next items"]')
                if next_button.count() and next_button.is_enabled():
                    next_button.click(timeout=timeout)
                thumbnails = thumbnail_images(panel, counter)
                if thumbnails.count() < index:
                    break

            thumbnails.nth(index - 1).click(timeout=timeout)
            panel.get_by_text(f"{index} of {total}", exact=True).wait_for(
                timeout=timeout
            )

        creative = clean_creative(panel.evaluate(READ_CREATIVE))
        if creative["image_urls"]:
            status = "Image URL"
        elif creative["player_source"]:
            status = "Player URL observed"
        else:
            status = "Preview only" if creative["preview_asset_url"] else "Unavailable"

        if index == 1:
            first_player = creative["player_source"]
            for field in (
                "primary_text", "headline", "description", "asset_url",
                "preview_asset_url", "destination_url",
            ):
                if creative[field]:
                    ad[field] = creative[field]
        elif first_player and creative["player_source"] == first_player:
            status = "Same player URL as version 1. Distinct MP4 unverified."
            creative["asset_url"] = ""

        versions[ad["library_id"], index] = version_row(
            ad, creative, index, total, status
        )


def read_locations(panel, ad, locations, timeout):
    control = panel.get_by_text("Transparency by location", exact=True)
    if not control.count():
        ad["transparency_status"] = "Location transparency unavailable"
        return

    control.first.click(timeout=timeout)
    found = False

    for region, tab_name in (("EU", "European Union"), ("UK", "United Kingdom")):
        tab = panel.get_by_role("tab", name=tab_name, exact=True)
        if not tab.count():
            continue

        tab.click(timeout=timeout)
        visible_panels = panel.get_by_role("tabpanel").filter(visible=True)
        scope = visible_panels.first if visible_panels.count() else panel
        try:
            scope.locator("table").filter(visible=True).first.wait_for(
                timeout=min(3000, timeout)
            )
        except TimeoutError:
            continue

        for table in scope.evaluate(READ_TABLES):
            header_index = next(
                (i for i, row in enumerate(table) if "Location" in row), None
            )
            if header_index is None:
                continue

            headers = table[header_index]
            location_column = headers.index("Location")

            for row in table[header_index + 1:]:
                if len(row) <= location_column:
                    continue
                location = row[location_column].strip(" \u200b")
                if not location:
                    continue

                source_type = ""
                if "Location type" in headers:
                    if len(row) < 3:
                        continue
                    direction = row[-1].strip().casefold()
                    if direction not in {"included", "excluded"}:
                        continue
                    kind = "Targeted" if direction == "included" else "Excluded"
                    type_column = headers.index("Location type")
                    source_type = row[type_column] if type_column < len(row) else ""
                elif "Reach" in headers:
                    kind = "Delivered"
                else:
                    continue

                key = (ad["library_id"], region, kind, location, source_type)
                locations[key] = dict(
                    library_id=ad["library_id"],
                    facebook_page_name=ad["facebook_page_name"],
                    region=region,
                    location_type=kind,
                    location=location,
                    source_location_type=source_type,
                )
                found = True

    ad["transparency_status"] = (
        "Locations collected" if found else "No location rows available"
    )


def inspect_ad(page, button, ad, versions, locations, timeout):
    panel = page.get_by_role("dialog", name="Ad details", exact=True)
    problems = []

    try:
        button.click(timeout=timeout)
        panel.get_by_text(
            f'Library ID: {ad["library_id"]}', exact=True
        ).first.wait_for(timeout=timeout)

        try:
            read_versions(panel, ad, versions, timeout)
        except Error:
            problems.append("Version control did not finish loading")

        try:
            read_locations(panel, ad, locations, timeout)
        except Error:
            ad["transparency_status"] = "Location controls did not finish loading"

        total = ad["version_count"]
        count = sum(
            (ad["library_id"], index) in versions
            for index in range(1, (total or 1) + 1)
        )
        if total and count < total:
            problems.append(f"Read {count} of {total} versions")
        elif not total:
            problems.append("Meta marked multiple versions without a visible count")

        ad["details_status"] = " | ".join(problems) or "Read"
    except Error:
        ad["details_status"] = "Ad details unavailable"
    finally:
        close_dialog(panel)


def make_sheet(book, title, fields, rows):
    sheet = book.create_sheet(title)
    sheet.append([
        field.replace("_", " ").title().replace(" Id", " ID").replace(" Url", " URL")
        for field in fields
    ])
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor="17365D")
        cell.font = Font(color="FFFFFF", bold=True)

    wrap_fields = {"primary_text", "headline", "description", "value", "details_status"}
    first_rows = {}
    previous_id = None
    shaded = False

    for row in rows:
        library_id = row.get("library_id")
        if library_id != previous_id:
            shaded = not shaded
            previous_id = library_id

        values = [row.get(field, "") for field in fields]
        sheet.append([
            ILLEGAL_CHARACTERS_RE.sub("", value)[:32767]
            if isinstance(value, str) else value
            for value in values
        ])
        if library_id:
            first_rows.setdefault(library_id, sheet.max_row)

        for cell, field in zip(sheet[sheet.max_row], fields):
            if isinstance(cell.value, str):
                cell.data_type = "s"
                if (
                    ("url" in field or field == "player_source")
                    and cell.value.startswith(("https://", "http://"))
                    and " | " not in cell.value
                ):
                    cell.hyperlink = cell.value
                    cell.style = "Hyperlink"

            if shaded:
                cell.fill = PatternFill("solid", fgColor="EDF4FB")
            cell.alignment = Alignment(vertical="top", wrap_text=field in wrap_fields)

    sheet.freeze_panes = "B2" if title != "Run" else "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for index, field in enumerate(fields, 1):
        width = 48 if field in wrap_fields else max(22, len(field) + 2)
        sheet.column_dimensions[get_column_letter(index)].width = width
    return first_rows


def link_to(cell, sheet, row):
    cell.hyperlink = Hyperlink(
        ref=cell.coordinate,
        location=f"'{sheet.replace(chr(39), chr(39) * 2)}'!A{row}",
        display=str(cell.value),
    )
    cell.style = "Hyperlink"


def save_workbook(path, ads, versions, locations, run, archive):
    ads = {key: ad for country_ads in archive.countries.values() for key, ad in country_ads.items()}
    version_rows = dict(versions)
    for library_id, ad in ads.items():
        if library_id in archive.existing_ids:
            continue
        version_rows.setdefault(
            (library_id, 1),
            version_row(ad, ad, 1, ad["version_count"], "Listing preview"),
        )

    version_counts = Counter(row["library_id"] for row in version_rows.values())
    location_counts = Counter(row["library_id"] for row in locations.values())
    ads_rows = [
        dict(
            ad,
            versions_exported=version_counts[library_id],
            locations_exported=location_counts[library_id],
        )
        for library_id, ad in ads.items()
    ]
    run_rows = dict(
        run,
        library_ids_exported=len(ads),
        version_rows_exported=len(version_rows),
        location_rows_exported=len(locations),
    )

    book = archive.book
    managed = list(archive.titles.values())
    managed += ["Versions", "Locations", "Repeated Ads", "Run", "Country Index"]
    for title in managed:
        if title in book.sheetnames:
            book.remove(book[title])
    rows_by_id = {row["library_id"]: row for row in ads_rows}
    ad_index = {}
    country_indexes = {}
    for country, country_ads in archive.countries.items():
        title = archive.titles[country]
        index = make_sheet(book, title, AD_FIELDS, (rows_by_id[key] for key in country_ads))
        country_indexes[title] = index
        for key, row in index.items():
            ad_index.setdefault(key, (title, row))
    make_sheet(book, "Repeated Ads", ["library_id", "countries"], (
        {"library_id": key, "countries": ", ".join(sorted(countries))}
        for key, countries in sorted(archive.memberships.items()) if len(countries) > 1
    ))
    make_sheet(book, "Country Index", ["country", "sheet"], (
        {"country": country, "sheet": title} for country, title in archive.titles.items()
    ))
    book["Country Index"].sheet_state = "hidden"
    version_index = make_sheet(
        book, "Versions", VERSION_FIELDS,
        (version_rows[key] for key in sorted(version_rows)),
    )
    location_index = make_sheet(
        book, "Locations", LOCATION_FIELDS,
        (locations[key] for key in sorted(locations)),
    )
    make_sheet(
        book, "Run", ["metric", "value"],
        ({"metric": key, "value": value} for key, value in run_rows.items()),
    )

    for tab_name, row_index, count_field in (
        ("Versions", version_index, "versions_exported"),
        ("Locations", location_index, "locations_exported"),
    ):
        count_column = AD_FIELDS.index(count_field) + 1
        for title, index in country_indexes.items():
            for library_id, ad_row in index.items():
                if library_id in row_index:
                    cell = book[title].cell(ad_row, count_column)
                    link_to(cell, tab_name, row_index[library_id])
        for cells in book[tab_name].iter_rows(min_row=2, max_col=1):
            cell = cells[0]
            if cell.value in ad_index:
                title, ad_row = ad_index[cell.value]
                link_to(cell, title, ad_row)

    temporary = path.with_name(path.stem + ".tmp.xlsx")
    book.save(temporary)
    temporary.replace(path)



def initial_page_problem(url, text):
    path = urlsplit(url).path.lower()
    body = text.casefold()
    if any(part in path for part in ('/checkpoint', '/challenge', '/captcha')):
        return 'Meta requires a security check on this server. Headless scraping cannot continue.'
    if path.startswith('/login') or 'you must log in to continue' in body:
        return 'Meta requires login on this server. The public Ad Library page was not available.'
    if any(phrase in body for phrase in ("you're temporarily blocked", 'you?re temporarily blocked',
                                         'too many requests', 'temporarily restricted', 'access denied')):
        return 'Meta restricted this request. Retry later; increasing the ad limit will not resolve this.'
    return ''


def save_page_diagnostics(page, output, reason):
    """Capture only the public page view; never cookies or browser storage."""
    try:
        info = {'reason': str(reason), 'url': page.url, 'title': page.title(),
                'visible_text': page.locator('body').inner_text(timeout=3000)[:8000]}
        with ZipFile(output.parent / 'diagnostics.zip', 'w') as bundle:
            bundle.writestr('page.json', json.dumps(info, indent=2, ensure_ascii=False))
            try:
                bundle.writestr('page.png', page.screenshot(timeout=5000))
            except Error as exc:
                bundle.writestr('screenshot-error.txt', str(exc).splitlines()[0])
        print('Page diagnostics saved. Download the diagnostic ZIP to inspect what Meta returned.', flush=True)
    except Exception as exc:
        print(f'Could not capture page diagnostics: {str(exc).splitlines()[0]}', flush=True)


def open_library(page, url, output, wait_seconds=60, attempts=2):
    try:
        for attempt in range(1, attempts + 1):
            check_cancel()
            event('starting', message=f'Opening Meta Ad Library (attempt {attempt}/{attempts})')
            try:
                response = page.goto(url, wait_until='domcontentloaded', timeout=60000)
                if response and response.status >= 400:
                    print(f'Meta returned HTTP {response.status}; checking whether the page can still render results.', flush=True)
            except TimeoutError:
                print('Navigation is slow; checking whether the search page has rendered.', flush=True)
            deadline = time.monotonic() + wait_seconds
            next_log = 0
            while time.monotonic() < deadline:
                try:
                    check_cancel()
                    if visible_ids(page):
                        return
                    empty = page.get_by_text(re.compile(r'^(No ads found|No results found|0 results)([ .]|$)', re.I))
                    if empty.count() and empty.first.is_visible():
                        return
                    try:
                        text = page.locator('body').inner_text(timeout=1500)
                    except TimeoutError:
                        text = ''
                    problem = initial_page_problem(page.url, text)
                    if problem:
                        raise Error(problem)
                    # Prefer declining optional cookies; do not accept tracking on the user's behalf.
                    for label in ('Decline optional cookies', 'Only allow essential cookies', 'Reject optional cookies'):
                        button = page.get_by_role('button', name=label, exact=True)
                        if button.count() and button.first.is_visible():
                            try:
                                button.first.click(timeout=2000)
                                print('Dismissed optional-cookie prompt.', flush=True)
                            except TimeoutError:
                                pass
                            break
                    if time.monotonic() >= next_log:
                        print(f'Waiting for initial ad cards (attempt {attempt}/{attempts}); page: {page.title()}', flush=True)
                        next_log = time.monotonic() + 10
                    page.wait_for_timeout(1000)
                except Error as exc:
                    if 'Execution context was destroyed' not in str(exc) and 'Cannot find context' not in str(exc):
                        raise
                    # Facebook can navigate again while its initial application loads.
                    page.wait_for_timeout(500)
            print('No ad cards yet; retrying the search page.' if attempt < attempts else
                  'No ad cards appeared after the initial loading retries.', flush=True)
        text = page.locator('body').inner_text(timeout=3000)
        excerpt = ' '.join(text.split())[:500]
        raise Error('Meta did not render readable ad cards or a confirmed empty result. '
                    'This may be a consent/login page, server access restriction, or changed layout. '
                    f'Page text: {excerpt or "(empty page)"}. Download diagnostics for the actual page.')
    except Exception as exc:
        save_page_diagnostics(page, output, exc)
        raise

def main():
    global EVENTS, CANCEL_FILE
    parser = argparse.ArgumentParser(description="Export Meta Ad Library to Excel.")
    parser.add_argument("--url", default=LIBRARY_URL, help="Facebook Library link; determines country")
    parser.add_argument("--query", help="Optional override for the link's search query")
    parser.add_argument("--max-ads", type=int, default=None,
                        help="IDs to consider, including repeats; 0 = unlimited. Default: DE 260, FR 270, others unlimited")
    parser.add_argument("--output", type=Path, default=Path("topinfoz_ads.xlsx"))
    parser.add_argument("--timeout", type=int, default=12, help="Seconds per UI action")
    parser.add_argument("--load-timeout", type=int, default=60,
                        help="Seconds per loading click/wait (default: 60)")
    parser.add_argument("--load-retries", type=int, default=20,
                        help="Consecutive attempts without new IDs (default: 20)")
    parser.add_argument("--expected-ads", type=int, default=0,
                        help="Minimum listing IDs required before accepting the end")
    parser.add_argument("--headless", action="store_true")
    parser.add_argument("--events", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--cancel-file", type=Path, help=argparse.SUPPRESS)
    args = parser.parse_args()
    EVENTS, CANCEL_FILE = args.events, args.cancel_file

    try:
        url, code, country = source_settings(args.url, args.query)
    except ValueError as exc:
        parser.error(str(exc))
    if args.max_ads is None:
        args.max_ads = COUNTRY_LIMITS.get(code, 0)
    if (min(args.timeout, args.load_timeout) < 1 or args.load_retries < 3
            or args.max_ads < 0 or args.expected_ads < 0
            or (args.max_ads and args.expected_ads > args.max_ads)
            or args.output.suffix.lower() != ".xlsx"):
        parser.error("Use positive timeouts, at least 3 retries, nonnegative counts, "
                     "expected ads no greater than the limit, and an .xlsx output.")

    timeout = args.timeout * 1000
    args.output.parent.mkdir(parents=True, exist_ok=True)
    try:
        archive = Archive(args.output, country, args.max_ads)
    except (OSError, ValueError) as exc:
        parser.exit(1, f"Cannot read existing workbook: {exc}\n")
    if args.output.exists():
        backup = args.output.with_name(
            f"{args.output.stem}.backup-{datetime.now():%Y%m%d-%H%M%S-%f}.xlsx")
        shutil.copy2(args.output, backup)
        print(f"Backup: {backup}", flush=True)
    ads = archive.countries[country]
    versions, locations = archive.versions, archive.locations
    run = {"source": url, "country": country, "max_ads": args.max_ads,
           "pagination": "Not started", "run_state": "In progress"}
    print(f"Country: {country} | Limit: {args.max_ads or 'unlimited'} | "
          f"Previously saved IDs: {len(archive.existing_ids)}", flush=True)

    try:
        check_cancel()
        event("starting", message="Starting headless browser")
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=args.headless)
            page = browser.new_page(
                locale="en-GB", viewport={"width": 1440, "height": 1000}
            )
            page.set_default_timeout(timeout)
            open_library(page, url, args.output)

            headings = page.get_by_role("heading", name=re.compile("results", re.I))
            run["meta_displayed_count"] = (
                headings.first.inner_text() if headings.count() else ""
            )
            def capture_listing():
                for raw in read_cards(page):
                    ad = make_ad(raw)
                    if ad:
                        archive.register(ad)

            run["pagination"] = (
                load_results(page, args.load_timeout * 1000,
                             args.load_retries, args.expected_ads, args.max_ads, capture_listing)
                if visible_ids(page) else "No results"
            )
            print(run["pagination"], flush=True)

            listing = read_cards(page)
            if visible_ids(page) and not listing:
                raise Error("Ads are visible but card selectors did not match. Meta's layout may have changed.")
            run["loaded_result_cards"] = len(listing)
            listing_ads = [make_ad(raw) for raw in listing]
            for ad in listing_ads:
                if ad:
                    archive.register(ad)

            summary = page.get_by_role("dialog", name="Summary data", exact=True)
            work = []
            queued = set()
            for index, ad in enumerate(listing_ads):
                if (ad and ad['library_id'] in archive.seen
                        and ad['library_id'] not in archive.existing_ids
                        and ad['library_id'] not in queued):
                    work.append((index, ad))
                    queued.add(ad['library_id'])
            detail_started = time.monotonic()
            for position, (index, listing_ad) in enumerate(work):
                check_cancel()
                event("scraping", completed=position, total=len(work),
                      ads=len(ads), elapsed=time.monotonic() - detail_started)
                if not listing_ad:
                    continue

                library_id = listing_ad["library_id"]
                if (library_id not in archive.seen or library_id in archive.existing_ids
                        or library_id in archive.attempted):
                    continue
                ad = ads[library_id]
                archive.attempted.add(library_id)
                card = ad_cards(page).nth(index)
                group_button = card.get_by_role(
                    "button", name="See summary details", exact=True
                )

                if group_button.count():
                    try:
                        group_button.click(timeout=timeout)
                        summary.get_by_text(
                            f"Library ID: {library_id}", exact=True
                        ).first.wait_for(timeout=timeout)

                        for member_index, raw_member in enumerate(read_cards(summary)):
                            child = make_ad(raw_member)
                            if not child:
                                continue

                            child = archive.register(child)
                            if child is None:
                                continue
                            if child["details_status"] == "Listing only":
                                archive.attempted.add(child["library_id"])
                                button = ad_cards(summary).nth(member_index).get_by_role(
                                    "button", name="See ad details", exact=True
                                )
                                inspect_ad(page, button, child, versions, locations, timeout)
                    except Error:
                        if ad["details_status"] == "Listing only":
                            ad["details_status"] = "Grouped results unavailable"
                    finally:
                        close_dialog(summary)
                elif ad["details_status"] == "Listing only":
                    button = card.get_by_role("button", name="See ad details", exact=True)
                    inspect_ad(page, button, ad, versions, locations, timeout)

                if (index + 1) % 25 == 0:
                    print(f"Processed {index + 1}/{len(listing)} cards", flush=True)
                    save_workbook(args.output, ads, versions, locations, run, archive)

                print(f"Ad {library_id}: {ad['details_status']}", flush=True)

            event("scraping", completed=len(work), total=len(work),
                  ads=len(ads), elapsed=time.monotonic() - detail_started)
            run["run_state"] = "Finished processing loaded results"
    except (Exception, KeyboardInterrupt) as exc:
        lines = str(exc).splitlines()
        run["run_state"] = "Interrupted: " + (lines[0] if lines else type(exc).__name__)
        event("cancelled" if isinstance(exc, Cancelled) else "error", message=run["run_state"])
    finally:
        event("saving", message="Saving Excel workbook")
        run["country_ids_considered"] = len(archive.seen)
        run["new_unique_ads"] = len(archive.seen - archive.existing_ids)
        run["previously_saved_ids_skipped"] = len(archive.seen & archive.existing_ids)
        save_workbook(args.output, ads, versions, locations, run, archive)
        print(f"{args.output.resolve()} | {len(ads)} ads | {run['run_state']}")
        warnings = sum(ad.get("details_status") != "Read" for ad in ads.values())
        event("finished", success=run["run_state"] == "Finished processing loaded results",
              path=str(args.output.resolve()), ads=len(ads), warnings=warnings,
              message=run["run_state"])
    return 0 if run["run_state"] == "Finished processing loaded results" else 1


if __name__ == "__main__":
    raise SystemExit(main())
