"""Manual end-to-end check. Uses a small real Meta query; not part of unit tests."""
from pathlib import Path
import json
import os
import socket
import subprocess
import sys
import time
from urllib.request import urlopen
from zipfile import ZipFile

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright, expect

ROOT=Path(__file__).resolve().parent
ARTIFACTS=ROOT/'exports'/'web-verification'
ARTIFACTS.mkdir(parents=True,exist_ok=True)
with socket.socket() as sock:
    sock.bind(('127.0.0.1',0)); port=sock.getsockname()[1]
base=f'http://127.0.0.1:{port}'
log=(ARTIFACTS/'server.log').open('w')
server=subprocess.Popen([sys.executable,'web_app.py'],cwd=ROOT,
                        env=dict(os.environ,PORT=str(port),HOST='127.0.0.1',DATA_DIR=str(ARTIFACTS/'jobs'),APP_PASSWORD=''),
                        stdout=log,stderr=log)
try:
    for _ in range(60):
        try:
            with urlopen(base+'/health',timeout=1) as response:
                if response.status==200:break
        except OSError:time.sleep(.3)
    with sync_playwright() as playwright:
        browser=playwright.chromium.launch(headless=True)
        page=browser.new_page(viewport={'width':1440,'height':1100},accept_downloads=True)
        errors=[];page.on('pageerror',lambda exc:errors.append(str(exc)))
        page.goto(base)
        page.locator('#url').fill('https://www.facebook.com/ads/library/?active_status=active&ad_type=all&country=ALL&media_type=all&q=topinfoz.com&search_type=keyword_unordered')
        page.locator('#limit').fill('1')
        page.locator('#start').click()
        page.locator('#job-content').wait_for(state='visible')
        expect(page.locator("#xlsx")).to_have_attribute("aria-disabled", "false", timeout=180000)
        state=page.evaluate("async()=>{const jobs=await (await fetch('/api/jobs')).json();return await(await fetch('/api/jobs/'+jobs[0].id)).json()}")
        print('SCRAPE',json.dumps({k:state['scrape'].get(k) for k in ('status','ads','warnings','message')}),flush=True)
        assert state['country']=='All countries'
        assert state['scrape']['status']=='complete',state['scrape']
        assert state['scrape'].get('ads',0)>0
        with page.expect_download() as download:
            page.locator('#xlsx').click()
        xlsx=ARTIFACTS/'all_ads.xlsx';download.value.save_as(xlsx)
        book=load_workbook(xlsx,read_only=True)
        assert 'All countries Ads' in book.sheetnames
        print('WORKBOOK', {s.title:s.max_row-1 for s in book},flush=True);book.close()
        with page.expect_download(timeout=180000) as download:
            page.locator('#assets').click()
        zip_path=ARTIFACTS/'all_assets.zip';download.value.save_as(zip_path)
        with ZipFile(zip_path) as archive:
            assert archive.testzip() is None
            summary=json.loads(archive.read('summary.json'))
            print('ZIP',json.dumps(summary),flush=True)
            assert summary['downloaded']>0,summary
        page.screenshot(path=str(ARTIFACTS/'desktop.png'),full_page=True)
        page.set_viewport_size({'width':390,'height':844})
        page.screenshot(path=str(ARTIFACTS/'mobile.png'),full_page=True)
        assert page.evaluate('document.documentElement.scrollWidth <= window.innerWidth'), 'Mobile overflow'
        page.reload()
        page.locator('#job-content').wait_for(state='visible')
        expect(page.locator("#zip-link")).to_be_visible()
        assert not errors,errors
        print('PASS: worldwide scrape, XLSX, ZIP download, refresh persistence, responsive layout and JS errors',flush=True)
        browser.close()
finally:
    server.terminate()
    try:server.wait(timeout=10)
    except subprocess.TimeoutExpired:server.kill()
    log.close()
