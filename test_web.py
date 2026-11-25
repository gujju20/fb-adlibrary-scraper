import io
import json
from pathlib import Path
import tempfile
import time
import unittest
from unittest.mock import patch
from zipfile import ZipFile
from openpyxl import Workbook
from starlette.testclient import TestClient

from asset_zip import build_zip, collect_assets, validate_media_url, AssetCancelled
from jobs import JobManager
from web_app import create_app


def fixture(path):
    book=Workbook(); book.remove(book.active)
    for country, rows in [('France', [('123', 'https://media.fbcdn.net/a.jpg'), ('124', '')]),
                          ('Germany', [('123', 'https://media.fbcdn.net/a.jpg'), ('456', 'https://media.fbcdn.net/b.mp4')])]:
        sheet=book.create_sheet(country+' Ads'); sheet.append(['Library ID','Asset URL'])
        for row in rows: sheet.append(row)
    sheet=book.create_sheet('Versions'); sheet.append(['Library ID','Version','Asset URL'])
    sheet.append(['123',1,'https://media.fbcdn.net/a.jpg'])
    sheet.append(['123',2,'https://media.fbcdn.net/c.jpg'])
    sheet.append(['999',1,'https://media.fbcdn.net/not-selected.jpg'])
    book.save(path); book.close()


def fake_download(url, path, **kwargs):
    if url.endswith('b.mp4'): raise OSError('Expired URL')
    path.write_bytes(b'\xff\xd8\xfftest')
    return '.jpg', path.stat().st_size


class AssetTests(unittest.TestCase):
    def test_all_countries_and_versions_are_collected_and_deduplicated(self):
        with tempfile.TemporaryDirectory() as directory:
            path=Path(directory)/'ads.xlsx'; fixture(path)
            assets,missing=collect_assets(path,'ALL')
            self.assertEqual(len(assets),3)
            self.assertEqual(len(missing),1)
            self.assertEqual({row['library_id'] for row in assets},{'123','456'})
            france,_=collect_assets(path,'France')
            self.assertEqual(len(france),2)

    def test_zip_contains_media_and_failure_report(self):
        with tempfile.TemporaryDirectory() as directory:
            path=Path(directory)/'ads.xlsx'; fixture(path)
            output=Path(directory)/'assets.zip'
            result=build_zip(path,output,downloader=fake_download)
            self.assertEqual(result['downloaded'],2)
            self.assertEqual(result['failed'],1)
            with ZipFile(output) as archive:
                self.assertIsNone(archive.testzip())
                self.assertIn('media/123_002.jpg',archive.namelist())
                report=archive.read('download_report.csv').decode('utf-8-sig')
                self.assertIn('Expired URL',report)
                self.assertIn('Missing',report)
                self.assertNotIn('not-selected',report)

    def test_cancel_generates_valid_partial_zip(self):
        with tempfile.TemporaryDirectory() as directory:
            path=Path(directory)/'ads.xlsx'; fixture(path)
            flag=Path(directory)/'cancel'; flag.touch()
            output=Path(directory)/'assets.zip'
            result=build_zip(path,output,cancel_file=flag,downloader=fake_download)
            self.assertTrue(result['cancelled'])
            self.assertEqual(result['downloaded'],0)
            with ZipFile(output) as archive:
                self.assertIn('Cancelled',archive.read('download_report.csv').decode())

    def test_public_media_hosts_only(self):
        for url in ['http://media.fbcdn.net/a','https://127.0.0.1/a','https://evil.test/a','https://fbcdn.net.evil.test/a','https://user@media.fbcdn.net/a']:
            with self.assertRaises(ValueError): validate_media_url(url)
        with patch('asset_zip.socket.getaddrinfo',return_value=[(2,1,6,'',('127.0.0.1',443))]):
            with self.assertRaises(ValueError):validate_media_url('https://media.fbcdn.net/a')


class WebTests(unittest.TestCase):
    def test_session_isolation_validation_and_downloads(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(JobManager,'_run'):
            app=create_app(directory)
            with TestClient(app) as client:
                self.assertEqual(client.get('/').status_code,200)
                self.assertEqual(client.post('/api/jobs',json={'url':'https://example.com'}).status_code,400)
                result=client.post('/api/jobs',json={'url':'https://www.facebook.com/ads/library/?country=ALL','limit':0})
                self.assertEqual(result.status_code,202)
                key=result.json()['id']; self.assertEqual(result.json()['country'],'All countries')
                manager=app.state.jobs
                folder=manager.directory(key)
                fixture(folder/'ads.xlsx')
                with manager.lock:
                    manager.jobs[key]['scrape'].update(status='complete',success=True)
                    manager._save(manager.jobs[key])
                response=client.get(f'/api/jobs/{key}/download/xlsx')
                self.assertEqual(response.status_code,200)
                self.assertIn('attachment',response.headers['content-disposition'])
                with ZipFile(io.BytesIO(response.content)) as archive:self.assertIsNone(archive.testzip())
                self.assertEqual(client.post(f'/api/jobs/{key}/assets',json={}).status_code,200)
                manager.jobs[key]['assets']['status']='complete'
                (folder/'assets.zip').write_bytes(b'zip test')
                response=client.get(f'/api/jobs/{key}/download/zip')
                self.assertEqual(response.headers['content-type'],'application/zip')
                self.assertEqual(response.content,b'zip test')
                self.assertEqual(client.post('/api/jobs',json={},headers={'Origin':'https://evil.test'}).status_code,403)
                client.cookies.clear()
                self.assertEqual(client.get(f'/api/jobs/{key}').status_code,404)
                self.assertEqual(client.get(f'/api/jobs/{key}/download/zip').status_code,404)
                self.assertEqual(client.get('/api/jobs').json(),[])

    def test_password_and_health(self):
        with tempfile.TemporaryDirectory() as directory, patch.dict('os.environ',{'APP_PASSWORD':'test-password'}):
            with TestClient(create_app(directory)) as client:
                self.assertEqual(client.get('/health').status_code,200)
                self.assertEqual(client.get('/').status_code,401)
                self.assertEqual(client.get('/',auth=('admin','test-password')).status_code,200)

    def test_restart_marks_queued_job_interrupted(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(JobManager,'_run'):
            manager=JobManager(directory)
            key=manager.create('owner','https://www.facebook.com/ads/library/?country=ALL','ALL','All countries',0)
            manager.close()
            restarted=JobManager(directory)
            self.assertEqual(restarted.get(key,'owner')['scrape']['status'],'interrupted')
            restarted.close()


if __name__=='__main__':unittest.main()
