import contextlib
import io
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import meta_ads_export as exporter
from openpyxl import load_workbook
from progress import estimate


class ExportTests(unittest.TestCase):
    def test_initial_page_restrictions(self):
        self.assertIn('login', exporter.initial_page_problem('https://www.facebook.com/login/', ''))
        self.assertIn('security check', exporter.initial_page_problem('https://www.facebook.com/checkpoint/', ''))
        self.assertIn('restricted', exporter.initial_page_problem(exporter.LIBRARY_URL, "You're temporarily blocked"))
        self.assertEqual(exporter.initial_page_problem(exporter.LIBRARY_URL, 'Log in  Ad Library'), '')

    def test_initial_page_retries_and_records_diagnostics(self):
        from unittest.mock import MagicMock
        page = MagicMock()
        page.goto.return_value.status = 200
        page.locator.return_value.inner_text.return_value = 'Unexpected landing page'
        with patch.object(exporter, 'check_cancel'), patch.object(exporter, 'save_page_diagnostics') as diagnostic:
            with self.assertRaisesRegex(exporter.Error, 'did not render'):
                exporter.open_library(page, exporter.LIBRARY_URL, Path('ads.xlsx'), wait_seconds=0)
            self.assertEqual(page.goto.call_count, 2)
            diagnostic.assert_called_once()

    def test_initial_navigation_can_recover_after_http_403(self):
        from unittest.mock import MagicMock
        page = MagicMock()
        page.goto.return_value.status = 403
        with patch.object(exporter, 'check_cancel'), patch.object(exporter, 'visible_ids',
                side_effect=[exporter.Error('Execution context was destroyed'), {'123'}]),                 patch.object(exporter, 'save_page_diagnostics') as diagnostic:
            exporter.open_library(page, exporter.LIBRARY_URL, Path('ads.xlsx'))
            page.wait_for_timeout.assert_called_with(500)
            diagnostic.assert_not_called()

    def test_url_validation(self):
        self.assertEqual(exporter.source_settings(exporter.LIBRARY_URL)[1], 'DE')
        self.assertEqual(exporter.source_settings('https://www.facebook.com/ads/library/?country=ALL')[2], 'All countries')
        for url in ['https://example.com/ads/library/?country=DE', 'file:///ads/library/?country=DE', 'https://www.facebook.com/ads/library/?country=INVALID']:
            with self.assertRaises(ValueError):
                exporter.source_settings(url)

    def test_estimate_requires_samples(self):
        self.assertIsNone(estimate(0, 10, 20))
        self.assertIsNone(estimate(2, 10, 20))
        self.assertEqual(estimate(5, 10, 20), 20)

    def test_failure_and_cancel_save_valid_partial_workbook(self):
        for cancel in (False, True):
            with self.subTest(cancel=cancel), tempfile.TemporaryDirectory() as directory:
                path = Path(directory) / 'result.xlsx'
                flag = Path(directory) / 'stop'
                if cancel:
                    flag.touch()
                args = ['exporter', '--url', exporter.LIBRARY_URL, '--max-ads', '0', '--events', '--headless', '--output', str(path), '--cancel-file', str(flag)]
                stream = io.StringIO()
                with patch('sys.argv', args), patch.object(exporter, 'sync_playwright', side_effect=RuntimeError('Browser unavailable')), contextlib.redirect_stdout(stream):
                    self.assertEqual(exporter.main(), 1)
                book = load_workbook(path)
                self.assertIn('Germany Ads', book.sheetnames)
                self.assertIn('Run', book.sheetnames)
                book.close()
                events = [json.loads(line[7:]) for line in stream.getvalue().splitlines() if line.startswith('@event ')]
                self.assertIn('cancelled' if cancel else 'error', [event['phase'] for event in events])
                self.assertFalse(events[-1]['success'])
                self.assertEqual(events[-1]['phase'], 'finished')


if __name__ == '__main__':
    unittest.main()
