"""The batch limit must reject oversize requests before any external reads."""
import io
import unittest
from unittest.mock import patch
import zipfile

import app


class BatchLimit(unittest.TestCase):
    def test_over_limit_is_rejected_before_generation(self):
        with patch.object(app, 'verified_generation_input') as verify:
            response = app.app.test_client().post(
                '/api/generate-batch', json=[{}] * (app.BATCH_DAY_LIMIT + 1))
        self.assertEqual(response.status_code, 400)
        self.assertIn(str(app.BATCH_DAY_LIMIT), response.json['error'])
        self.assertNotIn('Content-Disposition', response.headers)
        verify.assert_not_called()

    def test_limit_accepts_every_requested_day(self):
        days = [{'rocDate': f'115年11月{n}日'} for n in range(1, app.BATCH_DAY_LIMIT + 1)]
        with patch.object(app, 'verified_generation_input', side_effect=lambda day: day), \
             patch.object(app, 'generate_checked_word', side_effect=lambda day: io.BytesIO(b'checked')):
            response = app.app.test_client().post('/api/generate-batch', json=days)
        self.assertEqual(response.status_code, 200)
        with zipfile.ZipFile(io.BytesIO(response.data)) as archive:
            self.assertEqual(len(archive.namelist()), app.BATCH_DAY_LIMIT)
