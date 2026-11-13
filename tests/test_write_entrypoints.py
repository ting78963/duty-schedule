"""All persistence entry points must reject requests before touching services."""
import os
import unittest
from contextlib import ExitStack
from unittest.mock import patch
import app


class WriteEntryPointTests(unittest.TestCase):
    def test_all_writes_fail_closed_before_io(self):
        client = app.app.test_client()
        with ExitStack() as stack:
            effects = [stack.enter_context(patch.object(app, name)) for name in
                       ('sb_post', 'sb_patch', 'sb_delete', 'parse_excel',
                        'validate_template_bytes', 'template_bytes')]
            effects += [stack.enter_context(patch.object(app.req, method))
                        for method in ('get', 'post', 'patch', 'delete')]
            for configured, supplied in [('admin', ''), ('admin', 'wrong'),
                                         ('admin', '錯誤'), ('', 'admin')]:
                with patch.dict(os.environ, {'TEMPLATE_UPLOAD_TOKEN': configured}):
                    for method, path in [('POST', '/api/staff'),
                                         ('PUT', '/api/staff/1'),
                                         ('DELETE', '/api/staff/1'),
                                         ('POST', '/api/parse-excel'),
                                         ('POST', '/api/upload-template')]:
                        with self.subTest(path=path, configured=bool(configured), supplied=supplied):
                            response = client.open(path, method=method,
                                headers={'X-Template-Upload-Token': supplied})
                            self.assertEqual(response.status_code, 403)
                            self.assertIsInstance(response.json['error'], str)
            for effect in effects:
                effect.assert_not_called()

    def test_uploads_accept_unicode_admin_then_validate_file(self):
        client = app.app.test_client()
        with patch.dict(os.environ, {'TEMPLATE_UPLOAD_TOKEN': '管理密碼'}):
            for path in ('/api/parse-excel', '/api/upload-template'):
                response = client.post(path, headers={'X-Template-Upload-Token': '管理密碼'})
                self.assertEqual(response.status_code, 400)


if __name__ == '__main__':
    unittest.main()
