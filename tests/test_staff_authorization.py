"""Staff authorization tests using only an in-memory roster."""
import os
import unittest
from unittest.mock import patch
import app

class StaffAuthorizationTests(unittest.TestCase):
    def setUp(self):
        self.rows = {1: {'code': 1, 'name': '原姓名', 'type': '輪班'}}
        self.client = app.app.test_client()
        self.env = patch.dict(os.environ, {'TEMPLATE_UPLOAD_TOKEN': 'local-admin'})
        self.env.start()
        self.addCleanup(self.env.stop)
        def post(table, data):
            self.rows[data['code']] = dict(data)
            return [data]
        def update(table, key, code, data):
            self.rows[code].update(data)
            return [self.rows[code]]
        def delete(table, key, code):
            del self.rows[code]
            return 204
        self.mocks = []
        for name, fn in [('sb_post', post), ('sb_patch', update), ('sb_delete', delete)]:
            p = patch.object(app, name, side_effect=fn)
            self.mocks.append(p.start())
            self.addCleanup(p.stop)

    def test_missing_wrong_and_unconfigured_auth_never_write(self):
        for configured, supplied in [('local-admin', ''), ('local-admin', 'wrong'), ('', 'local-admin'), ('local-admin', '錯誤')]:
            with patch.dict(os.environ, {'TEMPLATE_UPLOAD_TOKEN': configured}):
                for method, url, data in [('POST', '/api/staff', {'code': 2, 'name': '新增'}), ('PUT', '/api/staff/1', {'name': '更改'}), ('DELETE', '/api/staff/1', None)]:
                    r = self.client.open(url, method=method, json=data, headers={'X-Template-Upload-Token': supplied})
                    self.assertEqual(r.status_code, 403)
                    self.assertIn('管理授權', r.json['error'])
        self.assertEqual(self.rows, {1: {'code': 1, 'name': '原姓名', 'type': '輪班'}})
        for mock in self.mocks:
            mock.assert_not_called()

    def test_existing_admin_can_add_update_delete(self):
        headers = {'X-Template-Upload-Token': 'local-admin'}
        self.assertEqual(self.client.post('/api/staff', json={'code': 2, 'name': '新增'}, headers=headers).status_code, 200)
        self.assertIn(2, self.rows)
        self.assertEqual(self.client.put('/api/staff/2', json={'name': '更新'}, headers=headers).status_code, 200)
        self.assertEqual(self.rows[2]['name'], '更新')
        self.assertEqual(self.client.delete('/api/staff/2', headers=headers).status_code, 200)
        self.assertNotIn(2, self.rows)

    def test_staff_query_still_needs_no_admin(self):
        with patch.object(app, 'SUPABASE_URL', 'https://local.invalid'), patch.object(app, 'sb_get', return_value=list(self.rows.values())):
            response = self.client.get('/api/staff')
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json[0]['name'], '原姓名')
