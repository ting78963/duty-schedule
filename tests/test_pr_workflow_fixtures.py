"""Execute the actual workflow Python blocks against disposable memory storage."""
import io
from contextlib import redirect_stdout
import os
from pathlib import Path
import re
import textwrap
import unittest
from unittest.mock import Mock, patch
from docx import Document
import requests
import app

ROOT = Path(__file__).resolve().parents[1]
WORKFLOWS = ('ascii-storage-key-check.yml', 'persistent-template-release-checks.yml')


class PRWorkflowFixtures(unittest.TestCase):
    def test_actual_workflow_blocks_without_network(self):
        for workflow in WORKFLOWS:
            with self.subTest(workflow=workflow):
                text = (ROOT / '.github/workflows' / workflow).read_text(encoding='utf-8')
                blocks = re.findall(r"python - <<'PY'\n(.*?)\n\s*PY(?:\n|$)", text, re.S)
                self.assertEqual(len(blocks), 2)
                objects, calls = {}, []
                def object_key(url):
                    self.assertTrue(url.startswith('https://isolated.invalid/storage/v1/object/'))
                    key = url.split('duty-templates/', 1)[1].split('?', 1)[0]
                    self.assertRegex(key, r'^qa_only_[a-f0-9]+\.docx$')
                    return key
                def post(url, **kwargs):
                    key = object_key(url)
                    self.assertNotIn(key, objects)
                    app.validate_template_bytes(kwargs['data'])
                    objects[key] = kwargs['data']; calls.append('upload')
                    return Mock(status_code=201)
                def get(url, **kwargs):
                    self.assertIn('/object/authenticated/', url)
                    key = object_key(url)
                    self.assertIn(key, objects)
                    calls.append('private-read')
                    return Mock(status_code=200, content=objects[key])
                def delete(url, **kwargs):
                    key = object_key(url)
                    self.assertIn(key, objects)
                    del objects[key]; calls.append('cleanup')
                    return Mock(status_code=200)
                # Block every Session request, including any accidentally
                # introduced HTTP verb not covered by the storage doubles.
                with patch.dict(os.environ, {'SUPABASE_URL':'https://isolated.invalid',
                                             'SUPABASE_KEY':'fake-offline-key'}), \
                     patch.object(requests.sessions.Session, 'request', side_effect=AssertionError('Network forbidden')), \
                     patch.object(requests, 'post', side_effect=post), \
                     patch.object(requests, 'get', side_effect=get), \
                     patch.object(requests, 'delete', side_effect=delete):
                    namespace = {'__name__': '__ci_fixture_test__'}
                    # Original job prints 'LIVE ... PASSED'. Do not emit that
                    # message for an offline run with storage doubles.
                    with redirect_stdout(io.StringIO()):
                        for block in blocks:
                            exec(compile(textwrap.dedent(block), workflow, 'exec'), namespace)
                    self.assertEqual(calls.count('upload'), 1)
                    self.assertGreaterEqual(calls.count('private-read'), 3)
                    self.assertEqual(calls.count('cleanup'), 1)
                    self.assertEqual(objects, {})
                    output = namespace['generated']
                    self.assertIn('中華民國115年11月2日【星期一】',
                                  '\n'.join(c.text for r in output.tables[0].rows for c in r.cells))
                    print('OFFLINE workflow passed:', workflow)

    def test_incomplete_fixtures_are_still_rejected(self):
        for missing in ('date', 'leave-rows'):
            doc = Document(); table = doc.add_table(rows=10, cols=35)
            table.cell(0,1).text = '中華民國115年10月1日【星期四】' if missing != 'date' else ''
            table.cell(4,30).text = '19'; table.cell(5,30).text = '29'
            for r, label in ((3,'輪   休'), (6,'休(補)假'), (8,'公假')):
                table.cell(r,34).text = label
            if missing == 'leave-rows':
                table.cell(8,34).text = ''; table.cell(9,34).text = '公假'
            b=io.BytesIO(); doc.save(b)
            with self.subTest(missing=missing), patch.object(app, 'template_bytes', return_value=b.getvalue()), \
                 self.assertRaises(ValueError):
                app.generate_word({'onDuty':[], 'offDuty':[], 'rocDate':'115年11月2日',
                                   'weekday':'一','isHoliday':False,'hasSake':False})
