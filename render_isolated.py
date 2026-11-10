"""Temporary Render WSGI entry point for the November batch acceptance test."""
import base64
import binascii
import hashlib
import hmac
import os
import re
import time
from types import SimpleNamespace

from flask import Request, abort, g, request

import app as duty

if any(os.environ.get(key) for key in ('SUPABASE_URL', 'SUPABASE_KEY', 'TEMPLATE_UPLOAD_TOKEN')):
    raise RuntimeError('Production configuration is forbidden in the isolated service')
token_digest = os.environ.get('TEST_TOKEN_SHA256', '')
if not re.fullmatch(r'[0-9a-f]{64}', token_digest):
    raise RuntimeError('A test token digest is required')

application = duty.app
application.config['MAX_CONTENT_LENGTH'] = 2 * 1024 * 1024
allowed_objects = {'source_excel/2026-11.xlsx', *duty.TEMPLATE_STORAGE_NAMES.values()}
fixture_prefix = 'https://fixture.invalid/storage/v1/object/authenticated/duty-templates/'


class FixtureRequest(Request):
    @property
    def json(self):
        return g.fixture_payload


application.request_class = FixtureRequest


class MemoryStorage:
    @staticmethod
    def get(url, **kwargs):
        if not url.startswith(fixture_prefix):
            raise RuntimeError('External reads are blocked')
        name = url.removeprefix(fixture_prefix).split('?', 1)[0]
        if name not in allowed_objects:
            raise RuntimeError('Fixture object is not allowed')
        time.sleep(g.fixture_delay)
        return SimpleNamespace(status_code=200, content=g.fixture_objects[name])


class FixtureRemarks:
    @staticmethod
    def get(key, default=None):
        return g.fixture_remarks.get(key, default)


duty.PERSISTENT_STAFF_REMARKS = FixtureRemarks()
duty.req = MemoryStorage()
duty.SUPABASE_URL = 'https://fixture.invalid'
duty.SUPABASE_KEY = 'isolated-memory-only'


@application.get('/healthz')
def health():
    return 'ok'


@application.before_request
def restrict_to_generation():
    if request.path == '/healthz' and request.method in ('GET', 'HEAD'):
        return
    if request.path not in ('/api/generate', '/api/generate-batch') or request.method != 'POST':
        abort(404)
    supplied = request.headers.get('Authorization', '').removeprefix('Bearer ')
    if not hmac.compare_digest(hashlib.sha256(supplied.encode()).hexdigest(), token_digest):
        abort(401)

    body = request.get_json()
    if not isinstance(body, dict) or not isinstance(body.get('objects'), dict):
        abort(400)
    objects = body['objects']
    payload = body.get('payload')
    dates = payload if request.path.endswith('-batch') else [payload]
    if (set(objects) != allowed_objects or not isinstance(dates, list)
            or not 1 <= len(dates) <= 30 or any(not isinstance(day, dict) for day in dates)):
        abort(400)
    delay = body.get('delay', 0)
    if type(delay) is not int or delay not in (0, 1):
        abort(400)
    try:
        if any(not isinstance(value, str) for value in objects.values()):
            abort(400)
        g.fixture_objects = {key: base64.b64decode(value, validate=True) for key, value in objects.items()}
    except (ValueError, binascii.Error):
        abort(400)
    remarks = body.get('remarks')
    if not isinstance(remarks, list) or len(remarks) > 63:
        abort(400)
    if any(not isinstance(row, list) or len(row) != 3 or type(row[0]) is not int
           or not isinstance(row[1], str) or not isinstance(row[2], str) for row in remarks):
        abort(400)
    g.fixture_remarks = {(code, name): note for code, name, note in remarks}
    g.fixture_payload = payload
    g.fixture_delay = delay


@application.after_request
def isolate_response(response):
    response.headers['Cache-Control'] = 'no-store'
    if 'fixture_payload' in g and os.name == 'posix':
        import resource
        peak = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
        application.logger.warning('isolated-test worker=%s status=%s peak_rss_kib=%s',
                                   os.getpid(), response.status_code, peak)
    return response
