"""Time one offline ZIP batch with explicitly supplied release files."""
import argparse
from collections import Counter
from contextlib import ExitStack
import ctypes
import gc
import hashlib
import io
import json
import os
from pathlib import Path
import sys
import time
from unittest.mock import Mock, patch
import zipfile

from docx import Document

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import app
import word_validation

KEYS = {
    'rocDate': 'roc_date', 'weekday': 'weekday', 'isHoliday': 'is_holiday',
    'hasSake': 'has_sake', 'sakeTimes': 'sake_times',
    'onDuty': 'on_duty', 'offDuty': 'off_duty',
}


def peak_rss_mib():
    if os.name == 'nt':
        # PROCESS_MEMORY_COUNTERS starts with two DWORDs, then SIZE_T fields.
        class Counters(ctypes.Structure):
            _fields_ = [
                ('size', ctypes.c_uint32), ('faults', ctypes.c_uint32),
                ('peak_working_set', ctypes.c_size_t), ('remaining', ctypes.c_size_t * 7),
            ]
        counters = Counters()
        counters.size = ctypes.sizeof(counters)
        if not ctypes.windll.psapi.GetProcessMemoryInfo(
                ctypes.c_void_p(-1), ctypes.byref(counters), counters.size):
            raise OSError('GetProcessMemoryInfo failed')
        return counters.peak_working_set / 2**20
    import resource
    scale = 2**20 if sys.platform == 'darwin' else 2**10
    return resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / scale


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('excel', type=Path)
    parser.add_argument('templates', type=Path, help='Directory of the five Storage-named templates')
    parser.add_argument('--storage-delay', type=float, default=0, help='Seconds per simulated read')
    parser.add_argument('--days', type=int, default=10, choices=range(1, 11))
    args = parser.parse_args()
    objects = {name: (args.templates / name).read_bytes()
               for name in app.TEMPLATE_STORAGE_NAMES.values()}
    excel = args.excel.read_bytes()
    objects['source_excel/2026-11.xlsx'] = excel
    parsed = app.parse_excel(excel)['monthlyData']
    days = [next(day for day in parsed.values() if day['roc_date'] == f'115年11月{n}日')
            for n in range(1, args.days + 1)]
    payload = [{key: day[field] for key, field in KEYS.items()} for day in days]
    elapsed = Counter()
    counts = Counter()
    per_day = {}

    def timed(name, function):
        def call(*args, **kwargs):
            start = time.perf_counter()
            try:
                return function(*args, **kwargs)
            finally:
                seconds = time.perf_counter() - start
                elapsed[name] += seconds
                counts[name] += 1
                if name in ('generate', 'validate'):
                    data = args[0] if name == 'generate' else args[1]
                    per_day.setdefault(data['rocDate'], {})[name] = round(seconds, 3)
        return call

    def read_storage(url, **kwargs):
        prefix = 'https://offline.invalid/storage/v1/object/authenticated/duty-templates/'
        assert url.startswith(prefix), url
        time.sleep(args.storage_delay)
        return Mock(status_code=200, content=objects[url.removeprefix(prefix).split('?')[0]])

    with ExitStack() as stack:
        stack.enter_context(patch.object(app, 'SUPABASE_URL', 'https://offline.invalid'))
        stack.enter_context(patch.object(app, 'SUPABASE_KEY', 'offline'))
        stack.enter_context(patch.object(app.req.sessions.Session, 'request',
                                         side_effect=AssertionError('Network forbidden')))
        stack.enter_context(patch.object(app.req, 'get', side_effect=timed('storage', read_storage)))
        for module, attr, label in (
            (app, 'parse_excel', 'parse'),
            (app, 'generate_word', 'generate'),
            (app, 'validate_generated_word', 'validate'),
            (word_validation, 'Document', 'validation_doc_open'),
            (word_validation, 'table_row_cells', 'validation_grid'),
            (gc, 'collect', 'gc'),
        ):
            stack.enter_context(patch.object(module, attr, timed(label, getattr(module, attr))))
        started = time.perf_counter()
        response = app.app.test_client().post('/api/generate-batch', json=payload)
        total = time.perf_counter() - started
        peak = peak_rss_mib()

    assert response.status_code == 200, response.get_json(silent=True)
    assert response.mimetype == 'application/zip'
    digest = hashlib.sha256()
    with zipfile.ZipFile(io.BytesIO(response.data)) as archive:
        names = archive.namelist()
        assert len(names) == len(set(names)) == args.days and archive.testzip() is None
        for name, data in zip(names, payload):
            assert name == f"勤務表_{data['rocDate']}.docx"
            content = archive.read(name)
            Document(io.BytesIO(content))
            with zipfile.ZipFile(io.BytesIO(content)) as doc:
                for part in sorted(doc.namelist()):
                    digest.update(part.encode())
                    digest.update(hashlib.sha256(doc.read(part)).digest())
    print(json.dumps({
        'http': response.status_code, 'days': len(names), 'wall_seconds': round(total, 3),
        'peak_rss_mib': round(peak, 2), 'docx_parts_sha256': digest.hexdigest(),
        'stage_seconds': {key: round(value, 3) for key, value in elapsed.items()},
        'calls': dict(counts), 'per_day_seconds': per_day,
    }, ensure_ascii=False))


if __name__ == '__main__':
    main()
