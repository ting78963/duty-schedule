import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';
import vm from 'node:vm';

const html = fs.readFileSync(new URL('../templates/index.html', import.meta.url), 'utf8');
const script = html.split('<script>')[1].split('</script>')[0].split('// 初始化')[0];

test('batch button prevents a second request and resets on error and success', async () => {
  const button = { innerHTML: '批次產生 ZIP', disabled: false };
  const alerts = [];
  let calls = 0;
  let finish;
  const context = vm.createContext({
    document: { getElementById: () => button },
    alert: message => alerts.push(message),
    fetch: () => {
      calls++;
      return new Promise(resolve => { finish = resolve; });
    },
  });
  vm.runInContext(script, context);
  vm.runInContext(`
    currentYear = 115;
    currentMonth = 11;
    multiSel.add(2);
    buildDayData = rocDate => ({rocDate});
  `, context);

  const pending = context.generateBatch();
  assert.equal(button.disabled, true);
  assert.equal(button.textContent, '產生中…');
  await context.generateBatch();
  assert.equal(calls, 1);
  finish({ ok: false, status: 500, json: async () => ({
    date: '115年11月2日', detail: '假別名單不符',
  }) });
  await pending;
  assert.equal(button.disabled, false);
  assert.equal(button.innerHTML, '批次產生 ZIP');
  assert.deepEqual(alerts, ['批次產生失敗：115年11月2日：假別名單不符']);

  let downloads = 0;
  context.downloadGeneratedFile = async () => { downloads++; };
  const retry = context.generateBatch();
  finish({ ok: true });
  await retry;
  assert.equal(downloads, 1);
  assert.equal(button.disabled, false);
  assert.equal(button.innerHTML, '批次產生 ZIP');

  vm.runInContext('multiSel = new Set(Array.from({length: 11}, (_, i) => i + 1))', context);
  await context.generateBatch();
  assert.equal(calls, 2);
  assert.equal(button.disabled, false);
  assert.match(alerts.at(-1), /每批最多 10 天/);
});
