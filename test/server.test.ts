import test from 'node:test';
import assert from 'node:assert/strict';
import { createApp } from '../src/server.js';
import { demoFiles, demoPlan, wordDemoFiles } from '../scripts/generate-demo.js';

test('本地服务完成导入、核对、导出，并拒绝未授权与跨站请求', async () => {
  const app = createApp(); const url = await app.listen();
  try {
    assert.equal((await fetch(url)).status, 200);
    assert.equal((await fetch(`${url}/app.js`)).headers.get('content-type'), 'application/javascript; charset=utf-8');
    assert.equal((await fetch(`${url}/api/demo`)).status, 403);
    const { token } = await (await fetch(`${url}/api/session`)).json() as { token: string };
    const headers = { 'x-session-token': token, 'Content-Type': 'application/json' };
    assert.equal((await fetch(`${url}/api/demo`, { headers: { ...headers, Origin: 'https://unrelated.invalid' } })).status, 403);
    assert.equal((await fetch(`${url}/package.json`)).status, 404);
    const files = (await demoFiles()).map((f) => ({ name: f.name, data: f.bytes.toString('base64') }));
    const wait = async (id: string) => {
      for (let i = 0; i < 300; i++) {
        const job = await (await fetch(`${url}/api/jobs/${id}`, { headers })).json() as { status: string; result: any; error?: string };
        if (job.status === 'completed') return job.result;
        if (job.status !== 'running') throw new Error(job.error ?? job.status);
        await new Promise((resolve) => setTimeout(resolve, 30));
      }
      throw new Error('服务任务未按时结束。');
    };
    const inspect = await (await fetch(`${url}/api/inspect`, { method: 'POST', headers, body: JSON.stringify({ files }) })).json() as { id: string };
    const batch = await wait(inspect.id);
    assert.equal(batch.batch.documents.length, 2);
    const started = await (await fetch(`${url}/api/check`, { method: 'POST', headers, body: JSON.stringify({ batchId: batch.batchId, plan: demoPlan }) })).json() as { id: string };
    const report = await wait(started.id);
    assert.equal(report.summary.different, 4);
    const exported = await fetch(`${url}/api/jobs/${started.id}/report?format=html`, { headers });
    assert.equal(exported.status, 200); assert.match(exported.headers.get('content-disposition')!, /attachment/);
    assert.match(await exported.text(), /发现差异/);
    const incomplete = await fetch(`${url}/api/plan/validate`, { method: 'POST', headers, body: JSON.stringify({ plan: { version: 99 } }) });
    assert.equal(incomplete.status, 400);
    const task = await (await fetch(`${url}/api/inspect`, { method: 'POST', headers, body: JSON.stringify({ files }) })).json() as { id: string };
    const cancelled = await (await fetch(`${url}/api/jobs/${task.id}/cancel`, { method: 'POST', headers, body: '{}' })).json() as { status: string };
    assert.equal(cancelled.status, 'cancelled');
    assert.equal((await fetch(`${url}/api/jobs/${task.id}/report`, { headers })).status, 400);
    const wordFiles = (await wordDemoFiles()).map((f) => ({ name: f.name, data: f.bytes.toString('base64') }));
    const wordStarted = await (await fetch(`${url}/api/word/inspect`, { method: 'POST', headers, body: JSON.stringify({ files: wordFiles }) })).json() as { id: string };
    const wordBatch = await wait(wordStarted.id);
    assert.equal(wordBatch.batch.documents.length, 2);
    const selections = wordBatch.batch.documents.map((doc: { id: string }, index: number) => ({ documentId: doc.id,
      category: 'softwareVersion', segmentId: 'body:t:1:r:3:c:2:0', value: index ? '2.0' : '1.0', subject: '演示软件', stage: index ? '后期' : '前期' }));
    const compared = await (await fetch(`${url}/api/word/compare`, { method: 'POST', headers, body: JSON.stringify({ batchId: wordBatch.batchId, selections }) })).json() as { id: string };
    const wordReport = await wait(compared.id);
    assert.equal(wordReport.judgment, 'human-review');
    assert.equal(wordReport.differences.length, 1);
    const wordExport = await fetch(`${url}/api/jobs/${compared.id}/report?format=html`, { headers });
    assert.match(await wordExport.text(), /人工判断/);
    const wrongMode = await fetch(`${url}/api/word/compare`, { method: 'POST', headers, body: JSON.stringify({ batchId: batch.batchId, selections }) });
    assert.equal(wrongMode.status, 400);
  } finally { await app.close(); }
});
