import { createServer, type IncomingMessage, type ServerResponse } from 'node:http';
import { randomBytes, randomUUID } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import { pathToFileURL } from 'node:url';
import { Worker } from 'node:worker_threads';
import { reportHtml, wordReportHtml } from './report.js';
import { documentLabels, validatePlan } from './engine.js';
import type { Batch, Report, UploadedFile } from './types.js';
import type { WordBatch, WordReport } from './word-types.js';

interface Job {
  id: string;
  kind: 'inspect' | 'check' | 'word-inspect' | 'word-compare';
  status: 'running' | 'completed' | 'failed' | 'cancelled';
  message: string;
  createdAt: number;
  result?: unknown;
  error?: string;
  worker?: Worker;
}
const MAX_BODY = 70 * 1024 * 1024;
const MAX_BATCH = 48 * 1024 * 1024;
const DEMO_FILES = ['演示台账.xlsx', '演示说明-甲.docx', '演示说明-乙.docx'];
const WORD_DEMO_FILES = ['演示文档-前期.docx', '演示文档-后期.docx'];
const json = (response: ServerResponse, status: number, data: unknown) => {
  response.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' });
  response.end(JSON.stringify(data));
};
async function readJson(request: IncomingMessage): Promise<unknown> {
  if (!String(request.headers['content-type']).startsWith('application/json')) throw new Error('只接受 application/json 请求。');
  if (Number(request.headers['content-length'] ?? 0) > MAX_BODY) throw new Error('请求超过 70 MB 上限。');
  const chunks: Buffer[] = [];
  let total = 0;
  for await (const chunk of request) {
    total += chunk.length;
    if (total > MAX_BODY) throw new Error('请求超过 70 MB 上限。');
    chunks.push(chunk);
  }
  return JSON.parse(Buffer.concat(chunks).toString('utf8'));
}
function validateUploads(value: unknown, mode: 'ledger' | 'word' = 'ledger'): UploadedFile[] {
  if (!Array.isArray(value) || value.length < 2 || value.length > (mode === 'word' ? 30 : 31)) throw new Error(mode === 'word' ? '请选择 2–30 份 Word 文档。' : '请选择一份 Excel 和 1–30 份 Word。');
  let total = 0;
  for (const file of value) {
    if (!file || typeof file.name !== 'string' || typeof file.data !== 'string' || !(mode === 'word' ? /\.docx?$/i : /\.(xlsx|docx)$/i).test(file.name)) throw new Error('文件类型无效。');
    if (file.data.length > 16 * 1024 * 1024 || file.data.length % 4 !== 0 || !/^[A-Za-z0-9+/]*={0,2}$/.test(file.data)) throw new Error('文件编码无效或单文件超过 12 MB。');
    total += Buffer.byteLength(file.data, 'base64');
  }
  if (total > MAX_BATCH) throw new Error('本批文件总大小超过 48 MB。');
  return value;
}
function preview(batch: Batch) {
  return { ...batch, documents: batch.documents.map((doc) => ({ ...doc, labels: documentLabels(doc) })) };
}

export function createApp() {
  const token = randomBytes(32).toString('hex');
  const jobs = new Map<string, Job>();
  const batches = new Map<string, { kind: 'ledger'; value: Batch; usedAt: number } | { kind: 'word'; value: WordBatch; usedAt: number }>();
  let port = 0;
  const cleanup = () => {
    const cutoff = Date.now() - 30 * 60 * 1000;
    for (const [id, batch] of batches) if (batch.usedAt < cutoff) batches.delete(id);
    for (const [id, job] of jobs) if (job.status !== 'running' && job.createdAt < cutoff) jobs.delete(id);
  };
  const start = (kind: Job['kind'], data: object): Job => {
    cleanup();
    if ([...jobs.values()].some((j) => j.status === 'running')) throw new Error('已有任务正在运行，请完成或取消后重试。');
    if (jobs.size >= 12) {
      const oldest = [...jobs.values()].find((j) => j.status !== 'running');
      if (oldest) jobs.delete(oldest.id);
    }
    const job: Job = { id: randomUUID(), kind, status: 'running', message: '任务已开始…', createdAt: Date.now() };
    const workerUrl = import.meta.url.endsWith('.ts') ? new URL('../dist/src/worker.js', import.meta.url) : new URL('./worker.js', import.meta.url);
    const worker = new Worker(workerUrl, { workerData: { task: kind, ...data }, resourceLimits: { maxOldGenerationSizeMb: 384 } });
    const deadline = setTimeout(() => {
      if (job.status === 'running') { job.status = 'failed'; job.error = '任务超过 60 秒限制，已停止；没有生成通过结论。'; void worker.terminate(); }
    }, 60000);
    deadline.unref();
    job.worker = worker;
    jobs.set(job.id, job);
    worker.on('message', (message) => {
      if (job.status !== 'running') return;
      if (message.type === 'progress') job.message = String(message.message);
      if (message.type === 'error') { job.status = 'failed'; job.error = String(message.error); }
      if (message.type === 'result') {
        job.status = 'completed'; job.message = '任务完成。';
        if (kind === 'inspect' || kind === 'word-inspect') {
          const batchId = randomUUID();
          if (batches.size >= 3) batches.delete(batches.keys().next().value!);
          if (kind === 'inspect') {
            batches.set(batchId, { kind: 'ledger', value: message.result, usedAt: Date.now() });
            job.result = { batchId, batch: preview(message.result) };
          } else {
            batches.set(batchId, { kind: 'word', value: message.result, usedAt: Date.now() });
            job.result = { batchId, batch: message.result };
          }
        } else job.result = message.result;
      }
    });
    worker.on('error', (error) => { if (job.status === 'running') { job.status = 'failed'; job.error = error.message; } });
    worker.on('exit', (code) => {
      clearTimeout(deadline);
      job.worker = undefined;
      if (job.status === 'running') { job.status = 'failed'; job.error = `任务意外结束（${code}）。未生成通过结果。`; }
    });
    return job;
  };
  const server = createServer(async (request, response) => {
    response.setHeader('X-Content-Type-Options', 'nosniff');
    response.setHeader('Referrer-Policy', 'no-referrer');
    response.setHeader('Content-Security-Policy', "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; connect-src 'self'; object-src 'none'; base-uri 'none'; frame-ancestors 'none'");
    try {
      const host = request.headers.host;
      if (host !== `127.0.0.1:${port}` && host !== `localhost:${port}`) return json(response, 403, { error: '只接受本机地址访问。' });
      const origin = request.headers.origin;
      if ((origin && origin !== `http://${host}`) || request.headers['sec-fetch-site'] === 'cross-site') return json(response, 403, { error: '不接受跨站请求。' });
      const url = new URL(request.url ?? '/', `http://${host}`);
      if (url.pathname === '/api/session' && request.method === 'GET') return json(response, 200, { token });
      if (url.pathname.startsWith('/api/')) {
        if (request.headers['x-session-token'] !== token) return json(response, 403, { error: '会话已失效，请刷新页面。' });
        if (url.pathname === '/api/demo' && request.method === 'GET') {
          const demoRoot = import.meta.url.endsWith('.ts') ? new URL('../examples/', import.meta.url) : new URL('../../examples/', import.meta.url);
          const files = await Promise.all(DEMO_FILES.map(async (name) => ({ name, data: (await readFile(new URL(name, demoRoot))).toString('base64') })));
          const plan = JSON.parse(await readFile(new URL('demo-plan.json', demoRoot), 'utf8'));
          return json(response, 200, { files, plan, notice: '所有演示文件均为虚构数据。预置差异为：版本不一致、数量不一致、明细多项、明细缺项。' });
        }
        if (url.pathname === '/api/word/demo' && request.method === 'GET') {
          const demoRoot = import.meta.url.endsWith('.ts') ? new URL('../examples/', import.meta.url) : new URL('../../examples/', import.meta.url);
          const files = await Promise.all(WORD_DEMO_FILES.map(async (name) => ({ name, data: (await readFile(new URL(name, demoRoot))).toString('base64') })));
          return json(response, 200, { files, notice: '演示文档和阶段均为虚构数据；版本差异仅供人工判断。' });
        }
        if (url.pathname === '/api/inspect' && request.method === 'POST') {
          const payload = await readJson(request) as { files?: unknown };
          const files = validateUploads(payload.files);
          return json(response, 202, { id: start('inspect', { files }).id });
        }
        if (url.pathname === '/api/check' && request.method === 'POST') {
          const payload = await readJson(request) as { batchId: string; plan: unknown };
          const batch = batches.get(payload.batchId);
          if (!batch || batch.kind !== 'ledger' || Date.now() - batch.usedAt > 30 * 60 * 1000) throw new Error('所选资料已过期，请重新导入。');
          batch.usedAt = Date.now();
          const plan = validatePlan(payload.plan);
          return json(response, 202, { id: start('check', { batch: batch.value, plan }).id });
        }
        if (url.pathname === '/api/word/inspect' && request.method === 'POST') {
          const payload = await readJson(request) as { files?: unknown };
          const files = validateUploads(payload.files, 'word');
          return json(response, 202, { id: start('word-inspect', { files }).id });
        }
        if (url.pathname === '/api/word/compare' && request.method === 'POST') {
          const payload = await readJson(request) as { batchId: string; selections: unknown };
          const batch = batches.get(payload.batchId);
          if (!batch || batch.kind !== 'word' || Date.now() - batch.usedAt > 30 * 60 * 1000) throw new Error('所选资料已过期，请重新导入。');
          if (!Array.isArray(payload.selections) || payload.selections.length > 300) throw new Error('字段选择数量无效。');
          batch.usedAt = Date.now();
          return json(response, 202, { id: start('word-compare', { batch: batch.value, selections: payload.selections }).id });
        }
        if (url.pathname === '/api/plan/validate' && request.method === 'POST') {
          const payload = await readJson(request) as { plan: unknown };
          return json(response, 200, { plan: validatePlan(payload.plan) });
        }
        const match = url.pathname.match(/^\/api\/jobs\/([a-f0-9-]+)(\/cancel|\/report)?$/);
        if (match) {
          const job = jobs.get(match[1]);
          if (!job) return json(response, 404, { error: '任务已过期。' });
          if (match[2] === '/cancel' && request.method === 'POST') {
            if (job.status === 'running') { job.status = 'cancelled'; job.message = '任务已取消，未生成核对结论。'; await job.worker?.terminate(); }
            return json(response, 200, { status: job.status });
          }
          if (match[2] === '/report' && request.method === 'GET') {
            if (!['check', 'word-compare'].includes(job.kind) || job.status !== 'completed') throw new Error('该任务没有完整生成的报告。');
            const report = job.result as Report | WordReport;
            const html = url.searchParams.get('format') === 'html';
            response.writeHead(200, { 'Content-Type': html ? 'text/html; charset=utf-8' : 'application/json; charset=utf-8',
              'Content-Disposition': `attachment; filename="ledger-check.${html ? 'html' : 'json'}"`, 'Cache-Control': 'no-store' });
            response.end(html ? (job.kind === 'word-compare' ? wordReportHtml(report as WordReport) : reportHtml(report as Report)) : JSON.stringify(report, null, 2)); return;
          }
          if (!match[2] && request.method === 'GET') return json(response, 200, { id: job.id, status: job.status, message: job.message, error: job.error, result: job.result });
        }
        return json(response, 404, { error: '接口不存在。' });
      }
      if (request.method !== 'GET') return json(response, 405, { error: '不支持该请求方式。' });
      const assets: Record<string, [string, string]> = { '/': ['index.html', 'text/html'], '/app.js': ['app.js', 'application/javascript'], '/word-mode.js': ['word-mode.js', 'application/javascript'], '/style.css': ['style.css', 'text/css'] };
      const asset = assets[url.pathname];
      if (!asset) return json(response, 404, { error: '页面不存在。' });
      const body = await readFile(new URL(`../public/${asset[0]}`, import.meta.url));
      response.writeHead(200, { 'Content-Type': `${asset[1]}; charset=utf-8`, 'Cache-Control': 'no-store' }); response.end(body);
    } catch (error) { json(response, 400, { error: error instanceof Error ? error.message : String(error) }); }
  });
  return {
    server,
    async listen(requestedPort = 0) {
      await new Promise<void>((resolve, reject) => { server.once('error', reject); server.listen(requestedPort, '127.0.0.1', () => { server.off('error', reject); resolve(); }); });
      port = (server.address() as { port: number }).port;
      return `http://127.0.0.1:${port}`;
    },
    async close() {
      await Promise.all([...jobs.values()].map(async (j) => { if (j.worker) { j.status = 'cancelled'; await j.worker.terminate(); } }));
      server.closeAllConnections();
      await new Promise<void>((resolve) => server.close(() => resolve()));
    },
  };
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const app = createApp();
  const portText = process.env.LEDGER_CHECK_PORT;
  const requestedPort = portText === undefined ? 4173 : Number(portText);
  if (!Number.isInteger(requestedPort) || requestedPort < 0 || requestedPort > 65535) throw new Error('LEDGER_CHECK_PORT 必须是有效端口号。');
  const url = await app.listen(requestedPort);
  console.log(`台账核对工具已启动：${url}\n文件仅在本机内存中处理。按 Ctrl+C 退出。`);
  process.once('SIGINT', async () => { await app.close(); process.exit(0); });
  process.once('SIGTERM', async () => { await app.close(); process.exit(0); });
}
