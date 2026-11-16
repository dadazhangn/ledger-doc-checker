import { parentPort, workerData } from 'node:worker_threads';
import { inspectFiles } from './parsers.js';
import { runCheck } from './engine.js';
import type { Batch, UploadedFile } from './types.js';

try {
  if (workerData.task === 'inspect') {
    parentPort!.postMessage({ type: 'progress', message: '正在读取所选 Excel 与 Word 的结构…' });
    const files = (workerData.files as UploadedFile[]).map((file) => ({ name: file.name, bytes: Buffer.from(file.data, 'base64') }));
    const batch = await inspectFiles(files);
    parentPort!.postMessage({ type: 'result', result: batch });
  } else if (workerData.task === 'check') {
    parentPort!.postMessage({ type: 'progress', message: '正在确认对应关系并逐项比较…' });
    const result = runCheck(workerData.batch as Batch, workerData.plan);
    parentPort!.postMessage({ type: 'result', result });
  } else throw new Error('未知任务。');
} catch (error) {
  parentPort!.postMessage({ type: 'error', error: error instanceof Error ? error.message : String(error) });
}
