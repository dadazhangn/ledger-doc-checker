import { parentPort, workerData } from 'node:worker_threads';
import { inspectFiles } from './parsers.js';
import { runCheck } from './engine.js';
import { compareWordSelections, inspectWordFiles } from './word-compare.js';
import type { Batch, UploadedFile } from './types.js';
import type { WordBatch } from './word-types.js';

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
  } else if (workerData.task === 'word-inspect') {
    parentPort!.postMessage({ type: 'progress', message: '正在提取所选文档中的可选内容…' });
    const files = (workerData.files as UploadedFile[]).map((file) => ({ name: file.name, bytes: Buffer.from(file.data, 'base64') }));
    const batch = await inspectWordFiles(files);
    parentPort!.postMessage({ type: 'result', result: batch });
  } else if (workerData.task === 'word-compare') {
    parentPort!.postMessage({ type: 'progress', message: '正在整理已选择的字段与来源…' });
    const result = compareWordSelections(workerData.batch as WordBatch, workerData.selections);
    parentPort!.postMessage({ type: 'result', result });
  } else throw new Error('未知任务。');
} catch (error) {
  parentPort!.postMessage({ type: 'error', error: error instanceof Error ? error.message : String(error) });
}
