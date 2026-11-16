import { randomUUID } from 'node:crypto';
import WordExtractor from 'word-extractor';
import { SaxesParser } from 'saxes';
import { hash, parseWord } from './parsers.js';
import { inspectZip } from './zip.js';
import type { WordBatch, WordFieldKind, WordRecord, WordReport, WordSegment, WordSelection, WordSource } from './word-types.js';

const categories: WordFieldKind[] = ['projectName', 'softwareVersion', 'documentVersion'];
const legacyNotice = '旧版 .doc 仅提取文本，表格位置、修订、域及完整性无法保证，请对照原文人工确认。';
const headerNotice = '页眉或页脚按文件部件提取，未确定所属节和实际页码，请人工确认。';
const maxSegments = 10000;
const maxSegmentLength = 2000;

function appendSegment(target: WordSegment[], id: string, location: string, text: string, notices: string[] = []) {
  if (!text.trim()) return;
  for (let offset = 0; offset < text.length; offset += maxSegmentLength) {
    if (target.length >= maxSegments) throw new Error('文档提取片段超过 10,000 个，请缩小文件范围。');
    const split = text.length > maxSegmentLength;
    target.push({ id: `${id}:${offset}`, location: split ? `${location}·提取字符 ${offset + 1}–${Math.min(offset + maxSegmentLength, text.length)}` : location,
      text: text.slice(offset, offset + maxSegmentLength), notices: split ? [...notices, '长内容按字符分段显示，请核对上下文。'] : [...notices] });
  }
}
function partNumber(name: string) { return name.match(/(\d+)\.xml$/)?.[1] ?? ''; }
function extractHeader(xml: Buffer, name: string): WordSegment[] {
  const parser = new SaxesParser({ xmlns: true });
  const result: WordSegment[] = [];
  const word = new Set(['http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'http://purl.oclc.org/ooxml/wordprocessingml/main']);
  const kind = name.includes('/header') ? '页眉' : '页脚';
  let depth = 0, current: { depth: number; text: string; reasons: string[] } | null = null;
  let textDepth = -1, skipDepth = -1, nodes = 0, paragraph = 0;
  parser.on('doctype', () => { throw new Error('不接受包含 DTD 的页眉页脚 XML。'); });
  parser.on('error', (error) => { throw error; });
  parser.on('opentag', (tag) => {
    if (++nodes > 150000 || ++depth > 100) throw new Error('页眉页脚 XML 结构超限。');
    const w = word.has(tag.uri);
    if (w && tag.local === 'p' && !current) current = { depth, text: '', reasons: [headerNotice] };
    if (!current) return;
    if (w && ['ins', 'del', 'moveFrom', 'moveTo'].includes(tag.local)) current.reasons.push('包含未处理的修订，请人工确认。');
    if (w && ['fldSimple', 'fldChar', 'instrText'].includes(tag.local)) current.reasons.push('包含域代码，显示值可能不是最新。');
    if ((w && ['drawing', 'pict', 'txbxContent', 'del', 'instrText'].includes(tag.local)) || tag.local === 'AlternateContent') {
      if (skipDepth < 0) skipDepth = depth;
      if (tag.local !== 'del' && tag.local !== 'instrText') current.reasons.push('包含图片或文本框，其内容未提取。');
    }
    if (skipDepth < 0 && w && tag.local === 't') textDepth = depth;
    if (skipDepth < 0 && w && tag.local === 'tab') current.text += '\t';
    if (skipDepth < 0 && w && ['br', 'cr'].includes(tag.local)) current.text += '\n';
  });
  parser.on('text', (text) => { if (current && textDepth >= 0 && skipDepth < 0) current.text += text; });
  parser.on('closetag', () => {
    if (textDepth === depth) textDepth = -1;
    if (skipDepth === depth) skipDepth = -1;
    if (current?.depth === depth) {
      paragraph++;
      appendSegment(result, `${name}:p:${paragraph}`, `${kind}${partNumber(name)}·提取段落 ${paragraph}`, current.text, [...new Set(current.reasons)]);
      current = null;
    }
    depth--;
  });
  parser.write(xml.toString('utf8')).close();
  return result;
}

export async function parseWordSource(data: Buffer, name: string, id = randomUUID()): Promise<WordSource> {
  if (data.length > 12 * 1024 * 1024) throw new Error('单文件超过 12 MB 上限。');
  if (/\.doc$/i.test(name)) {
    if (data.length < 512 || !data.subarray(0, 8).equals(Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]))) throw new Error('文件扩展名为 .doc，但不是受支持的旧版 Word 文件。');
    const document = await new WordExtractor().extract(data);
    const sections = [
      { id: 'body', name: '正文', text: document.getBody({ filterUnicode: false }) },
      { id: 'header', name: '页眉', text: document.getHeaders({ filterUnicode: false, includeFooters: false }) },
      { id: 'footer', name: '页脚', text: document.getFooters({ filterUnicode: false }) },
      { id: 'textbox', name: '文本框', text: document.getTextboxes({ filterUnicode: false }) },
    ];
    const segments: WordSegment[] = [];
    for (const section of sections) {
      const pieces = section.text.split(/[\t\n]/);
      pieces.forEach((piece, index) => appendSegment(segments, `${section.id}:${index + 1}`, `${section.name}提取片段 ${index + 1}`, piece, [legacyNotice]));
    }
    if (!segments.length) throw new Error('未提取到可供人工选择的文字；文件可能为空或受保护。');
    return { id, name, sha256: hash(data), format: 'doc', segments, notices: [legacyNotice] };
  }
  if (!/\.docx$/i.test(name)) throw new Error('只支持 .doc 和 .docx 文档。');
  const doc = await parseWord(data, name);
  const segments: WordSegment[] = [];
  doc.paragraphs.forEach((paragraph, index) => appendSegment(segments, `body:p:${index + 1}`, paragraph.location, paragraph.text, paragraph.reasons));
  for (const table of doc.tables) for (const row of table.rows) for (const cell of row) {
    appendSegment(segments, `body:t:${table.index}:r:${cell.row}:c:${cell.column}`, cell.location, cell.text, cell.reasons);
  }
  const parts = await inspectZip(data, (part) => /^word\/(header|footer)\d+\.xml$/.test(part));
  for (const [part, xml] of [...parts.entries()].sort(([a], [b]) => a.localeCompare(b))) {
    for (const segment of extractHeader(xml, part)) {
      if (segments.length >= maxSegments) throw new Error('文档提取片段超过 10,000 个，请缩小文件范围。');
      segments.push(segment);
    }
  }
  if (!segments.length) throw new Error('未提取到可供人工选择的文字；文件可能为空或只含图片。');
  return { id, name, sha256: hash(data), format: 'docx', segments,
    notices: ['图片、正文文本框、批注与实际页码未覆盖；存在修订、域或复杂表格的内容需人工确认。', ...(parts.size ? [headerNotice] : [])] };
}

export async function inspectWordFiles(files: { name: string; bytes: Buffer }[]): Promise<WordBatch> {
  if (files.length < 2 || files.length > 30) throw new Error('请选择 2–30 份 Word 文档。');
  const batch: WordBatch = { documents: [], files: [], errors: [] };
  for (const file of files) {
    const id = randomUUID();
    if (!file.name || file.name.length > 180 || /[\\/\x00-\x1f]/.test(file.name) || !/\.docx?$/i.test(file.name)) throw new Error('Word 文件名或类型无效。');
    batch.files.push({ id, name: file.name, sha256: hash(file.bytes) });
    try { batch.documents.push(await parseWordSource(file.bytes, file.name, id)); }
    catch (error) { batch.errors.push({ id, name: file.name, message: error instanceof Error ? error.message : String(error) }); }
  }
  return batch;
}

export function compareWordSelections(batch: WordBatch, input: unknown): WordReport {
  if (!Array.isArray(input) || input.length < 1 || input.length > 300) throw new Error('请手动选择 1–300 个需要查看的字段值。');
  const documents = new Map(batch.documents.map((doc) => [doc.id, doc]));
  const selections: WordSelection[] = [];
  const records: WordRecord[] = [];
  for (const raw of input) {
    if (!raw || typeof raw !== 'object' || Array.isArray(raw)) throw new Error('字段选择格式无效。');
    const entry = raw as Record<string, unknown>;
    if (typeof entry.documentId !== 'string' || typeof entry.segmentId !== 'string' ||
        !categories.includes(entry.category as WordFieldKind) || typeof entry.value !== 'string' ||
        !entry.value.trim() || entry.value.length > 200 || typeof entry.subject !== 'string' || entry.subject.length > 120 ||
        typeof entry.stage !== 'string' || entry.stage.length > 120) throw new Error('字段、来源、取值、对照组或阶段格式无效。');
    const document = documents.get(entry.documentId);
    const segment = document?.segments.find((item) => item.id === entry.segmentId);
    if (!document || !segment) throw new Error('文档或来源片段已失效，请重新选择。');
    if (!segment.text.includes(entry.value)) throw new Error('填写的值不在所选原文片段中，请重新核对来源。');
    const selection: WordSelection = { documentId: entry.documentId, category: entry.category as WordFieldKind, segmentId: entry.segmentId,
      value: entry.value, subject: entry.subject, stage: entry.stage };
    selections.push(selection);
    records.push({ ...selection, file: document.name, sha256: document.sha256, location: segment.location,
      notices: [...new Set([...document.notices, ...segment.notices])] });
  }
  const grouped = new Map<string, { category: WordFieldKind; subject: string; records: WordRecord[] }>();
  const uncompared: WordRecord[] = [];
  for (const record of records) {
    // 项目名称只做所选值的并列展示。版本需用户显式指定对照组，避免跨对象或文档类型自动关联。
    if (record.category !== 'projectName' && !record.subject.trim()) { uncompared.push(record); continue; }
    const subject = record.category === 'projectName' ? '' : record.subject;
    const key = JSON.stringify([record.category, subject]);
    const group = grouped.get(key) ?? { category: record.category, subject, records: [] };
    group.records.push(record); grouped.set(key, group);
  }
  const differences = [...grouped.values()].filter((group) => new Set(group.records.map((record) => record.value)).size > 1);
  for (const group of grouped.values()) if (group.records.length < 2) uncompared.push(...group.records);
  const coverage = batch.documents.map((doc) => {
    const selected = categories.filter((kind) => records.some((record) => record.documentId === doc.id && record.category === kind));
    return { documentId: doc.id, file: doc.name, selected, unselected: categories.filter((kind) => !selected.includes(kind)) };
  });
  // 文档 ID 是本次导入的临时标识。指纹绑定真实文件内容和来源位置，避免每次导入后无故变化。
  const stableSelections = records.map(({ category, value, subject, stage, sha256, location }) => ({ category, value, subject, stage, sha256, location }));
  return { schemaVersion: 1, engineVersion: '0.2.0', mode: 'word-comparison', judgment: 'human-review',
    selectionHash: hash(JSON.stringify(stableSelections)), files: batch.files, errors: batch.errors, records, differences, uncompared, coverage };
}
