import { createHash } from 'node:crypto';
import ExcelJS from 'exceljs';
import { SaxesParser } from 'saxes';
import { inspectZip } from './zip.js';
import type { Batch, CellValue, Ledger, WordCell, WordDocument, WordParagraph, WordTable } from './types.js';

export const hash = (data: Buffer | string) => createHash('sha256').update(data).digest('hex');
const wordNamespaces = new Set(['http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'http://purl.oclc.org/ooxml/wordprocessingml/main']);

function valueOf(value: unknown, numberFormat = ''): CellValue {
  if (value === undefined || value === null) return { kind: 'blank', raw: null, text: '' };
  if (value instanceof Date) return { kind: 'date', raw: value.toISOString(), text: value.toISOString().slice(0, 10) };
  if (typeof value === 'string') return { kind: 'text', raw: value, text: value };
  if (typeof value === 'boolean') return { kind: 'boolean', raw: value, text: String(value) };
  if (typeof value === 'number') {
    const text = Number.isInteger(value) && value >= 0 && /^0{2,}$/.test(numberFormat)
      ? String(value).padStart(numberFormat.length, '0') : String(value);
    return { kind: 'number', raw: value, text };
  }
  if (typeof value === 'object') {
    const entry = value as Record<string, unknown>;
    if ('formula' in entry || 'sharedFormula' in entry) return {
      kind: 'formula', raw: String(entry.formula ?? entry.sharedFormula),
      formula: String(entry.formula ?? entry.sharedFormula), text: '',
      ...('result' in entry ? { cached: valueOf(entry.result, numberFormat) } : {}),
    };
    if (Array.isArray(entry.richText)) {
      const text = entry.richText.map((part) => String(part.text ?? '')).join('');
      return { kind: 'text', raw: text, text };
    }
    if (typeof entry.text === 'string') return { kind: 'text', raw: entry.text, text: entry.text };
    if ('error' in entry) return { kind: 'error', raw: String(entry.error), text: String(entry.error) };
  }
  return { kind: 'error', raw: null, text: '暂不支持的单元格类型' };
}

export async function parseLedger(data: Buffer, name: string): Promise<Ledger> {
  await inspectZip(data, () => false);
  const book = new ExcelJS.Workbook();
  await book.xlsx.load(data as unknown as Parameters<typeof book.xlsx.load>[0]);
  if (book.worksheets.length > 20) throw new Error('工作表数量超过 20 张上限。');
  let totalCells = 0;
  const sheets = book.worksheets.map((sheet) => {
    if (sheet.rowCount > 10001 || sheet.columnCount > 100) throw new Error('工作表超过 10,000 条数据行或 100 列的首版限制。');
    const cells: Ledger['sheets'][number]['cells'] = [];
    sheet.eachRow((row) => row.eachCell((cell) => {
      totalCells++;
      if (totalCells > 250000) throw new Error('非空单元格数量超过 250,000 个上限。');
      cells.push({ address: cell.address, row: cell.fullAddress.row, column: cell.fullAddress.col,
        value: valueOf(cell.value, cell.numFmt), ...(cell.isMerged ? { merged: true } : {}) });
    }));
    return { name: sheet.name, rowCount: sheet.rowCount, columnCount: sheet.columnCount, cells };
  });
  return { name, sha256: hash(data), sheets };
}

interface XmlNode { local: string; uri: string; attributes: Record<string, string>; children: (XmlNode | string)[] }
const isWord = (node: XmlNode, name?: string) => wordNamespaces.has(node.uri) && (!name || node.local === name);
const nodes = (node: XmlNode): XmlNode[] => node.children.filter((child): child is XmlNode => typeof child !== 'string');
function descendants(node: XmlNode, predicate: (value: XmlNode) => boolean): XmlNode[] {
  const result: XmlNode[] = [];
  for (const child of nodes(node)) { if (predicate(child)) result.push(child); result.push(...descendants(child, predicate)); }
  return result;
}
function ownReasons(node: XmlNode): string[] {
  if (isWord(node) && ['ins', 'del', 'moveFrom', 'moveTo'].includes(node.local)) return ['包含未处理的修订，需人工确认。'];
  if (isWord(node) && ['fldSimple', 'fldChar', 'instrText'].includes(node.local)) return ['包含域代码，无法确认显示值是否最新。'];
  if ((isWord(node) && ['drawing', 'pict', 'txbxContent', 'altChunk'].includes(node.local)) || node.local === 'AlternateContent') return ['包含图片、文本框或外部内容，本版未解析。'];
  return [];
}
function reasons(node: XmlNode): string[] { return [...new Set([...ownReasons(node), ...nodes(node).flatMap(reasons)])]; }
function textOf(node: XmlNode): string {
  if ((isWord(node) && ['drawing', 'pict', 'txbxContent', 'del', 'instrText', 'altChunk'].includes(node.local)) || node.local === 'AlternateContent') return '';
  if (isWord(node, 't')) return node.children.filter((x): x is string => typeof x === 'string').join('');
  if (isWord(node, 'tab')) return '\t';
  if (isWord(node, 'br') || isWord(node, 'cr')) return '\n';
  return nodes(node).map(textOf).join('');
}
function structural(node: XmlNode, name: string): XmlNode[] {
  return nodes(node).flatMap((child) => {
    if (isWord(child, name)) return [child];
    if (isWord(child) && ['sdt', 'sdtContent', 'customXml', 'ins', 'del'].includes(child.local)) return structural(child, name);
    return [];
  });
}

export async function parseWord(data: Buffer, name: string): Promise<WordDocument> {
  const parts = await inspectZip(data, (part) => part === 'word/document.xml' || /^word\/(header|footer)\d+\.xml$/.test(part));
  const documentXml = parts.get('word/document.xml');
  if (!documentXml) throw new Error('DOCX 缺少 word/document.xml。');
  const xml = documentXml.toString('utf8');
  const root: XmlNode = { local: 'root', uri: '', attributes: {}, children: [] };
  const stack = [root];
  const parser = new SaxesParser({ xmlns: true });
  let nodeCount = 0;
  parser.on('doctype', () => { throw new Error('不接受包含 DTD 的 Office XML。'); });
  parser.on('error', (error) => { throw error; });
  parser.on('opentag', (tag) => {
    if (++nodeCount > 150000 || stack.length > 100) throw new Error('Word XML 的结构深度或节点数超限。');
    const attributes: Record<string, string> = Object.create(null);
    for (const attr of Object.values(tag.attributes)) attributes[attr.local] = attr.value;
    const node: XmlNode = { local: tag.local, uri: tag.uri, attributes, children: [] };
    stack.at(-1)!.children.push(node);
    stack.push(node);
  });
  parser.on('text', (text) => stack.at(-1)!.children.push(text));
  parser.on('closetag', () => { stack.pop(); });
  parser.write(xml).close();
  const body = descendants(root, (node) => isWord(node, 'body'))[0];
  if (!body) throw new Error('DOCX 缺少受支持的 Word 正文结构。');
  const paragraphs: WordParagraph[] = [];
  const tables: WordTable[] = [];
  const notices = ['核对范围为正文段落和普通表格；不覆盖页眉页脚、图片、文本框和版面页码。'];
  if (parts.size > 1) notices.push('文件含页眉或页脚，它们不在本次检查范围内。');
  const collect = (container: XmlNode, inherited: string[] = []) => {
    for (const child of nodes(container)) {
      const inheritedHere = [...new Set([...inherited, ...ownReasons(child)])];
      if (isWord(child, 'p')) paragraphs.push({ text: textOf(child), location: `正文·段落 ${paragraphs.length + 1}`, reasons: [...new Set([...inheritedHere, ...reasons(child)])] });
      else if (isWord(child, 'tbl')) {
        const index = tables.length + 1;
        const location = `正文·表格 ${index}`;
        const tableReasons = [...inheritedHere];
        if (descendants(child, (n) => isWord(n, 'tbl')).length) tableReasons.push('包含嵌套表格，本版不核对该表格。');
        if (descendants(child, (n) => isWord(n, 'vMerge')).length) tableReasons.push('包含纵向合并单元格，本版不核对该表格。');
        if (descendants(child, (n) => isWord(n) && ['ins', 'del', 'moveFrom', 'moveTo'].includes(n.local)).length) tableReasons.push('表格包含修订，需人工确认。');
        const rows = structural(child, 'tr').map((row, rowIndex) => {
          let column = Number(descendants(row, (n) => isWord(n, 'gridBefore'))[0]?.attributes.val ?? 0) + 1;
          return structural(row, 'tc').map((cell): WordCell => {
            const span = Number(descendants(cell, (n) => isWord(n, 'gridSpan'))[0]?.attributes.val ?? 1);
            if (!Number.isInteger(span) || span < 1 || span > 100) throw new Error('Word 表格跨度超出支持范围。');
            const result = { text: structural(cell, 'p').map(textOf).join('\n'), row: rowIndex + 1, column,
              columnSpan: span, location: `${location}·行 ${rowIndex + 1}·列 ${column}`, reasons: [...new Set([...tableReasons, ...reasons(cell)])] };
            column += span;
            return result;
          });
        });
        tables.push({ index, location, rows, reasons: [...new Set(tableReasons)] });
      } else if (!ownReasons(child).length) collect(child, inheritedHere);
      else notices.push(...inheritedHere);
    }
  };
  collect(body);
  return { name, sha256: hash(data), paragraphs, tables, notices: [...new Set(notices)] };
}

export async function inspectFiles(files: { name: string; bytes: Buffer }[]): Promise<Batch> {
  const batch: Batch = { ledger: null, documents: [], errors: [], files: [] };
  const seen = new Set<string>();
  if (files.length < 2 || files.length > 31) throw new Error('请选择 1 份 Excel 和 1–30 份 Word。');
  if (files.filter((file) => /\.xlsx$/i.test(file.name)).length !== 1) throw new Error('每批需要且只支持一份 .xlsx 台账。');
  for (const file of files) {
    if (!file.name || file.name.length > 180 || /[\\/\x00-\x1f]/.test(file.name)) throw new Error('文件名无效。');
    if (seen.has(file.name)) throw new Error('同一批次不接受重名文件，请先重命名。');
    seen.add(file.name);
    batch.files.push({ name: file.name, sha256: hash(file.bytes) });
    try {
      if (/\.xlsx$/i.test(file.name)) batch.ledger = await parseLedger(file.bytes, file.name);
      else if (/\.docx$/i.test(file.name)) batch.documents.push(await parseWord(file.bytes, file.name));
      else throw new Error('首版只支持 .xlsx 和 .docx。');
    } catch (error) { batch.errors.push({ name: file.name, message: error instanceof Error ? error.message : String(error) }); }
  }
  return batch;
}
