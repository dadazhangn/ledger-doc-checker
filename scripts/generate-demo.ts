import ExcelJS from 'exceljs';
import yazl from 'yazl';
import { mkdir, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import type { Plan } from '../src/types.js';

export const xmlEscape = (text: string) => text.replace(/[&<>"']/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[char]!);
export async function zipBuffer(parts: Record<string, string>): Promise<Buffer> {
  const zip = new yazl.ZipFile();
  for (const [name, text] of Object.entries(parts)) zip.addBuffer(Buffer.from(text), name, { mtime: new Date('2020-01-01T00:00:00Z') });
  const result = new Promise<Buffer>((resolve, reject) => {
    const chunks: Buffer[] = [];
    zip.outputStream.on('data', (chunk: Buffer) => chunks.push(chunk));
    zip.outputStream.on('error', reject);
    zip.outputStream.on('end', () => resolve(Buffer.concat(chunks)));
  });
  zip.end();
  return result;
}
export const paragraphXml = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${xmlEscape(text)}</w:t></w:r></w:p>`;
export const tableXml = (rows: string[][]) => `<w:tbl>${rows.map((row) => `<w:tr>${row.map((text) => `<w:tc>${paragraphXml(text)}</w:tc>`).join('')}</w:tr>`).join('')}</w:tbl>`;
export async function wordBuffer(body: string): Promise<Buffer> {
  return zipBuffer({
    '[Content_Types].xml': '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    '_rels/.rels': '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    'word/document.xml': `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr/></w:body></w:document>`,
  });
}
export const demoPlan: Plan = {
  version: 1,
  ledger: { sheet: '对象台账', headerRow: 1, keyColumn: 'A' },
  identityLabel: '对象名称',
  bindings: [{ document: '演示说明-甲.docx', key: '演示对象甲' }, { document: '演示说明-乙.docx', key: '演示对象乙' }],
  fields: [
    { sourceColumn: 'B', targetLabel: '版本', comparison: 'exact', required: true },
    { sourceColumn: 'C', targetLabel: '出具日期', comparison: 'date', required: true },
  ],
  details: { sheet: '明细', headerRow: 1, objectColumn: 'A', keyColumn: 'B', targetKeyHeader: '条目名称', columns: [
    { sourceColumn: 'C', targetHeader: '数量', comparison: 'number', tolerance: 0, required: true },
    { sourceColumn: 'D', targetHeader: '状态', comparison: 'exact', required: true },
  ] },
  allowCachedFormula: false,
};

export async function demoFiles(): Promise<{ name: string; bytes: Buffer }[]> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Ledger Doc Checker synthetic demo';
  workbook.created = new Date('2020-01-01T00:00:00Z');
  workbook.modified = workbook.created;
  const objects = workbook.addWorksheet('对象台账');
  objects.addRows([
    ['对象名称', '版本', '出具日期'],
    ['演示对象甲', '1.0', '2026-11-01'],
    ['演示对象乙', '2.1', '2026-11-02'],
  ]);
  const details = workbook.addWorksheet('明细');
  details.addRows([
    ['对象名称', '条目名称', '数量', '状态'],
    ['演示对象甲', '演示条目甲', 3, '已确认'],
    ['演示对象甲', '演示条目乙', 5, '待确认'],
    ['演示对象乙', '演示条目甲', 2, '已确认'],
    ['演示对象乙', '演示条目丙', 4, '待确认'],
  ]);
  for (const sheet of workbook.worksheets) {
    sheet.views = [{ state: 'frozen', ySplit: 1 }];
    sheet.columns.forEach((column) => { column.width = 22; });
    sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF246B60' } };
  }
  const heading = paragraphXml('虚构演示资料：仅用于验证工具功能，不代表真实项目或实际测试结果。');
  const fieldsA = tableXml([['字段', '取值'], ['对象名称', '演示对象甲'], ['版本', '1.1'], ['出具日期', '2026年11月1日']])
    .replace('<w:t xml:space="preserve">版本</w:t>', '<w:t>版</w:t></w:r><w:r><w:t>本</w:t>');
  const fieldsB = tableXml([['字段', '取值'], ['对象名称', '演示对象乙'], ['版本', '2.1'], ['出具日期', '2026年11月2日']]);
  return [
    { name: '演示台账.xlsx', bytes: Buffer.from(await workbook.xlsx.writeBuffer()) },
    { name: '演示说明-甲.docx', bytes: await wordBuffer(heading + fieldsA + tableXml([
      ['条目名称', '数量', '状态'], ['演示条目乙', '6', '待确认'], ['演示条目甲', '3', '已确认'], ['演示附加项', '1', '已确认'],
    ])) },
    { name: '演示说明-乙.docx', bytes: await wordBuffer(heading + fieldsB + tableXml([
      ['条目名称', '数量', '状态'], ['演示条目甲', '2', '已确认'],
    ])) },
  ];
}
if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const output = join(process.cwd(), 'examples');
  await mkdir(output, { recursive: true });
  for (const file of await demoFiles()) await writeFile(join(output, file.name), file.bytes);
  await writeFile(join(output, 'demo-plan.json'), JSON.stringify(demoPlan, null, 2) + '\n');
  console.log('已生成 3 份明确标注为虚构数据的演示文件和一份核对方案。');
}
