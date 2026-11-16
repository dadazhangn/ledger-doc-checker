import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import { demoFiles, demoPlan, paragraphXml, tableXml, wordBuffer, zipBuffer } from '../scripts/generate-demo.js';
import { inspectFiles, parseWord, hash } from '../src/parsers.js';
import { compareValues, runCheck, validatePlan } from '../src/engine.js';
import { reportHtml } from '../src/report.js';
import type { Batch } from '../src/types.js';

const demo = await demoFiles();
const batch = await inspectFiles(demo);
const plan = () => structuredClone(demoPlan);
const fresh = () => structuredClone(batch);
const fieldCell = (value: Batch, object: string, label: string) => {
  const doc = value.documents.find((d) => d.name.includes(object))!;
  return doc.tables[0].rows.find((row) => row[0]?.text === label)![1];
};

test('虚构演示完整检出四种已知差异，保留双方来源', () => {
  const report = runCheck(batch, plan());
  assert.equal(report.summary.different, 4);
  assert.equal(report.summary.unverifiable, 0);
  assert.equal(report.summary.fileErrors, 0);
  assert.equal(report.summary.complete, true);
  assert.equal(report.summary.verdict, 'different');
  assert.deepEqual(report.findings.filter((f) => f.status === 'different').map((f) => f.kind).sort(), ['detail', 'extra', 'field', 'missing']);
  const version = report.findings.find((f) => f.field === '版本' && f.status === 'different')!;
  assert.equal(version.expected?.location, '对象台账!B2');
  assert.equal(version.actual?.location, '正文·表格 1·行 3·列 2');
  assert.equal(version.expected?.value, '1.0');
  assert.equal(version.actual?.value, '1.1');
});

test('相同输入和方案产生相同结论与指纹，读取不改变源字节', async () => {
  const before = demo.map((f) => hash(f.bytes));
  const again = await inspectFiles(demo);
  assert.deepEqual(runCheck(again, plan()), runCheck(batch, plan()));
  assert.deepEqual(demo.map((f) => hash(f.bytes)), before);
});

test('正确文档经真实 DOCX 解析，在明细调序后仍全部一致', async () => {
  const correct = [demo[0],
    { name: '演示说明-甲.docx', bytes: await wordBuffer(tableXml([['对象名称', '演示对象甲'], ['版本', '1.0'], ['出具日期', '2026/11/1']]) + tableXml([['条目名称', '数量', '状态'], ['演示条目乙', '5.00', '待确认'], ['演示条目甲', '3', '已确认']])) },
    { name: '演示说明-乙.docx', bytes: await wordBuffer(tableXml([['对象名称', '演示对象乙'], ['版本', '2.1'], ['出具日期', '2026-11-02']]) + tableXml([['条目名称', '数量', '状态'], ['演示条目丙', '4', '待确认'], ['演示条目甲', '2', '已确认']])) },
  ];
  const report = runCheck(await inspectFiles(correct), plan());
  assert.equal(report.summary.verdict, 'equal');
  assert.equal(report.summary.different, 0);
  assert.equal(report.summary.unverifiable, 0);
  assert.equal(report.summary.equal, 14);
});

test('台账对象键重复时，不选第一行且不能给完整结论', () => {
  const value = fresh();
  const cells = value.ledger!.sheets[0].cells;
  cells.find((c) => c.address === 'A3')!.value.text = '演示对象甲';
  const report = runCheck(value, plan());
  assert.equal(report.summary.verdict, 'incomplete');
  assert(report.findings.some((f) => f.message.includes('对应多行')));
});

test('Word 对象与手工绑定不一致时，停止该文档后续检查', () => {
  const p = plan(); p.bindings[0].key = '演示对象乙';
  const report = runCheck(batch, p);
  assert.equal(report.summary.verdict, 'incomplete');
  assert(!report.findings.some((f) => f.document === p.bindings[0].document && f.kind === 'field'));
});

test('同名 Word 字段命中多处时记录歧义', () => {
  const value = fresh(); value.documents[0].paragraphs.push({ text: '版本:1.0', location: '正文·段落 2', reasons: [] });
  const report = runCheck(value, plan());
  assert(report.findings.some((f) => f.field === '版本' && f.status === 'unverifiable' && f.message.includes('2 处')));
});

test('Word 明细关联键重复时不拿任意一行参与比较', () => {
  const value = fresh(); const table = value.documents[0].tables[1]; table.rows.push(structuredClone(table.rows[1]));
  const report = runCheck(value, plan());
  assert.equal(report.summary.verdict, 'incomplete');
  assert(report.findings.some((f) => f.field.includes('演示条目乙') && f.message.includes('重复')));
  assert(!report.findings.some((f) => f.kind === 'detail' && f.field === '演示条目乙 · 数量'));
});

test('损坏的一个文件保留错误，其他文件可核对，整体不能通过', async () => {
  const value = await inspectFiles([demo[0], demo[1], { name: demo[2].name, bytes: Buffer.from('not a ZIP') }]);
  const report = runCheck(value, plan());
  assert.equal(report.summary.fileErrors, 1);
  assert.equal(report.summary.verdict, 'incomplete');
  assert.equal(report.fileErrors[0].name, demo[2].name);
});

test('公式默认阻止核对，显式采用缓存后保留时效提示', () => {
  const value = fresh(); const cell = value.ledger!.sheets[0].cells.find((c) => c.address === 'B2')!;
  cell.value = { kind: 'formula', raw: '"1.0"', formula: '"1.0"', text: '', cached: { kind: 'text', raw: '1.0', text: '1.0' } };
  assert(runCheck(value, plan()).findings.some((f) => f.field === '版本' && f.status === 'unverifiable'));
  const p = plan(); p.allowCachedFormula = true;
  assert(runCheck(value, p).findings.some((f) => f.field === '版本' && f.notices?.[0].includes('未重算')));
  delete cell.value.cached;
  assert(runCheck(value, p).findings.some((f) => f.field === '版本' && f.status === 'unverifiable'));
});

test('合并来源值、未处理修订与域代码不会静默通过', async () => {
  const merged = fresh(); merged.ledger!.sheets[0].cells.find((c) => c.address === 'B2')!.merged = true;
  assert.equal(runCheck(merged, plan()).summary.verdict, 'incomplete');
  const revision = await parseWord(await wordBuffer('<w:tbl><w:tr><w:tc>' + paragraphXml('版本') + '</w:tc><w:tc><w:p><w:ins><w:r><w:t>1.0</w:t></w:r></w:ins></w:p></w:tc></w:tr></w:tbl>'), '修订演示.docx');
  assert(revision.tables[0].rows[0][1].reasons.some((r) => r.includes('修订')));
  const field = await parseWord(await wordBuffer('<w:p><w:fldSimple w:instr="DATE"><w:r><w:t>2026-11-01</w:t></w:r></w:fldSimple></w:p>'), '域代码演示.docx');
  assert(field.paragraphs[0].reasons.some((r) => r.includes('域代码')));
});

test('纵向合并和嵌套表格有明确不支持原因', async () => {
  const vertical = await parseWord(await wordBuffer('<w:tbl><w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>' + paragraphXml('字段') + '</w:tc></w:tr></w:tbl>'), '合并演示.docx');
  assert(vertical.tables[0].reasons.some((r) => r.includes('纵向合并')));
  const nested = await parseWord(await wordBuffer('<w:tbl><w:tr><w:tc>' + tableXml([['字段', '值']]) + '</w:tc></w:tr></w:tbl>'), '嵌套演示.docx');
  assert(nested.tables[0].reasons.some((r) => r.includes('嵌套')));
});

test('数字模式按十进制比较，日期模式拒绝无效日期和含糊日期', () => {
  assert.equal(compareValues('0.3', '0.30', 'number', true).status, 'equal');
  assert.equal(compareValues('0.3', '0.31', 'number', true, 0.01).status, 'equal');
  assert.equal(compareValues('0.3', '0.32', 'number', true, 0.01).status, 'different');
  assert.equal(compareValues('001', '1', 'exact', true).status, 'different');
  assert.equal(compareValues(' v1 ', 'v1', 'exact', true).status, 'different');
  assert.equal(compareValues(' v1 ', 'v1', 'trim', true).status, 'equal');
  assert.equal(compareValues('2026-02-30', '2026-02-30', 'date', true).status, 'unverifiable');
  assert.equal(compareValues('01/02/2026', '2026-02-01', 'date', true).status, 'unverifiable');
  assert.equal(compareValues('', '', 'exact', true).status, 'different');
  assert.equal(compareValues('', '', 'exact', false).status, 'equal');
});

test('Excel 数字格式里的前导零保留为关联值', async () => {
  const book = new ExcelJS.Workbook(); const source = book.addWorksheet('对象');
  source.addRows([['对象键', '版本'], [7, '1.0']]); source.getCell('A2').numFmt = '000';
  const value = await inspectFiles([{ name: '前导零演示.xlsx', bytes: Buffer.from(await book.xlsx.writeBuffer()) }, { name: '前导零演示.docx', bytes: await wordBuffer(tableXml([['对象键', '007'], ['版本', '1.0']])) }]);
  const p = plan(); p.ledger = { sheet: '对象', headerRow: 1, keyColumn: 'A' }; p.identityLabel = '对象键'; p.bindings = [{ document: '前导零演示.docx', key: '007' }]; p.details = null; p.fields = [p.fields[0]];
  assert.equal(runCheck(value, p).summary.verdict, 'equal');
});

test('未知字段和未知明细表作为未完成项，不当作不存在差异', () => {
  const p = plan(); p.fields[0].targetLabel = '未提供的字段'; p.details!.targetKeyHeader = '未提供的列';
  const report = runCheck(batch, p);
  assert.equal(report.summary.verdict, 'incomplete');
  assert(report.summary.unverifiable >= 2);
});

test('没有比较规则、重复绑定和负容差被拒绝', () => {
  const p = plan(); p.fields = []; p.details = null;
  assert.throws(() => validatePlan(p), /至少配置/);
  const q = plan(); q.bindings.push(q.bindings[0]); assert.throws(() => validatePlan(q), /关联一次/);
  const r = plan(); r.fields[0].tolerance = -1; assert.throws(() => validatePlan(r), /容差/);
});

test('包含 DTD、损坏 XML 和重名文件的输入被拒绝', async () => {
  const dtd = await zipBuffer({ 'word/document.xml': '<!DOCTYPE w:document [<!ENTITY demo "text">]><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>' });
  await assert.rejects(parseWord(dtd, 'DTD演示.docx'), /DTD/);
  await assert.rejects(parseWord(await zipBuffer({ 'word/document.xml': '<broken>' }), '损坏演示.docx'));
  await assert.rejects(inspectFiles([demo[0], demo[1], demo[1]]), /重名/);
});

test('HTML 报告把业务内容当成文本，导出包含文件和方案指纹', () => {
  const value = fresh(); fieldCell(value, '甲', '版本').text = '<script>alert("demo")</script>';
  const report = runCheck(value, plan()); const html = reportHtml(report);
  assert(!html.includes('<script>'));
  assert(html.includes('&lt;script&gt;'));
  assert(html.includes(report.planHash));
  assert(html.includes(report.files[0].sha256));
});

test('源文件内容或方案变化会改变报告绑定指纹', () => {
  const p = plan(); p.fields[0].comparison = 'trim';
  assert.notEqual(runCheck(batch, p).planHash, runCheck(batch, plan()).planHash);
});
