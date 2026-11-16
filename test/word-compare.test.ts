import test from 'node:test';
import assert from 'node:assert/strict';
import { paragraphXml, wordBuffer, wordDemoFiles, zipBuffer } from '../scripts/generate-demo.js';
import { compareWordSelections, inspectWordFiles, parseWordSource } from '../src/word-compare.js';
import { wordReportHtml } from '../src/report.js';
import type { WordBatch, WordSelection } from '../src/word-types.js';

const demo = await wordDemoFiles();
const batch = await inspectWordFiles(demo);
function selected(value: WordBatch): WordSelection[] {
  return value.documents.flatMap((doc, index) => {
    const stage = index ? '后期' : '前期';
    return [
      { documentId: doc.id, category: 'projectName' as const, segmentId: 'body:t:1:r:2:c:2:0', value: '演示项目', subject: '', stage },
      { documentId: doc.id, category: 'softwareVersion' as const, segmentId: 'body:t:1:r:3:c:2:0', value: index ? '2.0' : '1.0', subject: '演示软件', stage },
      { documentId: doc.id, category: 'documentVersion' as const, segmentId: 'body:t:1:r:4:c:2:0', value: index ? 'B' : 'A', subject: '演示文档', stage },
    ];
  });
}

test('不同时期的软件版本和文档版本分别列出差异，保留人工判断', () => {
  const report = compareWordSelections(batch, selected(batch));
  assert.equal(report.judgment, 'human-review');
  assert.equal(report.differences.length, 2);
  assert.deepEqual(report.differences.map((g) => g.category), ['softwareVersion', 'documentVersion']);
  assert.deepEqual(report.differences[0].records.map((r) => r.stage), ['前期', '后期']);
  assert.deepEqual(report.differences[0].records.map((r) => r.value), ['1.0', '2.0']);
  assert(!('verdict' in report));
  assert(report.records.every((record) => record.location.includes('正文·表格')));
});

test('相同软件对照组的多条记录可以在同一份文档中属于不同阶段', () => {
  const source = batch.documents[0];
  const expanded = structuredClone(batch);
  const fake = expanded.documents[0].segments.find((x) => x.id === 'body:t:1:r:3:c:2:0')!;
  fake.text = '1.0 / 2.0';
  const values = [
    { documentId: source.id, category: 'softwareVersion', segmentId: fake.id, value: '1.0', subject: '同一对象', stage: '早期' },
    { documentId: source.id, category: 'softwareVersion', segmentId: fake.id, value: '2.0', subject: '同一对象', stage: '后期' },
  ];
  const report = compareWordSelections(expanded, values);
  assert.equal(report.differences.length, 1);
  assert.deepEqual(report.differences[0].records.map((x) => x.stage), ['早期', '后期']);
});

test('不同类别独立，未指定对照组的版本不会被自动比较', () => {
  const entries = selected(batch);
  entries.filter((x) => x.category !== 'projectName').forEach((x) => { x.subject = ''; });
  const report = compareWordSelections(batch, entries);
  assert.equal(report.differences.length, 0);
  assert.equal(report.uncompared.length, 4);
  entries.filter((x) => x.category !== 'projectName').forEach((x) => { x.subject = '同一名称'; });
  assert.deepEqual(compareWordSelections(batch, entries).differences.map((x) => x.category), ['softwareVersion', 'documentVersion']);
});

test('同名文件保持独立身份，用户可以标注不同阶段', async () => {
  const duplicated = await inspectWordFiles(demo.map((file) => ({ ...file, name: '演示文档.docx' })));
  assert.equal(duplicated.documents.length, 2);
  assert.notEqual(duplicated.documents[0].id, duplicated.documents[1].id);
  const report = compareWordSelections(duplicated, selected(duplicated));
  assert.equal(report.differences.length, 2);
  assert.equal(new Set(report.records.map((r) => r.file)).size, 1);
  assert.equal(new Set(report.records.map((r) => r.documentId)).size, 2);
});

test('取值必须来自用户选中的原文片段，不接受自行填写的新值', () => {
  const entries = selected(batch); entries[0].value = '片段中不存在的内容';
  assert.throws(() => compareWordSelections(batch, entries), /不在所选原文片段/);
});

test('未选择的字段和未分组的版本明确列出，不能转换成通过结论', () => {
  const report = compareWordSelections(batch, [selected(batch)[1]]);
  assert.deepEqual(report.coverage[0].unselected, ['projectName', 'documentVersion']);
  assert.deepEqual(report.coverage[1].unselected, ['projectName', 'softwareVersion', 'documentVersion']);
  assert.equal(report.uncompared.length, 1);
  assert.equal(report.judgment, 'human-review');
  assert.match(wordReportHtml(report), /未参与分组比较的内容/);
  assert.match(wordReportHtml(report), /未指定组名或组内只有一条记录/);
});

test('相同材料重新导入后报告指纹稳定，旧的临时文档标识不能复用', async () => {
  const next = await inspectWordFiles(demo);
  assert.notEqual(next.documents[0].id, batch.documents[0].id);
  assert.equal(compareWordSelections(batch, selected(batch)).selectionHash, compareWordSelections(next, selected(next)).selectionHash);
  assert.throws(() => compareWordSelections(next, selected(batch)), /已失效/);
});

test('无效、损坏、旧版 Word 文件与有效文件混合时保留读取失败信息', async () => {
  const mixed = await inspectWordFiles([demo[0], { name: '损坏的演示.doc', bytes: Buffer.from('not doc') }]);
  assert.equal(mixed.documents.length, 1);
  assert.equal(mixed.errors.length, 1);
  const report = compareWordSelections(mixed, [selected(mixed)[0]]);
  assert.equal(report.errors.length, 1);
  assert.equal(report.judgment, 'human-review');
});

test('页眉和页脚按文件部件提取，并标明无法确认节与页码', async () => {
  const base = '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + paragraphXml('演示正文') + '</w:body></w:document>';
  const header = '<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' + paragraphXml('演示页眉') + '</w:hdr>';
  const footer = '<?xml version="1.0"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' + paragraphXml('演示页脚') + '</w:ftr>';
  const doc = await parseWordSource(await zipBuffer({ 'word/document.xml': base, 'word/header1.xml': header, 'word/footer1.xml': footer }), '演示.docx');
  const fromHeader = doc.segments.find((segment) => segment.text === '演示页眉');
  assert(fromHeader?.location.includes('页眉1'));
  assert(fromHeader?.notices.some((notice) => notice.includes('未确定所属节')));
  assert(doc.segments.some((segment) => segment.text === '演示页脚'));
});

test('页眉中的域、修订与恶意 DTD 均不会被忽略', async () => {
  const base = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + paragraphXml('演示正文') + '</w:body></w:document>';
  const header = '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:fldSimple w:instr="DATE"><w:r><w:t>演示日期</w:t></w:r></w:fldSimple></w:p></w:hdr>';
  const doc = await parseWordSource(await zipBuffer({ 'word/document.xml': base, 'word/header1.xml': header }), '演示.docx');
  assert(doc.segments.find((item) => item.text === '演示日期')?.notices.some((notice) => notice.includes('域代码')));
  const dtd = '<!DOCTYPE w:hdr><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>';
  await assert.rejects(parseWordSource(await zipBuffer({ 'word/document.xml': base, 'word/header1.xml': dtd }), '演示.docx'), /DTD/);
});

test('长片段分段并提示用户复核上下文', async () => {
  const doc = await parseWordSource(await wordBuffer(paragraphXml('甲'.repeat(5000))), '演示.docx');
  assert.equal(doc.segments.length, 3);
  assert(doc.segments.every((segment) => segment.notices.some((notice) => notice.includes('上下文'))));
  assert.equal(doc.segments.map((segment) => segment.text).join(''), '甲'.repeat(5000));
});

test('报告对所选值做 HTML 转义，展示不同阶段和未选择字段', async () => {
  const content = '<script>alert(1)</script>';
  const modified = structuredClone(batch);
  const segment = modified.documents[0].segments.find((item) => item.id === 'body:t:1:r:2:c:2:0')!;
  segment.text = content;
  const entries = [selected(batch)[0]]; entries[0].value = content;
  const html = wordReportHtml(compareWordSelections(modified, entries));
  assert(!html.includes('<script>'));
  assert(html.includes('&lt;script&gt;'));
  assert(html.includes('未选择'));
  assert(html.includes('人工判断'));
});
