import type { Report } from './types.js';
import type { WordFieldKind, WordReport } from './word-types.js';

export const escapeHtml = (value: unknown) => String(value ?? '').replace(/[&<>"']/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[char]!);
export function reportHtml(report: Report): string {
  const statuses = { equal: '一致', different: '差异', unverifiable: '无法核对' };
  const title = report.summary.verdict === 'incomplete' ? '检查未完成' : report.summary.verdict === 'different' ? '发现差异' : '配置范围内一致';
  const evidence = (value: Report['findings'][number]['expected']) => value
    ? `<strong>${escapeHtml(value.value || '（空）')}</strong><small>${escapeHtml(value.file)} · ${escapeHtml(value.location)}</small>` : '—';
  return `<!doctype html><html lang="zh-CN"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>台账核对报告</title>
<style>body{font:15px/1.65 system-ui,sans-serif;color:#162c3e;background:#f3f6f8;margin:0}main{max-width:1240px;margin:40px auto;background:white;padding:36px;border-radius:18px}h1{margin:4px 0 20px}.eyebrow{color:#50716d}small{display:block;color:#64748b;overflow-wrap:anywhere}.summary{display:flex;gap:24px;flex-wrap:wrap;padding:20px;background:#eef7f5;border-radius:12px}table{width:100%;border-collapse:collapse;margin:24px 0}td,th{text-align:left;vertical-align:top;border-bottom:1px solid #dfe6eb;padding:12px;overflow-wrap:anywhere}th{background:#f3f6f8}strong{white-space:pre-wrap}.different{color:#b45309}.unverifiable{color:#7c3aed}.equal{color:#137969}details{margin-top:24px}pre{white-space:pre-wrap;word-break:break-word;font-size:12px}@media print{body{background:white}main{padding:0;margin:0}tr{break-inside:avoid}}@media(max-width:700px){main{margin:0;padding:16px}table{font-size:12px}td,th{padding:6px}}</style>
<main><div class="eyebrow">LEDGER CHECK · 本地核对结果</div><h1>${title}</h1><div class="summary"><span>一致 ${report.summary.equal}</span><span>差异 ${report.summary.different}</span><span>无法核对 ${report.summary.unverifiable}</span><span>文件错误 ${report.summary.fileErrors}</span></div>
<p>${escapeHtml(report.scope)}</p><small>引擎 ${escapeHtml(report.engineVersion)} · 方案 SHA-256 ${escapeHtml(report.planHash)}</small>
${report.fileErrors.length ? `<h2>文件错误</h2><ul>${report.fileErrors.map((e) => `<li>${escapeHtml(e.name)}：${escapeHtml(e.message)}</li>`).join('')}</ul>` : ''}
<table><thead><tr><th>结果 / 文档</th><th>检查项</th><th>台账依据</th><th>Word 内容</th><th>说明</th></tr></thead><tbody>${report.findings.map((f) => `<tr><td class="${f.status}">${statuses[f.status]}<small>${escapeHtml(f.document || '台账')}</small></td><td>${escapeHtml(f.field)}</td><td>${evidence(f.expected)}</td><td>${evidence(f.actual)}</td><td>${escapeHtml(f.message)}${f.notices?.map((n) => `<small>${escapeHtml(n)}</small>`).join('') ?? ''}</td></tr>`).join('')}</tbody></table>
<details><summary>来源文件与核对方案</summary><pre>${escapeHtml(JSON.stringify({ files: report.files, plan: report.plan }, null, 2))}</pre></details><small>本报告绑定上述文件内容与方案；文件修改后请重新导入核对。演示样本及其结果不代表真实业务资料的验证。</small></main></html>`;
}

export function wordReportHtml(report: WordReport): string {
  const labels: Record<WordFieldKind, string> = { projectName: '项目名称', softwareVersion: '软件版本', documentVersion: '文档版本' };
  const files = new Map(report.files.map((file, index) => [file.id, { ...file, ordinal: index + 1 }]));
  const groups = report.differences.map((group) => `<section><h2>${escapeHtml(labels[group.category])}${group.subject ? ` · ${escapeHtml(group.subject)}` : ''}</h2><table><thead><tr><th>取值</th><th>对照组 / 阶段</th><th>来源</th><th>人工复核提示</th></tr></thead><tbody>${group.records.map((record) => {
    const file = files.get(record.documentId);
    return `<tr><td>${escapeHtml(record.value)}</td><td>${escapeHtml(record.subject || '项目名称') }<small>${escapeHtml(record.stage || '阶段未标注')}</small></td><td>${escapeHtml(record.file)}（材料 ${file?.ordinal ?? '？'}）<small>${escapeHtml(record.location)}</small></td><td>${record.notices.map((notice) => `<small>${escapeHtml(notice)}</small>`).join('')}</td></tr>`;
  }).join('')}</tbody></table></section>`).join('');
  const all = report.records.map((record) => {
    const file = files.get(record.documentId);
    return `<tr><td>${escapeHtml(labels[record.category])}</td><td>${escapeHtml(record.value)}</td><td>${escapeHtml(record.subject || '未指定')}</td><td>${escapeHtml(record.stage || '未标注')}</td><td>${escapeHtml(record.file)}（材料 ${file?.ordinal ?? '？'}）<small>${escapeHtml(record.location)}</small></td></tr>`;
  }).join('');
  return `<!doctype html><html lang="zh-CN"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Word 文档差异清单</title>
<style>body{font:15px/1.65 system-ui,sans-serif;color:#162c3e;background:#f3f6f8;margin:0}main{max-width:1240px;margin:40px auto;background:white;padding:36px;border-radius:18px}h1{margin:4px 0 16px}small{display:block;color:#64748b;overflow-wrap:anywhere}.summary{padding:18px;background:#eef7f5;border-radius:12px}table{width:100%;border-collapse:collapse;margin:24px 0}td,th{text-align:left;vertical-align:top;border-bottom:1px solid #dfe6eb;padding:12px;overflow-wrap:anywhere;white-space:pre-wrap}th{background:#f3f6f8}details{margin-top:24px}pre{white-space:pre-wrap;word-break:break-word;font-size:12px}@media print{body{background:white}main{padding:0;margin:0}tr{break-inside:avoid}}@media(max-width:700px){main{margin:0;padding:16px}table{font-size:12px}td,th{padding:6px}}</style>
<main><div>WORD 文档互核 · 人工判断</div><h1>所选字段的差异清单</h1><p class="summary">发现 ${report.differences.length} 组不同取值。以下仅展示用户选中的内容，不判断对错、阶段顺序或版本有效性。</p>
<p>软件版本和文档版本只在用户明确填写了相同对照组时并列比较；项目名称按所选值展示。未选择的字段、未分组或只有一条记录的字段不会被当作已核对通过。</p>
${report.errors.length ? `<h2>读取失败</h2><ul>${report.errors.map((e) => `<li>${escapeHtml(e.name)}（材料 ${files.get(e.id)?.ordinal ?? '？'}）：${escapeHtml(e.message)}</li>`).join('')}</ul>` : ''}
${groups || '<p>已选取的对照组中没有发现不同取值；请结合已选范围和原文自行判断。</p>'}
<details><summary>全部已选内容（${report.records.length} 项）</summary><table><thead><tr><th>字段</th><th>取值</th><th>对照组</th><th>阶段</th><th>来源</th></tr></thead><tbody>${all}</tbody></table></details>
<details><summary>每份文档未选择的字段</summary><ul>${report.coverage.map((item) => `<li>${escapeHtml(item.file)}（材料 ${files.get(item.documentId)?.ordinal ?? '？'}）：${item.unselected.length ? item.unselected.map((kind) => escapeHtml(labels[kind])).join('、') : '三个字段均已选择'}</li>`).join('')}</ul></details>
<details><summary>未参与分组比较的内容</summary><ul>${report.uncompared.map((record) => `<li>${escapeHtml(record.file)}（材料 ${files.get(record.documentId)?.ordinal ?? '？'}）· ${escapeHtml(labels[record.category])}：${escapeHtml(record.value)}（未指定组名或组内只有一条记录）</li>`).join('') || '<li>没有此类记录。</li>'}</ul></details>
<details><summary>来源文件与校验信息</summary><pre>${escapeHtml(JSON.stringify({ files: report.files, selectionHash: report.selectionHash, engineVersion: report.engineVersion }, null, 2))}</pre></details>
<small>旧版 .doc 的文本、表格和修订信息可能不完整；请对照原文复核。文件修改后需重新选择并读取。</small></main></html>`;
}
