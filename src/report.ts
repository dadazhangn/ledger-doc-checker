import type { Report } from './types.js';

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
