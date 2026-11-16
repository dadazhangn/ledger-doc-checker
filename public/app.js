const $ = (id) => document.getElementById(id);
const state = { token: '', batch: null, batchId: '', plan: null, detailDraft: null, busy: false, job: '', resultJob: '', report: null, stale: false, filter: 'issues', search: '' };
const cellText = (cell) => cell?.value?.kind === 'formula' ? `=${cell.value.formula}` : cell?.value?.text ?? '';
function node(tag, text, className) { const value = document.createElement(tag); if (text !== undefined) value.textContent = text; if (className) value.className = className; return value; }
function showError(error) { $('error').textContent = error instanceof Error ? error.message : String(error); $('error').hidden = false; }
function clearError() { $('error').hidden = true; }
function notice(text) { $('notice').textContent = text; $('notice').hidden = !text; }
async function api(path, options = {}) {
  const response = await fetch(path, { ...options, headers: { 'x-session-token': state.token, ...(options.body ? { 'Content-Type': 'application/json' } : {}), ...options.headers } });
  if (!response.ok) { let text = '本地服务请求失败。'; try { text = (await response.json()).error || text; } catch {} throw new Error(text); }
  return response;
}
const post = async (path, body) => (await api(path, { method: 'POST', body: JSON.stringify(body) })).json();
function updateButtons() {
  $('inspect').disabled = state.busy || !$('ledger-file').files.length || !$('word-files').files.length;
  $('run-check').disabled = state.busy || !state.batch?.ledger;
  $('download-html').disabled = $('download-json').disabled = state.busy || state.stale || !state.resultJob;
}
function setBusy(busy, title = '') {
  state.busy = busy;
  document.querySelectorAll('button,input,select').forEach((element) => { element.disabled = busy; });
  if (!busy) document.querySelectorAll('input[data-number-only]').forEach((element) => { element.disabled = element.dataset.numberOnly !== 'true'; });
  $('cancel').disabled = !busy;
  $('progress').hidden = !busy;
  $('progress-title').textContent = title;
  updateButtons();
}
function markStale() {
  if (state.report) { state.stale = true; $('stale').hidden = false; }
  updateButtons();
}
async function runJob(path, body, title) {
  clearError(); setBusy(true, title);
  try {
    const { id } = await post(path, body);
    state.job = id;
    while (true) {
      const job = await (await api(`/api/jobs/${id}`)).json();
      $('progress-text').textContent = job.message;
      if (job.status === 'failed') throw new Error(job.error);
      if (job.status === 'cancelled') throw new Error('任务已取消，未生成新的核对结论。');
      if (job.status === 'completed') return { id, result: job.result };
      await new Promise((resolve) => setTimeout(resolve, 180));
    }
  } finally { state.job = ''; setBusy(false); }
}
function columnName(number) { let name = ''; while (number) { number--; name = String.fromCharCode(65 + number % 26) + name; number = Math.floor(number / 26); } return name; }
const sheet = (name) => state.batch?.ledger?.sheets.find((s) => s.name === name);
function columns(name, row) { return (sheet(name)?.cells ?? []).filter((c) => c.row === Number(row) && cellText(c).trim()).sort((a, b) => a.column - b.column).map((c) => ({ value: columnName(c.column), text: `${columnName(c.column)} · ${cellText(c)}` })); }
function fillSelect(element, options, current, placeholder = '请选择') {
  element.replaceChildren();
  const empty = node('option', placeholder); empty.value = ''; element.append(empty);
  for (const item of options) { const option = node('option', typeof item === 'string' ? item : item.text); option.value = typeof item === 'string' ? item : item.value; element.append(option); }
  if (current && !options.some((o) => (typeof o === 'string' ? o : o.value) === current)) { const missing = node('option', `已失效：${current}`); missing.value = current; missing.disabled = true; element.append(missing); }
  element.value = current ?? '';
}
function basePlan() {
  return { version: 1, ledger: { sheet: state.batch.ledger?.sheets[0]?.name ?? '', headerRow: 1, keyColumn: '' }, identityLabel: '', bindings: [], fields: [], details: null, allowCachedFormula: false };
}
function renderPreview() {
  const target = $('ledger-preview'); target.replaceChildren();
  const source = sheet(state.plan.ledger.sheet); if (!source) return;
  const headerRow = state.plan.ledger.headerRow;
  const selectedColumns = columns(source.name, headerRow).slice(0, 8);
  const table = node('table', undefined, 'data-table');
  const head = node('thead'); const headRow = node('tr'); headRow.append(node('th', '行'));
  selectedColumns.forEach((c) => headRow.append(node('th', c.text))); head.append(headRow); table.append(head);
  const body = node('tbody');
  const rows = [...new Set(source.cells.filter((c) => c.row > headerRow).map((c) => c.row))].slice(0, 4);
  for (const row of rows) { const tr = node('tr'); tr.append(node('td', String(row), 'row-index')); for (const col of selectedColumns) tr.append(node('td', cellText(source.cells.find((c) => c.row === row && columnName(c.column) === col.value)))); body.append(tr); }
  table.append(body); target.append(table);
}
function renderBindings() {
  $('bindings').replaceChildren();
  const source = sheet(state.plan.ledger.sheet);
  const keys = [...new Set((source?.cells ?? []).filter((c) => c.row > state.plan.ledger.headerRow && columnName(c.column) === state.plan.ledger.keyColumn && c.value.kind !== 'formula' && cellText(c).trim()).map(cellText))];
  for (const doc of state.batch.documents) {
    const row = node('div', undefined, 'binding-row');
    const name = node('div', doc.name, 'binding-name');
    const matches = doc.labels.filter((l) => l.label === state.plan.identityLabel.trim().replace(/[：:]$/, '').trim());
    name.append(node('small', !state.plan.identityLabel ? '请先指定 Word 对象字段标签' : matches.length === 1 ? `Word 中读到：${matches[0].text || '（空）'}` : matches.length ? `命中 ${matches.length} 处，存在歧义` : '未找到这个字段标签'));
    const choose = node('select'); choose.setAttribute('aria-label', `${doc.name} 对应的台账对象`);
    fillSelect(choose, keys, state.plan.bindings.find((b) => b.document === doc.name)?.key);
    choose.addEventListener('change', () => { state.plan.bindings = state.plan.bindings.filter((b) => b.document !== doc.name); if (choose.value) state.plan.bindings.push({ document: doc.name, key: choose.value }); markStale(); });
    row.append(name, node('span', '→'), choose); $('bindings').append(row);
  }
}
function control(label, element) { const wrapper = node('label', label); wrapper.append(element); return wrapper; }
function renderRuleList(container, rules, detail) {
  container.replaceChildren();
  const source = detail ? state.plan.details : state.plan.ledger;
  rules.forEach((rule, index) => {
    const row = node('div', undefined, 'rule-row');
    const sourceColumn = node('select'); fillSelect(sourceColumn, columns(detail ? source.sheet : source.sheet, detail ? source.headerRow : source.headerRow), rule.sourceColumn);
    sourceColumn.addEventListener('change', () => { rule.sourceColumn = sourceColumn.value; markStale(); });
    const target = node('input'); target.value = detail ? rule.targetHeader : rule.targetLabel; target.placeholder = detail ? 'Word 表头名称' : 'Word 字段标签'; if (!detail) target.setAttribute('list', 'word-labels');
    target.addEventListener('input', () => { rule[detail ? 'targetHeader' : 'targetLabel'] = target.value; markStale(); });
    const comparison = node('select'); fillSelect(comparison, [{ value: 'exact', text: '原文精确' }, { value: 'trim', text: '去首尾空白' }, { value: 'date', text: '日期' }, { value: 'number', text: '数字' }], rule.comparison);
    comparison.addEventListener('change', () => { rule.comparison = comparison.value; markStale(); renderRuleList(container, rules, detail); });
    const tolerance = node('input'); tolerance.type = 'number'; tolerance.min = '0'; tolerance.step = 'any'; tolerance.value = rule.tolerance ?? 0; tolerance.disabled = rule.comparison !== 'number'; tolerance.dataset.numberOnly = String(rule.comparison === 'number');
    tolerance.addEventListener('input', () => { rule.tolerance = Number(tolerance.value); markStale(); });
    const required = node('input'); required.type = 'checkbox'; required.checked = rule.required; required.addEventListener('change', () => { rule.required = required.checked; markStale(); });
    const remove = node('button', '×', 'remove-rule'); remove.setAttribute('aria-label', `移除${detail ? '明细列' : '固定字段'} ${index + 1}`);
    remove.addEventListener('click', () => { rules.splice(index, 1); markStale(); renderRuleList(container, rules, detail); });
    row.append(control('Excel 来源列', sourceColumn), control(detail ? 'Word 表头' : 'Word 字段', target), control('比较方式', comparison), control('容差', tolerance), control('必填', required), remove);
    container.append(row);
  });
  if (!rules.length) container.append(node('p', detail ? '添加需要比较的明细列。' : '尚未配置固定字段。可以添加字段，或启用明细核对。', 'help-text'));
}
function renderDetails() {
  const d = state.plan.details;
  $('details-enabled').checked = Boolean(d); $('detail-settings').hidden = !d;
  if (!d) return;
  fillSelect($('detail-sheet'), state.batch.ledger.sheets.map((s) => s.name), d.sheet);
  $('detail-header').value = d.headerRow;
  fillSelect($('detail-object'), columns(d.sheet, d.headerRow), d.objectColumn);
  fillSelect($('detail-key'), columns(d.sheet, d.headerRow), d.keyColumn);
  $('detail-target-key').value = d.targetKeyHeader;
  renderRuleList($('detail-columns'), d.columns, true);
}
function renderRules() {
  if (!state.batch?.ledger) { $('rules').hidden = true; return; }
  $('rules').hidden = false; $('empty-state').hidden = true;
  fillSelect($('object-sheet'), state.batch.ledger.sheets.map((s) => s.name), state.plan.ledger.sheet);
  $('object-header').value = state.plan.ledger.headerRow;
  fillSelect($('object-key'), columns(state.plan.ledger.sheet, state.plan.ledger.headerRow), state.plan.ledger.keyColumn);
  $('identity-label').value = state.plan.identityLabel;
  $('word-labels').replaceChildren(...[...new Set(state.batch.documents.flatMap((d) => d.labels.map((l) => l.label)))].map((label) => { const option = node('option'); option.value = label; return option; }));
  $('cached-formula').checked = state.plan.allowCachedFormula;
  renderPreview(); renderBindings(); renderRuleList($('field-rules'), state.plan.fields, false); renderDetails(); updateButtons();
}
function showBatch(result, plan) {
  state.batch = result.batch; state.batchId = result.batchId; state.detailDraft = null;
  state.plan = plan ? structuredClone(plan) : basePlan();
  const problems = $('file-problems'); problems.replaceChildren(); problems.hidden = !state.batch.errors.length;
  state.batch.errors.forEach((e) => problems.append(node('div', `${e.name}：${e.message}`)));
  renderRules();
}
async function fileData(file) {
  if (file.size > 12 * 1024 * 1024) throw new Error(`${file.name} 超过 12 MB。`);
  return new Promise((resolve, reject) => { const reader = new FileReader(); reader.onerror = () => reject(new Error(`无法读取 ${file.name}`)); reader.onload = () => resolve({ name: file.name, data: String(reader.result).split(',')[1] }); reader.readAsDataURL(file); });
}
function selectionChanged() {
  markStale(); state.batchId = ''; state.batch = null; $('rules').hidden = true; notice('');
  clearError(); $('file-problems').hidden = true; $('file-problems').replaceChildren();
  $('ledger-name').textContent = $('ledger-file').files[0]?.name ?? '点击选择 .xlsx 文件';
  const words = [...$('word-files').files]; $('word-names').textContent = words.length ? `${words.length} 份文档 · ${words.map((f) => f.name).join('、')}` : '点击选择 1–30 份 .docx 文件'; updateButtons();
}
$('ledger-file').addEventListener('change', selectionChanged); $('word-files').addEventListener('change', selectionChanged);
$('inspect').addEventListener('click', async () => {
  try {
    const chosen = [...$('ledger-file').files, ...$('word-files').files];
    if (chosen.reduce((sum, file) => sum + file.size, 0) > 48 * 1024 * 1024) throw new Error('文件总大小超过 48 MB。');
    setBusy(true, '准备所选资料');
    const files = await Promise.all(chosen.map(fileData));
    const job = await runJob('/api/inspect', { files }, '读取资料'); showBatch(job.result); $('rules').scrollIntoView({ behavior: 'smooth' });
  } catch (error) { setBusy(false); showError(error); }
});
$('load-demo').addEventListener('click', async () => {
  try {
    markStale(); clearError(); setBusy(true, '准备演示资料');
    const demo = await (await api('/api/demo')).json();
    const job = await runJob('/api/inspect', { files: demo.files }, '读取虚构演示资料');
    showBatch(job.result, demo.plan); $('ledger-name').textContent = '演示台账.xlsx（虚构）'; $('word-names').textContent = '演示说明-甲.docx、演示说明-乙.docx（虚构）';
    notice(demo.notice); $('rules').scrollIntoView({ behavior: 'smooth' });
  } catch (error) { setBusy(false); showError(error); }
});
$('object-sheet').addEventListener('change', () => { state.plan.ledger.sheet = $('object-sheet').value; state.plan.ledger.keyColumn = ''; state.plan.bindings = []; state.plan.fields = []; markStale(); renderRules(); });
$('object-header').addEventListener('change', () => { state.plan.ledger.headerRow = Number($('object-header').value); state.plan.ledger.keyColumn = ''; state.plan.bindings = []; state.plan.fields = []; markStale(); renderRules(); });
$('object-key').addEventListener('change', () => { state.plan.ledger.keyColumn = $('object-key').value; state.plan.bindings = []; markStale(); renderBindings(); });
$('identity-label').addEventListener('input', () => { state.plan.identityLabel = $('identity-label').value; markStale(); renderBindings(); });
$('cached-formula').addEventListener('change', () => { state.plan.allowCachedFormula = $('cached-formula').checked; markStale(); });
$('add-field').addEventListener('click', () => { state.plan.fields.push({ sourceColumn: '', targetLabel: '', comparison: 'exact', tolerance: 0, required: true }); markStale(); renderRuleList($('field-rules'), state.plan.fields, false); });
$('details-enabled').addEventListener('change', () => {
  if ($('details-enabled').checked) state.plan.details = state.detailDraft ?? { sheet: state.batch.ledger.sheets[1]?.name ?? state.batch.ledger.sheets[0].name, headerRow: 1, objectColumn: '', keyColumn: '', targetKeyHeader: '', columns: [] };
  else { state.detailDraft = state.plan.details; state.plan.details = null; }
  markStale(); renderDetails();
});
for (const id of ['detail-sheet', 'detail-header', 'detail-object', 'detail-key', 'detail-target-key']) {
  $(id).addEventListener('change', () => {
    const d = state.plan.details; if (!d) return;
    const property = { 'detail-sheet': 'sheet', 'detail-header': 'headerRow', 'detail-object': 'objectColumn', 'detail-key': 'keyColumn', 'detail-target-key': 'targetKeyHeader' }[id];
    d[property] = id === 'detail-header' ? Number($(id).value) : $(id).value;
    if (['detail-sheet', 'detail-header'].includes(id)) { d.objectColumn = ''; d.keyColumn = ''; d.columns = []; }
    markStale(); renderDetails();
  });
}
$('add-detail-column').addEventListener('click', () => { state.plan.details.columns.push({ sourceColumn: '', targetHeader: '', comparison: 'exact', tolerance: 0, required: true }); markStale(); renderDetails(); });
$('cancel').addEventListener('click', async () => { try { if (state.job) await post(`/api/jobs/${state.job}/cancel`, {}); } catch (error) { showError(error); } });
$('run-check').addEventListener('click', async () => {
  try {
    if (!state.plan.identityLabel || !state.plan.ledger.keyColumn) throw new Error('请先指定台账对象关联列和 Word 对象字段标签。');
    if (state.batch.documents.some((d) => !state.plan.bindings.find((b) => b.document === d.name))) throw new Error('请为每份 Word 明确选择对应的台账对象。');
    const plan = structuredClone(state.plan);
    const job = await runJob('/api/check', { batchId: state.batchId, plan }, '逐项核对');
    state.report = job.result; state.resultJob = job.id; state.stale = false; $('stale').hidden = true;
    renderReport(); updateButtons(); $('results').scrollIntoView({ behavior: 'smooth' });
  } catch (error) { markStale(); showError(error); }
});
function evidenceCell(evidence) {
  const td = node('td'); if (!evidence) { td.textContent = '—'; return td; }
  td.append(node('span', evidence.value || '（空）', 'value'), node('small', `${evidence.file} · ${evidence.location}`)); return td;
}
function renderReport() {
  $('results').hidden = false; $('empty-state').hidden = true;
  const s = state.report.summary;
  $('result-title').textContent = s.verdict === 'incomplete' ? '检查未完成，仍有待确认项' : s.verdict === 'different' ? `发现 ${s.different} 处差异` : '配置范围内一致';
  $('result-subtitle').textContent = `${s.documents} 份 Word · ${s.equal + s.different + s.unverifiable} 个检查项 · ${s.complete ? '配置范围已完成检查' : '无法核对与文件错误均未计为通过'}`;
  for (const [id, value] of [['count-equal', s.equal], ['count-different', s.different], ['count-unverifiable', s.unverifiable], ['count-errors', s.fileErrors]]) $(id).textContent = value;
  $('report-provenance').textContent = `${state.report.scope} 方案指纹：${state.report.planHash.slice(0, 16)}… · 文件与方案修改后需重新核对。`;
  renderFindings();
}
function renderFindings() {
  const target = $('findings'); target.replaceChildren();
  if (!state.report) return;
  for (const error of state.report.fileErrors) target.append(node('div', `${error.name}：${error.message}`, 'error-banner'));
  const filtered = state.report.findings.filter((f) => (state.filter === 'all' || (state.filter === 'issues' ? f.status !== 'equal' : f.status === state.filter)) && `${f.document} ${f.field} ${f.message}`.toLowerCase().includes(state.search.toLowerCase()));
  if (!filtered.length) { target.append(node('div', '当前筛选下没有检查项。', 'no-findings')); return; }
  const table = node('table', undefined, 'data-table'); const head = node('thead'); const tr = node('tr');
  for (const name of ['结果 / 文档', '检查项', '台账依据', 'Word 内容', '说明']) tr.append(node('th', name)); head.append(tr); table.append(head);
  const body = node('tbody');
  for (const f of filtered) {
    const row = node('tr'); const status = node('td'); status.append(node('span', { equal: '一致', different: '差异', unverifiable: '无法核对' }[f.status], `badge ${f.status}`), node('small', f.document || '台账'));
    const message = node('td', f.message); f.notices?.forEach((n) => message.append(node('small', n)));
    row.append(status, node('td', f.field), evidenceCell(f.expected), evidenceCell(f.actual), message); body.append(row);
  }
  table.append(body); target.append(table);
}
$('result-filters').addEventListener('click', (event) => { const button = event.target.closest('button[data-filter]'); if (!button) return; state.filter = button.dataset.filter; $('result-filters').querySelectorAll('button').forEach((b) => b.classList.toggle('selected', b === button)); renderFindings(); });
$('result-search').addEventListener('input', () => { state.search = $('result-search').value; renderFindings(); });
function saveBlob(blob, name) { const url = URL.createObjectURL(blob); const a = node('a'); a.href = url; a.download = name; document.body.append(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 1000); }
$('export-plan').addEventListener('click', () => { if (state.plan) saveBlob(new Blob([JSON.stringify(state.plan, null, 2)], { type: 'application/json' }), 'ledger-check-plan.json'); });
$('import-plan').addEventListener('change', async () => {
  try { const file = $('import-plan').files[0]; if (!file) return; if (file.size > 100000) throw new Error('方案文件超过 100 KB。'); const raw = JSON.parse(await file.text()); const { plan } = await post('/api/plan/validate', { plan: raw }); state.plan = plan; markStale(); renderRules(); notice('方案已载入，请核对本批文件和对应关系。'); } catch (error) { showError(error); } finally { $('import-plan').value = ''; }
});
for (const format of ['html', 'json']) $(`download-${format}`).addEventListener('click', async () => { try { const response = await api(`/api/jobs/${state.resultJob}/report?format=${format}`); saveBlob(await response.blob(), `台账核对报告.${format}`); } catch (error) { showError(error); } });
try { const session = await (await fetch('/api/session')).json(); if (!session.token) throw new Error('本地服务初始化失败。'); state.token = session.token; } catch (error) { showError(error); $('load-demo').disabled = true; }
updateButtons();
