const $ = (id) => document.getElementById(id);
const labels = { projectName: '项目名称', softwareVersion: '软件版本', documentVersion: '文档版本' };
const kinds = Object.keys(labels);
const blank = (category) => ({ category, segmentId: '', value: '', subject: '', stage: '' });
function element(tag, text, className) { const item = document.createElement(tag); if (text !== undefined) item.textContent = text; if (className) item.className = className; return item; }
function control(label, input) { const parent = element('label', label); parent.append(input); return parent; }

export function initWordMode({ api, post, runJob, fileData, saveBlob, getBusy, setBusy }) {
  const state = { files: [], batch: null, batchId: '', entries: new Map(), report: null, resultJob: '', stale: false, preparing: false, cancelPreparation: false };
  const showError = (error) => { $('word-error').textContent = error instanceof Error ? error.message : String(error); $('word-error').hidden = false; };
  const clearError = () => { $('word-error').hidden = true; };
  const notice = (text) => { $('word-notice').textContent = text; $('word-notice').hidden = !text; };
  function buttons() {
    $('word-inspect').disabled = getBusy() || state.files.length < 2 || state.files.length > 30;
    $('word-compare').disabled = getBusy() || !state.batch?.documents.length;
    $('word-download-html').disabled = $('word-download-json').disabled = getBusy() || state.stale || !state.resultJob;
  }
  function changed() {
    if (state.report) { state.stale = true; $('word-stale').hidden = false; }
    buttons();
  }
  function invalidateBatch() { state.batch = null; state.batchId = ''; $('word-rules').hidden = true; changed(); }
  function renderFiles() {
    const target = $('word-selected-files'); target.replaceChildren();
    if (!state.files.length) target.append(element('span', '尚未添加文件；也可以先体验虚构演示。', 'help-text'));
    state.files.forEach((file, index) => {
      const line = element('div', undefined, 'word-file-line');
      const remove = element('button', '×'); remove.type = 'button'; remove.setAttribute('aria-label', `移除材料 ${index + 1} ${file.name}`);
      remove.addEventListener('click', () => { state.files.splice(index, 1); invalidateBatch(); renderFiles(); });
      line.append(element('span', `材料 ${index + 1} · ${file.name}`), remove); target.append(line);
    });
    buttons();
  }
  function renderOptions(select, doc, entry, query) {
    const lowered = query.toLocaleLowerCase();
    const matches = doc.segments.filter((segment, index) => {
      if (!lowered) return true;
      const surrounding = [doc.segments[index - 1]?.text, segment.text, doc.segments[index + 1]?.text].join(' ');
      return `${segment.location} ${surrounding}`.toLocaleLowerCase().includes(lowered);
    });
    select.replaceChildren();
    const placeholder = element('option', matches.length > 80 ? `显示前 80 项，共 ${matches.length} 项；可搜索缩小范围` : '请选择原文片段'); placeholder.value = ''; select.append(placeholder);
    const selected = doc.segments.find((segment) => segment.id === entry.segmentId);
    const shown = matches.slice(0, 80);
    if (selected && !shown.some((segment) => segment.id === selected.id)) shown.unshift(selected);
    for (const segment of shown) {
      const preview = segment.text.replace(/\s+/g, ' ');
      const option = element('option', `${segment.location} · ${preview.slice(0, 85)}${preview.length > 85 ? '…' : ''}`);
      option.value = segment.id; select.append(option);
    }
    select.value = entry.segmentId || '';
  }
  function sourceNote(target, doc, entry) {
    target.replaceChildren();
    const segment = doc.segments.find((item) => item.id === entry.segmentId);
    if (!segment) { target.textContent = '先搜索并选择来源片段；选择后可把对照值缩小为片段中的原文。'; return; }
    const excerpt = segment.text.length > 300 ? `${segment.text.slice(0, 300)}…（仅预览前 300 字）` : segment.text;
    target.append(element('span', `来源：${segment.location}\n片段：${excerpt}`));
    if (segment.notices.length) target.append(element('span', `\n${segment.notices.join(' ')}`, 'word-source-notice'));
  }
  function renderEntry(doc, entry, index, rerender) {
    const wrapper = element('div', undefined, 'word-entry');
    const heading = element('div', undefined, 'word-entry-header');
    const remove = element('button', '×', 'word-remove'); remove.type = 'button'; remove.setAttribute('aria-label', `移除此文档的第 ${index + 1} 项选择`);
    remove.addEventListener('click', () => { state.entries.get(doc.id).splice(index, 1); changed(); rerender(); });
    heading.append(element('span', `字段选择 ${index + 1}`), remove); wrapper.append(heading);
    const grid = element('div', undefined, 'word-entry-grid');
    const category = element('select'); category.setAttribute('aria-label', `第 ${index + 1} 项字段类型`);
    for (const kind of kinds) { const option = element('option', labels[kind]); option.value = kind; category.append(option); } category.value = entry.category;
    category.addEventListener('change', () => { entry.category = category.value; changed(); rerender(); });
    const search = element('input'); search.type = 'search'; search.placeholder = '搜索关键词，也会查找前后片段'; search.setAttribute('aria-label', `搜索第 ${index + 1} 项来源片段`);
    const source = element('select'); source.setAttribute('aria-label', `第 ${index + 1} 项来源片段`);
    renderOptions(source, doc, entry, '');
    search.addEventListener('input', () => renderOptions(source, doc, entry, search.value));
    const value = element('textarea'); value.rows = 2; value.maxLength = 200; value.value = entry.value;
    value.placeholder = '选择片段后，可截取其中需要对照的原文'; value.setAttribute('aria-label', `第 ${index + 1} 项对照值`);
    const group = element('input'); group.type = 'text'; group.maxLength = 120; group.value = entry.subject; group.placeholder = '同一对象或同类文档填写完全相同的名称'; group.setAttribute('aria-label', `第 ${index + 1} 项对照组`);
    const stage = element('input'); stage.type = 'text'; stage.maxLength = 120; stage.value = entry.stage; stage.placeholder = '手动填写，可留空'; stage.setAttribute('aria-label', `第 ${index + 1} 项阶段`);
    const note = element('div', undefined, 'word-source-note'); sourceNote(note, doc, entry);
    source.addEventListener('change', () => {
      entry.segmentId = source.value;
      const segment = doc.segments.find((item) => item.id === source.value);
      entry.value = segment && segment.text.length <= 200 ? segment.text : '';
      value.value = entry.value; sourceNote(note, doc, entry); changed();
    });
    value.addEventListener('input', () => { entry.value = value.value; changed(); });
    group.addEventListener('input', () => { entry.subject = group.value; changed(); });
    stage.addEventListener('input', () => { entry.stage = stage.value; changed(); });
    grid.append(control('字段类型', category), control('搜索来源片段', search), control('来源片段', source), control('对照值（必须取自所选片段）', value));
    if (entry.category !== 'projectName') grid.append(control('对照组（留空则不参与分组）', group));
    grid.append(control('时期／阶段', stage)); wrapper.append(grid, note);
    return wrapper;
  }
  function renderDocument(doc, index) {
    const details = element('details', undefined, 'word-document');
    const summary = element('summary');
    summary.append(element('strong', `材料 ${index + 1} · ${doc.name}`), element('small', `${doc.format.toUpperCase()} · ${doc.segments.length} 个片段`));
    const content = element('div', undefined, 'word-doc-content');
    const warning = element('div', doc.notices.join(' '), 'word-doc-warning');
    const rows = element('div');
    const add = element('button', '＋ 添加字段', 'button subtle'); add.type = 'button';
    const renderRows = () => { rows.replaceChildren(...state.entries.get(doc.id).map((entry, i) => renderEntry(doc, entry, i, renderRows))); };
    add.addEventListener('click', () => { state.entries.get(doc.id).push(blank('softwareVersion')); changed(); renderRows(); });
    details.addEventListener('toggle', () => { if (details.open) renderRows(); else rows.replaceChildren(); });
    content.append(warning, rows, add); details.append(summary, content); return details;
  }
  function renderBatch() {
    $('word-rules').hidden = false; $('word-empty-state').hidden = true;
    $('word-document-cards').replaceChildren(...state.batch.documents.map((doc, index) => {
      const ordinal = state.batch.files.findIndex((f) => f.id === doc.id);
      return renderDocument(doc, ordinal >= 0 ? ordinal : index);
    }));
    $('word-file-errors').replaceChildren(...state.batch.errors.map((error) => element('div', `${error.name}：${error.message}。该文档未参与比较。`, 'error-banner word-file-error')));
    buttons();
  }
  function demoEntries(doc, index) {
    const stage = index === 0 ? '前期' : '后期';
    const specs = [
      ['projectName', 'body:t:1:r:2:c:2:0', '演示项目', ''],
      ['softwareVersion', 'body:t:1:r:3:c:2:0', index === 0 ? '1.0' : '2.0', '演示软件'],
      ['documentVersion', 'body:t:1:r:4:c:2:0', index === 0 ? 'A' : 'B', '演示文档'],
    ];
    return specs.map(([category, segmentId, value, subject]) => {
      const segment = doc.segments.find((item) => item.id === segmentId);
      return segment?.text.includes(value) ? { category, segmentId, value, subject, stage } : blank(category);
    });
  }
  function showBatch(result, demo = false) {
    state.batch = result.batch; state.batchId = result.batchId; state.entries = new Map();
    state.batch.documents.forEach((doc, index) => state.entries.set(doc.id, demo ? demoEntries(doc, index) : kinds.map(blank)));
    renderBatch(); $('word-rules').scrollIntoView({ behavior: 'smooth' });
  }
  $('word-input').addEventListener('change', () => {
    const selected = [...$('word-input').files]; $('word-input').value = '';
    if (!selected.length) return;
    if (state.files.length + selected.length > 30) { showError(new Error('一批最多添加 30 份 Word 文档。')); return; }
    clearError(); notice(''); state.files.push(...selected); invalidateBatch(); renderFiles();
  });
  $('cancel').addEventListener('click', () => { if (state.preparing) state.cancelPreparation = true; });
  $('word-inspect').addEventListener('click', async () => {
    try {
      clearError(); notice(''); state.preparing = true; state.cancelPreparation = false; setBusy(true, '准备所选文档');
      const total = state.files.reduce((sum, file) => sum + (file.size ?? Math.ceil(file.data.length * 0.75)), 0);
      if (total > 48 * 1024 * 1024) throw new Error('文件总大小超过 48 MB。');
      const files = await Promise.all(state.files.map((file) => file.data ? { name: file.name, data: file.data } : fileData(file)));
      state.preparing = false;
      if (state.cancelPreparation) throw new Error('读取已取消，未生成新的核对结果。');
      const job = await runJob('/api/word/inspect', { files }, '读取 Word 文档'); showBatch(job.result);
    } catch (error) { state.preparing = false; setBusy(false); showError(error); }
  });
  $('word-demo').addEventListener('click', async () => {
    try {
      clearError(); changed(); setBusy(true, '准备演示文档');
      const demo = await (await api('/api/word/demo')).json();
      state.files = demo.files; renderFiles();
      const job = await runJob('/api/word/inspect', { files: demo.files }, '读取虚构演示文档');
      showBatch(job.result, true); notice(demo.notice);
    } catch (error) { setBusy(false); showError(error); }
  });
  $('word-compare').addEventListener('click', async () => {
    try {
      clearError();
      const selections = [];
      for (const doc of state.batch.documents) for (const entry of state.entries.get(doc.id)) {
        if (!entry.segmentId && !entry.value && !entry.subject && !entry.stage) continue;
        if (!entry.segmentId || !entry.value.trim()) throw new Error(`${doc.name} 中有尚未选好来源或取值的字段。`);
        selections.push({ documentId: doc.id, ...entry, subject: entry.category === 'projectName' ? '' : entry.subject });
      }
      const job = await runJob('/api/word/compare', { batchId: state.batchId, selections }, '整理不同取值');
      state.report = job.result; state.resultJob = job.id; state.stale = false; $('word-stale').hidden = true;
      renderReport(); buttons(); $('word-results').scrollIntoView({ behavior: 'smooth' });
    } catch (error) { changed(); showError(error); }
  });
  function table(headers, rows) {
    const outer = element('div', undefined, 'table-scroll'); const table = element('table', undefined, 'data-table');
    const head = element('thead'); const tr = element('tr'); headers.forEach((heading) => tr.append(element('th', heading))); head.append(tr); table.append(head);
    const body = element('tbody'); rows.forEach((row) => { const tr = element('tr'); row.forEach((value) => { const td = element('td'); if (value instanceof Node) td.append(value); else td.textContent = value; tr.append(td); }); body.append(tr); });
    table.append(body); outer.append(table); return outer;
  }
  function describeRecord(record) {
    const file = state.batch.files.findIndex((item) => item.id === record.documentId) + 1;
    const cell = element('span', record.file, 'value');
    cell.append(element('small', `材料 ${file} · ${record.location}`));
    return cell;
  }
  function renderReport() {
    const report = state.report;
    $('word-results').hidden = false; $('word-empty-state').hidden = true;
    $('word-result-title').textContent = report.differences.length ? `发现 ${report.differences.length} 组不同取值` : '所选对照组未发现不同取值';
    $('word-differences').replaceChildren();
    if (!report.differences.length) $('word-differences').append(element('div', '仅根据已选内容比较；未选择、未分组或只有一条记录的内容需要人工处理。', 'card word-empty-difference'));
    for (const group of report.differences) {
      const card = element('article', undefined, 'card word-difference');
      card.append(element('h3', `${labels[group.category]}${group.subject ? ` · ${group.subject}` : ''}`), element('p', '以下为不同取值及其来源，阶段与版本含义请人工判断。'));
      card.append(table(['取值', '阶段', '来源', '复核提示'], group.records.map((record) => [record.value, record.stage || '未标注', describeRecord(record), record.notices.join(' ')])));
      $('word-differences').append(card);
    }
    $('word-selected-count').textContent = `（${report.records.length} 项）`;
    $('word-records').replaceChildren(table(['字段', '取值', '对照组', '阶段', '来源'], report.records.map((record) => [labels[record.category], record.value, record.subject || '未指定', record.stage || '未标注', describeRecord(record)])));
    $('word-coverage').replaceChildren(...report.coverage.map((item) => {
      const ordinal = state.batch.files.findIndex((file) => file.id === item.documentId) + 1;
      return element('div', `材料 ${ordinal} · ${item.file}：${item.unselected.length ? `未选择 ${item.unselected.map((kind) => labels[kind]).join('、')}` : '三个字段均已选择'}`, 'word-coverage-line');
    }));
    $('word-uncompared').replaceChildren(...(report.uncompared.length ? report.uncompared.map((record) => element('div', `${record.file} · ${labels[record.category]}：${record.value}（未指定组名或组内只有一条记录）`, 'word-coverage-line')) : [element('div', '没有此类记录。', 'word-coverage-line')]));
    $('word-provenance').textContent = `所有结果由人工判断 · 选择指纹 ${report.selectionHash.slice(0, 16)}… · 旧版 .doc 文本请对照原文复核 · 修改文件后请重新读取。`;
  }
  for (const format of ['html', 'json']) $(`word-download-${format}`).addEventListener('click', async () => {
    try { const response = await api(`/api/jobs/${state.resultJob}/report?format=${format}`); saveBlob(await response.blob(), `Word文档差异清单.${format}`); }
    catch (error) { showError(error); }
  });
  document.addEventListener('ledger-busy-change', buttons);
  renderFiles(); buttons();
}
