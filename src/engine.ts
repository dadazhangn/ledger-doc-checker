import { hash } from './parsers.js';
import type { Batch, CellValue, Comparison, Evidence, Finding, LedgerCell, LedgerSheet, Plan, Report, WordDocument, WordParagraph, WordTable } from './types.js';

export function columnNumber(value: string): number {
  if (!/^[A-Z]{1,3}$/.test(value)) return 0;
  return [...value].reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0);
}
const labelText = (value: string) => value.trim().replace(/[：:]$/, '').trim();
const emptyValue: CellValue = { kind: 'blank', raw: null, text: '' };

export function validatePlan(value: unknown): Plan {
  const bad = (message: string): never => { throw new Error(`核对方案无效：${message}`); };
  const object = (v: unknown): Record<string, unknown> => {
    if (!v || typeof v !== 'object' || Array.isArray(v)) bad('应为 JSON 对象。');
    return v as Record<string, unknown>;
  };
  const text = (v: unknown, name: string) => typeof v === 'string' && v.trim() && v.length <= 180 ? v : bad(`${name}不能为空或过长。`);
  const column = (v: unknown) => { const name = text(v, '列'); return columnNumber(name) > 0 && columnNumber(name) <= 100 ? name : bad('列必须是 A–CV。'); };
  const row = (v: unknown) => Number.isInteger(v) && Number(v) >= 1 && Number(v) <= 10000 ? Number(v) : bad('表头行必须为 1–10000 的整数。');
  const compare = (v: unknown): Comparison => ['exact', 'trim', 'date', 'number'].includes(String(v)) ? v as Comparison : bad('未知的比较方式。');
  const tolerance = (v: unknown) => v === undefined ? 0 : typeof v === 'number' && Number.isFinite(v) && v >= 0 && v <= 1000000 ? v : bad('数字容差必须是非负有限数。');
  const required = (v: unknown) => typeof v === 'boolean' ? v : bad('必须明确指定是否必填。');
  const root = object(value);
  if (root.version !== 1) bad('不支持这个方案版本。');
  const ledger = object(root.ledger);
  if (!Array.isArray(root.bindings) || root.bindings.length > 30) bad('文档对应关系必须是最多 30 项的列表。');
  const bindings = (root.bindings as unknown[]).map((v) => { const x = object(v); return { document: text(x.document, '文档名'), key: text(x.key, '对象值') }; });
  if (new Set(bindings.map((x) => x.document)).size !== bindings.length) bad('同一文档只能关联一次。');
  if (!Array.isArray(root.fields) || root.fields.length > 30) bad('固定字段最多 30 项。');
  const fields = (root.fields as unknown[]).map((v) => { const x = object(v); return {
    sourceColumn: column(x.sourceColumn), targetLabel: text(x.targetLabel, 'Word 字段标签'),
    comparison: compare(x.comparison), tolerance: tolerance(x.tolerance), required: required(x.required),
  }; });
  if (new Set(fields.map((x) => labelText(x.targetLabel))).size !== fields.length) bad('固定字段的目标标签重复。');
  let details: Plan['details'] = null;
  if (root.details !== null && root.details !== undefined) {
    const d = object(root.details);
    if (!Array.isArray(d.columns) || !d.columns.length || d.columns.length > 20) bad('明细比较列需要 1–20 项。');
    const columns = (d.columns as unknown[]).map((v) => { const x = object(v); return {
      sourceColumn: column(x.sourceColumn), targetHeader: text(x.targetHeader, 'Word 明细列名'),
      comparison: compare(x.comparison), tolerance: tolerance(x.tolerance), required: required(x.required),
    }; });
    const keyHeader = text(d.targetKeyHeader, 'Word 明细关联列');
    if (new Set([labelText(keyHeader), ...columns.map((x) => labelText(x.targetHeader))]).size !== columns.length + 1) bad('Word 明细列名存在重复。');
    details = { sheet: text(d.sheet, '明细工作表'), headerRow: row(d.headerRow), objectColumn: column(d.objectColumn),
      keyColumn: column(d.keyColumn), targetKeyHeader: keyHeader, columns };
  }
  if (!fields.length && !details) bad('至少配置一个固定字段或一组明细。');
  if (typeof root.allowCachedFormula !== 'boolean') bad('必须明确指定公式缓存使用方式。');
  return { version: 1, ledger: { sheet: text(ledger.sheet, '对象工作表'), headerRow: row(ledger.headerRow), keyColumn: column(ledger.keyColumn) },
    identityLabel: text(root.identityLabel, 'Word 对象标签'), bindings, fields, details, allowCachedFormula: root.allowCachedFormula as boolean };
}

export function documentLabels(document: WordDocument): { label: string; text: string; location: string; reasons: string[] }[] {
  const result: ReturnType<typeof documentLabels> = [];
  for (const table of document.tables) for (const row of table.rows) {
    if (row.length < 2 || row.length % 2) continue;
    for (let i = 0; i < row.length; i += 2) {
      const label = labelText(row[i].text);
      if (label && label.length <= 100) result.push({ label, text: row[i + 1].text, location: row[i + 1].location,
        reasons: [...new Set([...table.reasons, ...row[i].reasons, ...row[i + 1].reasons])] });
    }
  }
  for (const paragraph of document.paragraphs) {
    const match = paragraph.text.match(/^([^:：\n]{1,100})[:：]([\s\S]*)$/);
    if (match) result.push({ label: labelText(match[1]), text: match[2], location: paragraph.location, reasons: paragraph.reasons });
  }
  return result;
}

type IndexedSheet = { sheet: LedgerSheet; rows: Map<number, Map<number, LedgerCell>> };
function indexSheet(sheet: LedgerSheet): IndexedSheet {
  const rows: IndexedSheet['rows'] = new Map();
  for (const cell of sheet.cells) { if (!rows.has(cell.row)) rows.set(cell.row, new Map()); rows.get(cell.row)!.set(cell.column, cell); }
  return { sheet, rows };
}
function getCell(sheet: IndexedSheet, row: number, column: string) { return sheet.rows.get(row)?.get(columnNumber(column)); }
function sourceEvidence(batch: Batch, sheet: IndexedSheet, row: number, column: string, value?: string): Evidence {
  return { file: batch.ledger!.name, location: `${sheet.sheet.name}!${column}${row}`, value: value ?? getCell(sheet, row, column)?.value.text ?? '' };
}
function readValue(cell: LedgerCell | undefined, allowCached: boolean): { value: CellValue; reason?: string; notices?: string[] } {
  if (!cell) return { value: emptyValue };
  if (cell.merged) return { value: cell.value, reason: '来源是合并单元格，无法确定每条记录的独立取值。' };
  if (cell.value.kind === 'error') return { value: cell.value, reason: '来源单元格是错误值或不支持的类型。' };
  if (cell.value.kind === 'formula') {
    if (!allowCached || !cell.value.cached || cell.value.cached.kind === 'error' || cell.value.cached.kind === 'blank') return { value: cell.value, reason: '来源是公式；本工具不计算公式，当前没有经方案允许使用的有效缓存值。' };
    return { value: cell.value.cached, notices: ['按方案采用 Excel 公式缓存值，未重算且无法验证缓存时效。'] };
  }
  if (cell.value.kind === 'number' && typeof cell.value.raw === 'number' && Number.isInteger(cell.value.raw) && !Number.isSafeInteger(cell.value.raw)) return { value: cell.value, reason: '来源数字超出安全整数范围，无法保证精度。' };
  return { value: cell.value };
}

function dateValue(text: string): string | null {
  const match = text.match(/^(\d{4})(?:-|\/|年)(\d{1,2})(?:-|\/|月)(\d{1,2})日?$/);
  if (!match) return null;
  const [, y, m, d] = match;
  const date = new Date(Date.UTC(Number(y), Number(m) - 1, Number(d)));
  return date.getUTCFullYear() === Number(y) && date.getUTCMonth() + 1 === Number(m) && date.getUTCDate() === Number(d)
    ? `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}` : null;
}
function decimal(text: string): { value: bigint; scale: number } | null {
  if (text.length > 100 || !/^[+-]?\d+(?:\.\d+)?$/.test(text)) return null;
  const negative = text.startsWith('-');
  const [whole, fraction = ''] = text.replace(/^[+-]/, '').split('.');
  return { value: BigInt(whole + fraction) * (negative ? -1n : 1n), scale: fraction.length };
}
export function compareValues(expected: string, actual: string, kind: Comparison, required: boolean, tolerance = 0): Pick<Finding, 'status' | 'message'> {
  if (required && (!expected.trim() || !actual.trim())) return { status: 'different', message: !expected.trim() ? '台账必填值为空。' : 'Word 必填值为空。' };
  if (!expected && !actual && !required) return { status: 'equal', message: '双方均为空，且方案允许为空。' };
  let equal = false;
  if (kind === 'exact') equal = expected === actual;
  else if (kind === 'trim') equal = expected.trim() === actual.trim();
  else if (kind === 'date') {
    const left = dateValue(expected), right = dateValue(actual);
    if (!left || !right) return { status: 'unverifiable', message: '日期无法按明确的年/月/日格式解析。' };
    equal = left === right;
  } else {
    const left = decimal(expected), right = decimal(actual), allowed = decimal(String(tolerance));
    if (!left || !right || !allowed) return { status: 'unverifiable', message: '数字无法按十进制定点格式解析；不自动推断千分位、单位或百分比。' };
    const scale = Math.max(left.scale, right.scale, allowed.scale);
    const delta = left.value * 10n ** BigInt(scale - left.scale) - right.value * 10n ** BigInt(scale - right.scale);
    equal = (delta < 0n ? -delta : delta) <= allowed.value * 10n ** BigInt(scale - allowed.scale);
  }
  return { status: equal ? 'equal' : 'different', message: equal ? '按所选规则一致。' : '按所选规则不一致。' };
}

export function runCheck(batch: Batch, rawPlan: unknown): Report {
  const plan = validatePlan(rawPlan);
  const findings: Finding[] = [];
  const addBlock = (document: string, field: string, message: string, expected?: Evidence, actual?: Evidence) => findings.push({ status: 'unverifiable', kind: 'coverage', document, field, message, expected, actual });
  if (!batch.ledger) throw new Error('Excel 台账未能解析，无法开始核对。');
  const index = new Map(batch.ledger.sheets.map((sheet) => [sheet.name, indexSheet(sheet)]));
  const objectSheet = index.get(plan.ledger.sheet);
  if (!objectSheet) throw new Error('方案指定的对象工作表不存在。');
  const sourceColumns = [plan.ledger.keyColumn, ...plan.fields.map((x) => x.sourceColumn)];
  for (const column of sourceColumns) if (!getCell(objectSheet, plan.ledger.headerRow, column)?.value.text.trim()) throw new Error(`对象工作表 ${column} 列的表头为空或不存在。`);
  const objects = new Map<string, number[]>();
  for (const [row, cells] of objectSheet.rows) {
    if (row <= plan.ledger.headerRow || ![...cells.values()].some((c) => c.value.kind !== 'blank')) continue;
    const source = readValue(getCell(objectSheet, row, plan.ledger.keyColumn), false);
    if (source.reason || !source.value.text.trim()) { addBlock('', '台账对象关联键', source.reason ?? '台账存在关联键为空的数据行。', sourceEvidence(batch, objectSheet, row, plan.ledger.keyColumn)); continue; }
    const key = source.value.text;
    objects.set(key, [...objects.get(key) ?? [], row]);
  }
  for (const [key, rows] of objects) if (rows.length > 1) addBlock('', '台账对象关联键', `对象“${key}”对应多行台账，禁止自动选取。`, sourceEvidence(batch, objectSheet, rows[0], plan.ledger.keyColumn));
  const detailSheet = plan.details ? index.get(plan.details.sheet) : undefined;
  if (plan.details && !detailSheet) throw new Error('方案指定的明细工作表不存在。');
  const estimatedChecks = batch.documents.length * (plan.fields.length + 1 + (detailSheet ? detailSheet.sheet.rowCount * ((plan.details?.columns.length ?? 0) + 1) : 0));
  if (estimatedChecks > 100000) throw new Error('本批潜在检查项超过 100,000 个，请减少文件或明细范围后分批核对。');
  if (detailSheet && plan.details) for (const column of [plan.details.objectColumn, plan.details.keyColumn, ...plan.details.columns.map((x) => x.sourceColumn)]) {
    if (!getCell(detailSheet, plan.details.headerRow, column)?.value.text.trim()) throw new Error(`明细工作表 ${column} 列的表头为空或不存在。`);
  }
  if (detailSheet && plan.details) for (const [row, cells] of detailSheet.rows) {
    if (row <= plan.details.headerRow || ![...cells.values()].some((c) => c.value.kind !== 'blank')) continue;
    const object = readValue(getCell(detailSheet, row, plan.details.objectColumn), false);
    if (object.reason || !objects.has(object.value.text)) addBlock('', '明细所属对象', object.reason ?? '该明细的所属对象为空或未出现在对象台账中。', sourceEvidence(batch, detailSheet, row, plan.details.objectColumn));
  }
  const bindings = new Map(plan.bindings.map((b) => [b.document, b.key]));
  const documentNames = new Set(batch.documents.map((d) => d.name));
  for (const binding of plan.bindings) if (!documentNames.has(binding.document)) addBlock(binding.document, '文档关联', '方案中的文档未包含在本批成功解析的文件中。');
  for (const document of batch.documents) {
    const key = bindings.get(document.name);
    if (!key) { addBlock(document.name, '文档关联', '尚未明确选择对应的台账对象。'); continue; }
    const rows = objects.get(key);
    if (!rows || rows.length !== 1) { addBlock(document.name, '文档关联', '对应的台账对象不存在或不是唯一记录。'); continue; }
    const row = rows[0];
    const labels = documentLabels(document);
    const locate = (label: string): WordParagraph | string => {
      const candidates = labels.filter((item) => item.label === labelText(label));
      if (!candidates.length) return '未找到该字段标签；不自动全文猜测字段值。';
      if (candidates.length > 1) return `字段标签命中 ${candidates.length} 处，需要消除歧义。`;
      return candidates[0];
    };
    const identity = locate(plan.identityLabel);
    if (typeof identity === 'string' || identity.reasons.length) { addBlock(document.name, '文档关联', typeof identity === 'string' ? identity : identity.reasons.join(' ')); continue; }
    const sourceKey = sourceEvidence(batch, objectSheet, row, plan.ledger.keyColumn, key);
    const targetKey = { file: document.name, location: identity.location, value: identity.text };
    if (identity.text !== key) {
      findings.push({ status: 'different', kind: 'binding', document: document.name, field: '文档关联', message: 'Word 对象与所选台账对象不一致。', expected: sourceKey, actual: targetKey });
      addBlock(document.name, '后续检查', '文档关联未成立，已停止该文档的字段和明细核对。');
      continue;
    }
    findings.push({ status: 'equal', kind: 'binding', document: document.name, field: '文档关联', message: '文档与唯一台账对象一致。', expected: sourceKey, actual: targetKey });
    for (const rule of plan.fields) {
      const source = readValue(getCell(objectSheet, row, rule.sourceColumn), plan.allowCachedFormula);
      const expected = sourceEvidence(batch, objectSheet, row, rule.sourceColumn, source.value.text);
      const target = locate(rule.targetLabel);
      if (source.reason || typeof target === 'string' || target.reasons.length) {
        addBlock(document.name, rule.targetLabel, source.reason ?? (typeof target === 'string' ? target : target.reasons.join(' ')), expected,
          typeof target === 'string' ? undefined : { file: document.name, location: target.location, value: target.text });
        continue;
      }
      findings.push({ ...compareValues(source.value.text, target.text, rule.comparison, rule.required, rule.tolerance), kind: 'field',
        document: document.name, field: rule.targetLabel, expected, actual: { file: document.name, location: target.location, value: target.text }, notices: source.notices });
    }
    if (detailSheet && plan.details) {
      const rule = plan.details;
      const wantedHeaders = [rule.targetKeyHeader, ...rule.columns.map((c) => c.targetHeader)].map(labelText);
      const candidates = document.tables.filter((table) => {
        const headers = table.rows[0]?.map((c) => labelText(c.text)) ?? [];
        return wantedHeaders.every((name) => headers.includes(name));
      });
      if (candidates.length !== 1) { addBlock(document.name, '明细表', candidates.length ? '命中多个相同表头的明细表，需人工消除歧义。' : '未找到与已配置列名完整对应的明细表。'); continue; }
      const table = candidates[0];
      const headers = table.rows[0];
      if (table.reasons.length || headers.some((h) => h.columnSpan !== 1 || h.reasons.length) || wantedHeaders.some((name) => headers.filter((c) => labelText(c.text) === name).length !== 1)) {
        addBlock(document.name, '明细表', table.reasons.join(' ') || '明细表头有合并、修订、域代码或重名列，无法可靠定位。'); continue;
      }
      const keyColumn = headers.find((c) => labelText(c.text) === labelText(rule.targetKeyHeader))!.column;
      const sourceRows = new Map<string, number[]>();
      for (const [detailRow, cells] of detailSheet.rows) {
        if (detailRow <= rule.headerRow || ![...cells.values()].some((c) => c.value.kind !== 'blank')) continue;
        if (getCell(detailSheet, detailRow, rule.objectColumn)?.value.text !== key) continue;
        const cell = readValue(getCell(detailSheet, detailRow, rule.keyColumn), false);
        if (cell.reason || !cell.value.text.trim()) { addBlock(document.name, '明细关联键', cell.reason ?? '台账明细关联键为空。', sourceEvidence(batch, detailSheet, detailRow, rule.keyColumn)); continue; }
        sourceRows.set(cell.value.text, [...sourceRows.get(cell.value.text) ?? [], detailRow]);
      }
      const targetRows = new Map<string, WordTable['rows']>();
      for (const targetRow of table.rows.slice(1)) {
        if (!targetRow.some((c) => c.text.trim() || c.reasons.length)) continue;
        const cell = targetRow.find((c) => c.column === keyColumn);
        if (!cell || !cell.text.trim() || cell.columnSpan !== 1 || cell.reasons.length) { addBlock(document.name, '明细关联键', cell?.reasons.join(' ') || 'Word 明细关联键为空或不支持。', undefined, cell ? { file: document.name, location: cell.location, value: cell.text } : undefined); continue; }
        targetRows.set(cell.text, [...targetRows.get(cell.text) ?? [], targetRow]);
      }
      if (!sourceRows.size && !targetRows.size) findings.push({ status: 'equal', kind: 'detail-set', document: document.name, field: '明细条目集合', message: '台账与 Word 的该对象明细均为空。' });
      for (const itemKey of new Set([...sourceRows.keys(), ...targetRows.keys()])) {
        const left = sourceRows.get(itemKey) ?? [], right = targetRows.get(itemKey) ?? [];
        if (left.length > 1 || right.length > 1) { addBlock(document.name, `${itemKey} · 关联键`, '明细关联键重复，禁止自动选择其中一行。'); continue; }
        if (!left.length || !right.length) {
          findings.push({ status: 'different', kind: left.length ? 'missing' : 'extra', document: document.name, field: itemKey,
            message: left.length ? '台账有该条目，Word 明细缺项。' : 'Word 明细有该条目，台账中不存在。',
            expected: left.length ? sourceEvidence(batch, detailSheet, left[0], rule.keyColumn) : undefined,
            actual: right.length ? { file: document.name, location: right[0].find((c) => c.column === keyColumn)!.location, value: itemKey } : undefined });
          continue;
        }
        for (const column of rule.columns) {
          const source = readValue(getCell(detailSheet, left[0], column.sourceColumn), plan.allowCachedFormula);
          const expected = sourceEvidence(batch, detailSheet, left[0], column.sourceColumn, source.value.text);
          const targetColumn = headers.find((c) => labelText(c.text) === labelText(column.targetHeader))!.column;
          const target = right[0].find((c) => c.column === targetColumn);
          if (source.reason || !target || target.columnSpan !== 1 || target.reasons.length) {
            addBlock(document.name, `${itemKey} · ${column.targetHeader}`, source.reason ?? (target?.reasons.join(' ') || 'Word 明细单元格缺失或存在合并。'), expected,
              target ? { file: document.name, location: target.location, value: target.text } : undefined); continue;
          }
          findings.push({ ...compareValues(source.value.text, target.text, column.comparison, column.required, column.tolerance), kind: 'detail',
            document: document.name, field: `${itemKey} · ${column.targetHeader}`, expected,
            actual: { file: document.name, location: target.location, value: target.text }, notices: source.notices });
        }
      }
    }
  }
  if (!batch.documents.length) addBlock('', '检查覆盖', '没有成功解析的 Word 文档。');
  const equal = findings.filter((f) => f.status === 'equal').length;
  const different = findings.filter((f) => f.status === 'different').length;
  const unverifiable = findings.filter((f) => f.status === 'unverifiable').length;
  const complete = unverifiable === 0 && batch.errors.length === 0 && findings.length > 0;
  return { schemaVersion: 1, engineVersion: '0.2.0', planHash: hash(JSON.stringify(plan)), files: batch.files,
    scope: '仅核对方案指定的字段和首行为表头的正文普通明细表；明细按所属对象覆盖该对象全部数据行（含隐藏/筛选行）。页眉页脚、图片、文本框、嵌套/纵向合并表格与真实页码不在支持范围。',
    plan, findings, fileErrors: batch.errors,
    summary: { equal, different, unverifiable, fileErrors: batch.errors.length, documents: batch.documents.length, complete,
      verdict: !complete ? 'incomplete' : different ? 'different' : 'equal' } };
}
