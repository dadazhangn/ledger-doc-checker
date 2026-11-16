import { test, expect } from '@playwright/test';
import { mkdir } from 'node:fs/promises';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';

test('演示的真实界面流程、报告下载、规则变化使旧结果过期', async ({ page }) => {
  const errors: string[] = []; page.on('pageerror', (error) => errors.push(error.message));
  await page.setViewportSize({ width: 1440, height: 1000 });
  await page.goto('/');
  await page.locator('#tab-excel').click();
  await expect(page.getByRole('heading', { name: '每一处差异，都有据可查。' })).toBeVisible();
  await mkdir('.cache', { recursive: true });
  await page.screenshot({ path: '.cache/desktop-start.png', fullPage: true });
  await page.getByRole('button', { name: /体验演示/ }).click();
  await expect(page.locator('#rules')).toBeVisible();
  await expect(page.locator('#notice')).toContainText('虚构数据');
  await page.getByRole('button', { name: /开始核对/ }).click();
  await expect(page.locator('#result-title')).toHaveText('发现 4 处差异');
  await expect(page.locator('#count-different')).toHaveText('4');
  await expect(page.locator('#count-unverifiable')).toHaveText('0');
  await expect(page.locator('#findings tbody tr')).toHaveCount(4);
  await expect(page.locator('#findings')).toContainText('对象台账!B2');
  await page.screenshot({ path: '.cache/desktop-results.png', fullPage: true });
  await page.locator('#result-search').fill('版本');
  await expect(page.locator('#findings tbody tr')).toHaveCount(1);
  await page.locator('#result-search').fill('');
  const download = page.waitForEvent('download');
  await page.getByRole('button', { name: /导出报告/ }).click();
  await (await download).saveAs('.cache/demo-report.html');
  await page.locator('#identity-label').fill('对象名称 ');
  await expect(page.locator('#stale')).toBeVisible();
  await expect(page.locator('#download-html')).toBeDisabled();
  await page.getByRole('button', { name: /开始核对/ }).click();
  await expect(page.locator('#stale')).toBeHidden();
  await expect(page.locator('#download-html')).toBeEnabled();
  expect(errors).toEqual([]);
});

test('窄屏首屏可操作且页面本身不横向溢出', async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto('/');
  await page.locator('#tab-excel').click();
  await expect(page.locator('#load-demo')).toBeVisible();
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth)).toBe(true);
  await page.screenshot({ path: '.cache/mobile-start.png', fullPage: true });
  await page.getByRole('button', { name: /体验演示/ }).click();
  await page.getByRole('button', { name: /开始核对/ }).click();
  await expect(page.locator('#result-title')).toHaveText('发现 4 处差异');
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth)).toBe(true);
  await page.screenshot({ path: '.cache/mobile-results.png', fullPage: true });
});

test('用户选文件并从空白配置建立方案，再保存和重载', async ({ page }) => {
  await page.goto('/');
  await page.locator('#tab-excel').click();
  await page.locator('#ledger-file').setInputFiles(resolve('examples/演示台账.xlsx'));
  await page.locator('#word-files').setInputFiles([resolve('examples/演示说明-甲.docx'), resolve('examples/演示说明-乙.docx')]);
  await page.getByRole('button', { name: /读取资料/ }).click();
  await expect(page.locator('#rules')).toBeVisible();
  await page.locator('#object-key').selectOption('A');
  await page.locator('#identity-label').fill('对象名称');
  await page.getByLabel('演示说明-甲.docx 对应的台账对象').selectOption('演示对象甲');
  await page.getByLabel('演示说明-乙.docx 对应的台账对象').selectOption('演示对象乙');
  await page.getByRole('button', { name: '＋ 添加字段' }).click();
  const field = page.locator('#field-rules .rule-row').first();
  await field.getByLabel('Excel 来源列').selectOption('B');
  await field.getByLabel('Word 字段', { exact: true }).fill('版本');
  await page.getByRole('button', { name: /开始核对/ }).click();
  await expect(page.locator('#result-title')).toHaveText('发现 1 处差异');
  const saved = page.waitForEvent('download');
  await page.getByRole('button', { name: '保存方案' }).click();
  await (await saved).saveAs('.cache/saved-plan.json');
  await field.getByLabel('Word 字段', { exact: true }).fill('找不到的字段');
  await page.locator('#import-plan').setInputFiles(resolve('.cache/saved-plan.json'));
  await expect(page.locator('#field-rules .rule-row').first().getByLabel('Word 字段', { exact: true })).toHaveValue('版本');
  await page.getByRole('button', { name: /开始核对/ }).click();
  await expect(page.locator('#result-title')).toHaveText('发现 1 处差异');
  await expect(page.locator('#stale')).toBeHidden();
});

test('Word 互核演示将软件与文档版本分开，差异由人工判断', async ({ page }) => {
  const errors: string[] = []; page.on('pageerror', (error) => errors.push(error.message));
  await page.setViewportSize({ width: 1440, height: 1000 });
  await page.goto('/');
  await expect(page.getByRole('heading', { name: '把不同时期的文档，放在一起看。' })).toBeVisible();
  await mkdir('.cache', { recursive: true });
  await page.screenshot({ path: '.cache/word-start.png', fullPage: true });
  await page.locator('#word-demo').click();
  await expect(page.locator('#word-rules')).toBeVisible();
  await expect(page.locator('#word-notice')).toContainText('虚构');
  await page.locator('#word-compare').click();
  await expect(page.locator('#word-result-title')).toHaveText('发现 2 组不同取值');
  await expect(page.locator('#word-differences')).toContainText('软件版本');
  await expect(page.locator('#word-differences')).toContainText('文档版本');
  await expect(page.locator('#word-differences')).toContainText('前期');
  await expect(page.locator('#word-differences')).toContainText('后期');
  await page.screenshot({ path: '.cache/word-results.png', fullPage: true });
  const download = page.waitForEvent('download');
  await page.locator('#word-download-html').click();
  await (await download).saveAs('.cache/word-report.html');
  await page.locator('.word-document summary').first().click();
  await page.locator('.word-document').first().getByLabel('第 2 项阶段').fill('用户修改的阶段');
  await expect(page.locator('#word-stale')).toBeVisible();
  await expect(page.locator('#word-download-html')).toBeDisabled();
  expect(errors).toEqual([]);
});

test('同名 Word 可分批添加，人工选值和对照组后只列出不同取值', async ({ page }) => {
  await page.goto('/');
  const first = await readFile(resolve('examples/演示文档-前期.docx'));
  const second = await readFile(resolve('examples/演示文档-后期.docx'));
  await page.locator('#word-input').setInputFiles({ name: '演示说明.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: first });
  await page.locator('#word-input').setInputFiles({ name: '演示说明.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: second });
  await expect(page.locator('#word-selected-files')).toContainText('材料 1');
  await expect(page.locator('#word-selected-files')).toContainText('材料 2');
  await page.locator('#word-inspect').click();
  await expect(page.locator('#word-rules')).toBeVisible();
  const documents = page.locator('.word-document');
  for (let i = 0; i < 2; i++) {
    const doc = documents.nth(i); await doc.locator('summary').click();
    await doc.getByLabel('第 2 项来源片段', { exact: true }).selectOption('body:t:1:r:3:c:2:0');
    await doc.getByLabel('第 2 项对照组').fill('演示软件');
    await doc.getByLabel('第 2 项阶段').fill(i ? '后期' : '前期');
  }
  await page.locator('#word-compare').click();
  await expect(page.locator('#word-result-title')).toHaveText('发现 1 组不同取值');
  await expect(page.locator('#word-differences')).toContainText('1.0');
  await expect(page.locator('#word-differences')).toContainText('2.0');
  await expect(page.locator('#word-coverage')).toContainText('未选择');
});

test('Word 互核在窄屏上可操作', async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto('/');
  await page.locator('#word-demo').click();
  await expect(page.locator('#word-rules')).toBeVisible();
  await page.locator('#word-compare').click();
  await expect(page.locator('#word-result-title')).toHaveText('发现 2 组不同取值');
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth)).toBe(true);
  await page.screenshot({ path: '.cache/word-mobile.png', fullPage: true });
});
