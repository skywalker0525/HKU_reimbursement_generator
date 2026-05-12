const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');
const ExcelJS = require('exceljs');
const JSZip = require('jszip');

const {
  buildExcel,
  buildPdfRenderHtml,
  buildWord,
  getImageDimensions,
  normalizeCurrencyCode,
  normalizeLayoutOptions,
  preflightExchangeRates,
} = require('../src/main');

const tinyPng = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAIAAAADCAIAAADZc7J/AAAACXBIWXMAAAsTAAALEwEAmpwYAAAADklEQVR4nGP8z8DAwMAAAAwAAQGAtJkAAAAASUVORK5CYII=',
  'base64',
);

function receipt(overrides = {}) {
  return {
    index: 1,
    originalName: 'receipt.pdf',
    sourcePath: 'receipt.pdf',
    invoiceDate: '2026-05-12',
    invoiceNumber: 'INV001',
    description: 'Receipt Expense',
    currency: 'OTHER',
    originalAmount: '100.00',
    hkabRate: '',
    hkdAmount: '',
    warning: '',
    ...overrides,
  };
}

test('currency codes normalize for HKAB lookup', () => {
  assert.equal(normalizeCurrencyCode('RMB'), 'CNY');
  assert.equal(normalizeCurrencyCode('usd'), 'USD');
  assert.equal(normalizeCurrencyCode('HKD'), 'HKD');
  assert.equal(normalizeCurrencyCode('BTC'), 'OTHER');
});

test('missing manual rate blocks generation preflight', async () => {
  const result = await preflightExchangeRates([receipt()]);
  assert.equal(result.missingRates.length, 1);
  assert.equal(result.receipts[0].missingRate, true);
});

test('manual rate calculates HKD amount with HKAB per-100 formula', async () => {
  const result = await preflightExchangeRates([
    receipt({ currency: 'USD', originalAmount: '200', hkabRate: '780' }),
  ]);
  assert.equal(result.missingRates.length, 0);
  assert.equal(result.receipts[0].rateCode, 'USD');
  assert.equal(result.receipts[0].hkdAmount, '1560.00');
});

test('HKD receipts do not require exchange rates', async () => {
  const result = await preflightExchangeRates([
    receipt({ currency: 'HKD', originalAmount: '88.5' }),
  ]);
  assert.equal(result.missingRates.length, 0);
  assert.equal(result.receipts[0].hkdAmount, '88.50');
});

test('PDF render HTML uses PDF.js canvas without viewer chrome', () => {
  const html = buildPdfRenderHtml('file:///pdf.mjs', 'file:///worker.mjs', 'file:///receipt.pdf');
  assert.match(html, /pdfjsLib\.getDocument/);
  assert.match(html, /canvas\.toDataURL\('image\/png'\)/);
  assert.doesNotMatch(html, /toolbar|navpanes|scrollbar|capturePage|pdf-viewer/i);
});

test('image dimensions preserve full-page aspect ratio', () => {
  assert.deepEqual(getImageDimensions(tinyPng, '.png'), { width: 2, height: 3 });
});

test('Excel total formula spans all generated receipt rows', async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'hku-excel-'));
  const excelPath = path.join(dir, 'out.xlsx');
  await buildExcel([
    receipt({ hkdAmount: '10' }),
    receipt({ hkdAmount: '20' }),
  ], excelPath, { claimantName: 'A', department: 'B', staffStudentNo: 'C', telephoneNo: 'D' });
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(excelPath);
  assert.deepEqual(workbook.getWorksheet(1).getCell('K55').value, { formula: 'SUM(K29:K31)' });
});

test('Word layout controls page breaks and evidence order', async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'hku-word-'));
  const imagePath = path.join(dir, 'receipt.png');
  const ratePath = path.join(dir, 'rate.png');
  await fs.writeFile(imagePath, tinyPng);
  await fs.writeFile(ratePath, tinyPng);
  const wordPath = path.join(dir, 'out.docx');
  await buildWord([
    receipt({ receiptScreenshot: imagePath, rateScreenshot: ratePath, currency: 'USD', rateCode: 'USD', hkabRate: '780', hkdAmount: '780' }),
    receipt({ index: 2, receiptScreenshot: imagePath, rateScreenshot: ratePath, currency: 'USD', rateCode: 'USD', hkabRate: '780', hkdAmount: '780' }),
  ], wordPath, { claimantName: 'A', department: 'B', staffStudentNo: 'C', telephoneNo: 'D' }, normalizeLayoutOptions({
    receiptStartsNewPage: true,
    exchangeEvidencePosition: 'beforeReceipt',
  }));
  const zip = await JSZip.loadAsync(await fs.readFile(wordPath));
  const documentXml = await zip.file('word/document.xml').async('string');
  assert.match(documentXml, /w:type="page"/);
  assert.ok(documentXml.indexOf('HKAB exchange-rate evidence') < documentXml.indexOf('Original receipt image'));
});
