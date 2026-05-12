const { app, BrowserWindow, dialog, ipcMain, nativeImage } = require('electron');
const path = require('path');
const fs = require('fs/promises');
const fss = require('fs');
const { pathToFileURL } = require('url');
const https = require('https');
const pdfParse = require('pdf-parse');
const Tesseract = require('tesseract.js');
const ExcelJS = require('exceljs');
const {
  AlignmentType,
  Document,
  ExternalHyperlink,
  HeadingLevel,
  ImageRun,
  Packer,
  Paragraph,
  TextRun,
} = require('docx');

const ROOT = path.resolve(__dirname, '..', '..');
const DEFAULT_TEMPLATE = path.join(__dirname, '..', 'templates', '615_404_Reimbursement_form.xlsx');
const HKAB_URL = 'https://www.hkab.org.hk/en/rates/exchange-rates';
const HKAB_API_BASE = 'https://www.hkab.org.hk/api/member/public/getExrate';

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1220,
    height: 820,
    minWidth: 980,
    minHeight: 680,
    title: 'HKU_reimbursement_generator',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
    },
  });
  mainWindow.loadFile(path.join(__dirname, 'renderer.html'));
}

app.whenReady().then(createWindow);
app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});

ipcMain.handle('select-files', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: '选择票据 / Select Receipts',
    properties: ['openFile', 'multiSelections'],
    filters: [
      { name: 'Receipts', extensions: ['pdf', 'jpg', 'jpeg', 'png'] },
      { name: 'All Files', extensions: ['*'] },
    ],
  });
  return result.canceled ? [] : result.filePaths;
});

ipcMain.handle('select-output-folder', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: '选择输出文件夹 / Select Output Folder',
    properties: ['openDirectory', 'createDirectory'],
  });
  return result.canceled ? '' : result.filePaths[0];
});

ipcMain.handle('analyze-files', async (_event, filePaths) => {
  const receipts = [];
  for (let i = 0; i < filePaths.length; i += 1) {
    receipts.push(await analyzeReceipt(filePaths[i], i + 1));
  }
  return receipts;
});

ipcMain.handle('generate-package', async (event, payload) => generatePackage(payload, (progress) => {
  event.sender.send('generation-progress', progress);
}));

async function analyzeReceipt(filePath, index) {
  const ext = path.extname(filePath).toLowerCase();
  const stat = await fs.stat(filePath);
  let text = '';
  const warnings = [];

  try {
    if (ext === '.pdf') {
      const data = await fs.readFile(filePath);
      const parsed = await pdfParse(data);
      text = parsed.text || '';
    } else {
      const result = await Tesseract.recognize(filePath, 'eng+chi_sim');
      text = result.data.text || '';
    }
  } catch (error) {
    warnings.push(`Text extraction failed: ${error.message}`);
  }

  const invoiceNumber = extractInvoiceNumber(text, filePath);
  const invoiceDate = extractInvoiceDate(text) || toDateInput(stat.mtime);
  const amount = extractAmount(text);
  const currency = extractCurrency(text, filePath);
  const rawDescription = extractDescription(text, filePath);
  const description = toEnglishDescription(rawDescription || path.basename(filePath, ext));

  if (!amount) warnings.push('Amount needs review.');
  if (!invoiceNumber) warnings.push('Invoice number needs review.');
  if (!text.trim()) warnings.push('No embedded text found. Please review fields manually.');

  return {
    id: `${Date.now()}-${index}`,
    index,
    sourcePath: filePath,
    originalName: path.basename(filePath),
    invoiceNumber,
    invoiceDate,
    description,
    currency,
    originalAmount: amount || '',
    hkabRate: '',
    hkdAmount: '',
    confidence: warnings.length ? 'Needs review' : 'Good',
    warning: warnings.join(' '),
  };
}

function extractInvoiceNumber(text, filePath) {
  const haystack = `${text}\n${path.basename(filePath)}`;
  const patterns = [
    /(?:Invoice|发票|票据|号码|No\.?|Number)[^\d]{0,20}(\d{8,24})/i,
    /_(\d{14,24})/,
    /\b(\d{20})\b/,
    /\b(\d{8})\b/,
  ];
  for (const pattern of patterns) {
    const match = haystack.match(pattern);
    if (match) return match[1];
  }
  return '';
}

function extractInvoiceDate(text) {
  const patterns = [
    /(\d{4})[年\/\-.](\d{1,2})[月\/\-.](\d{1,2})日?/,
    /(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/,
  ];
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (!match) continue;
    if (match[1].length === 4) {
      return `${match[1]}-${pad2(match[2])}-${pad2(match[3])}`;
    }
    return `${match[3]}-${pad2(match[2])}-${pad2(match[1])}`;
  }
  const pdfDate = text.match(/D:(\d{4})(\d{2})(\d{2})/);
  if (pdfDate) return `${pdfDate[1]}-${pdfDate[2]}-${pdfDate[3]}`;
  return '';
}

function extractAmount(text) {
  const compact = text.replace(/\s+/g, ' ');
  const currencyAmounts = [...compact.matchAll(/(?:RMB|CNY|HKD|HK\$|¥|￥)\s*([0-9][0-9,]*\.?\d{0,2})/gi)]
    .map((match) => Number(match[1].replace(/,/g, '')))
    .filter((value) => Number.isFinite(value) && value > 0 && value < 1000000);
  if (currencyAmounts.length) return Math.max(...currencyAmounts).toFixed(2);
  const focusedPatterns = [
    /(?:价税合计|小写|合计|Total|Amount|Subtotal)[^0-9¥￥HK$]{0,30}(?:RMB|CNY|HKD|HK\$|¥|￥)?\s*([0-9][0-9,]*\.?\d{0,2})/i,
    /(?:RMB|CNY|HKD|HK\$|¥|￥)\s*([0-9][0-9,]*\.?\d{0,2})/i,
  ];
  for (const pattern of focusedPatterns) {
    const match = compact.match(pattern);
    if (match) return normalizeAmount(match[1]);
  }
  const candidates = [...compact.matchAll(/([0-9][0-9,]*\.\d{2})/g)]
    .map((match) => Number(match[1].replace(/,/g, '')))
    .filter((value) => Number.isFinite(value) && value > 0 && value < 1000000);
  if (!candidates.length) return '';
  return Math.max(...candidates).toFixed(2);
}

function extractCurrency(text, filePath) {
  const haystack = `${text}\n${filePath}`;
  if (/HKD|HK\$|港币|港幣/i.test(haystack)) return 'HKD';
  if (/USD|US\$|美元/i.test(haystack)) return 'USD';
  if (/RMB|CNY|¥|￥|人民币|人民幣|China Tax/i.test(haystack)) return 'RMB';
  return 'RMB';
}

function extractDescription(text, filePath) {
  const known = [
    'Mouse Accessories',
    'Phone Clamp and Holder',
    'Computer and Phone Accessories',
    'USB Hub',
    'Power Bank',
    'Tripod',
    'Waterproof Material',
    'Electric Drill',
    'Casters',
    'International Freight Forwarding Service Fee',
    'Delivery Fee',
    'Training Fee',
  ];
  for (const item of known) {
    if (new RegExp(item, 'i').test(text)) return item;
  }
  const chineseHints = [
    [/充电器|充电宝|三脚架|手机支架|USB|插座|电脑|笔记本/i, 'Computer and Phone Accessories'],
    [/鼠标|键盘/i, 'Mouse Accessories'],
    [/运费|货运|物流|快递|运输/i, 'Delivery Fee'],
    [/防水材料|防水/i, 'Waterproof Material'],
    [/电钻|工具/i, 'Electric Drill'],
    [/脚轮|万向轮/i, 'Casters'],
    [/培训|课程|安全证书/i, 'Training Fee'],
  ];
  for (const [pattern, description] of chineseHints) {
    if (pattern.test(text)) return description;
  }
  const chineseLines = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => /[\u4e00-\u9fff]/.test(line) && line.length >= 3 && line.length <= 40);
  const productLine = chineseLines.find((line) => /货物|服务|材料|配件|工具|运输|培训|办公|电脑|电子|五金|夹|支架|钻|轮|螺丝/.test(line));
  return productLine || path.basename(filePath, path.extname(filePath));
}

function toEnglishDescription(value) {
  const text = String(value || '').trim();
  const dictionary = [
    [/鼠标|滑鼠/i, 'Mouse Accessories'],
    [/手机.*夹|夹具|支架|云台/i, 'Phone Clamp and Holder'],
    [/电脑|计算机|USB|拓展坞|扩展坞|充电宝|三脚架/i, 'Computer and Phone Accessories'],
    [/防水|防水材料/i, 'Waterproof Material'],
    [/电钻|钻/i, 'Electric Drill'],
    [/脚轮|万向轮|轮子/i, 'Casters'],
    [/运费|运输|货运|物流|快递/i, 'Delivery Fee'],
    [/培训|课程|安全/i, 'Training Fee'],
    [/螺母|螺栓|螺丝|五金/i, 'Steel Nut and Bolt'],
    [/尼龙扎带|扎带/i, 'Nylon Cable Tie'],
  ];
  for (const [pattern, replacement] of dictionary) {
    if (pattern.test(text)) return replacement;
  }
  const cleaned = text
    .replace(/[_-]*\d{8,24}[_-]*/g, ' ')
    .replace(/[^\w\s()+,.-]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  return titleCase(cleaned || 'Receipt Expense');
}

async function generatePackage({ receipts, outputBase, claimantInfo }, reportProgress = () => {}) {
  if (!receipts?.length) throw new Error('No receipts to generate.');
  const claimant = normalizeClaimantInfo(claimantInfo);
  const today = new Date();
  const claimantSlug = sanitizePersonName(claimant.claimantName) || 'CLAIMANT';
  const outputRoot = uniquePath(path.join(outputBase || ROOT, `Reimbursement_${formatDate(today)}_${claimantSlug}`));
  const receiptsDir = path.join(outputRoot, 'receipts');
  const rateDir = path.join(outputRoot, 'exchange_rates');
  const receiptShotDir = path.join(outputRoot, 'receipt_screenshots');
  await fs.mkdir(receiptsDir, { recursive: true });
  await fs.mkdir(rateDir, { recursive: true });
  await fs.mkdir(receiptShotDir, { recursive: true });
  reportProgress({ percent: 8, message: 'Created output folder / 已创建输出文件夹' });

  const enriched = [];
  const rateCache = new Map();
  for (let i = 0; i < receipts.length; i += 1) {
    const receipt = { ...receipts[i], index: i + 1 };
    reportProgress({ percent: 10 + Math.round((i / receipts.length) * 55), message: `Processing receipt ${i + 1}/${receipts.length} / 正在处理票据 ${i + 1}/${receipts.length}` });
    const amount = Number(String(receipt.originalAmount).replace(/,/g, '')) || 0;
    if (receipt.currency === 'RMB') {
      let rateInfo = rateCache.get(receipt.invoiceDate || 'current');
      if (!rateInfo) {
        rateInfo = await captureExchangeRate(receipt.invoiceDate, rateDir);
        rateCache.set(receipt.invoiceDate || 'current', rateInfo);
      }
      receipt.hkabRate = receipt.hkabRate || rateInfo.rate || '';
      receipt.rateScreenshot = rateInfo.screenshotPath;
      receipt.rateSource = HKAB_URL;
      const rate = Number(receipt.hkabRate);
      receipt.hkdAmount = rate ? (amount * rate / 100).toFixed(2) : receipt.hkdAmount || '';
      if (!rate) receipt.warning = joinWarning(receipt.warning, rateInfo.warning || 'HKAB CNY Selling rate could not be parsed automatically for the invoice date.');
    } else if (receipt.currency === 'HKD') {
      receipt.hkdAmount = amount.toFixed(2);
    } else {
      receipt.warning = joinWarning(receipt.warning, 'Non-HKD/RMB currency needs manual confirmation.');
    }
    receipt.copiedPath = await copyRenamedReceipt(receipt, receiptsDir);
    receipt.receiptScreenshot = await captureReceiptImage(receipt.copiedPath, receiptShotDir, receipt.index);
    enriched.push(receipt);
  }

  const personName = claimantSlug || 'CLAIMANT';
  const excelPath = path.join(outputRoot, `615_404 Reimbursement form_${personName}.xlsx`);
  const wordPath = path.join(outputRoot, `Reimbursement_${personName}.docx`);
  const previewPath = path.join(outputRoot, 'receipt_preview.json');

  reportProgress({ percent: 72, message: 'Writing Excel / 正在写入 Excel' });
  await buildExcel(enriched, excelPath, claimant);
  reportProgress({ percent: 84, message: 'Writing Word document / 正在生成 Word 文档' });
  await buildWord(enriched, wordPath, claimant);
  reportProgress({ percent: 94, message: 'Saving preview data / 正在保存预览数据' });
  await fs.writeFile(previewPath, JSON.stringify(enriched, null, 2), 'utf8');
  reportProgress({ percent: 100, message: 'Done / 完成' });

  return { outputRoot, excelPath, wordPath, previewPath, receipts: enriched };
}

async function captureExchangeRate(invoiceDate, rateDir) {
  const safeDate = invoiceDate || formatDate(new Date());
  const screenshotPath = path.join(rateDir, `HKAB_CNY_${safeDate}.png`);
  const apiRate = await fetchHkabApiRate(safeDate);
  if (apiRate.rate) {
    const evidenceHtml = buildHkabEvidenceHtml({
      requestedDate: safeDate,
      rateDate: apiRate.rateDate,
      lastUpdated: apiRate.lastUpdated,
      rate: apiRate.rate,
      sourceUrl: `${HKAB_API_BASE}/${safeDate}`,
      rowHtml: buildHkabApiRowHtml(apiRate),
      warning: apiRate.holiday ? 'HKAB marks this date as a non-working day.' : '',
    });
    await captureHtmlEvidence(evidenceHtml, screenshotPath);
    return { rate: apiRate.rate, screenshotPath };
  }

  const interactive = await tryCaptureInteractiveHkabRate(safeDate, screenshotPath);
  if (interactive.rate || interactive.screenshotPath) return interactive;

  const page = await fetchHkabExchangePage();
  const parsed = parseHkabPage(page.html);
  if (parsed.rate && (!invoiceDate || parsed.rateDate === invoiceDate)) {
    const evidenceHtml = buildHkabEvidenceHtml({
      requestedDate: safeDate,
      rateDate: parsed.rateDate,
      lastUpdated: parsed.lastUpdated,
      rate: parsed.rate,
      sourceUrl: HKAB_URL,
      rowHtml: parsed.rowHtml,
    });
    await captureHtmlEvidence(evidenceHtml, screenshotPath);
    return { rate: parsed.rate, screenshotPath };
  }

  const evidenceHtml = buildHkabEvidenceHtml({
    requestedDate: safeDate,
    rateDate: parsed.rateDate,
    lastUpdated: parsed.lastUpdated,
    rate: parsed.rate || '',
    sourceUrl: HKAB_URL,
    rowHtml: parsed.rowHtml,
    warning: `HKAB page did not provide the requested invoice date (${safeDate}). The current page date was ${parsed.rateDate || 'not detected'}, so no automatic rate was applied.`,
  });
  await captureHtmlEvidence(evidenceHtml, screenshotPath);
  return {
    rate: '',
    screenshotPath,
    warning: `HKAB did not show the requested invoice date ${safeDate}; please enter the correct CNY Selling rate manually.`,
  };
}

async function tryCaptureInteractiveHkabRate(invoiceDate, screenshotPath) {
  const win = new BrowserWindow({
    width: 1365,
    height: 1100,
    show: false,
    webPreferences: { sandbox: true },
  });
  try {
    await win.loadURL(HKAB_URL);
    await new Promise((resolve) => setTimeout(resolve, 3500));
    await win.webContents.executeJavaScript(`
      (() => {
        const input = document.querySelector('input[type="date"]');
        if (input) {
          const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
          setter.call(input, ${JSON.stringify(invoiceDate)});
          input.dispatchEvent(new Event('input', { bubbles: true }));
          input.dispatchEvent(new Event('change', { bubbles: true }));
          input.blur();
        }
      })()
    `);
    await new Promise((resolve) => setTimeout(resolve, 3500));
    const bodyText = await win.webContents.executeJavaScript(`
      (() => {
        const rows = [...document.querySelectorAll('[role="row"], tr, div')];
        const row = rows.find((el) => /\\bCNY\\b/.test(el.innerText || '') && /Selling/i.test(el.innerText || ''));
        if (row) {
          row.scrollIntoView({ block: 'center', inline: 'nearest' });
          row.style.outline = '4px solid #d97706';
          row.style.outlineOffset = '4px';
        }
        document.body.innerText;
      })()
    `);
    await new Promise((resolve) => setTimeout(resolve, 700));
    const image = await win.webContents.capturePage({ x: 0, y: 0, width: 1365, height: 900 });
    await fs.writeFile(screenshotPath, image.toPNG());
    const displayedDate = extractHkabDisplayedDate(bodyText);
    const rate = parseCnySellingRate(bodyText);
    if (!rate) return { rate: '', screenshotPath: '' };
    if (invoiceDate && displayedDate && displayedDate !== invoiceDate) {
      return {
        rate: '',
        screenshotPath,
        warning: `HKAB page displayed ${displayedDate}, not requested invoice date ${invoiceDate}; please enter the correct CNY Selling rate manually.`,
      };
    }
    if (invoiceDate && !displayedDate) {
      return {
        rate: '',
        screenshotPath,
        warning: `HKAB page date could not be verified for invoice date ${invoiceDate}; please enter the correct CNY Selling rate manually.`,
      };
    }
    return { rate, screenshotPath };
  } catch (_error) {
    return { rate: '', screenshotPath: '' };
  } finally {
    win.destroy();
  }
}

function parseCnySellingRate(text) {
  const lines = text.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
  for (let i = 0; i < lines.length; i += 1) {
    if (/^(?:Currency Code:\s*)?CNY$/i.test(lines[i]) || /Onshore RMB \/ Chinese Yuan/i.test(lines[i])) {
      const nearby = lines.slice(Math.max(0, i - 2), i + 12).join(' ');
      const selling = nearby.match(/Selling:?\s*(\d{2,3}\.\d{2,6})/i);
      if (selling) return selling[1];
      const numbers = nearby.match(/\b\d{2,3}\.\d{2,6}\b/g);
      if (numbers?.length) return numbers[0];
    }
  }
  const compact = text.replace(/\s+/g, ' ');
  const broad = compact.match(/(?:Currency Code:\s*)?CNY\s+Currency:\s+Onshore RMB \/ Chinese Yuan\s+Selling:?\s*(\d{2,3}\.\d{2,6})/i);
  return broad ? broad[1] : '';
}

function extractHkabDisplayedDate(text) {
  const match = text.match(/exchange rates as on\s+(\d{4}-\d{2}-\d{2})/i)
    || text.match(/as on\s+(\d{4}-\d{2}-\d{2})/i)
    || text.match(/Last updated:\s*(\d{4}-\d{2}-\d{2})/i);
  return match ? match[1] : '';
}

function fetchHkabExchangePage() {
  return new Promise((resolve) => {
    const request = https.get(HKAB_URL, { timeout: 20000 }, (response) => {
      let html = '';
      response.setEncoding('utf8');
      response.on('data', (chunk) => { html += chunk; });
      response.on('end', () => resolve({ status: response.statusCode, html }));
    });
    request.on('timeout', () => {
      request.destroy(new Error('HKAB request timed out.'));
    });
    request.on('error', (error) => resolve({ status: 0, html: '', error: error.message }));
  });
}

async function fetchHkabApiRate(date) {
  const url = `${HKAB_API_BASE}/${date || ''}`;
  const response = await fetchText(url);
  if (!response.text || response.status < 200 || response.status >= 300) {
    return { rate: '', warning: `HKAB API request failed for ${date}.` };
  }
  try {
    const data = JSON.parse(response.text);
    if (data?.holiday) {
      return {
        rate: '',
        holiday: true,
        rateDate: data.RateDate || date,
        warning: `HKAB marks ${date} as a non-working day.`,
      };
    }
    return {
      rate: data?.CNYSelling || '',
      buyingTT: data?.CNYBuyingTT || '',
      buyingDD: data?.CNYBuyingOD || '',
      rateDate: data?.RateDate || date,
      lastUpdated: data?.updated_at || '',
      holiday: Boolean(data?.holiday),
    };
  } catch (error) {
    return { rate: '', warning: `HKAB API JSON parse failed: ${error.message}` };
  }
}

function fetchText(url) {
  return new Promise((resolve) => {
    const request = https.get(url, { timeout: 20000 }, (response) => {
      let text = '';
      response.setEncoding('utf8');
      response.on('data', (chunk) => { text += chunk; });
      response.on('end', () => resolve({ status: response.statusCode, text }));
    });
    request.on('timeout', () => request.destroy(new Error('Request timed out.')));
    request.on('error', (error) => resolve({ status: 0, text: '', error: error.message }));
  });
}

function buildHkabApiRowHtml(rateInfo) {
  return `
    <div class="rate-row">
      <div><span class="mobile_change_layout_table_cell_mobile_title">Currency Code:</span><div>CNY</div></div>
      <div><span class="mobile_change_layout_table_cell_mobile_title">Currency:</span><div>Onshore RMB / Chinese Yuan</div></div>
      <div><span class="mobile_change_layout_table_cell_mobile_title">Selling:</span><div>${escapeText(rateInfo.rate)}</div></div>
      <div><span class="mobile_change_layout_table_cell_mobile_title">Buying TT:</span><div>${escapeText(rateInfo.buyingTT || '')}</div></div>
      <div><span class="mobile_change_layout_table_cell_mobile_title">Buying D/D:</span><div>${escapeText(rateInfo.buyingDD || '')}</div></div>
    </div>`;
}

function parseHkabPage(html) {
  if (!html) return { rate: '', rateDate: '', lastUpdated: '', rowHtml: '' };
  const rateDate = decodeHtml((html.match(/as on\s*<strong>([^<]+)<\/strong>/i) || [])[1] || '');
  const lastUpdated = decodeHtml((html.match(/Last updated:\s*([^<]+)</i) || [])[1] || '');
  const cnyIndex = html.search(/<div[^>]*>\s*CNY\s*<\/div>/i);
  let rowHtml = '';
  if (cnyIndex >= 0) {
    const rowStart = html.lastIndexOf('<div role="row" class="general_table_row exchange_rate"', cnyIndex);
    const nextRow = html.indexOf('<div role="row" class="general_table_row exchange_rate"', cnyIndex + 20);
    if (rowStart >= 0) {
      rowHtml = html.slice(rowStart, nextRow > rowStart ? nextRow : cnyIndex + 1400);
    }
  }
  const rowText = stripHtml(rowHtml);
  const rateMatch = rowText.match(/Selling:\s*(\d{2,3}\.\d{2,6})/i)
    || html.match(/<div[^>]*>\s*CNY\s*<\/div>[\s\S]{0,600}?Selling:[\s\S]{0,120}?(\d{2,3}\.\d{2,6})/i)
    || html.match(/CNYSelling":\d+[\s\S]{0,600}?"(\d{2,3}\.\d{2,6})"/i);
  return {
    rate: rateMatch ? rateMatch[1] : '',
    rateDate,
    lastUpdated,
    rowHtml,
  };
}

function buildHkabEvidenceHtml({ requestedDate, rateDate, lastUpdated, rate, sourceUrl, rowHtml, warning = '' }) {
  const cleanRow = rowHtml || `
    <div class="rate-row">
      <div><b>Currency Code:</b> CNY</div>
      <div><b>Currency:</b> Onshore RMB / Chinese Yuan</div>
      <div><b>Selling:</b> ${escapeText(rate)}</div>
    </div>`;
  return `<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { margin: 0; background: #f4f7fb; font-family: Arial, sans-serif; color: #16324f; }
    .page { width: 1180px; margin: 0 auto; padding: 44px; }
    .source { background: white; border: 1px solid #c8d6e5; padding: 26px; }
    h1 { margin: 0 0 10px; font-size: 34px; }
    .meta { font-size: 20px; line-height: 1.55; color: #334155; margin-bottom: 24px; }
    .notice { background: #e8f3f7; border-left: 8px solid #2870c5; padding: 16px 20px; margin: 20px 0; font-size: 19px; }
    .warning { background: #fff7ed; border-left: 8px solid #d97706; padding: 16px 20px; margin: 20px 0; font-size: 19px; color: #7c2d12; }
    .row-title, .rate-row { display: grid; grid-template-columns: 180px 1fr 150px 150px 150px; gap: 16px; align-items: center; }
    .row-title { background: #2870c5; color: white; padding: 15px; font-size: 18px; font-weight: bold; }
    .rate-row, [role="row"].exchange_rate { display: grid !important; grid-template-columns: 180px 1fr 150px 150px 150px; gap: 16px; padding: 18px 15px; border: 4px solid #d97706; background: #fff7ed; font-size: 20px; }
    .mobile_change_layout_table_cell_mobile_title { display: block; color: #64748b; font-size: 14px; }
    a { color: #1d4ed8; }
  </style>
</head>
<body>
  <div class="page">
    <div class="source">
      <h1>Exchange Rates | The Hong Kong Association of Banks</h1>
      <div class="meta">
        Source: <a>${escapeText(sourceUrl)}</a><br>
        HKAB rate date shown on page: <b>${escapeText(rateDate || 'Current page')}</b><br>
        Requested invoice date: <b>${escapeText(requestedDate)}</b><br>
        Last updated: <b>${escapeText(lastUpdated || 'Not shown')}</b>
      </div>
      <div class="notice">Hong Kong Dollars to 100 units of Foreign Currency</div>
      ${warning ? `<div class="warning">${escapeText(warning)}</div>` : ''}
      <div class="row-title"><div>Currency Code</div><div>Currency</div><div>Selling</div><div>Buying TT</div><div>Buying D/D</div></div>
      ${cleanRow}
    </div>
  </div>
</body>
</html>`;
}

async function captureHtmlEvidence(html, screenshotPath) {
  const htmlPath = screenshotPath.replace(/\.png$/i, '.html');
  await fs.writeFile(htmlPath, html, 'utf8');
  const win = new BrowserWindow({
    width: 1280,
    height: 780,
    show: false,
    webPreferences: { sandbox: true },
  });
  try {
    await win.loadURL(pathToFileURL(htmlPath).toString());
    await new Promise((resolve) => setTimeout(resolve, 500));
    const image = await win.webContents.capturePage();
    await fs.writeFile(screenshotPath, image.toPNG());
  } finally {
    win.destroy();
  }
}

function stripHtml(html) {
  return decodeHtml(String(html || '').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim());
}

function decodeHtml(value) {
  return String(value || '')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}

function escapeText(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

async function copyRenamedReceipt(receipt, receiptsDir) {
  const ext = path.extname(receipt.sourcePath).toLowerCase();
  const amount = receipt.originalAmount || '0.00';
  const name = `${String(receipt.index).padStart(2, '0')}_${sanitizeFileName(receipt.description)}_${receipt.currency}${amount}${ext}`;
  const dest = uniquePath(path.join(receiptsDir, name));
  await fs.copyFile(receipt.sourcePath, dest);
  return dest;
}

async function captureReceiptImage(filePath, outputDir, index) {
  const ext = path.extname(filePath).toLowerCase();
  if (['.png', '.jpg', '.jpeg'].includes(ext)) return filePath;

  const outputPath = path.join(outputDir, `${String(index).padStart(2, '0')}_receipt.png`);
  const win = new BrowserWindow({
    width: 1800,
    height: 1300,
    show: false,
    webPreferences: { sandbox: true },
  });
  try {
    await win.loadURL(`${pathToFileURL(filePath).toString()}#toolbar=0&navpanes=0&scrollbar=0&view=FitH`);
    await new Promise((resolve) => setTimeout(resolve, 3000));
    const canvasPng = await extractPdfViewerCanvas(win);
    if (canvasPng) {
      await fs.writeFile(outputPath, canvasPng);
      return outputPath;
    }
    await win.webContents.executeJavaScript(`
      document.body.style.background = "white";
      document.body.style.margin = "0";
      document.documentElement.style.overflow = "hidden";
    `);
    const image = await win.webContents.capturePage();
    await fs.writeFile(outputPath, image.toPNG());
    return outputPath;
  } catch (_error) {
    await createTextPng(`Receipt screenshot failed.\nSource file: ${filePath}`, outputPath);
    return outputPath;
  } finally {
    win.destroy();
  }
}

async function extractPdfViewerCanvas(win) {
  try {
    const dataUrl = await win.webContents.executeJavaScript(`
      (() => {
        const seen = new Set();
        const all = [];
        function walk(root) {
          if (!root || seen.has(root)) return;
          seen.add(root);
          if (root.querySelectorAll) {
            root.querySelectorAll('*').forEach((el) => {
              if (el.tagName === 'CANVAS') all.push(el);
              if (el.shadowRoot) walk(el.shadowRoot);
            });
          }
        }
        walk(document);
        const canvases = all.filter((c) => c.width > 300 && c.height > 250);
        if (!canvases.length) return '';
        const pageCanvases = canvases.filter((c) => {
          const rect = c.getBoundingClientRect();
          return rect.width > 300 && rect.height > 250 && rect.top > 50;
        });
        const target = (pageCanvases.length ? pageCanvases : canvases)
          .sort((a, b) => (b.width * b.height) - (a.width * a.height))[0];
        return target.toDataURL('image/png');
      })()
    `);
    if (!dataUrl || !dataUrl.startsWith('data:image/png;base64,')) return null;
    return Buffer.from(dataUrl.split(',')[1], 'base64');
  } catch (_error) {
    return null;
  }
}

function cropPdfViewerToReceipt(image) {
  const size = image.getSize();
  const bitmap = image.toBitmap();
  const sample = 4;
  const sw = Math.floor(size.width / sample);
  const sh = Math.floor(size.height / sample);
  const isPageLike = (sx, sy) => {
    const x = sx * sample;
    const y = sy * sample;
    const idx = (y * size.width + x) * 4;
    const b = bitmap[idx];
    const g = bitmap[idx + 1];
    const r = bitmap[idx + 2];
    const bright = r > 218 && g > 218 && b > 218;
    const redInvoiceLine = r > 130 && g < 110 && b < 110;
    const darkText = r < 80 && g < 80 && b < 80;
    return bright || redInvoiceLine || darkText;
  };

  const visited = new Uint8Array(sw * sh);
  const stack = [];
  let best = null;
  for (let sy = Math.floor(70 / sample); sy < sh; sy += 1) {
    for (let sx = 0; sx < sw; sx += 1) {
      const pos = sy * sw + sx;
      if (visited[pos] || !isPageLike(sx, sy)) continue;
      let count = 0;
      let minX = sx;
      let maxX = sx;
      let minY = sy;
      let maxY = sy;
      visited[pos] = 1;
      stack.push([sx, sy]);
      while (stack.length) {
        const [cx, cy] = stack.pop();
        count += 1;
        if (cx < minX) minX = cx;
        if (cx > maxX) maxX = cx;
        if (cy < minY) minY = cy;
        if (cy > maxY) maxY = cy;
        for (const [nx, ny] of [[cx + 1, cy], [cx - 1, cy], [cx, cy + 1], [cx, cy - 1]]) {
          if (nx < 0 || ny < 0 || nx >= sw || ny >= sh) continue;
          const npos = ny * sw + nx;
          if (visited[npos] || !isPageLike(nx, ny)) continue;
          visited[npos] = 1;
          stack.push([nx, ny]);
        }
      }
      const width = (maxX - minX + 1) * sample;
      const height = (maxY - minY + 1) * sample;
      const area = width * height;
      if (width > 350 && height > 250 && (!best || area > best.area)) {
        best = { minX, maxX, minY, maxY, area };
      }
    }
  }
  if (!best) return null;
  const crop = {
    x: Math.max(0, best.minX * sample - 10),
    y: Math.max(0, best.minY * sample - 10),
    width: Math.min(size.width - Math.max(0, best.minX * sample - 10), (best.maxX - best.minX + 1) * sample + 20),
    height: Math.min(size.height - Math.max(0, best.minY * sample - 10), (best.maxY - best.minY + 1) * sample + 20),
  };
  if (crop.width < 300 || crop.height < 250) return null;
  return nativeImage.createFromBuffer(image.toPNG()).crop(crop);
}

function cropWhiteReceiptPage(image) {
  const size = image.getSize();
  const bitmap = image.toBitmap();
  const sample = 3;
  const sw = Math.floor(size.width / sample);
  const sh = Math.floor(size.height / sample);
  const visited = new Uint8Array(sw * sh);
  const stack = [];
  const isWhite = (sx, sy) => {
    const x = sx * sample;
    const y = sy * sample;
    const idx = (y * size.width + x) * 4;
    const b = bitmap[idx];
    const g = bitmap[idx + 1];
    const r = bitmap[idx + 2];
    return r > 235 && g > 235 && b > 235;
  };
  let best = null;
  for (let sy = Math.floor(70 / sample); sy < sh; sy += 1) {
    for (let sx = 0; sx < sw; sx += 1) {
      const pos = sy * sw + sx;
      if (visited[pos] || !isWhite(sx, sy)) continue;
      visited[pos] = 1;
      stack.push([sx, sy]);
      let count = 0;
      let minX = sx;
      let maxX = sx;
      let minY = sy;
      let maxY = sy;
      while (stack.length) {
        const [cx, cy] = stack.pop();
        count += 1;
        if (cx < minX) minX = cx;
        if (cx > maxX) maxX = cx;
        if (cy < minY) minY = cy;
        if (cy > maxY) maxY = cy;
        for (const [nx, ny] of [[cx + 1, cy], [cx - 1, cy], [cx, cy + 1], [cx, cy - 1]]) {
          if (nx < 0 || ny < 0 || nx >= sw || ny >= sh) continue;
          const npos = ny * sw + nx;
          if (visited[npos] || !isWhite(nx, ny)) continue;
          visited[npos] = 1;
          stack.push([nx, ny]);
        }
      }
      const width = (maxX - minX + 1) * sample;
      const height = (maxY - minY + 1) * sample;
      const area = width * height;
      const plausiblePage = width > 600 && height > 350 && area > 250000;
      if (plausiblePage && (!best || area > best.area)) {
        best = { minX, maxX, minY, maxY, area };
      }
    }
  }
  if (!best) return null;
  const x = Math.max(0, best.minX * sample - 24);
  const y = Math.max(0, best.minY * sample - 24);
  const crop = {
    x,
    y,
    width: Math.min(size.width - x, (best.maxX - best.minX + 1) * sample + 48),
    height: Math.min(size.height - y, (best.maxY - best.minY + 1) * sample + 48),
  };
  return nativeImage.createFromBuffer(image.toPNG()).crop(crop);
}

async function buildExcel(receipts, excelPath, claimant) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(DEFAULT_TEMPLATE);
  const sheet = workbook.getWorksheet(1);
  sheet.getCell('F14').value = claimant.claimantName;
  sheet.getCell('C15').value = claimant.department;
  sheet.getCell('E17').value = claimant.staffStudentNo;
  sheet.getCell('D20').value = claimant.telephoneNo;

  const startRow = 29;
  const defaultTotalRow = 55;
  const maxTemplateItems = Math.floor((defaultTotalRow - startRow) / 2);
  const extraItems = Math.max(0, receipts.length - maxTemplateItems);
  if (extraItems > 0) {
    sheet.duplicateRow(defaultTotalRow - 2, extraItems * 2, true);
  }
  const totalRow = defaultTotalRow + extraItems * 2;
  const lastAmountRow = receipts.length ? startRow + (receipts.length - 1) * 2 : startRow;

  receipts.forEach((receipt, i) => {
    const row = startRow + i * 2;
    sheet.getCell(`A${row}`).value = receipt.description;
    sheet.getCell(`J${row}`).value = String(i + 1);
    const amountCell = sheet.getCell(`K${row}`);
    amountCell.value = Number(receipt.hkdAmount) || 0;
    amountCell.numFmt = '"HK$" #,##0.00';
  });
  const totalCell = sheet.getCell(`K${totalRow}`);
  totalCell.value = { formula: `SUM(K${startRow}:K${lastAmountRow})` };
  totalCell.numFmt = '"HK$" #,##0.00';
  await workbook.xlsx.writeFile(excelPath);
}

async function buildWord(receipts, wordPath, claimant) {
  const children = [
    new Paragraph({
      text: 'Reimbursement Supporting Documents',
      heading: HeadingLevel.TITLE,
      alignment: AlignmentType.CENTER,
    }),
    new Paragraph(`Claimant / Payee Name: ${claimant.claimantName}`),
    new Paragraph(`Department: ${claimant.department}`),
    new Paragraph(`Staff/Student No.: ${claimant.staffStudentNo}`),
    new Paragraph(`Telephone No.: ${claimant.telephoneNo}`),
    new Paragraph(`Generated Date: ${formatDate(new Date())}`),
    new Paragraph(''),
  ];

  for (const receipt of receipts) {
    children.push(new Paragraph({
      text: `${receipt.index}. ${receipt.description}`,
      heading: HeadingLevel.HEADING_1,
    }));
    children.push(new Paragraph(`Invoice Number: ${receipt.invoiceNumber || 'Please verify'}`));
    children.push(new Paragraph(`Invoice Date: ${receipt.invoiceDate || 'Please verify'}`));
    children.push(new Paragraph(`Voucher: ${receipt.index}`));
    children.push(new Paragraph(`Amount: HK$ ${formatMoney(receipt.hkdAmount)}`));
    if (receipt.currency === 'RMB') {
      children.push(new Paragraph(`Original Amount: RMB ${formatMoney(receipt.originalAmount)}`));
      children.push(new Paragraph(`Exchange rate ${receipt.hkabRate || 'Please verify'}`));
      children.push(new Paragraph(`RMB ${formatMoney(receipt.originalAmount)} × ${receipt.hkabRate || 'rate'} / 100 = HK$ ${formatMoney(receipt.hkdAmount)}`));
      children.push(new Paragraph({
        children: [
          new TextRun('Exchange-rate source: '),
          new ExternalHyperlink({
            link: HKAB_URL,
            children: [new TextRun({ text: 'Exchange Rates | The Hong Kong Association of Banks', style: 'Hyperlink' })],
          }),
        ],
      }));
    }
    if (receipt.warning) children.push(new Paragraph(`Review Note: ${receipt.warning}`));
    children.push(new Paragraph('Original receipt screenshot:'));
    await addImageIfPossible(children, receipt.receiptScreenshot || receipt.copiedPath, 620);
    if (receipt.rateScreenshot) {
      children.push(new Paragraph('HKAB exchange-rate screenshot:'));
      await addImageIfPossible(children, receipt.rateScreenshot, 620);
    }
    children.push(new Paragraph(''));
  }

  const doc = new Document({
    sections: [{ properties: {}, children }],
  });
  const buffer = await Packer.toBuffer(doc);
  await fs.writeFile(wordPath, buffer);
}

async function addImageIfPossible(children, filePath, width) {
  try {
    const ext = path.extname(filePath).toLowerCase();
    if (!['.png', '.jpg', '.jpeg'].includes(ext)) {
      children.push(new Paragraph(`Source file: ${path.basename(filePath)}`));
      return;
    }
    const image = await fs.readFile(filePath);
    children.push(new Paragraph({
      children: [
        new ImageRun({
          data: image,
          transformation: { width, height: Math.round(width * 0.72) },
        }),
      ],
    }));
  } catch (_error) {
    children.push(new Paragraph(`Source file: ${path.basename(filePath)}`));
  }
}

async function createTextPng(text, outputPath) {
  // Tiny valid PNG fallback; the explanatory text is stored beside it.
  const base64 = 'iVBORw0KGgoAAAANSUhEUgAAAoAAAAHgCAIAAADhM9qrAAAACXBIWXMAAAsTAAALEwEAmpwYAAAFGElEQVR4nO3VMQEAIAzAsIF/z0MNDkQK6d6Z2QMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA4J4B3lQAAQwplXwAAAAASUVORK5CYII=';
  await fs.writeFile(outputPath, Buffer.from(base64, 'base64'));
  await fs.writeFile(outputPath.replace(/\.png$/i, '.txt'), text, 'utf8');
  return outputPath;
}

function normalizeAmount(value) {
  const number = Number(String(value).replace(/,/g, ''));
  return Number.isFinite(number) ? number.toFixed(2) : '';
}

function formatMoney(value) {
  const number = Number(String(value || 0).replace(/,/g, ''));
  return Number.isFinite(number)
    ? number.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
    : '0.00';
}

function formatDate(date) {
  return `${date.getFullYear()}${pad2(date.getMonth() + 1)}${pad2(date.getDate())}`;
}

function toDateInput(date) {
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
}

function pad2(value) {
  return String(value).padStart(2, '0');
}

function titleCase(value) {
  return value.toLowerCase().replace(/\b[a-z]/g, (letter) => letter.toUpperCase());
}

function sanitizeFileName(value) {
  return toEnglishDescription(value).replace(/[<>:"/\\|?*]/g, '').replace(/\s+/g, '_').slice(0, 80);
}

function sanitizePersonName(value) {
  return String(value || '')
    .trim()
    .replace(/[<>:"/\\|?*]/g, '')
    .replace(/[^\w\s.-]/g, '')
    .replace(/\s+/g, '_')
    .toUpperCase()
    .slice(0, 80);
}

function uniquePath(candidate) {
  if (!fss.existsSync(candidate)) return candidate;
  const parsed = path.parse(candidate);
  let i = 2;
  while (true) {
    const next = path.join(parsed.dir, `${parsed.name}_${i}${parsed.ext}`);
    if (!fss.existsSync(next)) return next;
    i += 1;
  }
}

function joinWarning(existing, addition) {
  return [existing, addition].filter(Boolean).join(' ');
}

function normalizeClaimantInfo(info = {}) {
  return {
    claimantName: String(info.claimantName || '').trim(),
    department: String(info.department || '').trim(),
    staffStudentNo: String(info.staffStudentNo || '').trim(),
    telephoneNo: String(info.telephoneNo || '').trim(),
  };
}
