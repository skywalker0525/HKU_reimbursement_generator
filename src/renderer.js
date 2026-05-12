const state = {
  files: [],
  outputBase: '',
  receipts: [],
};

const el = {
  selectFiles: document.getElementById('selectFiles'),
  selectOutput: document.getElementById('selectOutput'),
  analyze: document.getElementById('analyze'),
  generate: document.getElementById('generate'),
  fileCount: document.getElementById('fileCount'),
  outputPath: document.getElementById('outputPath'),
  status: document.getElementById('status'),
  receiptBody: document.getElementById('receiptBody'),
  claimantName: document.getElementById('claimantName'),
  department: document.getElementById('department'),
  staffStudentNo: document.getElementById('staffStudentNo'),
  telephoneNo: document.getElementById('telephoneNo'),
  progressText: document.getElementById('progressText'),
  progressPct: document.getElementById('progressPct'),
  progressBar: document.getElementById('progressBar'),
};

window.reimburse.onProgress((payload) => {
  const pct = Math.max(0, Math.min(100, Number(payload.percent) || 0));
  el.progressBar.value = pct;
  el.progressPct.textContent = `${Math.round(pct)}%`;
  el.progressText.textContent = payload.message || 'Working';
});

el.selectFiles.addEventListener('click', async () => {
  const files = await window.reimburse.selectFiles();
  if (!files.length) return;
  state.files = files;
  state.receipts = [];
  el.fileCount.textContent = `${files.length} file(s)`;
  setStatus('已选择票据。 / Receipt files selected.');
  renderTable();
  updateButtons();
});

el.selectOutput.addEventListener('click', async () => {
  const folder = await window.reimburse.selectOutputFolder();
  if (!folder) return;
  state.outputBase = folder;
  el.outputPath.textContent = folder;
  setStatus('已选择输出位置。 / Output folder selected.');
  updateButtons();
});

el.analyze.addEventListener('click', async () => {
  try {
    setBusy(true, '正在识别票据，请稍等。 / Analyzing receipts, please wait.');
    state.receipts = await window.reimburse.analyzeFiles(state.files);
    renderTable();
    el.progressBar.value = 100;
    el.progressPct.textContent = '100%';
    el.progressText.textContent = 'Analysis complete / 识别完成';
    setStatus('识别完成，请核对并编辑预览表。 / Analysis complete. Please review and edit the preview table.');
  } catch (error) {
    setStatus(`识别失败 / Analysis failed: ${error.message}`, true);
  } finally {
    setBusy(false);
    updateButtons();
  }
});

el.generate.addEventListener('click', async () => {
  syncReceiptsFromTable();
  try {
    setBusy(true, '正在生成 Excel、Word、票据副本和汇率截图。 / Generating Excel, Word, receipt copies, and exchange-rate screenshots.');
    const result = await window.reimburse.generatePackage({
      receipts: state.receipts,
      outputBase: state.outputBase,
      claimantInfo: {
        claimantName: el.claimantName.value.trim(),
        department: el.department.value.trim(),
        staffStudentNo: el.staffStudentNo.value.trim(),
        telephoneNo: el.telephoneNo.value.trim(),
      },
    });
    state.receipts = result.receipts;
    renderTable();
    setStatus(`生成完成 / Generated: ${result.outputRoot}`);
  } catch (error) {
    setStatus(`生成失败 / Generation failed: ${error.message}`, true);
  } finally {
    setBusy(false);
    updateButtons();
  }
});

function updateButtons() {
  el.analyze.disabled = !state.files.length;
  el.generate.disabled = !state.receipts.length || !state.outputBase;
}

function setBusy(isBusy, message = '') {
  for (const button of [el.selectFiles, el.selectOutput, el.analyze, el.generate]) {
    button.disabled = isBusy;
  }
  if (message) setStatus(message);
  if (isBusy) {
    el.progressBar.value = 3;
    el.progressPct.textContent = '3%';
    el.progressText.textContent = message;
  }
  if (!isBusy) updateButtons();
}

function setStatus(message, warn = false) {
  el.status.textContent = message;
  el.status.classList.toggle('warn', warn);
}

function renderTable() {
  if (!state.receipts.length) {
    el.receiptBody.innerHTML = `<tr class="empty"><td colspan="10">${
      state.files.length
        ? '点击“识别票据”。 / Click “Analyze Receipts”.'
        : '请先选择票据。 / Select receipt files first.'
    }</td></tr>`;
    return;
  }
  el.receiptBody.innerHTML = state.receipts.map((receipt, idx) => `
    <tr data-index="${idx}">
      <td>${idx + 1}</td>
      <td class="fileCell">${escapeHtml(receipt.originalName)}</td>
      <td><input data-field="description" value="${escapeAttr(receipt.description)}" /></td>
      <td><input data-field="invoiceNumber" value="${escapeAttr(receipt.invoiceNumber)}" /></td>
      <td><input data-field="invoiceDate" type="date" value="${escapeAttr(receipt.invoiceDate)}" /></td>
      <td>
        <select data-field="currency">
          ${['RMB', 'HKD', 'USD', 'OTHER'].map((currency) => `<option ${receipt.currency === currency ? 'selected' : ''}>${currency}</option>`).join('')}
        </select>
      </td>
      <td><input data-field="originalAmount" value="${escapeAttr(receipt.originalAmount)}" /></td>
      <td><input data-field="hkabRate" value="${escapeAttr(receipt.hkabRate)}" placeholder="Auto" /></td>
      <td><input data-field="hkdAmount" value="${escapeAttr(receipt.hkdAmount)}" placeholder="Auto" /></td>
      <td>${escapeHtml(receipt.warning || receipt.confidence || 'Good')}</td>
    </tr>
  `).join('');
}

function syncReceiptsFromTable() {
  for (const row of [...el.receiptBody.querySelectorAll('tr[data-index]')]) {
    const receipt = state.receipts[Number(row.dataset.index)];
    for (const input of row.querySelectorAll('[data-field]')) {
      receipt[input.dataset.field] = input.value.trim();
    }
  }
}

function escapeHtml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function escapeAttr(value) {
  return escapeHtml(value).replace(/`/g, '&#096;');
}

updateButtons();
