const excelInput = document.querySelector('#excel-input');
const pdfInput = document.querySelector('#pdf-input');
const excelLabel = document.querySelector('#excel-label');
const pdfList = document.querySelector('#pdf-list');
const pdfCount = document.querySelector('#pdf-count');
const runButton = document.querySelector('#run-button');
const clearButton = document.querySelector('#clear-button');
const dropZone = document.querySelector('#drop-zone');
const toleranceInput = document.querySelector('#tolerance');
const statusBadge = document.querySelector('#status-badge');
const emptyState = document.querySelector('#empty-state');
const runningState = document.querySelector('#running-state');
const resultState = document.querySelector('#result-state');
const errorState = document.querySelector('#error-state');
const logWindow = document.querySelector('#log-window');
const downloadButton = document.querySelector('#download-button');

let excelFile = null;
let pdfFiles = [];
let currentJobId = null;
let startedAt = null;
let pollTimer = null;
let clockTimer = null;

function formatSize(bytes) {
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

function refreshIcons() {
  if (window.lucide) window.lucide.createIcons({ attrs: { 'stroke-width': 1.8 } });
}

function updateReadyState() {
  runButton.disabled = !excelFile || pdfFiles.length === 0 || currentJobId !== null;
}

function escapeHtml(value) {
  const div = document.createElement('div');
  div.textContent = value;
  return div.innerHTML;
}

function renderFiles() {
  excelLabel.textContent = excelFile ? excelFile.name : '选择一份序时账';
  pdfCount.textContent = `${pdfFiles.length} 份 PDF`;
  pdfList.innerHTML = '';
  pdfFiles.forEach((file, index) => {
    const row = document.createElement('div');
    row.className = 'file-row';
    row.innerHTML = `<i data-lucide="file-text"></i><span title="${escapeHtml(file.name)}">${escapeHtml(file.name)}</span><small>${formatSize(file.size)}</small><button type="button" title="移除文件" data-index="${index}"><i data-lucide="x"></i></button>`;
    pdfList.appendChild(row);
  });
  pdfList.querySelectorAll('button').forEach(button => {
    button.addEventListener('click', () => {
      pdfFiles.splice(Number(button.dataset.index), 1);
      renderFiles();
    });
  });
  updateReadyState();
  refreshIcons();
}

function addPdfFiles(files) {
  const existing = new Set(pdfFiles.map(file => `${file.name}:${file.size}:${file.lastModified}`));
  [...files].filter(file => file.name.toLowerCase().endsWith('.pdf')).forEach(file => {
    const key = `${file.name}:${file.size}:${file.lastModified}`;
    if (!existing.has(key)) {
      pdfFiles.push(file);
      existing.add(key);
    }
  });
  renderFiles();
}

excelInput.addEventListener('change', () => {
  excelFile = excelInput.files[0] || null;
  renderFiles();
});
pdfInput.addEventListener('change', () => addPdfFiles(pdfInput.files));

['dragenter', 'dragover'].forEach(eventName => dropZone.addEventListener(eventName, event => {
  event.preventDefault();
  dropZone.classList.add('dragging');
}));
['dragleave', 'drop'].forEach(eventName => dropZone.addEventListener(eventName, event => {
  event.preventDefault();
  dropZone.classList.remove('dragging');
}));
dropZone.addEventListener('drop', event => addPdfFiles(event.dataTransfer.files));

clearButton.addEventListener('click', () => {
  if (currentJobId) return;
  excelFile = null;
  pdfFiles = [];
  excelInput.value = '';
  pdfInput.value = '';
  resetActivity();
  renderFiles();
});

document.querySelector('#decrease-tolerance').addEventListener('click', () => {
  toleranceInput.value = Math.max(0, Number(toleranceInput.value || 0) - 1);
});
document.querySelector('#increase-tolerance').addEventListener('click', () => {
  toleranceInput.value = Math.min(366, Number(toleranceInput.value || 0) + 1);
});

function setStatus(label, className) {
  statusBadge.textContent = label;
  statusBadge.className = `status-badge ${className}`;
}

function showState(state) {
  emptyState.classList.toggle('hidden', state !== 'empty');
  runningState.classList.toggle('hidden', state !== 'running');
  resultState.classList.toggle('hidden', state !== 'result');
  errorState.classList.toggle('hidden', state !== 'error');
}

function resetActivity() {
  currentJobId = null;
  showState('empty');
  setStatus('等待文件', 'idle');
  logWindow.innerHTML = '';
  clearInterval(pollTimer);
  clearInterval(clockTimer);
}

function renderLogs(logs) {
  const current = [...logWindow.querySelectorAll('.log-line')].map(item => item.textContent);
  if (current.length === logs.length && current.every((value, index) => value === logs[index])) return;
  logWindow.innerHTML = logs.map(line => `<div class="log-line">${escapeHtml(line)}</div>`).join('');
  logWindow.scrollTop = logWindow.scrollHeight;
  const latest = logs.at(-1);
  if (latest) document.querySelector('#progress-title').textContent = latest.startsWith('OCR识别') ? '正在识别扫描对账单' : '正在核对流水';
}

function startClock() {
  startedAt = Date.now();
  clearInterval(clockTimer);
  clockTimer = setInterval(() => {
    const seconds = Math.floor((Date.now() - startedAt) / 1000);
    document.querySelector('#elapsed-time').textContent = `${String(Math.floor(seconds / 60)).padStart(2, '0')}:${String(seconds % 60).padStart(2, '0')}`;
  }, 1000);
}

function renderResult(job) {
  const stats = job.stats;
  document.querySelector('#statement-count').textContent = stats.statement_count.toLocaleString();
  document.querySelector('#exact-count').textContent = stats.exact_count.toLocaleString();
  document.querySelector('#tolerance-count').textContent = stats.tolerance_count.toLocaleString();
  document.querySelector('#unmatched-count').textContent = stats.unmatched_count.toLocaleString();
  document.querySelector('#result-name').textContent = job.result_name;
  const warnings = [];
  if (stats.unrecognized_pdf_count) warnings.push(`${stats.unrecognized_pdf_count} 份 PDF 未识别`);
  if (stats.unknown_bank_count) warnings.push(`${stats.unknown_bank_count} 笔流水未识别银行`);
  document.querySelector('#result-warnings').innerHTML = warnings.map(item => `<p class="warning-line"><i data-lucide="triangle-alert"></i>${item}</p>`).join('');
  downloadButton.href = `/api/jobs/${job.id}/download`;
  downloadButton.dataset.jobId = job.id;
  refreshIcons();
}

downloadButton.addEventListener('click', async event => {
  event.preventDefault();
  const label = downloadButton.querySelector('span');
  const saveNote = document.querySelector('#save-note');
  const originalLabel = label.textContent;
  label.textContent = '正在保存';
  saveNote.textContent = '';
  downloadButton.classList.add('busy');
  try {
    const response = await fetch(`/api/jobs/${downloadButton.dataset.jobId}/save-local`, { method: 'POST' });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || '保存失败');
    label.textContent = '结果已保存';
    saveNote.textContent = '已保存到桌面“流水核对结果”文件夹';
  } catch (error) {
    label.textContent = '保存失败，请重试';
    saveNote.textContent = error.message;
  } finally {
    downloadButton.classList.remove('busy');
    window.setTimeout(() => { label.textContent = originalLabel; }, 1800);
  }
});

async function pollJob() {
  try {
    const response = await fetch(`/api/jobs/${currentJobId}`);
    const job = await response.json();
    if (!response.ok) throw new Error(job.error || '无法读取任务状态');
    renderLogs(job.logs || []);
    if (job.status === 'completed') {
      clearInterval(pollTimer);
      clearInterval(clockTimer);
      showState('result');
      setStatus('已完成', 'completed');
      renderResult(job);
      currentJobId = null;
      updateReadyState();
    } else if (job.status === 'failed') {
      clearInterval(pollTimer);
      clearInterval(clockTimer);
      showState('error');
      setStatus('失败', 'failed');
      document.querySelector('#error-message').textContent = job.error || job.logs?.at(-1) || '未知错误';
      currentJobId = null;
      updateReadyState();
    } else {
      setStatus(job.status === 'queued' ? '排队中' : '处理中', 'running');
    }
  } catch (error) {
    clearInterval(pollTimer);
    clearInterval(clockTimer);
    showState('error');
    setStatus('连接失败', 'failed');
    document.querySelector('#error-message').textContent = error.message;
    currentJobId = null;
    updateReadyState();
  }
}

runButton.addEventListener('click', async () => {
  if (!excelFile || !pdfFiles.length) return;
  const formData = new FormData();
  formData.append('excel', excelFile);
  pdfFiles.forEach(file => formData.append('pdfs', file));
  formData.append('date_tolerance', toleranceInput.value || '62');
  showState('running');
  setStatus('上传中', 'running');
  document.querySelector('#progress-title').textContent = '正在传入本机服务';
  logWindow.innerHTML = '';
  runButton.disabled = true;
  startClock();
  try {
    const response = await fetch('/api/jobs', { method: 'POST', body: formData });
    const job = await response.json();
    if (!response.ok) throw new Error(job.error || '无法创建任务');
    currentJobId = job.id;
    renderLogs(job.logs || []);
    setStatus('排队中', 'running');
    pollTimer = setInterval(pollJob, 800);
    await pollJob();
  } catch (error) {
    clearInterval(clockTimer);
    showState('error');
    setStatus('失败', 'failed');
    document.querySelector('#error-message').textContent = error.message;
    currentJobId = null;
    updateReadyState();
  }
});

window.addEventListener('DOMContentLoaded', () => {
  refreshIcons();
  renderFiles();
});
