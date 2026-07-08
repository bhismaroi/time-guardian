// MOL Time Guardian — Cloudflare Pages UI controller.
// Drives the form, summary KPIs, download link, and warnings for the
// vanilla-JS attendance bundle that uses exceljs + browserCompiler.

const form = document.getElementById('upload-form');
const fingerprintInput = document.getElementById('fingerprint');
const onlineInput = document.getElementById('online');
const generateButton = document.getElementById('generate-btn');
const resetButton = document.getElementById('reset-btn');
const statusEl = document.getElementById('status');
const statusPill = document.getElementById('status-pill');
const statusPillText = document.getElementById('status-pill-text');
const summaryEl = document.getElementById('summary');
const downloadEl = document.getElementById('download');
const warningsEl = document.getElementById('warnings');
const footerCopy = document.getElementById('footer-copy');

let lastDownloadUrl = null;

if (footerCopy) {
  footerCopy.textContent = `© ${new Date().getFullYear()} MOL Group · All rights reserved`;
}

function setBusy(isBusy) {
  generateButton.disabled = isBusy;
  resetButton.disabled = isBusy;
  const label = generateButton.querySelector('.btn__label');
  if (label) {
    label.textContent = isBusy ? 'Compiling…' : 'Compile Attendance';
  }
}

function setStatus(message) {
  statusEl.textContent = message;
}

function setStatusPill(visible, label) {
  if (!statusPill) return;
  statusPill.hidden = !visible;
  if (label && statusPillText) {
    statusPillText.textContent = label;
  }
}

function clearOutput() {
  summaryEl.innerHTML = '';
  downloadEl.innerHTML = '';
  warningsEl.innerHTML = '';
  warningsEl.hidden = true;
  setStatusPill(false);
  if (lastDownloadUrl) {
    URL.revokeObjectURL(lastDownloadUrl);
    lastDownloadUrl = null;
  }
}

function renderSummary(summary) {
  const items = [
    { label: 'Month', value: summary.month, muted: !summary.month },
    { label: 'Employees', value: summary.employees, muted: false },
    { label: 'Matched', value: summary.matchedEmployees, muted: false },
    {
      label: 'Fingerprint only',
      value: summary.fingerprintOnlyEmployees,
      muted: !summary.fingerprintOnlyEmployees,
    },
    {
      label: 'Online only',
      value: summary.onlineOnlyEmployees,
      muted: !summary.onlineOnlyEmployees,
    },
  ];

  summaryEl.innerHTML = items
    .map(
      (item) => `
        <div class="kpi${item.muted ? ' kpi--muted' : ''}">
          <div class="kpi__label">${item.label}</div>
          <div class="kpi__value">${item.value || 0}</div>
        </div>`,
    )
    .join('');
}

function renderDownload(fileName) {
  const link = document.createElement('a');
  link.href = lastDownloadUrl;
  link.download = fileName;
  link.innerHTML = `
    <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
      <polyline points="7 10 12 15 17 10" />
      <line x1="12" y1="15" x2="12" y2="3" />
    </svg>
    <span>Download ${fileName}</span>
  `;
  downloadEl.appendChild(link);
}

function renderWarnings(warnings) {
  if (!warnings.length) {
    warningsEl.hidden = true;
    warningsEl.innerHTML = '';
    return;
  }
  warningsEl.hidden = false;
  warningsEl.innerHTML = warnings
    .map((warning) => `<div class="warning-item">${warning}</div>`)
    .join('');
}

form.addEventListener('submit', async (event) => {
  event.preventDefault();

  if (!fingerprintInput.files[0] || !onlineInput.files[0]) {
    setStatus('Please choose both Excel files first.');
    return;
  }

  const fingerprintFile = fingerprintInput.files[0];
  const onlineFile = onlineInput.files[0];

  if (!fingerprintFile.name.match(/\.(xlsx|xls)$/i) || !onlineFile.name.match(/\.(xlsx|xls)$/i)) {
    setStatus('Please select valid Excel files (.xlsx or .xls).');
    return;
  }

  setBusy(true);
  clearOutput();
  setStatus('Compiling attendance workbook in your browser…');

  try {
    const payload = await window.AttendanceCompiler.buildCompiledWorkbookFromFiles(
      fingerprintFile,
      onlineFile,
      { onProgress: (msg) => { setStatus(msg); } },
    );
    setStatus('Generating Excel file…');
    const buffer = await payload.workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });

    if (lastDownloadUrl) {
      URL.revokeObjectURL(lastDownloadUrl);
    }
    lastDownloadUrl = URL.createObjectURL(blob);

    setStatus(`Report ready — ${payload.summary.month || 'attendance period'}.`);
    setStatusPill(true, `${payload.summary.employees} employees`);
    renderSummary(payload.summary);
    renderDownload(payload.fileName);
    renderWarnings(payload.warnings);
  } catch (error) {
    setStatus(error.message || 'Something went wrong while compiling the report.');
  } finally {
    setBusy(false);
  }
});

resetButton.addEventListener('click', () => {
  // form.reset() runs first (type="reset"); our handler cleans output state.
  clearOutput();
  setStatus('Waiting for files.');
});