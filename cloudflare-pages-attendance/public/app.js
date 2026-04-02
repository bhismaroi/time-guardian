const form = document.getElementById('upload-form');
const fingerprintInput = document.getElementById('fingerprint');
const onlineInput = document.getElementById('online');
const generateButton = document.getElementById('generate-btn');
const resetButton = document.getElementById('reset-btn');
const statusEl = document.getElementById('status');
const summaryEl = document.getElementById('summary');
const downloadEl = document.getElementById('download');
const warningsEl = document.getElementById('warnings');

form.addEventListener('submit', async (event) => {
  event.preventDefault();

  if (!fingerprintInput.files[0] || !onlineInput.files[0]) {
    statusEl.textContent = 'Please choose both Excel files first.';
    return;
  }

  setBusy(true);
  clearOutput();
  statusEl.textContent = 'Compiling attendance workbook in your browser...';

  try {
    const payload = await window.AttendanceCompiler.buildCompiledWorkbookFromFiles(
      fingerprintInput.files[0],
      onlineInput.files[0],
    );
    const buffer = await payload.workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const downloadUrl = URL.createObjectURL(blob);

    statusEl.textContent = 'Report ready.';
    summaryEl.innerHTML = [
      `Month: <strong>${payload.summary.month}</strong>`,
      `Employees: <strong>${payload.summary.employees}</strong>`,
      `Matched: <strong>${payload.summary.matchedEmployees}</strong>`,
      `Fingerprint only: <strong>${payload.summary.fingerprintOnlyEmployees}</strong>`,
      `Online only: <strong>${payload.summary.onlineOnlyEmployees}</strong>`,
    ].join(' | ');

    const link = document.createElement('a');
    link.href = downloadUrl;
    link.download = payload.fileName;
    link.textContent = `Download ${payload.fileName}`;
    downloadEl.appendChild(link);

    if (payload.warnings.length) {
      for (const warning of payload.warnings) {
        const item = document.createElement('div');
        item.className = 'warning-item';
        item.textContent = warning;
        warningsEl.appendChild(item);
      }
    } else {
      const item = document.createElement('div');
      item.className = 'warning-item';
      item.textContent = 'No matching warnings found in these files.';
      warningsEl.appendChild(item);
    }
  } catch (error) {
    statusEl.textContent = error.message;
  } finally {
    setBusy(false);
  }
});

resetButton.addEventListener('click', async () => {
  form.reset();
  clearOutput();
  statusEl.textContent = 'Waiting for files.';
});

function setBusy(isBusy) {
  generateButton.disabled = isBusy;
  resetButton.disabled = isBusy;
}

function clearOutput() {
  summaryEl.innerHTML = '';
  downloadEl.innerHTML = '';
  warningsEl.innerHTML = '';
}
