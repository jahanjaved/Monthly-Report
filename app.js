let manifest = null;
let uploadedFiles = new Map();

const fileInput = document.getElementById('fileInput');
const validateBtn = document.getElementById('validateBtn');
const generateBtn = document.getElementById('generateBtn');
const expectedList = document.getElementById('expectedList');
const statusBox = document.getElementById('statusBox');
const results = document.getElementById('results');
const dropzone = document.getElementById('dropzone');

async function loadManifest() {
  const paths = ['assets/manifest.json', 'manifest.json'];
  let lastErr = null;
  for (const path of paths) {
    try {
      const res = await fetch(path);
      if (!res.ok) throw new Error(`HTTP ${res.status} for ${path}`);
      manifest = await res.json();
      renderExpectedList();
      return;
    } catch (err) {
      lastErr = err;
    }
  }
  throw lastErr || new Error('Manifest could not be loaded.');
}

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function normalizeName(name) {
  return String(name || '')
    .toLowerCase()
    .replace(/\.xlsx$/i, '')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function looksLikeBaseTemplate(name) {
  const n = normalizeName(name);
  return n.includes('orignal') || n.includes('original') || n.includes('master') || n.includes('reference');
}

function renderExpectedList() {
  expectedList.innerHTML = '';
  if (!manifest || !Array.isArray(manifest.expectedFiles) || !manifest.expectedFiles.length) {
    const li = document.createElement('li');
    li.textContent = 'Flexible mode active. Upload one or more monthly Excel files.';
    li.className = 'ok';
    expectedList.appendChild(li);
    return;
  }

  manifest.expectedFiles.forEach(name => {
    const li = document.createElement('li');
    li.dataset.name = name;
    li.textContent = name;
    li.className = uploadedFiles.has(name) ? 'ok' : 'missing';
    expectedList.appendChild(li);
  });

  const flex = document.createElement('li');
  flex.className = 'ok';
  flex.innerHTML = 'Flexible mode: exact names are <strong>not required</strong>. The website can start from any uploaded workbook and build a master draft.';
  expectedList.appendChild(flex);
}

function setStatus(text, cls = 'status-idle') {
  statusBox.className = 'status ' + cls;
  statusBox.textContent = text;
}

function addResult(title, body) {
  const card = document.createElement('div');
  card.className = 'result-card';
  card.innerHTML = `<h3>${escapeHtml(title)}</h3><p>${body}</p>`;
  results.appendChild(card);
}

function refreshUploaded(files) {
  for (const file of files) {
    if (file.name.toLowerCase().endsWith('.xlsx')) {
      uploadedFiles.set(file.name, file);
    }
  }
  renderExpectedList();
  results.innerHTML = '';
  const count = uploadedFiles.size;
  if (count === 0) {
    setStatus('No Excel files loaded yet.', 'status-warn');
    generateBtn.disabled = true;
    return;
  }
  setStatus(`${count} Excel file(s) loaded. Flexible mode is ready.`, 'status-ok');
  generateBtn.disabled = false;
}

fileInput.addEventListener('change', (e) => refreshUploaded(e.target.files));

['dragenter', 'dragover'].forEach(evt => {
  dropzone.addEventListener(evt, e => {
    e.preventDefault();
    dropzone.classList.add('drag');
  });
});

['dragleave', 'drop'].forEach(evt => {
  dropzone.addEventListener(evt, e => {
    e.preventDefault();
    dropzone.classList.remove('drag');
  });
});

dropzone.addEventListener('drop', e => {
  refreshUploaded(e.dataTransfer.files);
});

async function inspectWorkbook(file) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await file.arrayBuffer());
  return wb;
}

function scoreWorkbook(fileName, workbook) {
  let score = workbook.worksheets.length * 1000;
  if (looksLikeBaseTemplate(fileName)) score += 100000;
  for (const ws of workbook.worksheets) {
    score += (ws.actualRowCount || 0) * 2;
    score += (ws.actualColumnCount || 0);
  }
  return score;
}

validateBtn.addEventListener('click', async () => {
  results.innerHTML = '';
  if (uploadedFiles.size === 0) {
    setStatus('Please upload at least one Excel file.', 'status-bad');
    return;
  }

  setStatus('Validating uploaded workbook pack...', 'status-idle');
  try {
    const inspected = [];
    for (const [name, file] of uploadedFiles.entries()) {
      const wb = await inspectWorkbook(file);
      inspected.push({ name, wb });
      addResult(name, `Loaded successfully with ${wb.worksheets.length} sheet(s).`);
    }

    inspected.sort((a, b) => scoreWorkbook(b.name, b.wb) - scoreWorkbook(a.name, a.wb));
    const base = inspected[0];
    addResult('Base workbook candidate', `${escapeHtml(base.name)} was selected as the best starting point for the master workbook.`);
    setStatus('Validation successful. You can now generate the master workbook.', 'status-ok');
  } catch (err) {
    console.error(err);
    setStatus('Validation failed while opening one of the workbooks.', 'status-bad');
    addResult('Open error', escapeHtml(String(err)));
  }
});

async function generate() {
  results.innerHTML = '';
  if (uploadedFiles.size === 0) {
    setStatus('Please upload at least one Excel file.', 'status-bad');
    return;
  }

  setStatus('Opening uploaded workbooks...', 'status-idle');

  const loaded = [];
  for (const [name, file] of uploadedFiles.entries()) {
    const wb = await inspectWorkbook(file);
    loaded.push({ name, file, wb });
  }

  loaded.sort((a, b) => scoreWorkbook(b.name, b.wb) - scoreWorkbook(a.name, a.wb));
  const baseEntry = loaded[0];
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await baseEntry.file.arrayBuffer());

  addResult('Base workbook', `${escapeHtml(baseEntry.name)} was used as the base master workbook.`);

  const donorEntries = loaded.slice(1);
  let addedSheets = 0;
  const existingSheetNames = new Set(workbook.worksheets.map(ws => ws.name));

  for (const donor of donorEntries) {
    for (const sourceWs of donor.wb.worksheets) {
      if (existingSheetNames.has(sourceWs.name)) continue;
      const targetWs = workbook.addWorksheet(sourceWs.name, {
        properties: { ...sourceWs.properties },
        state: sourceWs.state,
        pageSetup: { ...sourceWs.pageSetup },
        views: sourceWs.views ? JSON.parse(JSON.stringify(sourceWs.views)) : undefined,
      });

      sourceWs.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        const targetRow = targetWs.getRow(rowNumber);
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const targetCell = targetRow.getCell(colNumber);
          if (cell.value && typeof cell.value === 'object' && cell.value.formula) {
            targetCell.value = { formula: cell.value.formula, result: cell.value.result };
          } else {
            targetCell.value = cell.value;
          }
          if (cell.style) {
            targetCell.style = JSON.parse(JSON.stringify(cell.style));
          }
        });
        targetRow.commit();
      });

      sourceWs.columns?.forEach((col, idx) => {
        if (!col) return;
        const targetCol = targetWs.getColumn(idx + 1);
        targetCol.width = col.width;
        targetCol.hidden = col.hidden;
        if (col.style) targetCol.style = JSON.parse(JSON.stringify(col.style));
      });

      existingSheetNames.add(sourceWs.name);
      addedSheets += 1;
    }
  }

  addResult('Merge summary', `${donorEntries.length} additional workbook(s) were checked and ${addedSheets} missing sheet(s) were added into the master draft.`);

  let applied = 0;
  if (manifest && Array.isArray(manifest.patches)) {
    for (const patch of manifest.patches) {
      const ws = workbook.getWorksheet(patch.sheet);
      if (!ws) continue;
      const cell = ws.getCell(patch.cell);
      if (patch.formula !== undefined) {
        cell.value = { formula: patch.formula };
      } else if (patch.value !== undefined) {
        cell.value = patch.value;
      }
      applied += 1;
    }
  }

  addResult('Optional patch set', `${applied} known correction(s) were applied where matching sheets existed.`);

  setStatus('Preparing download...', 'status-idle');
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'orignal_combined_browser_flexible.xlsx';
  a.click();
  URL.revokeObjectURL(url);

  setStatus('Done. Flexible master workbook generated.', 'status-ok');
  addResult('Important note', 'This version no longer requires orignal(2).xlsx and no longer forces a fixed number of contractor files. It builds a master draft from whatever monthly workbooks you upload.');
}

generateBtn.addEventListener('click', () => {
  generate().catch(err => {
    console.error(err);
    setStatus('Generation failed.', 'status-bad');
    addResult('Error', escapeHtml(String(err)));
  });
});

loadManifest().catch(err => {
  console.error(err);
  setStatus('Manifest could not be loaded, but flexible upload mode is still available.', 'status-warn');
  manifest = { expectedFiles: [], patches: [] };
  renderExpectedList();
});
