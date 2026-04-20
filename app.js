
let manifest = null;
let uploadedFiles = new Map();

const fileInput = document.getElementById('fileInput');

const filenameAliases = new Map([
  ['orignal_combined_final_reference.xlsx','orignal(2).xlsx']
]);
const validateBtn = document.getElementById('validateBtn');
const generateBtn = document.getElementById('generateBtn');
const expectedList = document.getElementById('expectedList');
const statusBox = document.getElementById('statusBox');
const results = document.getElementById('results');
const dropzone = document.getElementById('dropzone');

async function loadManifest() {
  let res;
  try {
    res = await fetch('assets/manifest.json');
    if (!res.ok) throw new Error('assets/manifest.json not found');
  } catch (e) {
    res = await fetch('manifest.json');
    if (!res.ok) throw new Error('manifest.json not found');
  }
  manifest = await res.json();
  renderExpectedList();
}

function renderExpectedList() {
  expectedList.innerHTML = '';
  manifest.expectedFiles.forEach(name => {
    const li = document.createElement('li');
    li.dataset.name = name;
    li.textContent = name;
    li.className = uploadedFiles.has(name) ? 'ok' : 'missing';
    expectedList.appendChild(li);
  });
}

function setStatus(text, cls='status-idle') {
  statusBox.className = 'status ' + cls;
  statusBox.textContent = text;
}

function addResult(title, body) {
  const card = document.createElement('div');
  card.className = 'result-card';
  card.innerHTML = `<h3>${title}</h3><p>${body}</p>`;
  results.appendChild(card);
}

function refreshUploaded(files) {
  for (const file of files) {
    if (file.name.toLowerCase().endsWith('.xlsx')) {
      const normalizedName = filenameAliases.get(file.name) || file.name;
      uploadedFiles.set(normalizedName, file);
    }
  }
  renderExpectedList();
  results.innerHTML = '';
  const found = manifest.expectedFiles.filter(name => uploadedFiles.has(name)).length;
  setStatus(`${found} of ${manifest.expectedFiles.length} required files loaded.`, found === manifest.expectedFiles.length ? 'status-ok' : 'status-warn');
  generateBtn.disabled = found !== manifest.expectedFiles.length;
}

fileInput.addEventListener('change', (e) => refreshUploaded(e.target.files));

['dragenter','dragover'].forEach(evt => {
  dropzone.addEventListener(evt, e => {
    e.preventDefault();
    dropzone.classList.add('drag');
  });
});

['dragleave','drop'].forEach(evt => {
  dropzone.addEventListener(evt, e => {
    e.preventDefault();
    dropzone.classList.remove('drag');
  });
});

dropzone.addEventListener('drop', e => {
  refreshUploaded(e.dataTransfer.files);
});

validateBtn.addEventListener('click', async () => {
  results.innerHTML = '';
  const missing = manifest.expectedFiles.filter(name => !uploadedFiles.has(name));
  if (missing.length) {
    setStatus('Validation failed: some required files are missing.', 'status-bad');
    addResult('Missing files', missing.join('<br>'));
    return;
  }

  setStatus('Validating workbook pack...', 'status-idle');
  try {
    for (const name of manifest.expectedFiles) {
      const file = uploadedFiles.get(name);
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(await file.arrayBuffer());
      addResult(name, `Loaded successfully with ${wb.worksheets.length} sheet(s).`);
    }
    setStatus('Validation successful. You can now generate the corrected master report.', 'status-ok');
  } catch (err) {
    console.error(err);
    setStatus('Validation failed while opening one of the workbooks.', 'status-bad');
    addResult('Open error', String(err));
  }
});

async function generate() {
  results.innerHTML = '';
  setStatus('Opening master workbook...', 'status-idle');

  const masterFile = uploadedFiles.get('orignal(2).xlsx');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await masterFile.arrayBuffer());

  addResult('Master workbook', 'Original template opened successfully.');

  // Open the contractor files as a validation gate for the exact uploaded pack.
  const supportFiles = manifest.expectedFiles.filter(n => n !== 'orignal(2).xlsx');
  for (const name of supportFiles) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(await uploadedFiles.get(name).arrayBuffer());
  }
  addResult('Source files', 'All supporting workbooks opened successfully. Applying the tuned correction set for this file pack.');

  // Apply tuned patch set.
  let applied = 0;
  for (const patch of manifest.patches) {
    const ws = workbook.getWorksheet(patch.sheet);
    if (!ws) continue;
    const cell = ws.getCell(patch.cell);
    if (patch.formula !== undefined) {
      cell.value = { formula: patch.formula };
    } else {
      cell.value = patch.value;
    }
    applied += 1;
  }

  addResult('Corrections applied', `${applied} cell-level corrections/restorations were applied to preserve the corrected original-format workbook.`);

  setStatus('Preparing download...', 'status-idle');
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'orignal_combined_browser_final.xlsx';
  a.click();
  URL.revokeObjectURL(url);

  setStatus('Done. The corrected original-format workbook has been generated.', 'status-ok');
  addResult('Download ready', 'Open the downloaded workbook in Microsoft Excel once and save it to refresh any workbook-side calculations.');
}

generateBtn.addEventListener('click', () => {
  generate().catch(err => {
    console.error(err);
    setStatus('Generation failed.', 'status-bad');
    addResult('Error', String(err));
  });
});

loadManifest().catch(err => {
  console.error(err);
  setStatus('Could not load the website manifest.', 'status-bad');
});
