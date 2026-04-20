let uploadedFiles = new Map();

const fileInput = document.getElementById('fileInput');
const reviewBtn = document.getElementById('reviewBtn');
const generateBtn = document.getElementById('generateBtn');
const loadedList = document.getElementById('loadedList');
const statusBox = document.getElementById('statusBox');
const results = document.getElementById('results');
const dropzone = document.getElementById('dropzone');

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

function canonical(text) {
  return String(text || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function refreshLoadedList() {
  loadedList.innerHTML = '';
  for (const file of uploadedFiles.values()) {
    const li = document.createElement('li');
    li.className = 'ok';
    li.textContent = file.name;
    loadedList.appendChild(li);
  }
  if (!uploadedFiles.size) {
    const li = document.createElement('li');
    li.className = 'missing';
    li.textContent = 'No Excel files loaded yet.';
    loadedList.appendChild(li);
  }
  generateBtn.disabled = uploadedFiles.size === 0;
  setStatus(
    uploadedFiles.size ? `${uploadedFiles.size} Excel file(s) loaded. Generation is enabled.` : 'Waiting for Excel files.',
    uploadedFiles.size ? 'status-ok' : 'status-idle'
  );
}

function addFiles(files) {
  for (const file of files) {
    if (file.name.toLowerCase().endsWith('.xlsx')) {
      uploadedFiles.set(file.name, file);
    }
  }
  results.innerHTML = '';
  refreshLoadedList();
}

fileInput.addEventListener('change', e => addFiles(e.target.files));
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
dropzone.addEventListener('drop', e => addFiles(e.dataTransfer.files));

function findTemplateCandidate(books) {
  const named = books.find(x => /orignal|original|master|reference/i.test(x.file.name));
  if (named) return named;
  return books.slice().sort((a,b) => (b.workbook.worksheets.length - a.workbook.worksheets.length) || (b.score - a.score))[0];
}

function getSheetAliasKeys(sheetName) {
  const c = canonical(sheetName);
  const map = {
    '2026oshsubmissionschedule': ['2026oshsubmissionschedule'],
    'oshdescriptionsummary': ['oshdescriptionsummary'],
    'oshstatistics': ['oshstatistics'],
    'oshkpi': ['oshkpi'],
    'manpowermanhoursbreakdown': ['manpowermanhoursbreakdown'],
    'dailybreakdowndetails': ['dailybreakdowndetails', 'dailymanhoursbreakdown'],
    'trainingbreakdown': ['trainingbreakdown', 'oshtrainingbreakdownsummary', 'oshinductionbreakdown', 'oshinductiondailybreakdown'],
  };
  return map[c] || [c];
}

function workbookSheetMap(workbook) {
  const map = new Map();
  workbook.worksheets.forEach(ws => map.set(canonical(ws.name), ws));
  return map;
}

function numericValue(v) {
  if (v == null) return null;
  if (typeof v === 'number' && Number.isFinite(v)) return v;
  if (typeof v === 'object') {
    if (typeof v.result === 'number' && Number.isFinite(v.result)) return v.result;
    if (typeof v.value === 'number' && Number.isFinite(v.value)) return v.value;
  }
  return null;
}

function textValue(v) {
  if (typeof v === 'string' && v.trim()) return v.trim();
  if (v && typeof v === 'object' && typeof v.richText === 'object') {
    return v.richText.map(x => x.text || '').join('').trim();
  }
  return null;
}

function cellHasFormula(cell) {
  return !!(cell && cell.value && typeof cell.value === 'object' && Object.prototype.hasOwnProperty.call(cell.value, 'formula'));
}

function safeWorksheetBounds(ws) {
  return { rowCount: ws.rowCount || ws.actualRowCount || 0, columnCount: ws.columnCount || ws.actualColumnCount || 0 };
}

function composeManhourNarrative(sourceSheets) {
  const parts = [];
  for (const src of sourceSheets) {
    const cell = src.sheet.getCell('B5');
    const txt = textValue(cell.value);
    if (txt && !parts.includes(txt)) parts.push(txt);
  }
  return parts.join('\n');
}

function addMergeSummary(workbook, summaryRows) {
  const existing = workbook.getWorksheet('Merge_Summary');
  if (existing) workbook.removeWorksheet(existing.id);
  const ws = workbook.addWorksheet('Merge_Summary');
  ws.getCell('A1').value = 'Flexible Master Monthly Report Merge Summary';
  ws.getCell('A3').value = 'Generated On';
  ws.getCell('B3').value = new Date().toLocaleString();
  ws.getCell('A5').value = 'Sheet';
  ws.getCell('B5').value = 'Source Sheets Used';
  ws.getCell('C5').value = 'Numeric Cells Merged';
  ws.getCell('D5').value = 'Notes';
  let row = 6;
  for (const item of summaryRows) {
    ws.getCell(`A${row}`).value = item.sheet;
    ws.getCell(`B${row}`).value = item.sources;
    ws.getCell(`C${row}`).value = item.merged;
    ws.getCell(`D${row}`).value = item.note;
    row += 1;
  }
  ws.columns = [
    { width: 32 },
    { width: 58 },
    { width: 18 },
    { width: 60 }
  ];
}

async function loadWorkbook(file) {
  const workbook = new ExcelJS.Workbook();
  const buf = await file.arrayBuffer();
  await workbook.xlsx.load(buf);
  return workbook;
}

reviewBtn.addEventListener('click', async () => {
  results.innerHTML = '';
  if (!uploadedFiles.size) {
    setStatus('Please load at least one Excel file.', 'status-warn');
    return;
  }
  setStatus('Reviewing uploaded files...', 'status-idle');
  try {
    const books = [];
    for (const file of uploadedFiles.values()) {
      const wb = await loadWorkbook(file);
      books.push({ file, workbook: wb, score: wb.worksheets.reduce((n, ws) => n + ws.rowCount + ws.columnCount, 0) });
      addResult(file.name, `Opened successfully. Sheets: ${wb.worksheets.length}`);
    }
    const template = findTemplateCandidate(books);
    addResult('Template selected', template.file.name);
    setStatus('Files opened successfully. You can now generate the master workbook.', 'status-ok');
  } catch (err) {
    console.error(err);
    setStatus('A workbook could not be opened.', 'status-bad');
    addResult('Open error', String(err));
  }
});

generateBtn.addEventListener('click', async () => {
  results.innerHTML = '';
  if (!uploadedFiles.size) {
    setStatus('Please load at least one Excel file.', 'status-warn');
    return;
  }

  setStatus('Opening workbooks and preparing merge...', 'status-idle');

  try {
    const books = [];
    for (const file of uploadedFiles.values()) {
      const wb = await loadWorkbook(file);
      books.push({ file, workbook: wb, score: wb.worksheets.reduce((n, ws) => n + ws.rowCount + ws.columnCount, 0) });
    }

    const templateEntry = findTemplateCandidate(books);
    const templateWorkbook = templateEntry.workbook;
    const summaryRows = [];

    const sourceBooks = books.filter(b => b !== templateEntry).map(b => ({
      file: b.file,
      workbook: b.workbook,
      sheets: workbookSheetMap(b.workbook)
    }));

    addResult('Template workbook', templateEntry.file.name);
    addResult('Loaded source files', sourceBooks.map(x => x.file.name).join('\n') || 'No extra source files. The template itself will be downloaded.');

    for (const targetSheet of templateWorkbook.worksheets) {
      const aliasKeys = getSheetAliasKeys(targetSheet.name);
      const candidateSheets = [];

      for (const source of sourceBooks) {
        for (const key of aliasKeys) {
          const found = source.sheets.get(key);
          if (found) {
            candidateSheets.push({ file: source.file.name, sheet: found });
            break;
          }
        }
      }

      let mergedNumericCells = 0;
      const targetBounds = safeWorksheetBounds(targetSheet);

      for (let r = 1; r <= targetBounds.rowCount; r++) {
        const row = targetSheet.getRow(r);
        for (let c = 1; c <= targetBounds.columnCount; c++) {
          const cell = row.getCell(c);
          if (cellHasFormula(cell)) continue;

          const numbers = [];
          const texts = [];

          for (const src of candidateSheets) {
            const bounds = safeWorksheetBounds(src.sheet);
            if (r > bounds.rowCount || c > bounds.columnCount) continue;
            const v = src.sheet.getRow(r).getCell(c).value;
            const n = numericValue(v);
            if (n !== null) numbers.push(n);
            else {
              const t = textValue(v);
              if (t) texts.push(t);
            }
          }

          const currentNum = numericValue(cell.value);
          if (numbers.length && (currentNum !== null || cell.value == null)) {
            cell.value = numbers.reduce((a,b) => a + b, 0);
            mergedNumericCells += 1;
          } else if ((cell.value == null || cell.value === '') && texts.length) {
            cell.value = texts[0];
          }
        }
      }

      if (canonical(targetSheet.name) === 'oshdescriptionsummary') {
        const descSources = candidateSheets.filter(x => canonical(x.sheet.name) === 'oshdescriptionsummary');
        const narrative = composeManhourNarrative(descSources);
        if (narrative) targetSheet.getCell('B5').value = narrative;
      }

      summaryRows.push({
        sheet: targetSheet.name,
        sources: candidateSheets.map(x => `${x.file} -> ${x.sheet.name}`).join(' | ') || 'No matching source sheet found',
        merged: mergedNumericCells,
        note: candidateSheets.length ? 'Matching numeric cells were summed into the template sheet.' : 'No matching sheet in uploaded source files.'
      });
    }

    addMergeSummary(templateWorkbook, summaryRows);

    setStatus('Preparing download...', 'status-idle');
    const buffer = await templateWorkbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'master_monthly_report_merged.xlsx';
    a.click();
    URL.revokeObjectURL(url);

    const totalMerged = summaryRows.reduce((n, x) => n + x.merged, 0);
    addResult('Merge complete', `Generated master_monthly_report_merged.xlsx\nTotal numeric cells merged: ${totalMerged}`);
    setStatus('Done. The master workbook has been generated.', 'status-ok');
  } catch (err) {
    console.error(err);
    setStatus('Generation failed.', 'status-bad');
    addResult('Error', String(err));
  }
});

refreshLoadedList();
