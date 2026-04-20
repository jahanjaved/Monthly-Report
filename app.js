let uploadedFiles = [];
let workbookCache = new Map();

const fileInput = document.getElementById('fileInput');
const validateBtn = document.getElementById('validateBtn');
const generateBtn = document.getElementById('generateBtn');
const loadedList = document.getElementById('loadedList');
const statusBox = document.getElementById('statusBox');
const results = document.getElementById('results');
const dropzone = document.getElementById('dropzone');

const CANONICAL_ORDER = ['schedule','description','statistics','kpi','manpower','daily','training'];

function normalizeText(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function canonicalSheetName(name) {
  const n = normalizeText(name);
  if (n.includes('submissionschedule')) return 'schedule';
  if (n.includes('descriptionsummary')) return 'description';
  if (n === 'oshstatistics' || n.includes('oshstatistics')) return 'statistics';
  if (n.includes('oshkpi')) return 'kpi';
  if (n.includes('manpowermanhoursbreakdown')) return 'manpower';
  if (n.includes('dailybreakdowndetails') || n.includes('dailymanhoursbreakdown')) return 'daily';
  if (n.includes('trainingbreakdown') || n.includes('trainingbreakdownsummary') || n.includes('inductionbreakdown')) return 'training';
  return n || 'unknown';
}

function isFormulaValue(value) {
  return value && typeof value === 'object' && (
    Object.prototype.hasOwnProperty.call(value, 'formula') ||
    Object.prototype.hasOwnProperty.call(value, 'sharedFormula')
  );
}

function isDateValue(value) {
  return value instanceof Date;
}

function isNumericValue(value) {
  return typeof value === 'number' && Number.isFinite(value);
}

function roundishEqual(a, b) {
  return Math.abs(a - b) < 1e-9;
}

function summarizeKinds(info) {
  const parts = [];
  if (info.isTemplateHint) parts.push('template hint');
  if (info.isReportLike) parts.push('monthly report');
  if (!parts.length) parts.push('extra workbook');
  return parts.join(' • ');
}

function setStatus(text, cls='status-idle') {
  statusBox.className = 'status ' + cls;
  statusBox.textContent = text;
}

function addResult(title, body) {
  const card = document.createElement('div');
  card.className = 'result-card';
  card.innerHTML = `<h3>${escapeHtml(title)}</h3><p>${escapeHtml(body)}</p>`;
  results.appendChild(card);
}

function escapeHtml(text) {
  return String(text)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;');
}

async function loadWorkbook(file) {
  const cacheKey = `${file.name}__${file.size}__${file.lastModified}`;
  if (workbookCache.has(cacheKey)) {
    return workbookCache.get(cacheKey);
  }
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await file.arrayBuffer());
  workbookCache.set(cacheKey, wb);
  return wb;
}

function workbookInfo(file, wb) {
  const sheetMeta = wb.worksheets.map(ws => ({
    name: ws.name,
    canonical: canonicalSheetName(ws.name),
    rows: ws.rowCount || 0,
    cols: ws.columnCount || 0,
  }));

  const monthlyMatches = sheetMeta.filter(x => CANONICAL_ORDER.includes(x.canonical)).length;
  const lowerName = file.name.toLowerCase();
  const isTemplateHint = /(original|orignal|master|template)/.test(lowerName);
  const isReportLike = monthlyMatches >= 3;

  let templateScore = 0;
  if (isTemplateHint) templateScore += 100;
  templateScore += monthlyMatches * 10;
  templateScore += sheetMeta.length;

  return {
    file,
    wb,
    sheetMeta,
    monthlyMatches,
    isTemplateHint,
    isReportLike,
    templateScore,
  };
}

function renderLoadedList(items, templateInfo = null) {
  loadedList.innerHTML = '';
  if (!items.length) {
    loadedList.innerHTML = '<li class="empty">No files loaded yet.</li>';
    return;
  }

  items.forEach(info => {
    const li = document.createElement('li');
    let cls = 'extra';
    if (templateInfo && info.file.name === templateInfo.file.name) cls = 'template';
    else if (info.isReportLike) cls = 'report';
    li.className = cls;

    const canonicalText = info.sheetMeta.slice(0, 7).map(x => x.canonical).join(', ');
    li.innerHTML = `
      <strong>${escapeHtml(info.file.name)}</strong>
      <span class="meta">${escapeHtml(summarizeKinds(info))}</span>
      <span class="meta">${info.wb.worksheets.length} sheet(s) • canonical matches: ${info.monthlyMatches}</span>
      <span class="meta">${escapeHtml(canonicalText || 'no recognizable monthly sheets')}</span>
    `;
    loadedList.appendChild(li);
  });
}

function chooseTemplate(items) {
  if (!items.length) return null;
  return [...items].sort((a, b) => b.templateScore - a.templateScore)[0];
}

function getWorkbookInfos() {
  return Promise.all(uploadedFiles.map(async file => workbookInfo(file, await loadWorkbook(file))));
}

function refreshFiles(fileList) {
  const incoming = Array.from(fileList || []).filter(f => /\.(xlsx|xlsm)$/i.test(f.name));
  const seen = new Map();
  for (const file of [...uploadedFiles, ...incoming]) {
    seen.set(`${file.name}__${file.size}__${file.lastModified}`, file);
  }
  uploadedFiles = [...seen.values()];
  renderLoadedList([]);
  results.innerHTML = '';
  generateBtn.disabled = uploadedFiles.length === 0;
  setStatus(uploadedFiles.length ? `${uploadedFiles.length} Excel file(s) loaded.` : 'Waiting for Excel files.', uploadedFiles.length ? 'status-warn' : 'status-idle');
}

fileInput.addEventListener('change', e => refreshFiles(e.target.files));

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
  refreshFiles(e.dataTransfer.files);
});

validateBtn.addEventListener('click', async () => {
  results.innerHTML = '';
  if (!uploadedFiles.length) {
    setStatus('Please upload at least one Excel file.', 'status-bad');
    return;
  }

  setStatus('Reviewing uploaded workbooks...', 'status-idle');
  try {
    const infos = await getWorkbookInfos();
    const templateInfo = chooseTemplate(infos);
    renderLoadedList(infos, templateInfo);

    const reportCount = infos.filter(x => x.isReportLike).length;
    addResult('Template workbook', templateInfo ? templateInfo.file.name : 'No template candidate found');
    addResult('Report-like workbooks', `${reportCount} workbook(s) look like monthly report files.`);
    infos.forEach(info => {
      addResult(info.file.name, [
        `Kind: ${summarizeKinds(info)}`,
        `Sheets: ${info.wb.worksheets.length}`,
        `Canonical matches: ${info.monthlyMatches}`,
      ].join('\n'));
    });

    setStatus(`Review complete. ${uploadedFiles.length} file(s) loaded. Ready to generate the master workbook.`, 'status-ok');
    generateBtn.disabled = false;
  } catch (err) {
    console.error(err);
    setStatus('A workbook could not be opened. Please check the Excel file and try again.', 'status-bad');
    addResult('Open error', String(err));
  }
});

function findSheetByCanonical(workbook, canonical) {
  const candidates = workbook.worksheets.filter(ws => canonicalSheetName(ws.name) === canonical);
  return candidates[0] || null;
}

function compatibilityScore(baseWs, srcWs) {
  const baseRows = Math.max(baseWs.rowCount || 0, 1);
  const baseCols = Math.max(baseWs.columnCount || 0, 1);
  const srcRows = Math.max(srcWs.rowCount || 0, 1);
  const srcCols = Math.max(srcWs.columnCount || 0, 1);

  const rowRatio = Math.min(baseRows, srcRows) / Math.max(baseRows, srcRows);
  const colRatio = Math.min(baseCols, srcCols) / Math.max(baseCols, srcCols);
  return (rowRatio * 0.7) + (colRatio * 0.3);
}

function ensureSummarySheet(workbook) {
  let ws = workbook.getWorksheet('Merge_Summary');
  if (ws) workbook.removeWorksheet(ws.id);
  ws = workbook.addWorksheet('Merge_Summary');
  ws.getCell('A1').value = 'Flexible Master Monthly Report Builder';
  ws.getCell('A2').value = 'Generated';
  ws.getCell('B2').value = new Date();
  ws.getCell('A3').value = 'Note';
  ws.getCell('B3').value = 'Compatible sheets were merged by matching sheet type and cell position. Formulas were preserved from the template workbook.';
  ws.columns = [
    { key:'a', width: 18 },
    { key:'b', width: 42 },
    { key:'c', width: 18 },
    { key:'d', width: 18 },
    { key:'e', width: 18 },
    { key:'f', width: 48 },
  ];
  return ws;
}

async function generateMasterWorkbook() {
  results.innerHTML = '';
  if (!uploadedFiles.length) {
    setStatus('Please upload at least one Excel file.', 'status-bad');
    return;
  }

  setStatus('Opening uploaded workbooks...', 'status-idle');
  const infos = await getWorkbookInfos();
  const templateInfo = chooseTemplate(infos);
  if (!templateInfo) {
    setStatus('Could not determine a base workbook.', 'status-bad');
    return;
  }

  renderLoadedList(infos, templateInfo);

  const reportInfos = infos.filter(x => x.isReportLike);
  const mergeSources = reportInfos.length ? reportInfos : infos;

  addResult('Template selected', templateInfo.file.name);
  addResult('Files considered for merge', mergeSources.map(x => x.file.name).join('\n'));

  const outWb = new ExcelJS.Workbook();
  await outWb.xlsx.load(await templateInfo.file.arrayBuffer());
  outWb.calcProperties.fullCalcOnLoad = true;
  outWb.calcProperties.forceFullCalc = true;

  const summaryWs = ensureSummarySheet(outWb);
  let summaryRow = 6;
  summaryWs.getCell(`A${summaryRow}`).value = 'File';
  summaryWs.getCell(`B${summaryRow}`).value = 'Role';
  summaryWs.getCell(`C${summaryRow}`).value = 'Sheets';
  summaryWs.getCell(`D${summaryRow}`).value = 'Matches';
  summaryWs.getCell(`E${summaryRow}`).value = 'Template score';
  summaryWs.getCell(`F${summaryRow}`).value = 'Notes';
  summaryRow += 1;

  infos.forEach(info => {
    summaryWs.getCell(`A${summaryRow}`).value = info.file.name;
    summaryWs.getCell(`B${summaryRow}`).value = info.file.name === templateInfo.file.name ? 'Template' : (info.isReportLike ? 'Merged source' : 'Extra workbook');
    summaryWs.getCell(`C${summaryRow}`).value = info.wb.worksheets.length;
    summaryWs.getCell(`D${summaryRow}`).value = info.monthlyMatches;
    summaryWs.getCell(`E${summaryRow}`).value = info.templateScore;
    summaryWs.getCell(`F${summaryRow}`).value = summarizeKinds(info);
    summaryRow += 1;
  });

  summaryRow += 1;
  summaryWs.getCell(`A${summaryRow}`).value = 'Merged sheet coverage';
  summaryRow += 1;

  let totalMergedCells = 0;
  let totalTouchedSheets = 0;

  for (const baseWs of outWb.worksheets) {
    if (baseWs.name === 'Merge_Summary') continue;
    const canonical = canonicalSheetName(baseWs.name);
    if (!CANONICAL_ORDER.includes(canonical)) continue;

    const compatibleSheets = [];
    for (const info of mergeSources) {
      const srcWs = findSheetByCanonical(info.wb, canonical);
      if (!srcWs) continue;
      const score = compatibilityScore(baseWs, srcWs);
      if (score >= 0.72) {
        compatibleSheets.push({ info, ws: srcWs, score });
      }
    }

    const skippedSheets = [];
    for (const info of mergeSources) {
      const srcWs = findSheetByCanonical(info.wb, canonical);
      if (!srcWs) continue;
      const score = compatibilityScore(baseWs, srcWs);
      if (score < 0.72) {
        skippedSheets.push(`${info.file.name} (${score.toFixed(2)})`);
      }
    }

    let mergedHere = 0;

    if (compatibleSheets.length) {
      const maxRow = Math.max(...compatibleSheets.map(x => x.ws.rowCount || 0), baseWs.rowCount || 0);
      const maxCol = Math.max(...compatibleSheets.map(x => x.ws.columnCount || 0), baseWs.columnCount || 0);

      for (let r = 1; r <= maxRow; r++) {
        const baseRow = baseWs.getRow(r);
        for (let c = 1; c <= maxCol; c++) {
          const baseCell = baseRow.getCell(c);

          if (baseCell.isMerged) continue;
          if (isFormulaValue(baseCell.value)) continue;
          if (isDateValue(baseCell.value)) continue;

          const numbers = [];
          for (const item of compatibleSheets) {
            const srcCell = item.ws.getRow(r).getCell(c);
            const value = srcCell ? srcCell.value : null;
            if (isNumericValue(value)) numbers.push(value);
          }

          if (!numbers.length) continue;

          const unique = [];
          numbers.forEach(num => {
            if (!unique.some(v => roundishEqual(v, num))) unique.push(num);
          });

          const newValue = unique.length === 1 ? unique[0] : numbers.reduce((a, b) => a + b, 0);
          if (baseCell.value !== newValue) {
            baseCell.value = newValue;
            mergedHere += 1;
          }
        }
      }
    }

    totalMergedCells += mergedHere;
    totalTouchedSheets += compatibleSheets.length ? 1 : 0;

    summaryWs.getCell(`A${summaryRow}`).value = baseWs.name;
    summaryWs.getCell(`B${summaryRow}`).value = compatibleSheets.length ? compatibleSheets.map(x => x.info.file.name).join(' | ') : 'No compatible source sheet';
    summaryWs.getCell(`C${summaryRow}`).value = compatibleSheets.length;
    summaryWs.getCell(`D${summaryRow}`).value = mergedHere;
    summaryWs.getCell(`E${summaryRow}`).value = compatibleSheets.map(x => x.score.toFixed(2)).join(', ');
    summaryWs.getCell(`F${summaryRow}`).value = skippedSheets.length ? `Skipped: ${skippedSheets.join(' | ')}` : 'No skipped sheets';
    summaryRow += 1;
  }

  addResult('Merge completed', `${totalMergedCells} cell(s) updated across ${totalTouchedSheets} sheet(s).`);

  setStatus('Preparing Excel download...', 'status-idle');
  const buffer = await outWb.xlsx.writeBuffer();
  const blob = new Blob(
    [buffer],
    { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }
  );
  const fileName = `master_monthly_report_${new Date().toISOString().slice(0,10)}.xlsx`;
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  a.click();
  URL.revokeObjectURL(url);

  setStatus(`Done. ${fileName} has been generated. Open it in Excel and save once to refresh workbook calculations.`, 'status-ok');
}

generateBtn.addEventListener('click', () => {
  generateMasterWorkbook().catch(err => {
    console.error(err);
    setStatus('Generation failed. Please check the workbook pack and try again.', 'status-bad');
    addResult('Error', String(err));
  });
});
