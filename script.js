function normalizeText(text) {
  return String(text ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase();
}

function hslToArgbHex(h, s, l) {
  s /= 100; l /= 100;
  const c = (1 - Math.abs(2 * l - 1)) * s;
  const x = c * (1 - Math.abs(((h / 60) % 2) - 1));
  const m = l - c / 2;
  let r = 0, g = 0, b = 0;
  if (h < 60) [r, g, b] = [c, x, 0];
  else if (h < 120) [r, g, b] = [x, c, 0];
  else if (h < 180) [r, g, b] = [0, c, x];
  else if (h < 240) [r, g, b] = [0, x, c];
  else if (h < 300) [r, g, b] = [x, 0, c];
  else [r, g, b] = [c, 0, x];
  const toHex = (v) => Math.round((v + m) * 255).toString(16).padStart(2, '0');
  return 'FF' + toHex(r) + toHex(g) + toHex(b);
}

function generateRandomColor() {
  const h = Math.floor(Math.random() * 360);
  const s = 55 + Math.random() * 15;
  const l = 78 + Math.random() * 10;
  return hslToArgbHex(h, s, l);
}

async function processStreetsExcel(file, options = {}) {
  if (!file) throw new Error('Nenhum arquivo foi fornecido.');
  if (typeof ExcelJS === 'undefined') {
    throw new Error('Não foi possível carregar a biblioteca ExcelJS (verifique sua conexão).');
  }

  const outputFileName = options.outputFileName || 'ruas_processado.xlsx';
  const targetStreetColumn = options.streetColumnName || 'nome trecho';
  const targetFinalColumn = options.finalColumnName || 'final';

  const buffer = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const sheet = options.sheetName ? workbook.getWorksheet(options.sheetName) : workbook.worksheets[0];
  if (!sheet) throw new Error('Não encontrei nenhuma planilha dentro do arquivo.');

  const headerRow = sheet.getRow(1);
  const columns = [];
  headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    columns.push({ index: colNumber, name: String(cell.value ?? '').trim() });
  });

  if (columns.length === 0) {
    throw new Error('Não encontrei uma linha de cabeçalho na primeira linha da planilha.');
  }

  const findColumn = (target) => {
    const col = columns.find((c) => normalizeText(c.name) === normalizeText(target));
    if (!col) {
      throw new Error(`Não encontrei a coluna "${target}" no arquivo.`);
    }
    return col.index;
  };

  const idxStreetName = findColumn(targetStreetColumn);
  const idxCep = findColumn('cep');
  
  // Find the custom 'final' column safely (optional, won't throw if not found)
  let idxFinal = null;
  try {
    idxFinal = findColumn(targetFinalColumn);
  } catch (e) {
    idxFinal = null; 
  }

  // Pass 1: Map all unique CEPs for each street name
  const cepsByStreet = new Map();
  const allCeps = new Set();
  let totalOriginalRows = 0;

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    totalOriginalRows++;

    const streetVal = row.getCell(idxStreetName).value;
    const streetName = normalizeText(streetVal && typeof streetVal === 'object' ? streetVal.result : streetVal);

    const cepVal = row.getCell(idxCep).value;
    const cep = String((cepVal && typeof cepVal === 'object' ? cepVal.result : cepVal) ?? '').trim();

    if (!cepsByStreet.has(streetName)) {
      cepsByStreet.set(streetName, new Set());
    }
    cepsByStreet.get(streetName).add(cep);
    allCeps.add(cep);
  });

  let totalStreetsWithMultipleCeps = 0;
  for (const ceps of cepsByStreet.values()) {
    if (ceps.size > 1) totalStreetsWithMultipleCeps++;
  }
  
  const cepColors = new Map();
  const getColorForCep = (cep) => {
    if (!cepColors.has(cep)) {
      cepColors.set(cep, generateRandomColor());
    }
    return cepColors.get(cep);
  };

  let totalDuplicatedGroups = 0;
  let totalCopiedRows = 0;
  
  const rowCount = sheet.rowCount;
  
  for (let i = rowCount; i >= 2; i--) {
    const row = sheet.getRow(i);
    
    if (idxFinal) {
      const finalCell = row.getCell(idxFinal);
      const cellValue = finalCell.value && typeof finalCell.value === 'object' ? finalCell.value.result : finalCell.value;
      
      if (cellValue === null || cellValue === undefined || String(cellValue).trim() === '') {
        finalCell.value = 99999;
      }
    }

    const streetVal = row.getCell(idxStreetName).value;
    const streetName = normalizeText(streetVal && typeof streetVal === 'object' ? streetVal.result : streetVal);

    const cepVal = row.getCell(idxCep).value;
    const currentCep = String((cepVal && typeof cepVal === 'object' ? cepVal.result : cepVal) ?? '').trim();

    const streetCeps = Array.from(cepsByStreet.get(streetName) || []);
    const otherCeps = streetCeps.filter(c => c !== currentCep);

    if (otherCeps.length > 0) {
      totalDuplicatedGroups++;
    }

    // Insert new cloned rows immediately below the current row (i + 1)
    for (const targetCep of otherCeps) {
      const color = getColorForCep(targetCep);

      sheet.spliceRows(i + 1, 0, []); 
      const newRow = sheet.getRow(i + 1);
      totalCopiedRows++;

      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const targetCell = newRow.getCell(colNumber);
        targetCell.value = cell.value; 
        targetCell.style = cell.style; 
      });

      newRow.height = row.height;
      newRow.getCell(idxCep).value = targetCep;

      newRow.eachCell({ includeEmpty: true }, (cell) => {
        const currentStyle = cell.style || {};
        cell.style = {
          ...currentStyle,
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: color } }
        };
      });
    }
  }

  const outputArrayBuffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([outputArrayBuffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });

  return {
    blob,
    fileName: outputFileName,
    summary: {
      totalOriginalRows,
      totalStreets: cepsByStreet.size,
      totalDistinctCeps: allCeps.size,
      totalStreetsWithMultipleCeps,
      totalDuplicatedGroups,
      totalCopiedRows,
      totalOutputRows: totalOriginalRows + totalCopiedRows,
    },
  };
}

// =====================================================================
// UI Logic
// =====================================================================
const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const fileNameEl = document.getElementById('fileName');
const statusEl = document.getElementById('status');
const statusTextEl = document.getElementById('statusText');
const errorBox = document.getElementById('errorBox');
const results = document.getElementById('results');
const downloadLink = document.getElementById('downloadLink');
const resetBtn = document.getElementById('resetBtn');

let currentObjectUrl = null;

function resetUI() {
  if (currentObjectUrl) { URL.revokeObjectURL(currentObjectUrl); currentObjectUrl = null; }
  fileInput.value = '';
  fileNameEl.hidden = true;
  statusEl.hidden = true;
  errorBox.hidden = true;
  results.hidden = true;
  dropzone.hidden = false;
}

function showError(message) {
  statusEl.hidden = true;
  results.hidden = true;
  errorBox.hidden = false;
  errorBox.textContent = message;
}

async function handleFile(file) {
  if (!file) return;
  errorBox.hidden = true;
  results.hidden = true;
  fileNameEl.hidden = false;
  fileNameEl.textContent = file.name;
  statusEl.hidden = false;
  statusTextEl.textContent = 'Processando planilha…';

  try {
    const baseName = file.name.replace(/\.xlsx$/i, '');
        
    const streetColName = document.getElementById('colNameInput').value.trim() || 'nome trecho';
    const finalColName = document.getElementById('colFinalInput').value.trim() || 'final';

    const { blob, summary } = await processStreetsExcel(file, {
      outputFileName: `${baseName}_processado.xlsx`,
      streetColumnName: streetColName,
      finalColumnName: finalColName
    });

    document.getElementById('statRuas').textContent = summary.totalStreets;
    document.getElementById('statCeps').textContent = summary.totalDistinctCeps;
    document.getElementById('statRuasMulti').textContent = summary.totalStreetsWithMultipleCeps;
    document.getElementById('statGrupos').textContent = summary.totalDuplicatedGroups;
    document.getElementById('statCopiadas').textContent = summary.totalCopiedRows;
    document.getElementById('statTotal').textContent = summary.totalOutputRows;

    if (currentObjectUrl) URL.revokeObjectURL(currentObjectUrl);
    currentObjectUrl = URL.createObjectURL(blob);
    downloadLink.href = currentObjectUrl;
    downloadLink.download = `${baseName}_processado.xlsx`;

    statusEl.hidden = true;
    results.hidden = false;
  } catch (err) {
    showError(err.message || 'Ocorreu um erro inesperado ao processar o arquivo.');
  }
}

dropzone.addEventListener('click', () => fileInput.click());
dropzone.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileInput.click(); }
});
fileInput.addEventListener('change', (e) => handleFile(e.target.files[0]));

['dragover', 'dragenter'].forEach((evt) =>
  dropzone.addEventListener(evt, (e) => { e.preventDefault(); dropzone.classList.add('dragover'); })
);
['dragleave', 'drop'].forEach((evt) =>
  dropzone.addEventListener(evt, (e) => { e.preventDefault(); dropzone.classList.remove('dragover'); })
);
dropzone.addEventListener('drop', (e) => {
  const file = e.dataTransfer.files[0];
  if (file) handleFile(file);
});

resetBtn.addEventListener('click', resetUI);