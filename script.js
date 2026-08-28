// ---------------------------------------------------------------------
// Processing logic
// ---------------------------------------------------------------------
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
  if (!file) throw new Error('Nenhum arquivo foi fornecido.'); // Keeping UI messages in PT-BR
  if (typeof ExcelJS === 'undefined') {
    throw new Error('Não foi possível carregar a biblioteca ExcelJS (verifique sua conexão).');
  }

  const outputFileName = options.outputFileName || 'ruas_processado.xlsx';

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
      throw new Error(
        `Não encontrei a coluna "${target}" no arquivo. Verifique se o cabeçalho da planilha tem uma coluna com esse nome (maiúsculas/minúsculas não importam).`
      );
    }
    return col.index;
  };

  const idxStreetName = findColumn('nome trecho');
  const idxCep = findColumn('cep');
  const finalColumn = columns.find((c) => normalizeText(c.name) === 'final');
  const idxFinal = finalColumn ? finalColumn.index : null;
  const totalColumns = columns.reduce((max, c) => Math.max(max, c.index), 0);

  const rows = [];
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const values = [];
    for (let c = 1; c <= totalColumns; c++) values.push(row.getCell(c).value ?? null);
    rows.push(values);
  });
  if (rows.length === 0) throw new Error('A planilha não tem nenhuma linha de dados abaixo do cabeçalho.');

  if (idxFinal) {
    for (const values of rows) {
      const v = values[idxFinal - 1];
      if (v === null || v === undefined || String(v).trim() === '') values[idxFinal - 1] = 99999;
    }
  }

  rows.sort((a, b) => {
    const nameA = normalizeText(a[idxStreetName - 1]);
    const nameB = normalizeText(b[idxStreetName - 1]);
    if (nameA !== nameB) return nameA < nameB ? -1 : 1;
    const cepA = String(a[idxCep - 1] ?? '');
    const cepB = String(b[idxCep - 1] ?? '');
    return cepA.localeCompare(cepB, 'pt-BR', { numeric: true });
  });

  const cepsByStreet = new Map();
  for (const values of rows) {
    const key = normalizeText(values[idxStreetName - 1]);
    const cep = String(values[idxCep - 1] ?? '').trim();
    if (!cepsByStreet.has(key)) cepsByStreet.set(key, []);
    const list = cepsByStreet.get(key);
    if (!list.includes(cep)) list.push(cep);
  }

  const blocks = [];
  let currentBlock = null;
  for (const values of rows) {
    const streetKey = normalizeText(values[idxStreetName - 1]);
    const cep = String(values[idxCep - 1] ?? '').trim();
    if (currentBlock && currentBlock.streetKey === streetKey && currentBlock.cep === cep) {
      currentBlock.rows.push(values);
    } else {
      currentBlock = { streetKey, cep, rows: [values] };
      blocks.push(currentBlock);
    }
  }

  const finalRows = [];
  let totalDuplicatedGroups = 0;
  let totalCopiedRows = 0;
  let totalStreetsWithMultipleCeps = 0;
  for (const ceps of cepsByStreet.values()) if (ceps.length > 1) totalStreetsWithMultipleCeps++;
  const allCeps = new Set(rows.map((v) => String(v[idxCep - 1] ?? '').trim()));

  for (const block of blocks) {
    for (const values of block.rows) finalRows.push({ values, color: null });
    const allBlockCeps = cepsByStreet.get(block.streetKey) || [];
    const otherCeps = allBlockCeps.filter((c) => c !== block.cep);
    for (const targetCep of otherCeps) {
      const color = generateRandomColor();
      totalDuplicatedGroups++;
      for (const originalValues of block.rows) {
        const copy = originalValues.slice();
        copy[idxCep - 1] = targetCep;
        finalRows.push({ values: copy, color });
        totalCopiedRows++;
      }
    }
  }

  const newWorkbook = new ExcelJS.Workbook();
  const newSheet = newWorkbook.addWorksheet(sheet.name || 'Ruas');
  const headerNames = [];
  for (let c = 1; c <= totalColumns; c++) {
    const col = columns.find((cc) => cc.index === c);
    headerNames.push(col ? col.name : '');
  }
  newSheet.addRow(headerNames);
  newSheet.getRow(1).font = { bold: true };

  for (const { values, color } of finalRows) {
    const row = newSheet.addRow(values);
    if (color) {
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
      });
    }
  }
  newSheet.columns.forEach((col) => { col.width = 18; });

  const outputArrayBuffer = await newWorkbook.xlsx.writeBuffer();
  const blob = new Blob([outputArrayBuffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });

  return {
    blob,
    fileName: outputFileName,
    summary: {
      totalOriginalRows: rows.length,
      totalStreets: cepsByStreet.size,
      totalDistinctCeps: allCeps.size,
      totalStreetsWithMultipleCeps,
      totalDuplicatedGroups,
      totalCopiedRows,
      totalOutputRows: finalRows.length,
    },
  };
}

// ---------------------------------------------------------------------
// UI Logic
// ---------------------------------------------------------------------
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
    const { blob, summary } = await processStreetsExcel(file, {
      outputFileName: `${baseName}_processado.xlsx`,
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