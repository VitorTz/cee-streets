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


function getPlainValue(cell) {
  const v = cell.value;
  return v && typeof v === 'object' ? v.result : v;
}

function nextTick() {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

async function processStreetsExcel(file, options = {}) {
  if (!file) throw new Error('Nenhum arquivo foi fornecido.');
  if (typeof ExcelJS === 'undefined') {
    throw new Error('Não foi possível carregar a biblioteca ExcelJS (verifique sua conexão).');
  }

  const outputFileName = options.outputFileName || 'ruas_processado.xlsx';
  const targetStreetColumn = options.streetColumnName || 'nome trecho';
  const targetFinalColumn = options.finalColumnName || 'final';
  const targetOrderColumn = options.orderColumnName || 'ordem';

  let skipRows = parseInt(options.skipRows, 10);
  if (!Number.isFinite(skipRows) || skipRows < 0) skipRows = 0;

  const buffer = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const sheet = options.sheetName ? workbook.getWorksheet(options.sheetName) : workbook.worksheets[0];
  if (!sheet) throw new Error('Não encontrei nenhuma planilha dentro do arquivo.');

  const headerRow = sheet.getRow(1 + skipRows);
  const columns = [];
  headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    columns.push({ index: colNumber, name: String(cell.value ?? '').trim() });
  });

  if (columns.length === 0) {
    throw new Error(`Não encontrei uma linha de cabeçalho na linha ${1 + skipRows} da planilha (a primeira linha depois das ${skipRows} ignoradas).`);
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
    
  let idxFinal = null;
  try {
    idxFinal = findColumn(targetFinalColumn);
  } catch (e) {
    idxFinal = null; 
  }
  
  let idxOrdem = null;
  try {
    idxOrdem = findColumn(targetOrderColumn);
  } catch (e) {
    idxOrdem = null;
  }

  const rowCount = sheet.rowCount;
  
  const dataStartRow = 2 + skipRows;
  if (dataStartRow > rowCount + 1) {
    throw new Error(
      `O número de linhas a ignorar (${skipRows}) é maior do que a quantidade de linhas de dados da planilha (${Math.max(rowCount - 1, 0)}).`
    );
  }
  
  const cepsByStreet = new Map();
  const allCeps = new Set();
  let totalOriginalRows = 0;
  
  const lastFinalByStreet = new Map();

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber < dataStartRow) return; // linha de cabeçalho ou ignorada
    totalOriginalRows++;

    const streetVal = row.getCell(idxStreetName).value;
    const streetName = normalizeText(streetVal && typeof streetVal === 'object' ? streetVal.result : streetVal);

    const cepVal = row.getCell(idxCep).value;
    const cep = String((cepVal && typeof cepVal === 'object' ? cepVal.result : cepVal) ?? '').trim();

    if (!streetName || !cep) {
      return;
    }

    if (!cepsByStreet.has(streetName)) {
      cepsByStreet.set(streetName, new Set());
    }
    cepsByStreet.get(streetName).add(cep);
    allCeps.add(cep);

    if (idxFinal) {
      const finalNum = Number(getPlainValue(row.getCell(idxFinal)));
      if (Number.isFinite(finalNum)) {
        const atual = lastFinalByStreet.get(streetName);
        if (!atual || finalNum > atual.maxFinal) {
          lastFinalByStreet.set(streetName, { maxFinal: finalNum, rows: new Set([rowNumber]) });
        } else if (finalNum === atual.maxFinal) {
          atual.rows.add(rowNumber);
        }
      }
    }
  });
  
  const forcedFinalRows = new Set();
  for (const { rows } of lastFinalByStreet.values()) {
    for (const r of rows) forcedFinalRows.add(r);
  }

  let totalStreetsWithMultipleCeps = 0;
  for (const ceps of cepsByStreet.values()) {
    if (ceps.size > 1) totalStreetsWithMultipleCeps++;
  }
  
  let totalDuplicatedGroups = 0;
  let totalCopiedRows = 0;

  {    
    // Agrupa por bloco
    const blocks = [];
    let currentBlock = null;
    for (let i = dataStartRow; i <= rowCount; i++) {
      const row = sheet.getRow(i);
      const streetVal = row.getCell(idxStreetName).value;
      const streetNameRaw = String(getPlainValue(row.getCell(idxStreetName)) ?? '').trim();
      const streetName = normalizeText(streetVal && typeof streetVal === 'object' ? streetVal.result : streetVal);
      const cepVal = row.getCell(idxCep).value;
      const cep = String((cepVal && typeof cepVal === 'object' ? cepVal.result : cepVal) ?? '').trim();
      
      if (!streetName || !cep) {
        continue;
      }

      if (currentBlock && currentBlock.streetName === streetName && currentBlock.cep === cep) {
        currentBlock.endRow = i;
      } else {
        currentBlock = { startRow: i, endRow: i, streetName, streetNameRaw, cep };
        blocks.push(currentBlock);
      }
    }
    
    const totalBlocks = blocks.length;
    let blocosProcessados = 0;
    let ultimoRespiro = Date.now();

    if (typeof options.onProgress === 'function') {
      options.onProgress({ current: 0, total: totalBlocks, streetName: null });
    }

    for (let b = blocks.length - 1; b >= 0; b--) {
      const block = blocks[b];      
      if (idxFinal) {
        for (let r = block.startRow; r <= block.endRow; r++) {
          if (forcedFinalRows.has(r)) {
            sheet.getRow(r).getCell(idxFinal).value = 99999;
          }
        }
      }

      blocosProcessados++;
      if (typeof options.onProgress === 'function') {
        options.onProgress({
          current: blocosProcessados,
          total: totalBlocks,
          streetName: block.streetNameRaw || block.streetName,
        });
      }
      
      if (Date.now() - ultimoRespiro > 100) {
        await nextTick();
        ultimoRespiro = Date.now();
      }

      const streetCeps = Array.from(cepsByStreet.get(block.streetName) || []);
      const otherCeps = streetCeps.filter((c) => c !== block.cep);
      if (otherCeps.length === 0) continue;
      
      const snapshot = [];
      for (let r = block.startRow; r <= block.endRow; r++) {
        const row = sheet.getRow(r);
        const cells = [];
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          cells.push({ colNumber, value: cell.value, style: cell.style });
        });
        snapshot.push({ height: row.height, cells });
      }
            
      const reversedTargets = [...otherCeps].reverse();
      for (const targetCep of reversedTargets) {
        const color = generateRandomColor();
        totalDuplicatedGroups++;
        totalCopiedRows += snapshot.length;

        // Insere N linhas em branco de uma vez, logo após o bloco original
        sheet.spliceRows(block.endRow + 1, 0, ...snapshot.map(() => []));

        snapshot.forEach((snapRow, k) => {
          const newRow = sheet.getRow(block.endRow + 1 + k);
          newRow.height = snapRow.height;
          for (const c of snapRow.cells) {
            const targetCell = newRow.getCell(c.colNumber);
            targetCell.value = c.value;
            targetCell.style = c.style;
          }
          newRow.getCell(idxCep).value = targetCep;
          newRow.eachCell({ includeEmpty: true }, (cell) => {
            const currentStyle = cell.style || {};
            cell.style = {
              ...currentStyle,
              fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: color } }
            };
          });
        });
      }
    }
  }

  // Ordenação final (opcional)
  if (options.sortOutput === true) {
    const finalRowCount = sheet.rowCount;
    const records = [];
    
    for (let r = dataStartRow; r <= finalRowCount; r++) {
      const row = sheet.getRow(r);
      const cells = [];
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        cells.push({ colNumber, value: cell.value, style: cell.style });
      });
      
      const streetValue = String(getPlainValue(row.getCell(idxStreetName)) ?? '').trim();
      const cepValue = String(getPlainValue(row.getCell(idxCep)) ?? '').trim();
      const orderValue = idxOrdem ? Number(getPlainValue(row.getCell(idxOrdem))) : NaN;
      
      records.push({ height: row.height, cells, streetValue, cepValue, orderValue });
    }

    records.sort((a, b) => {
      // ordena pelo nome da rua (asc)
      const byStreet = a.streetValue.localeCompare(b.streetValue, 'pt-BR', { numeric: true });
      if (byStreet !== 0) return byStreet;

      // ordena pelo cep (asc)
      const byCep = a.cepValue.localeCompare(b.cepValue, 'pt-BR', { numeric: true });
      if (byCep !== 0) return byCep;
      
      // ordena pela ordem (asc)
      if (Number.isFinite(a.orderValue) && Number.isFinite(b.orderValue)) {
        return a.orderValue - b.orderValue;
      }
      return 0;
    });

    records.forEach((record, idx) => {
      const row = sheet.getRow(dataStartRow + idx);
      row.height = record.height;
      for (const c of record.cells) {
        const targetCell = row.getCell(c.colNumber);
        targetCell.value = c.value;
        targetCell.style = c.style;
      }
    });
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
      skippedRows: skipRows,
      totalStreets: cepsByStreet.size,
      totalDistinctCeps: allCeps.size,
      totalStreetsWithMultipleCeps,
      totalDuplicatedGroups,
      totalCopiedRows,
      totalOutputRows: totalOriginalRows + skipRows + totalCopiedRows,
    },
  };
}

// UI
const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const fileNameEl = document.getElementById('fileName');
const statusEl = document.getElementById('status');
const statusTextEl = document.getElementById('statusText');
const progressTrackEl = document.getElementById('progressTrack');
const progressFillEl = document.getElementById('progressFill');
const progressDetailEl = document.getElementById('progressDetail');
const errorBox = document.getElementById('errorBox');
const results = document.getElementById('results');
const downloadLink = document.getElementById('downloadLink');
const resetBtn = document.getElementById('resetBtn');


const historySection = document.getElementById('historySection');
const historyList = document.getElementById('historyList');

const fileGenerationTracker = new Map();

function updateProgress({ current, total, streetName }) {  
  const pct = total > 0 ? Math.min(100, Math.round((current / total) * 100)) : 0;
  progressFillEl.style.width = pct + '%';
  progressTrackEl.setAttribute('aria-valuenow', String(pct));
  progressDetailEl.textContent = streetName
    ? `${pct}% - trecho ${current} de ${total}: "${streetName}"`
    : `${pct}% - ${total} trechos encontrados`;
}

function resetUI() {
  fileInput.value = '';
  fileNameEl.hidden = true;
  statusEl.hidden = true;
  errorBox.hidden = true;
  results.hidden = true;
  dropzone.hidden = false;
  progressFillEl.style.width = '0%';
  progressTrackEl.setAttribute('aria-valuenow', '0');
  progressDetailEl.textContent = '';
}

function showError(message) {
  statusEl.hidden = true;
  results.hidden = true;
  errorBox.hidden = false;
  errorBox.textContent = message;
}

function formatBytes(bytes, decimals = 2) {
  if (!+bytes) return '0 Bytes';
  const k = 1024;
  const dm = decimals < 0 ? 0 : decimals;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return `${parseFloat((bytes / Math.pow(k, i)).toFixed(dm))} ${sizes[i]}`;
}

async function handleFile(file) {
  if (!file) return;
  errorBox.hidden = true;
  results.hidden = true;
  fileNameEl.hidden = false;
  fileNameEl.textContent = file.name;
  statusEl.hidden = false;
  statusTextEl.textContent = 'Processando planilha…';
  progressFillEl.style.width = '0%';
  progressTrackEl.setAttribute('aria-valuenow', '0');
  progressDetailEl.textContent = 'Lendo arquivo…';

  try {
    const baseName = file.name.replace(/\.xlsx$/i, '');
    const processedBaseName = `${baseName}_processado`;
    
    // nome do arquivo (com versionamento)
    const count = fileGenerationTracker.get(processedBaseName) || 0;
    const finalDownloadName = count === 0 
      ? `${processedBaseName}.xlsx` 
      : `${processedBaseName}_v${count}.xlsx`;
    
    fileGenerationTracker.set(processedBaseName, count + 1);

    const streetColName = document.getElementById('colNameInput').value.trim() || 'nome trecho';
    const finalColName = document.getElementById('colFinalInput').value.trim() || 'final';
    const orderColName = document.getElementById('colOrderInput').value.trim() || 'ordem';
    const skipRows = document.getElementById('skipRowsInput').value.trim() || '0';
    const sortOutput = document.getElementById('sortOutputInput').checked;
    
    const startProcessingTime = performance.now();

    const { blob, summary } = await processStreetsExcel(file, {
      outputFileName: finalDownloadName,
      streetColumnName: streetColName,
      finalColumnName: finalColName,
      orderColumnName: orderColName,
      skipRows,
      sortOutput,
      onProgress: updateProgress
    });
    
    const endProcessingTime = performance.now();
    const durationSeconds = ((endProcessingTime - startProcessingTime) / 1000).toFixed(2);

    document.getElementById('statRuas').textContent = summary.totalStreets;
    document.getElementById('statCeps').textContent = summary.totalDistinctCeps;
    document.getElementById('statRuasMulti').textContent = summary.totalStreetsWithMultipleCeps;
    document.getElementById('statGrupos').textContent = summary.totalDuplicatedGroups;
    document.getElementById('statCopiadas').textContent = summary.totalCopiedRows;
    document.getElementById('statTotal').textContent = summary.totalOutputRows;
    
    const currentObjectUrl = URL.createObjectURL(blob);
    downloadLink.href = currentObjectUrl;
    downloadLink.download = finalDownloadName;
        
    const creationTime = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
    const humanReadableSize = formatBytes(blob.size);
        
    const historyItem = document.createElement('li');
        
   const svgIcon = `
  <svg
    class="history-icon"
    viewBox="0 0 24 24"
    fill="none"
    xmlns="http://www.w3.org/2000/svg"
    aria-hidden="true"
  >
    <!-- Excel background -->
    <path
      d="M5 3C5 2.45 5.45 2 6 2H14L20 8V21C20 21.55 19.55 22 19 22H6C5.45 22 5 21.55 5 21V3Z"
      fill="#217346"
    />

    <!-- Fold -->
    <path
      d="M14 2V8H20"
      fill="#185C37"
    />

    <!-- Spreadsheet -->
    <rect
      x="8"
      y="10"
      width="9"
      height="9"
      rx="0.8"
      fill="white"
      fill-opacity="0.95"
    />

    <!-- Spreadsheet grid -->
    <path
      d="M11 10V19
         M14 10V19
         M8 13H17
         M8 16H17"
      stroke="#217346"
      stroke-width="1"
    />

    <!-- Excel X -->
    <path
      d="M3.5 7.5L7.5 11.5
         M7.5 7.5L3.5 11.5"
      stroke="#107C41"
      stroke-width="2"
      stroke-linecap="round"
    />
  </svg>
`;
        
    const historyAnchor = document.createElement('a');
    historyAnchor.href = currentObjectUrl;
    historyAnchor.download = finalDownloadName;
    historyAnchor.textContent = finalDownloadName;
    historyAnchor.className = 'history-link';
    historyAnchor.title = finalDownloadName;
        
    const metaContainer = document.createElement('div');
    metaContainer.className = 'history-meta';
    
    const sizeSpan = document.createElement('span');
    sizeSpan.textContent = humanReadableSize;
    
    const timeSpan = document.createElement('span');
    timeSpan.textContent = `${creationTime}`;
        
    metaContainer.appendChild(sizeSpan);
    metaContainer.appendChild(timeSpan);
    
    historyItem.innerHTML = svgIcon;
    historyItem.appendChild(historyAnchor);
    historyItem.appendChild(metaContainer);
        
    historyList.prepend(historyItem);
    historySection.hidden = false;

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
