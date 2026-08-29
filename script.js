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

// Lê o valor "de verdade" de uma célula, mesmo quando ela vem de uma fórmula
// (nesse caso o ExcelJS retorna um objeto { formula, result, ... } em vez do
// valor puro).
function getPlainValue(cell) {
  const v = cell.value;
  return v && typeof v === 'object' ? v.result : v;
}

// Devolve o controle pro navegador por um instante. Sem isso, o loop de
// processamento (que é síncrono) trava a repintura da tela até acabar tudo
// — a barra de progresso "pularia" de 0% pra 100% de uma vez só, em vez de
// se mover de verdade.
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
  
  // Find the custom 'final' column safely (optional, won't throw if not found)
  let idxFinal = null;
  try {
    idxFinal = findColumn(targetFinalColumn);
  } catch (e) {
    idxFinal = null; 
  }

  // Find the custom 'ordem' column safely (optional, won't throw if not found).
  // Sem ela, a regra "última ordem do último cep da rua" não tem como ser
  // calculada, então ela simplesmente não é aplicada (o preenchimento de
  // "final" vazio continua funcionando normalmente).
  let idxOrdem = null;
  try {
    idxOrdem = findColumn(targetOrderColumn);
  } catch (e) {
    idxOrdem = null;
  }

  const rowCount = sheet.rowCount;

  // As `skipRows` primeiras linhas são ignoradas por completo — nem viram
  // cabeçalho, nem dado (não entram em nenhum cálculo, não são
  // recoloridas, e não são tocadas de forma alguma) — mas continuam no
  // arquivo final, exatamente como estavam. O cabeçalho de verdade é a
  // primeira linha logo depois delas (linha `1 + skipRows`), e os dados
  // começam na linha seguinte a essa.
  const dataStartRow = 2 + skipRows;
  if (dataStartRow > rowCount + 1) {
    throw new Error(
      `O número de linhas a ignorar (${skipRows}) é maior do que a quantidade de linhas de dados da planilha (${Math.max(rowCount - 1, 0)}).`
    );
  }

  // Pass 1: Map all unique CEPs for each street name
  const cepsByStreet = new Map();
  const allCeps = new Set();
  let totalOriginalRows = 0;

  // Para cada rua, guarda o maior valor de "ordem" já visto e o(s)
  // número(s) de linha que atingem esse máximo — é isso que define "a
  // última ordem do último cep da rua" (o maior ordem da rua inteira,
  // somando todos os ceps dela).
  const lastOrdemByStreet = new Map(); // streetName -> { maxOrdem, rows: Set<rowNumber> }

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

    if (idxOrdem) {
      const ordemNum = Number(getPlainValue(row.getCell(idxOrdem)));
      if (Number.isFinite(ordemNum)) {
        const atual = lastOrdemByStreet.get(streetName);
        if (!atual || ordemNum > atual.maxOrdem) {
          lastOrdemByStreet.set(streetName, { maxOrdem: ordemNum, rows: new Set([rowNumber]) });
        } else if (ordemNum === atual.maxOrdem) {
          atual.rows.add(rowNumber);
        }
      }
    }
  });

  // Conjunto plano com o número de TODAS as linhas que representam "a
  // última ordem do último cep" de alguma rua — usado nas Passes 2 pra
  // decidir quando forçar o "final" para 99999 mesmo que já tenha valor.
  const forcedFinalRows = new Set();
  for (const { rows } of lastOrdemByStreet.values()) {
    for (const r of rows) forcedFinalRows.add(r);
  }

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

  {
    // ---------------------------------------------------------------------
    // Agrupa por bloco: primeiro identifica blocos contíguos (mesma rua +
    // mesmo cep, linhas em sequência), depois insere cada cópia como um
    // bloco inteiro logo abaixo do bloco original.
    //
    // Não fazemos um sort global antes disso (pra preservar a formatação
    // e a ordem original do arquivo) — os blocos são detectados na ordem
    // física em que já aparecem na planilha. Se o arquivo não tiver as
    // linhas de cada rua+cep já em sequência, blocos "quebrados" viram
    // blocos separados.
    // ---------------------------------------------------------------------
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

    // Processa os blocos de baixo para cima: como cada bloco insere linhas
    // sempre abaixo do seu próprio fim original, isso nunca desloca os
    // índices dos blocos que ainda faltam processar (que estão acima).
    const totalBlocks = blocks.length;
    let blocosProcessados = 0;
    let ultimoRespiro = Date.now();

    if (typeof options.onProgress === 'function') {
      options.onProgress({ current: 0, total: totalBlocks, streetName: null });
    }

    for (let b = blocks.length - 1; b >= 0; b--) {
      const block = blocks[b];

      // Regra do "final": força 99999 na última ordem do último cep da rua
      // (mesmo que já tenha valor); nas demais linhas, só preenche se
      // estiver vazio — igual ao modo intercalado.
      if (idxFinal) {
        for (let r = block.startRow; r <= block.endRow; r++) {
          const finalCell = sheet.getRow(r).getCell(idxFinal);
          if (forcedFinalRows.has(r)) {
            finalCell.value = 99999;
          } else {
            const cellValue = getPlainValue(finalCell);
            if (cellValue === null || cellValue === undefined || String(cellValue).trim() === '') {
              finalCell.value = 99999;
            }
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
      // Só cede o controle ao navegador de vez em quando (no máximo a cada
      // 100ms de trabalho contínuo) — ceder a cada bloco individualmente
      // deixaria o processamento bem mais lento em planilhas grandes.
      if (Date.now() - ultimoRespiro > 100) {
        await nextTick();
        ultimoRespiro = Date.now();
      }

      const streetCeps = Array.from(cepsByStreet.get(block.streetName) || []);
      const otherCeps = streetCeps.filter((c) => c !== block.cep);
      if (otherCeps.length === 0) continue;

      // Retrato (snapshot) das linhas originais do bloco, tiradas ANTES de
      // qualquer inserção (a inserção desloca índices de linha abaixo dela).
      const snapshot = [];
      for (let r = block.startRow; r <= block.endRow; r++) {
        const row = sheet.getRow(r);
        const cells = [];
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          cells.push({ colNumber, value: cell.value, style: cell.style });
        });
        snapshot.push({ height: row.height, cells });
      }

      // Insere os blocos-cópia em ordem reversa de "otherCeps": como cada
      // inserção acontece sempre logo após o bloco original (mesma
      // posição), a última cópia inserida acaba ficando mais perto do
      // original — inserindo de trás pra frente, a ordem final das cópias
      // sai crescente, igual à ordem de "otherCeps".
      const reversedTargets = [...otherCeps].reverse();
      for (const targetCep of reversedTargets) {
        const color = getColorForCep(targetCep);
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

  // Ordenação final (opcional): reordena TODA a região processada (da
  // primeira linha de dados até a última linha do arquivo, já incluindo as
  // cópias inseridas) por CEP crescente e, em caso de empate, por "ordem"
  // crescente. As linhas ignoradas (`skipRows`) e o cabeçalho ficam de
  // fora — nunca são tocadas nem se movem.
  //
  // Como não dá pra simplesmente "mover" linhas no ExcelJS, tiramos um
  // retrato de cada linha da região (valores + estilo, igual já fazemos
  // pra duplicar blocos), ordenamos esse retrato em memória, e
  // reescrevemos cada linha física na nova ordem — cor e formatação vêm
  // junto porque fazem parte do retrato.
  if (options.sortOutput === true) {
    const finalRowCount = sheet.rowCount;
    const registros = [];
    for (let r = dataStartRow; r <= finalRowCount; r++) {
      const row = sheet.getRow(r);
      const cells = [];
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        cells.push({ colNumber, value: cell.value, style: cell.style });
      });
      const cepValor = String(getPlainValue(row.getCell(idxCep)) ?? '').trim();
      const ordemValor = idxOrdem ? Number(getPlainValue(row.getCell(idxOrdem))) : NaN;
      registros.push({ height: row.height, cells, cepValor, ordemValor });
    }

    // Array.prototype.sort é estável (garantido desde ES2019): quando o
    // comparador devolve 0 (ex.: sem coluna "ordem" pra desempatar), a
    // ordem relativa original é preservada em vez de virar aleatória.
    registros.sort((a, b) => {
      const porCep = a.cepValor.localeCompare(b.cepValor, 'pt-BR', { numeric: true });
      if (porCep !== 0) return porCep;
      if (Number.isFinite(a.ordemValor) && Number.isFinite(b.ordemValor)) {
        return a.ordemValor - b.ordemValor;
      }
      return 0;
    });

    registros.forEach((registro, idx) => {
      const row = sheet.getRow(dataStartRow + idx);
      row.height = registro.height;
      for (const c of registro.cells) {
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

// =====================================================================
// UI Logic
// =====================================================================
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

// New DOM elements for the history feature
const historySection = document.getElementById('historySection');
const historyList = document.getElementById('historyList');

// Map to track how many times a base file name has been generated
const fileGenerationTracker = new Map();

function updateProgress({ current, total, streetName }) {
  const pct = total > 0 ? Math.round((current / total) * 100) : 0;
  progressFillEl.style.width = pct + '%';
  progressTrackEl.setAttribute('aria-valuenow', String(pct));
  progressDetailEl.textContent = streetName
    ? `${pct}% - street ${current} of ${total}: "${streetName}"`
    : `${pct}% - ${total} streets found`;
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
  statusTextEl.textContent = 'Processing spreadsheet...';
  progressFillEl.style.width = '0%';
  progressTrackEl.setAttribute('aria-valuenow', '0');
  progressDetailEl.textContent = 'Reading file...';

  try {
    const baseName = file.name.replace(/\.xlsx$/i, '');
    const processedBaseName = `${baseName}_processado`;
    
    // Determine the versioned file name
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

    // Start performance timer
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

    // End performance timer and calculate duration in seconds
    const endProcessingTime = performance.now();
    const durationSeconds = ((endProcessingTime - startProcessingTime) / 1000).toFixed(2);

    document.getElementById('statRuas').textContent = summary.totalStreets;
    document.getElementById('statCeps').textContent = summary.totalDistinctCeps;
    document.getElementById('statRuasMulti').textContent = summary.totalStreetsWithMultipleCeps;
    document.getElementById('statGrupos').textContent = summary.totalDuplicatedGroups;
    document.getElementById('statCopiadas').textContent = summary.totalCopiedRows;
    document.getElementById('statTotal').textContent = summary.totalOutputRows;

    // Create object URL for download
    const currentObjectUrl = URL.createObjectURL(blob);
    downloadLink.href = currentObjectUrl;
    downloadLink.download = finalDownloadName;
    
    // Get current time and human-readable file size
    const creationTime = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
    const humanReadableSize = formatBytes(blob.size);
    
    // Build the history list item (<li>)
    const historyItem = document.createElement('li');
    
    // Inline SVG for the Excel Icon
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
    
    // Create the anchor link for the file
    const historyAnchor = document.createElement('a');
    historyAnchor.href = currentObjectUrl;
    historyAnchor.download = finalDownloadName;
    historyAnchor.textContent = finalDownloadName;
    historyAnchor.className = 'history-link';
    historyAnchor.title = finalDownloadName; // Tooltip for truncated names
    
    // Create the metadata container (Size, Processing Time, Creation Time)
    const metaContainer = document.createElement('div');
    metaContainer.className = 'history-meta';
    
    const sizeSpan = document.createElement('span');
    sizeSpan.textContent = humanReadableSize;
    
    const timeSpan = document.createElement('span');
    timeSpan.textContent = `${creationTime}`;
    
    // Assemble the elements
    metaContainer.appendChild(sizeSpan);
    metaContainer.appendChild(timeSpan);
    
    historyItem.innerHTML = svgIcon;
    historyItem.appendChild(historyAnchor);
    historyItem.appendChild(metaContainer);
    
    // Prepend to show the most recent files at the top
    historyList.prepend(historyItem);
    historySection.hidden = false;

    statusEl.hidden = true;
    results.hidden = false;
  } catch (err) {
    showError(err.message || 'An unexpected error occurred while processing the file.');
  }
}

// Event Listeners remain the same
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