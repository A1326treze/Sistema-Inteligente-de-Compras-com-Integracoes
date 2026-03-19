/***************************************
 * Sync AppSheet <-> Google Sheets
 * - PC (Preços (pesquisa)) -> App (Preços app)
 * - App (Preços app) -> PC (Preços (pesquisa))
 *
 * Match NO APP por chave natural:
 * Data + Item + Marca + Qntd + Loja
 * Mantém o ID existente (Key do AppSheet = ID).
 ***************************************/

/**
 * Configurações de mapeamento para a sincronização.
 * @constant
 */
const SYNC_CFG = {
  // Sheets
  sheetPesquisa: 'Preços (pesquisa)',
  sheetInventario: 'Inventário', // só pra “canonizar” itens (opcional, mas ajuda)
  sheetApp: 'Preços app',

  // Preços (pesquisa)
  dateCellPesquisa: 'B1',
  headerRowPesquisa: 2,
  firstDataRowPesquisa: 3,
  colCheck: 1,      // A (checkbox)
  colItem: 2,       // B
  colMarca: 3,      // C
  colQtd: 4,        // D
  colTipo: 5,       // E (fórmula, NÃO escrever)
  storeStartCol: 6, // F (primeira loja)

  // Marcação visual
  noteApp: 'APP',
  dupBorderColor: '#ff0000',

  // Origem (AppSheet)
  originPC: 'PC',
  originApp: 'App',
};

// =======================
// Entradas públicas
// =======================

/**
 * Lê os preços na aba de pesquisa (PC) e envia para a aba conectada ao aplicativo (App).
 * Realiza operação de upsert baseada em chaves naturais compostas.
 * @returns {void}
 */
function syncPcParaApp() {
  const ss = SpreadsheetApp.getActive();
  const sh = mustGetSheet_(ss, SYNC_CFG.sheetPesquisa);
  const app = mustGetSheet_(ss, SYNC_CFG.sheetApp);

  const appCols = getAppCols_(app); // encontra colunas por nome (ID, Data, Item, Marca, Qntd, Loja, Preço, Origem)
  ensureAppHeadersOk_(appCols);

  const { dateObj, dateKey } = getPesquisaDate_(sh);

  const { storeEndCol, stores } = getStoresFromPesquisa_(sh);
  if (stores.length === 0) throw new Error('Não encontrei colunas de lojas válidas na aba Preços (pesquisa).');

  const lastRow = sh.getLastRow();
  if (lastRow < SYNC_CFG.firstDataRowPesquisa) {
    ss.toast('Nada para importar (sem linhas na pesquisa).', 'Sync PC → App', 5);
    return;
  }

  // Lê dados da pesquisa (até storeEndCol)
  const rg = sh.getRange(
    SYNC_CFG.firstDataRowPesquisa, 1,
    lastRow - SYNC_CFG.firstDataRowPesquisa + 1,
    storeEndCol
  );
  const data = rg.getValues();
  const notes = rg.getNotes();


  // Monta registros PC (um por loja com preço válido)
  // chave natural: dateKey|item|marca|qtd|loja
  const regs = [];
  for (const row of data) {
    const item  = cleanText_(row[SYNC_CFG.colItem - 1]);
    const marca = cleanText_(row[SYNC_CFG.colMarca - 1]);
    const qtdV  = row[SYNC_CFG.colQtd - 1];
    const qtd   = cleanQtd_(qtdV);
    const tipo  = cleanText_(row[SYNC_CFG.colTipo - 1]); // só pra validar existência

    if (!item || !qtd || !tipo) continue;

    for (let i = 0; i < stores.length; i++) {
      const loja = stores[i];
      const priceCell = row[(SYNC_CFG.storeStartCol - 1) + i];
      const preco = toNum_(priceCell);
      if (!loja || !isFinite(preco) || preco <= 0) continue;

      const key = makeNaturalKey_(dateKey, item, marca, qtd, loja);
      regs.push({
        key,
        dateObj,
        dateKey,
        item,
        marca,
        qtd,
        loja,
        preco
      });
    }
  }

  if (regs.length === 0) {
    ss.toast('Nada para importar (sem preços válidos).', 'Sync PC → App', 6);
    return;
  }

  // Indexa APP (por chave natural)
  const appIndex = buildAppIndex_(app, appCols);

  // Upserts
  const updates = []; // {row, values[]}
  const inserts = []; // values[]

  for (const r of regs) {
    const hit = appIndex.get(r.key);

if (hit) {
      const rowNum = hit.row;
      const id = hit.id;

      const keepApp = normalize_(hit.origem) === normalize_(SYNC_CFG.originApp);
      const origemFinal = keepApp ? SYNC_CFG.originApp : (r.origin || SYNC_CFG.originPC);

      const newRow = makeAppRow_(appCols, id, r.dateObj, r.item, r.marca, r.qtd, r.loja, r.preco, origemFinal);
      updates.push({ row: rowNum, values: newRow });

    } else {
      const id = Utilities.getUuid();
      const origemFinal = r.origin || SYNC_CFG.originPC;

      const newRow = makeAppRow_(appCols, id, r.dateObj, r.item, r.marca, r.qtd, r.loja, r.preco, origemFinal);
      inserts.push(newRow);
    }

  }

  // Escreve updates
  for (const u of updates) {
    app.getRange(u.row, 1, 1, appCols.lastCol).setValues([u.values]);
  }

  // Escreve inserts (append)
  if (inserts.length > 0) {
    const start = app.getLastRow() + 1;
    app.getRange(start, 1, inserts.length, appCols.lastCol).setValues(inserts);
  }

  ss.toast(`OK: ${updates.length} atualizados | ${inserts.length} inseridos (${dateKey})`, 'Sync PC → App', 8);
}

/**
 * Lê os registros criados via aplicativo (App) e preenche os preços 
 * nas colunas de lojas corretas na aba de pesquisa (PC).
 * @returns {void}
 */
function syncAppParaPc() {
  const ss = SpreadsheetApp.getActive();
  const sh = mustGetSheet_(ss, SYNC_CFG.sheetPesquisa);
  const app = mustGetSheet_(ss, SYNC_CFG.sheetApp);

  const appCols = getAppCols_(app);
  ensureAppHeadersOk_(appCols);

  const { dateKey: targetDateKey } = getPesquisaDate_(sh);

  // lojas da matriz (F.. antes de Custo/Mult)
  const { storeEndCol, stores } = getStoresFromPesquisa_(sh);
  const storeStartCol = SYNC_CFG.storeStartCol;
  const storeCols = storeEndCol - storeStartCol + 1;

  if (storeCols <= 0) throw new Error('Não encontrei lojas na linha de cabeçalho da pesquisa.');

  const storeMap = new Map();
  for (let i = 0; i < stores.length; i++) {
    const name = cleanText_(stores[i]);
    if (!name) continue;
    storeMap.set(normalize_(name), storeStartCol + i);
  }
  if (storeMap.size === 0) throw new Error('Não encontrei lojas na linha de cabeçalho da pesquisa.');

  // canonização de item (dropdown)
  const canonItem = buildCanonicalItemMap_();

  // lê AppSheet (Preços app)
  const appLastRow = app.getLastRow();
  if (appLastRow < 2) {
    ss.toast('Nada no App para exportar.', 'Sync App → PC', 5);
    return;
  }
  const appVals = app.getRange(2, 1, appLastRow - 1, appCols.lastCol).getValues();

  // prepara lista de operações (somente Origem=App e Data=B1)
  const ops = [];
  const needRows = new Map(); // rowKey -> {item, marca, qtd}
  for (const row of appVals) {
    const data   = row[appCols.colData - 1];
    const origem = cleanText_(row[appCols.colOrigem - 1]);

    if (normalize_(origem) !== normalize_(SYNC_CFG.originApp)) continue;

    const dk = dateKeyFromAny_(data);
    if (dk !== targetDateKey) continue;

    const itemRaw = cleanText_(row[appCols.colItem - 1]);
    const item    = canonize_(canonItem, itemRaw);

    const marca = cleanText_(row[appCols.colMarca - 1]);
    const qtd   = cleanQtd_(row[appCols.colQtd - 1]);
    const loja  = cleanText_(row[appCols.colLoja - 1]);
    const preco = toNum_(row[appCols.colPreco - 1]);

    if (!item || !qtd || !loja || !isFinite(preco) || preco <= 0) continue;

    const lojaCol = storeMap.get(normalize_(loja));
    if (!lojaCol) continue;

    const rowKey = makePesquisaRowKey_(item, marca, qtd);

    needRows.set(rowKey, { item, marca, qtd });
    ops.push({ rowKey, lojaCol, preco });
  }

  if (ops.length === 0) {
    ss.toast(`Nada do App para exportar em ${targetDateKey}.`, 'Sync App → PC', 7);
    return;
  }

  // ===== indexa linhas existentes (B:C:D) sem varrer a planilha inteira
  const firstDataRow = SYNC_CFG.firstDataRowPesquisa;
  const headerRow = SYNC_CFG.headerRowPesquisa;

  let lastItemRow = lastUsedRowInCol_(sh, SYNC_CFG.colItem, firstDataRow); // última linha com Item
  if (lastItemRow < firstDataRow) lastItemRow = firstDataRow - 1;

  const index = new Map(); // rowKey -> rowNum
  if (lastItemRow >= firstDataRow) {
    const n = lastItemRow - firstDataRow + 1;
    const baseVals = sh.getRange(firstDataRow, SYNC_CFG.colItem, n, 3).getValues(); // B:C:D
    for (let i = 0; i < baseVals.length; i++) {
      const item  = cleanText_(baseVals[i][0]);
      const marca = cleanText_(baseVals[i][1]);
      const qtd   = cleanQtd_(baseVals[i][2]);
      if (!item || !qtd) continue;
      index.set(makePesquisaRowKey_(item, marca, qtd), firstDataRow + i);
    }
  }

  // ===== cria em lote as linhas que faltam, mas REUTILIZA linhas em branco
  const missing = [];
  for (const [k, v] of needRows.entries()) {
    if (!index.has(k)) missing.push({ key: k, ...v });
  }

  if (missing.length > 0) {
  const insertAfter = Math.max(lastItemRow, headerRow);
  const startRow = insertAfter + 1;
  const need = missing.length;

  // garante que existem linhas físicas suficientes (sem “inflar” à toa)
  const lastNeededRow = startRow + need - 1;
  const maxRows = sh.getMaxRows();
  if (lastNeededRow > maxRows) {
    sh.insertRowsAfter(maxRows, lastNeededRow - maxRows);
  }

  // SEMPRE copie de uma linha-modelo fixa (linha 3) — deve ter fórmulas/validações
  const templateRow = SYNC_CFG.firstDataRowPesquisa; // 3
  sh.getRange(templateRow, 1, 1, storeEndCol)
    .copyTo(sh.getRange(startRow, 1, need, storeEndCol), { contentsOnly: false });

  // limpa preços e os campos B:C:D (pra garantir “linha zerada”)
  sh.getRange(startRow, storeStartCol, need, storeCols).clearContent();
  sh.getRange(startRow, SYNC_CFG.colItem, need, 3).clearContent();

  // escreve Item/Marca/Qntd (B:C:D) em lote
  const out = missing.map(r => [r.item, r.marca || '', r.qtd]);
  sh.getRange(startRow, SYNC_CFG.colItem, need, 3).setValues(out);

  // atualiza index
  for (let i = 0; i < missing.length; i++) {
    index.set(missing[i].key, startRow + i);
  }

  lastItemRow = startRow + need - 1;
  }

  // ===== aplica preços em MEMÓRIA e escreve em BLOCO
  const totalRows = Math.max(0, lastItemRow - firstDataRow + 1);
  const storeRange = sh.getRange(firstDataRow, storeStartCol, totalRows, storeCols);

  const storeVals = storeRange.getValues();
  const storeNotes = storeRange.getNotes();

  const itemNotesRange = sh.getRange(firstDataRow, SYNC_CFG.colItem, totalRows, 1);
  const itemNotes = itemNotesRange.getNotes();

  let changed = 0;
  const dupRows = new Set();
  const dupCells = [];

  for (const op of ops) {
    const rowNum = index.get(op.rowKey);
    if (!rowNum) continue;

    const r = rowNum - firstDataRow;           // offset em storeVals
    const c = op.lojaCol - storeStartCol;      // offset em storeVals

    const old = toNum_(storeVals[r][c]);
    if (isFinite(old) && old > 0 && Math.abs(old - op.preco) > 1e-9) {
      dupRows.add(rowNum);
      dupCells.push([rowNum, op.lojaCol]);
    }

    storeVals[r][c] = op.preco;

    storeNotes[r][c] = appendNoteStr_(storeNotes[r][c], SYNC_CFG.noteApp);
    itemNotes[r][0]  = appendNoteStr_(itemNotes[r][0],  SYNC_CFG.noteApp);

    changed++;
  }

  storeRange.setValues(storeVals);
  storeRange.setNotes(storeNotes);
  itemNotesRange.setNotes(itemNotes);

  // bordas vermelhas só onde houve conflito (geralmente poucos)
  for (const rowNum of dupRows) {
    sh.getRange(rowNum, 1, 1, 5).setBorder(
      true, true, true, true, true, true,
      SYNC_CFG.dupBorderColor,
      SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );
  }
  for (const [rowNum, col] of dupCells) {
    sh.getRange(rowNum, col).setBorder(
      true, true, true, true, true, true,
      SYNC_CFG.dupBorderColor,
      SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );
  }
  SpreadsheetApp.flush();
  try { atualizarGradienteTudo(); } catch (e) {}
  try { aplicarSeparadoresItens(); } catch (e) {}
  try { adicionarBordasPesquisa(); } catch (e) {}
  ss.toast(`OK: ${changed} preços aplicados | ${dupCells.length} conflitos marcados.`, 'Sync App → PC', 8);
}


// =======================
// Helpers (AppSheet sheet)
// =======================

/**
 * Verifica se a nota contém a marcação de edição via App.
 * @private
 * @param {string} s - O texto da nota
 * @returns {boolean}
 */
function noteHasApp_(s) {
  return normalize_(String(s || '')).includes(normalize_(SYNC_CFG.noteApp));
}

/**
 * Normaliza textos removendo acentos e caracteres invisíveis.
 * @private
 * @param {any} v - Valor a ser normalizado
 * @returns {string}
 */
// normalização forte (remove acento + espaços invisíveis)
function normText_(v) {
  return String(v ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')        // acentos
    .replace(/[\u200B-\u200D\uFEFF]/g, '') // zero-width / BOM
    .trim()
    .toLowerCase();
}

/**
 * Localiza a última linha preenchida em uma coluna específica.
 * @private
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - Aba alvo.
 * @param {number} col - Índice da coluna.
 * @param {number} startRow - Linha de início da busca.
 * @returns {number}
 */
function lastUsedRowInCol_(sh, col, startRow) {
  const lr = sh.getLastRow();
  if (lr < startRow) return startRow - 1;

  const n = lr - startRow + 1;
  const vals = sh.getRange(startRow, col, n, 1).getDisplayValues();

  for (let i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0] ?? '').trim() !== '') return startRow + i;
  }
  return startRow - 1;
}

/**
 * Retorna um objeto com os índices mapeados de cada coluna na aba do AppSheet.
 * Busca dinamicamente comparando cabeçalhos com variações de nomes.
 * @private
 * @param {GoogleAppsScript.Spreadsheet.Sheet} appSheet - Aba do AppSheet.
 * @returns {Object}
 */
function getAppCols_(appSheet) {
  const headers = appSheet.getRange(1, 1, 1, appSheet.getLastColumn()).getValues()[0].map(h => String(h || '').trim());
  const norm = headers.map(h => normalize_(h));

  const colID     = findHeader_(norm, ['id']);
  const colData   = findHeader_(norm, ['data', 'date']);
  const colItem   = findHeader_(norm, ['item']);
  const colMarca  = findHeader_(norm, ['marca', 'brand']);
  const colQtd    = findHeader_(norm, ['qntd', 'qtd', 'quantidade', 'quantity']);
  const colLoja   = findHeader_(norm, ['loja', 'store']);
  const colPreco  = findHeader_(norm, ['preco', 'preço', 'price']);
  const colOrigem = findHeader_(norm, ['origem', 'source']);

  return {
    lastCol: appSheet.getLastColumn(),
    colID, colData, colItem, colMarca, colQtd, colLoja, colPreco, colOrigem,
    headers
  };
}

/**
 * Verifica se todas as colunas obrigatórias foram localizadas na aba do App.
 * @private
 * @param {Object} c - Objeto gerado por getAppCols_
 * @throws {Error} - Se faltar alguma coluna vital
 */
function ensureAppHeadersOk_(c) {
  const missing = [];
  if (!c.colID) missing.push('ID');
  if (!c.colData) missing.push('Data');
  if (!c.colItem) missing.push('Item');
  if (!c.colMarca) missing.push('Marca');
  if (!c.colQtd) missing.push('Qntd');
  if (!c.colLoja) missing.push('Loja');
  if (!c.colPreco) missing.push('Preço');
  if (!c.colOrigem) missing.push('Origem');
  if (missing.length) throw new Error('Preços app: faltando colunas: ' + missing.join(', '));
}

/**
 * Constrói o array que representa a linha a ser inserida/atualizada na aba do App.
 * @private
 * @param {Object} appCols - Índices das colunas
 * @param {string} id - ID único (UUID)
 * @param {Date|string} dateObj - Data alvo
 * @param {string} item - Nome do Item
 * @param {string} marca - Marca
 * @param {number|string} qtd - Quantidade
 * @param {string} loja - Nome da loja
 * @param {number} preco - Valor monetário
 * @param {string} origem - PC ou App
 * @returns {Array} Linha formatada.
 */
function makeAppRow_(appCols, id, dateObj, item, marca, qtd, loja, preco, origem) {
  const row = new Array(appCols.lastCol).fill('');
  row[appCols.colID - 1] = id;
  row[appCols.colData - 1] = dateObj instanceof Date ? dateObj : new Date(dateObj);
  row[appCols.colItem - 1] = item;
  row[appCols.colMarca - 1] = marca || '';
  row[appCols.colQtd - 1] = typeof qtd === 'number' ? qtd : toNum_(qtd);
  row[appCols.colLoja - 1] = loja;
  row[appCols.colPreco - 1] = preco;
  row[appCols.colOrigem - 1] = origem;
  return row;
}

/**
 * Constrói um Map em memória representando os registros existentes na aba do App.
 * Chaveada pela chave natural (makeNaturalKey_).
 * @private
 * @param {GoogleAppsScript.Spreadsheet.Sheet} appSheet - Aba alvo.
 * @param {Object} appCols - Índices das colunas.
 * @returns {Map<string, Object>}
 */
function buildAppIndex_(appSheet, appCols) {
  const idx = new Map();
  const lr = appSheet.getLastRow();
  if (lr < 2) return idx;

  const vals = appSheet.getRange(2, 1, lr - 1, appCols.lastCol).getValues();

  for (let i = 0; i < vals.length; i++) {
    const row = vals[i];
    const id = String(row[appCols.colID - 1] || '').trim();
    const data = row[appCols.colData - 1];
    const item = cleanText_(row[appCols.colItem - 1]);
    const marca = cleanText_(row[appCols.colMarca - 1]);
    const qtd = cleanQtd_(row[appCols.colQtd - 1]);
    const loja = cleanText_(row[appCols.colLoja - 1]);
    const origem = cleanText_(row[appCols.colOrigem - 1]);

    if (!id || !item || !qtd || !loja) continue;

    const dateKey = dateKeyFromAny_(data);
    if (!dateKey) continue;

    const key = makeNaturalKey_(dateKey, item, marca, qtd, loja);
    idx.set(key, { row: i + 2, id, origem });
  }

  return idx;
}


// =======================
// Helpers (Pesquisa sheet)
// =======================

/**
 * Concatena uma string de nota preservando notas existentes.
 * @private
 * @param {string} cur - Nota atual
 * @param {string} text - Texto a adicionar
 * @returns {string} Nota concatenada
 */
function appendNoteStr_(cur, text) {
  const a = String(cur || '').trim();
  const t = String(text || '').trim();
  if (!t) return a;
  if (!a) return t;
  if (a.toLowerCase().includes(t.toLowerCase())) return a;
  return a + '\n' + t;
}

/**
 * Identifica e valida a data presente no cabeçalho da pesquisa (ex: B1).
 * Suporta formatos avulsos ou leitura externa via `hist_getPesquisaDateRangeFromB1_`.
 * @private
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - Aba de pesquisa.
 * @returns {Object} Objeto com a data nativa e a string dateKey formatada.
 */
function getPesquisaDate_(sh) {
  // Se existir o parser do Histórico (B1 texto/período), reaproveita ele
  if (typeof hist_getPesquisaDateRangeFromB1_ === 'function') {
    const dr = hist_getPesquisaDateRangeFromB1_(sh);
    if (dr && dr.error) throw new Error(dr.error);

    // Sync usa a data INICIAL do período
    return {
      dateObj: dr.start,
      dateKey: dr.startKey
    };
  }

  // Fallback local (B1 lido como texto)
  const rawDisplay = String(sh.getRange(SYNC_CFG.dateCellPesquisa).getDisplayValue() || '').trim();
  const parsed = parsePesquisaDateText_(rawDisplay);

  if (!parsed) {
    throw new Error(
      `Data inválida em ${SYNC_CFG.dateCellPesquisa}. Use, por exemplo: 18/02/2026 ou 18-25/02/2026`
    );
  }

  return {
    dateObj: parsed.start,
    dateKey: parsed.startKey
  };
}

/**
 * Transforma uma string de período de data (ex: 18-25/02) em objetos de Data.
 * @private
 * @param {string} txt - O texto bruto
 * @returns {Object|null} Objeto contendo dados processados de Data.
 */
function parsePesquisaDateText_(txt) {
  const TZ = Session.getScriptTimeZone();
  const s0 = String(txt || '').trim();
  if (!s0) return null;

  // normaliza espaços e conectores
  const s = s0
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/\bat[eé]\b/gi, ' a ')
    .trim();

  let m;

  // 1) Data única: dd/mm/yyyy
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    const d = +m[1];
    const mo = +m[2];
    const y = +m[3];
    const dt = makeValidDate_(y, mo, d);
    if (!dt) return null;

    const startKey = Utilities.formatDate(dt, TZ, 'yyyy-MM-dd');
    return { start: dt, end: dt, startKey, endKey: startKey, days: 1 };
  }

  // 2) Período completo: dd/mm/yyyy a dd/mm/yyyy
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s*(?:a|-)\s*(\d{1,2})\/(\d{1,2})\/(\d{4})$/i);
  if (m) {
    const d1 = +m[1], m1 = +m[2], y1 = +m[3];
    const d2 = +m[4], m2 = +m[5], y2 = +m[6];

    let start = makeValidDate_(y1, m1, d1);
    let end = makeValidDate_(y2, m2, d2);
    if (!start || !end) return null;

    if (start.getTime() > end.getTime()) {
      const tmp = start;
      start = end;
      end = tmp;
    }

    const startKey = Utilities.formatDate(start, TZ, 'yyyy-MM-dd');
    const endKey = Utilities.formatDate(end, TZ, 'yyyy-MM-dd');
    const days = diffDaysInclusive_(start, end);

    return { start, end, startKey, endKey, days };
  }

  // 3) Atalho: dd-dd/mm/yyyy  (ex: 18-25/02/2026)
  m = s.match(/^(\d{1,2})\s*-\s*(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    const d1 = +m[1];
    const d2 = +m[2];
    const mo = +m[3];
    const y = +m[4];

    let start = makeValidDate_(y, mo, d1);
    let end = makeValidDate_(y, mo, d2);
    if (!start || !end) return null;

    if (start.getTime() > end.getTime()) {
      const tmp = start;
      start = end;
      end = tmp;
    }

    const startKey = Utilities.formatDate(start, TZ, 'yyyy-MM-dd');
    const endKey = Utilities.formatDate(end, TZ, 'yyyy-MM-dd');
    const days = diffDaysInclusive_(start, end);

    return { start, end, startKey, endKey, days };
  }

  // 4) Atalho: dd/mm a dd/mm/yyyy  (ex: 18/02 a 25/02/2026)
  m = s.match(/^(\d{1,2})\/(\d{1,2})\s*(?:a|-)\s*(\d{1,2})\/(\d{1,2})\/(\d{4})$/i);
  if (m) {
    const d1 = +m[1], m1 = +m[2];
    const d2 = +m[3], m2 = +m[4], y = +m[5];

    let start = makeValidDate_(y, m1, d1);
    let end = makeValidDate_(y, m2, d2);
    if (!start || !end) return null;

    if (start.getTime() > end.getTime()) {
      const tmp = start;
      start = end;
      end = tmp;
    }

    const startKey = Utilities.formatDate(start, TZ, 'yyyy-MM-dd');
    const endKey = Utilities.formatDate(end, TZ, 'yyyy-MM-dd');
    const days = diffDaysInclusive_(start, end);

    return { start, end, startKey, endKey, days };
  }

  // 5) Atalho: dd a dd/mm/yyyy  (ex: 18 a 25/02/2026)
  m = s.match(/^(\d{1,2})\s*(?:a|-)\s*(\d{1,2})\/(\d{1,2})\/(\d{4})$/i);
  if (m) {
    const d1 = +m[1];
    const d2 = +m[2];
    const mo = +m[3];
    const y = +m[4];

    let start = makeValidDate_(y, mo, d1);
    let end = makeValidDate_(y, mo, d2);
    if (!start || !end) return null;

    if (start.getTime() > end.getTime()) {
      const tmp = start;
      start = end;
      end = tmp;
    }

    const startKey = Utilities.formatDate(start, TZ, 'yyyy-MM-dd');
    const endKey = Utilities.formatDate(end, TZ, 'yyyy-MM-dd');
    const days = diffDaysInclusive_(start, end);

    return { start, end, startKey, endKey, days };
  }

  return null;
}

/**
 * Auxilia a criação de instâncias Date robustas (valida falsos meses, ex 31/02)
 * @private
 * @param {number} y - Ano
 * @param {number} m - Mês
 * @param {number} d - Dia
 * @returns {Date|null}
 */
function makeValidDate_(y, m, d) {
  if (!Number.isFinite(y) || !Number.isFinite(m) || !Number.isFinite(d)) return null;
  if (y < 1900 || y > 3000) return null;
  if (m < 1 || m > 12) return null;
  if (d < 1 || d > 31) return null;

  const dt = new Date(y, m - 1, d);
  if (isNaN(dt.getTime())) return null;

  // validação real (evita 31/02 virar 03/03)
  if (
    dt.getFullYear() !== y ||
    dt.getMonth() !== (m - 1) ||
    dt.getDate() !== d
  ) return null;

  return dt;
}

/**
 * Calcula a diferença em dias entre duas datas (inclusivo).
 * @private
 */
function diffDaysInclusive_(start, end) {
  const a = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  const b = new Date(end.getFullYear(), end.getMonth(), end.getDate());
  const msPerDay = 24 * 60 * 60 * 1000;
  return Math.floor((b.getTime() - a.getTime()) / msPerDay) + 1;
}

/**
 * Identifica a coluna onde acabam as lojas na aba de pesquisa,
 * parando antes das colunas "Custo" ou "Mult".
 * @private
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - Aba de pesquisa.
 * @returns {Object} Limite de colunas e nomes das lojas detectadas.
 */
// acha onde acabam as lojas (antes de “Custo”/“Mult”)
function getStoresFromPesquisa_(sh) {
  const lastCol = sh.getLastColumn();
  const hdr = sh.getRange(SYNC_CFG.headerRowPesquisa, 1, 1, lastCol).getValues()[0].map(v => normalize_(v));

  let colCusto = 0, colMult = 0;
  for (let c = 0; c < hdr.length; c++) {
    const h = hdr[c];
    if (!h) continue;
    if (!colCusto && h.includes('custo')) colCusto = c + 1;
    if (!colMult && (h === 'mult' || h.includes('mult'))) colMult = c + 1;
  }

  let storeEndCol = lastCol;
  const candidates = [colCusto, colMult].filter(v => v && v > SYNC_CFG.storeStartCol);
  if (candidates.length) storeEndCol = Math.min(...candidates) - 1;

  const storeCols = storeEndCol - SYNC_CFG.storeStartCol + 1;
  const stores = (storeCols > 0)
    ? sh.getRange(SYNC_CFG.headerRowPesquisa, SYNC_CFG.storeStartCol, 1, storeCols)
        .getValues()[0]
        .map(v => String(v || '').trim())
    : [];

  return { storeEndCol, stores };
}

/**
 * Mapeia as lojas com seus respectivos índices de coluna.
 * @private
 */
function getStoresMapFromPesquisa_(sh) {
  const { storeEndCol, stores } = getStoresFromPesquisa_(sh);
  const map = new Map();
  for (let i = 0; i < stores.length; i++) {
    const name = stores[i];
    if (!name) continue;
    map.set(normalize_(name), SYNC_CFG.storeStartCol + i);
  }
  return { storeEndCol, storeMap: map };
}

/**
 * Cria um índice das linhas da aba de pesquisa mapeando Item|Marca|Quantidade
 * para os números das linhas em que se encontram.
 * @private
 */
// Indexa linhas existentes por (Item|Marca|Qntd)
function buildPesquisaIndex_(sh) {
  const lastRow = sh.getLastRow();
  const idx = new Map();
  if (lastRow < SYNC_CFG.firstDataRowPesquisa) return idx;

  const vals = sh.getRange(
    SYNC_CFG.firstDataRowPesquisa, 1,
    lastRow - SYNC_CFG.firstDataRowPesquisa + 1,
    5 // A:E
  ).getValues();

  for (let i = 0; i < vals.length; i++) {
    const rowNum = SYNC_CFG.firstDataRowPesquisa + i;
    const item = cleanText_(vals[i][SYNC_CFG.colItem - 1]);
    const marca = cleanText_(vals[i][SYNC_CFG.colMarca - 1]);
    const qtd = cleanQtd_(vals[i][SYNC_CFG.colQtd - 1]);
    if (!item || !qtd) continue;

    const k = makePesquisaRowKey_(item, marca, qtd);
    // guarda o último (pra inserir no bloco certo)
    idx.set(k, rowNum);
  }

  return idx;
}

/**
 * Verifica se um Item+Marca+Qtd já existe e o atualiza; se não, cria
 * a respectiva linha de forma segura.
 * @private
 */
// cria/acha linha de (Item, Marca, Qntd) sem tocar em Tipo (E)
function upsertPesquisaRow_(sh, index, item, marca, qtd) {
  const k = makePesquisaRowKey_(item, marca, qtd);
  const existingRow = index.get(k);
  if (existingRow) return existingRow;

  let insertAt = findInsertRowForItem_(sh, item);

  // Se insertAt estourar o limite, adiciona no fim de forma segura
  if (insertAt > sh.getMaxRows()) {
    sh.insertRowAfter(sh.getMaxRows());          // cria 1 linha no fim
    insertAt = sh.getMaxRows();                  // a linha nova é a última
  } else {
    sh.insertRowBefore(insertAt);
  }

  // escolhe uma linha modelo para copiar formatação/validações
  const lr = sh.getLastRow();
  let modelRow = null;

  if (insertAt < lr) modelRow = insertAt + 1;          
  else if (insertAt - 1 >= SYNC_CFG.firstDataRowPesquisa) modelRow = insertAt - 1;

  if (modelRow) {
    sh.getRange(modelRow, 1, 1, sh.getLastColumn())
      .copyTo(sh.getRange(insertAt, 1, 1, sh.getLastColumn()), { contentsOnly: false });
    clearRowPrices_(sh, insertAt);
  }

  // escreve só B,C,D (não mexe na E - Tipo)
  sh.getRange(insertAt, SYNC_CFG.colItem).setValue(item);
  sh.getRange(insertAt, SYNC_CFG.colMarca).setValue(marca || '');
  sh.getRange(insertAt, SYNC_CFG.colQtd).setValue(qtd);

  index.set(k, insertAt);
  return insertAt;
}


/**
 * Encontra a linha onde um novo item recém-sincronizado deve ser inserido.
 * - se o item já existe: insere logo abaixo do último daquele item
 * - se não existe: insere no fim dos itens (calculado pela coluna do Item)
 * @private
 */
// - se o item já existe: insere logo abaixo do último daquele item
// - se não existe: insere no fim dos itens (calculado pela coluna do Item)
function findInsertRowForItem_(sh, item) {
  const key = normText_(item);
  if (!key) return SYNC_CFG.firstDataRowPesquisa;

  const start = SYNC_CFG.firstDataRowPesquisa;
  const lastItemRow = lastUsedRowInCol_(sh, SYNC_CFG.colItem, start);

  // se não tem itens ainda
  if (lastItemRow < start) return start;

  const n = lastItemRow - start + 1;
  const items = sh.getRange(start, SYNC_CFG.colItem, n, 1).getDisplayValues();

  let lastMatch = -1;
  for (let i = 0; i < items.length; i++) {
    if (normText_(items[i][0]) === key) lastMatch = start + i;
  }

  // achou o item -> logo abaixo do último match
  if (lastMatch !== -1) return lastMatch + 1;

  // não achou -> fim dos itens (pela coluna do item)
  return lastItemRow + 1;
}

/**
 * Reseta o conteúdo das colunas de preços para uma dada linha.
 * @private
 */
function clearRowPrices_(sh, row) {
  // limpa colunas de lojas (F..até antes de custo/mult)
  const { storeEndCol } = getStoresFromPesquisa_(sh);
  if (storeEndCol >= SYNC_CFG.storeStartCol) {
    sh.getRange(row, SYNC_CFG.storeStartCol, 1, storeEndCol - SYNC_CFG.storeStartCol + 1).clearContent();
  }
}

// =======================
// Duplicados / marcações
// =======================

/**
 * Pinta de vermelho os itens em conflito ao sincronizar do App para o PC.
 * @private
 */
function markDuplicatePesquisa_(sh, rowNum, lojaCol) {
  // borda vermelha em A:E (inclui checkbox + item + marca + qntd + tipo)
  sh.getRange(rowNum, 1, 1, 5).setBorder(
    true, true, true, true, true, true,
    SYNC_CFG.dupBorderColor,
    SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );

  // borda vermelha também na célula do preço
  sh.getRange(rowNum, lojaCol).setBorder(
    true, true, true, true, true, true,
    SYNC_CFG.dupBorderColor,
    SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
}

/**
 * Evita repetição e duplicação da nota visual do APP.
 * @private
 */
function addNote_(cell, text) {
  const cur = String(cell.getNote() || '').trim();
  if (!cur) {
    cell.setNote(text);
    return;
  }
  // evita duplicar
  if (!cur.toLowerCase().includes(String(text).toLowerCase())) {
    cell.setNote(cur + '\n' + text);
  }
}

// =======================
// Canonização de Item (dropdown)
// =======================

/**
 * Cria um map lendo os itens da aba Inventário para que a entrada via 
 * App seja canonizada/padronizada de acordo com os dropdowns oficiais.
 * @private
 */
function buildCanonicalItemMap_() {
  const ss = SpreadsheetApp.getActive();
  const inv = ss.getSheetByName(SYNC_CFG.sheetInventario);
  const map = new Map();
  if (!inv) return map;

  // Item no Inventário está na coluna C, dados a partir da linha 3 (pelo seu padrão)
  const lastRow = inv.getLastRow();
  if (lastRow < 3) return map;

  const vals = inv.getRange(3, 3, lastRow - 2, 1).getDisplayValues().flat();
  for (const v of vals) {
    const s = String(v || '').trim();
    if (!s) continue;
    map.set(normalize_(s), s);
  }
  return map;
}

/**
 * Compara o texto da entrada com o array oficial para forçar correção.
 * @private
 */
function canonize_(canonMap, item) {
  const key = normalize_(item);
  return canonMap.get(key) || item;
}

// =======================
// Util
// =======================

/**
 * Lança erro caso uma aba não seja encontrada para interromper o fluxo.
 * @private
 */
function mustGetSheet_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Não encontrei a aba "${name}".`);
  return sh;
}

/**
 * Mapeia e localiza o número (index) de colunas dinamicamente.
 * @private
 */
function findHeader_(normHeaders, variants) {
  const vset = variants.map(normalize_);
  for (let i = 0; i < normHeaders.length; i++) {
    const h = normHeaders[i];
    if (!h) continue;
    if (vset.includes(h)) return i + 1;
  }
  return 0;
}

/**
 * Gera a string formatada para a chave natural de identificação.
 * @private
 */
function makeNaturalKey_(dateKey, item, marca, qtd, loja) {
  return [
    String(dateKey),
    normalize_(item),
    normalize_(marca || ''),
    String(qtd).trim(),
    normalize_(loja)
  ].join('|');
}

/**
 * Gera a string formatada de identificação em nível de linha de pesquisa.
 * @private
 */
function makePesquisaRowKey_(item, marca, qtd) {
  return [
    normalize_(item),
    normalize_(marca || ''),
    String(qtd).trim()
  ].join('|');
}

/**
 * Limpa blocos de espaço de um texto avulso.
 * @private
 */
function cleanText_(v) {
  return String(v ?? '').replace(/\s+/g, ' ').trim();
}

/**
 * Limpa formatações numéricas e string para extrair valor QTD bruto.
 * @private
 */
function cleanQtd_(v) {
  // AppSheet pode mandar número; planilha pode ter texto
  if (typeof v === 'number' && isFinite(v)) return v;
  const s = String(v ?? '').trim();
  if (!s) return '';
  // aceita "1", "1.5", "1,5"
  const n = Number(s.replace(',', '.'));
  return isFinite(n) ? n : s;
}

/**
 * Interpreta preços com texto, vírgula e sufixos financeiros, devolvendo um Float.
 * @private
 */
function toNum_(v) {
  if (typeof v === 'number') return v;
  const s = String(v ?? '').trim();
  if (!s) return NaN;
  const cleaned = s
    .replace(/\s/g, '')
    .replace(/^R\$/i, '')
    .replace(/\./g, '')
    .replace(',', '.');
  const n = Number(cleaned);
  return isFinite(n) ? n : NaN;
}

/**
 * Normaliza os textos convertendo para Low Case e removendo qualquer acento ou espaços excessivos.
 * @private
 */
function normalize_(s) {
  return String(s ?? '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}

/**
 * Obtém o string index de uma string ou objeto Data para bater contra targetDate.
 * @private
 */
function dateKeyFromAny_(d) {
  if (!d) return '';
  let dt = d;
  if (!(dt instanceof Date)) {
    const parsed = Date.parse(String(d));
    if (!isNaN(parsed)) dt = new Date(parsed);
  }
  if (!(dt instanceof Date) || isNaN(dt.getTime())) return '';
  return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * Função gatilho para a Sincronização Bidirecional completa com LockService.
 * @returns {void}
 */
function syncPcEApp() {
  const ss = SpreadsheetApp.getActive();
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(30000)) {
    ss.toast('Já existe uma sincronização em andamento.', 'Sync', 5);
    return;
  }

  try {
    ss.toast('Sincronizando: App → PC...', 'Sync', 5);
    syncAppParaPc();   // sua função existente

    ss.toast('Sincronizando: PC → App...', 'Sync', 5);
    syncPcParaApp();   // sua função existente

    ss.toast('Sincronização completa (App ⇄ PC).', 'Sync', 6);
  } catch (err) {
    ss.toast('Erro na sincronização: ' + (err && err.message ? err.message : err), 'Sync', 10);
    throw err;
  } finally {
    lock.releaseLock();
  }
}
