/***********************
 * Histórico (v5) + período em B1 + normalização
 *
 * O campo B1 (CFG.CELLS.PESQ.DATA_PESQUISA) é interpretado SEMPRE como texto (display).
 * Formatos aceitos:
 * - 18/02/2026
 * - 18/02/2026 a 25/02/2026
 * - 18/02/2026 ate 25/02/2026  (também aceita "até")
 * - 18/02/2026 - 25/02/2026
 * - 18-25/02/2026              (atalho: mesmo mês/ano do final)
 * - 18/02 - 25/02/2026         (atalho: ano do final)
 *
 * Mantém lógica:
 * - incremental: não apaga ausentes
 * - sync: apaga ausentes só dentro das séries presentes (item+marca+tipo+qtd)
 * - Travar=TRUE protege linha contra alterações no salvar e na normalização
 * - Timestamp em K ("Salvo em") sempre atualizado em updates e inserido em inserts
 *
 * Undo/Redo (compacto):
 * - keepLast = 3
 * - deltas ficam numa aba oculta (_HIST_TXN); propriedades guardam só metadados
 ***********************/

/**
 * Configurações para a pilha de Undo/Redo e transações.
 * @constant {Object}
 */
const HIST_STACK = {
  undoKey: 'HIST_UNDO_STACK_V5',
  redoKey: 'HIST_REDO_STACK_V5',
  keepLast: 3,

  txnSheetName: '_HIST_TXN',         // aba oculta de deltas
  txnHeaderCols: 10,                 // A:J
  txnRebuildThresholdRows: 50000,    // se passar disso, compacta para só o que está nos stacks

  maxOpsToStore: 20000               // limite de segurança para undo (não limita salvar)
};

// ====== ENTRADAS PÚBLICAS ======

/**
 * Salva os preços do dia atual de forma incremental.
 * Adiciona novos dados ou atualiza existentes sem apagar itens que não foram lançados hoje.
 * @returns {void}
 */
function salvarPrecosDoDia() {
  // incremental (não apaga ausentes)
  salvarPrecosDoPeriodo_(false);
}

/**
 * Sincroniza os preços do dia atual.
 * Apaga preços antigos do histórico que pertencem à mesma série (item+marca+tipo+qtd) 
 * mas não foram relançados na pesquisa atual (útil para correção).
 * @returns {void}
 */
function sincronizarPrecosDoDia() {
  // sync (apaga ausentes dentro das séries presentes)
  salvarPrecosDoPeriodo_(true);
}

/**
 * Desfaz a última operação de salvar ou sincronizar no Histórico.
 * Recupera o estado da aba oculta de transações ou dos metadados antigos.
 * @returns {void}
 */
function desfazerUltimaAcaoHistorico() {
  const ss = SpreadsheetApp.getActive();
  let hist;
  try {
    hist = CFG.mustGetSheet(ss, CFG.SHEETS.HIST);
  } catch (e) {
    return ss.toast(String(e.message || e), 'Undo', 6);
  }

  const undo = hist_loadUndo_();
  if (undo.length === 0) return ss.toast('Nada para desfazer.', 'Undo', 6);

  const txn = undo.pop();
  hist_saveUndo_(undo);
  hist_pushRedo_(txn);

  const { headerRowHist, firstDataRowHist } = hist_findHeader_(hist);

  // Compatibilidade: txn compacto (novo) ou legado (antigo)
  if (hist_isCompactTxn_(txn)) {
    const ops = hist_readCompactTxnOps_(ss, txn);
    if (!ops) return; // toast já foi dado

    // Agrupa por tipo
    const insByDate = new Map();
    const updByDate = new Map();
    const delRows = [];

    for (const op of ops) {
      if (op.op === 'INS') {
        if (!insByDate.has(op.dateKey)) insByDate.set(op.dateKey, []);
        insByDate.get(op.dateKey).push(op);
      } else if (op.op === 'UPD') {
        if (!updByDate.has(op.dateKey)) updByDate.set(op.dateKey, []);
        updByDate.get(op.dateKey).push(op);
      } else if (op.op === 'DEL') {
        if (op.rowValues) delRows.push(op.rowValues);
      }
    }

    // 1) remover inserted (ignora Travar=true, igual comportamento antigo)
    for (const [dateKey, list] of insByDate.entries()) {
      let keyToRow = hist_buildKeyRowMap_(hist, dateKey, false, headerRowHist, firstDataRowHist);
      const rowsToDelete = [];
      for (const it of list) {
        const row = keyToRow.get(it.key)?.row;
        if (row) rowsToDelete.push(row);
      }
      rowsToDelete.sort((a, b) => b - a);
      for (const r of rowsToDelete) hist.deleteRow(r);
    }

    // 2) reverter updates (Preço + Timestamp) (ignora Travar=true, igual comportamento antigo)
    for (const [dateKey, list] of updByDate.entries()) {
      const keyToRow = hist_buildKeyRowMap_(hist, dateKey, false, headerRowHist, firstDataRowHist);
      for (const u of list) {
        const row = keyToRow.get(u.key)?.row;
        if (!row) continue;

        // Preço (G)
        if (isFinite(u.oldPrice)) hist.getRange(row, CFG.COL.HIST.PRECO).setValue(u.oldPrice);
        else hist.getRange(row, CFG.COL.HIST.PRECO).setValue('');

        // Timestamp (K)
        if (u.oldSavedAtMs) hist.getRange(row, CFG.COL.HIST.TIMESTAMP).setValue(new Date(u.oldSavedAtMs));
        else hist.getRange(row, CFG.COL.HIST.TIMESTAMP).setValue('');
      }
    }

    // 3) restaurar deletes (insere no topo)
    if (delRows.length > 0) {
      hist.insertRowsBefore(firstDataRowHist, delRows.length);

      // A:H
      const out = delRows.map(v => ([
        hist_dateFromKey_(v[0]),
        v[1], v[2], v[3], v[4],
        v[5], v[6], v[7]
      ]));
      hist.getRange(firstDataRowHist, 1, out.length, 8).setValues(out);

      // K
      const tsOut = delRows.map(v => [v[8] ? new Date(v[8]) : '']);
      hist.getRange(firstDataRowHist, CFG.COL.HIST.TIMESTAMP, tsOut.length, 1).setValues(tsOut);
    }

    ss.toast(`Undo ${txn.label || 'ok'}.`, 'Undo', 8);
    return;
  }

  // Legado: suporta dayTxns ou dateKey único
  const dayTxns = txn.dayTxns ? txn.dayTxns : [txn];
  const label = txn.dayTxns
    ? `${hist_formatBr_(dayTxns[0].dateKey)} a ${hist_formatBr_(dayTxns[dayTxns.length - 1].dateKey)}`
    : hist_formatBr_(txn.dateKey);

  for (const t of dayTxns) {
    const dateKey = t.dateKey;

    // remove inserted
    let keyToRow = hist_buildKeyRowMap_(hist, dateKey, false, headerRowHist, firstDataRowHist);
    const rowsToDelete = [];
    for (const r of (t.inserted || [])) {
      const key = hist_makeKey_(r[0], r[1], r[2], r[3], r[4], r[5]);
      const row = keyToRow.get(key)?.row;
      if (row) rowsToDelete.push(row);
    }
    rowsToDelete.sort((a, b) => b - a);
    for (const row of rowsToDelete) hist.deleteRow(row);

    // rebuild map
    keyToRow = hist_buildKeyRowMap_(hist, dateKey, false, headerRowHist, firstDataRowHist);

    // revert updates
    for (const u of (t.updated || [])) {
      const row = keyToRow.get(u.key)?.row;
      if (!row) continue;

      if (isFinite(u.oldPrice)) hist.getRange(row, CFG.COL.HIST.PRECO).setValue(u.oldPrice);
      else hist.getRange(row, CFG.COL.HIST.PRECO).setValue('');

      if (u.oldSavedAtMs) hist.getRange(row, CFG.COL.HIST.TIMESTAMP).setValue(new Date(u.oldSavedAtMs));
      else hist.getRange(row, CFG.COL.HIST.TIMESTAMP).setValue('');
    }

    // restore deleted
    const deleted = t.deleted || [];
    if (deleted.length > 0) {
      hist.insertRowsBefore(firstDataRowHist, deleted.length);

      const out = deleted.map(d => [
        hist_dateFromKey_(d.values[0]),
        d.values[1], d.values[2], d.values[3], d.values[4],
        d.values[5], d.values[6], d.values[7]
      ]);
      hist.getRange(firstDataRowHist, 1, out.length, 8).setValues(out);

      const tsOut = deleted.map(d => [d.values[8] ? new Date(d.values[8]) : '']);
      hist.getRange(firstDataRowHist, CFG.COL.HIST.TIMESTAMP, tsOut.length, 1).setValues(tsOut);
    }
  }

  ss.toast(`Undo ${label}, ok.`, 'Undo', 8);
}

/**
 * Refaz a última ação desfeita no Histórico.
 * Recupera as alterações guardadas na pilha de Redo.
 * @returns {void}
 */
function refazerUltimaAcaoHistorico() {
  const ss = SpreadsheetApp.getActive();
  let hist;
  try {
    hist = CFG.mustGetSheet(ss, CFG.SHEETS.HIST);
  } catch (e) {
    return ss.toast(String(e.message || e), 'Redo', 6);
  }

  const redo = hist_loadRedo_();
  if (redo.length === 0) return ss.toast('Nada para refazer.', 'Redo', 6);

  const txn = redo.pop();
  hist_saveRedo_(redo);

  const { headerRowHist, firstDataRowHist } = hist_findHeader_(hist);

  if (hist_isCompactTxn_(txn)) {
    const ops = hist_readCompactTxnOps_(ss, txn);
    if (!ops) return;

    // Agrupa
    const delByDate = new Map();
    const updByDate = new Map();
    const insByDate = new Map();

    for (const op of ops) {
      if (op.op === 'DEL') {
        if (!delByDate.has(op.dateKey)) delByDate.set(op.dateKey, []);
        delByDate.get(op.dateKey).push(op);
      } else if (op.op === 'UPD') {
        if (!updByDate.has(op.dateKey)) updByDate.set(op.dateKey, []);
        updByDate.get(op.dateKey).push(op);
      } else if (op.op === 'INS') {
        if (!insByDate.has(op.dateKey)) insByDate.set(op.dateKey, []);
        insByDate.get(op.dateKey).push(op);
      }
    }

    // 1) deletes (inclui Travar=true, igual comportamento antigo do redo)
    for (const [dateKey, list] of delByDate.entries()) {
      const keyToRow = hist_buildKeyRowMap_(hist, dateKey, true, headerRowHist, firstDataRowHist);
      const rows = [];
      for (const d of list) {
        const r = keyToRow.get(d.key)?.row;
        if (r) rows.push(r);
      }
      rows.sort((a, b) => b - a);
      for (const r of rows) hist.deleteRow(r);
    }

    // 2) updates
    for (const [dateKey, list] of updByDate.entries()) {
      const keyToRow = hist_buildKeyRowMap_(hist, dateKey, true, headerRowHist, firstDataRowHist);
      for (const u of list) {
        const row = keyToRow.get(u.key)?.row;
        if (!row) continue;

        if (isFinite(u.newPrice)) hist.getRange(row, CFG.COL.HIST.PRECO).setValue(u.newPrice);
        else hist.getRange(row, CFG.COL.HIST.PRECO).setValue('');

        if (u.newSavedAtMs) hist.getRange(row, CFG.COL.HIST.TIMESTAMP).setValue(new Date(u.newSavedAtMs));
        else hist.getRange(row, CFG.COL.HIST.TIMESTAMP).setValue('');
      }
    }

    // 3) inserts (evita duplicar)
    for (const [dateKey, list] of insByDate.entries()) {
      const keyToRow = hist_buildKeyRowMap_(hist, dateKey, true, headerRowHist, firstDataRowHist);

      const toInsert = [];
      for (const it of list) {
        if (!keyToRow.has(it.key) && it.rowValues) toInsert.push(it.rowValues);
      }

      if (toInsert.length > 0) {
        hist.insertRowsBefore(firstDataRowHist, toInsert.length);

        const out = toInsert.map(v => ([
          hist_dateFromKey_(v[0]),
          v[1], v[2], v[3], v[4],
          v[5], v[6], v[7]
        ]));
        hist.getRange(firstDataRowHist, 1, out.length, 8).setValues(out);

        const tsOut = toInsert.map(v => [v[8] ? new Date(v[8]) : '']);
        hist.getRange(firstDataRowHist, CFG.COL.HIST.TIMESTAMP, tsOut.length, 1).setValues(tsOut);
      }
    }

    // redo volta a ser undoável
    hist_pushUndo_(txn);
    ss.toast(`Redo ${txn.label || 'ok'}.`, 'Redo', 8);
    return;
  }

  // Legado
  const dayTxns = txn.dayTxns ? txn.dayTxns : [txn];
  const label = txn.dayTxns
    ? `${hist_formatBr_(dayTxns[0].dateKey)} a ${hist_formatBr_(dayTxns[dayTxns.length - 1].dateKey)}`
    : hist_formatBr_(txn.dateKey);

  for (const t of dayTxns) {
    const dateKey = t.dateKey;

    // apply deletes
    let keyToRow = hist_buildKeyRowMap_(hist, dateKey, true, headerRowHist, firstDataRowHist);
    const rowsToDelete = [];
    for (const d of (t.deleted || [])) {
      const v = d.values;
      const key = hist_makeKey_(v[0], v[1], v[2], v[3], v[4], v[5]);
      const row = keyToRow.get(key)?.row;
      if (row) rowsToDelete.push(row);
    }
    rowsToDelete.sort((a, b) => b - a);
    for (const r of rowsToDelete) hist.deleteRow(r);

    // rebuild map
    keyToRow = hist_buildKeyRowMap_(hist, dateKey, true, headerRowHist, firstDataRowHist);

    // apply updates
    for (const u of (t.updated || [])) {
      const row = keyToRow.get(u.key)?.row;
      if (!row) continue;

      if (isFinite(u.newPrice)) hist.getRange(row, CFG.COL.HIST.PRECO).setValue(u.newPrice);
      else hist.getRange(row, CFG.COL.HIST.PRECO).setValue('');

      if (u.newSavedAtMs) hist.getRange(row, CFG.COL.HIST.TIMESTAMP).setValue(new Date(u.newSavedAtMs));
      else hist.getRange(row, CFG.COL.HIST.TIMESTAMP).setValue('');
    }

    // apply inserts
    const inserted = t.inserted || [];
    if (inserted.length > 0) {
      keyToRow = hist_buildKeyRowMap_(hist, dateKey, true, headerRowHist, firstDataRowHist);

      const toInsert = [];
      for (const r of inserted) {
        const key = hist_makeKey_(r[0], r[1], r[2], r[3], r[4], r[5]);
        if (!keyToRow.has(key)) toInsert.push(r);
      }

      if (toInsert.length > 0) {
        hist.insertRowsBefore(firstDataRowHist, toInsert.length);

        const out = toInsert.map(r => [
          hist_dateFromKey_(r[0]),
          r[1], r[2], r[3], r[4],
          r[5], r[6], r[7]
        ]);
        hist.getRange(firstDataRowHist, 1, out.length, 8).setValues(out);

        const tsOut = toInsert.map(r => [r[8] ? new Date(r[8]) : '']);
        hist.getRange(firstDataRowHist, CFG.COL.HIST.TIMESTAMP, tsOut.length, 1).setValues(tsOut);
      }
    }
  }

  hist_pushUndo_(txn);
  ss.toast(`Redo ${label}, ok.`, 'Redo', 8);
}

/**
 * RETROATIVO:
 * Normaliza todo o Histórico (A:H e K), limpando textos, arredondando quantidades/preços,
 * e garantindo formatação correta. Ignora as linhas onde Travar=TRUE.
 * @returns {void}
 */
function normalizarHistoricoCompleto() {
  const ss = SpreadsheetApp.getActive();
  let hist;
  try {
    hist = CFG.mustGetSheet(ss, CFG.SHEETS.HIST);
  } catch (e) {
    return ss.toast(String(e.message || e), 'Normalizar', 6);
  }

  const { headerRowHist, firstDataRowHist } = hist_findHeader_(hist);
  const lr = hist.getLastRow();
  if (lr < firstDataRowHist) return ss.toast('Histórico vazio (sem dados).', 'Normalizar', 4);

  // garante cabeçalho K
  const tsHead = hist.getRange(headerRowHist, CFG.COL.HIST.TIMESTAMP).getValue();
  if (!String(tsHead || '').trim()) hist.getRange(headerRowHist, CFG.COL.HIST.TIMESTAMP).setValue('Salvo em');

  const nRows = lr - firstDataRowHist + 1;

  // Lê só A:H e K (não encosta em I:J)
  const rangeAH = hist.getRange(firstDataRowHist, 1, nRows, 8); // A:H
  const valsAH = rangeAH.getValues();

  const rangeK = hist.getRange(firstDataRowHist, CFG.COL.HIST.TIMESTAMP, nRows, 1); // K
  const valsK = rangeK.getValues();

  let changedRows = 0;
  let changedCells = 0;

  for (let i = 0; i < nRows; i++) {
    const row = valsAH[i];     // A:H
    const travar = row[CFG.COL.HIST.TRAVAR - 1] === true; // H
    if (travar) continue;

    let changed = false;

    // Índices dentro de A:H
    const idxItem  = CFG.COL.HIST.ITEM - 1;  // 1
    const idxMarca = CFG.COL.HIST.MARCA - 1; // 2
    const idxLoja  = CFG.COL.HIST.LOJA - 1;  // 3
    const idxTipo  = CFG.COL.HIST.TIPO - 1;  // 4
    const idxQtd   = CFG.COL.HIST.QTD - 1;   // 5
    const idxPreco = CFG.COL.HIST.PRECO - 1; // 6

    // Textos
    const item0  = row[idxItem];
    const marca0 = row[idxMarca];
    const loja0  = row[idxLoja];
    const tipo0  = row[idxTipo];

    const item1  = hist_cleanTextKeepCase_(item0);
    const marca1 = hist_cleanTextKeepCase_(marca0);
    const loja1  = hist_cleanTextKeepCase_(loja0);
    const tipo1  = hist_cleanTextKeepCase_(tipo0);

    if (item1 !== String(item0 ?? ''))   { row[idxItem] = item1; changed = true; changedCells++; }
    if (marca1 !== String(marca0 ?? '')) { row[idxMarca] = marca1; changed = true; changedCells++; }
    if (loja1 !== String(loja0 ?? ''))   { row[idxLoja] = loja1; changed = true; changedCells++; }
    if (tipo1 !== String(tipo0 ?? ''))   { row[idxTipo] = tipo1; changed = true; changedCells++; }

    // Quantidade
    const qtd0 = row[idxQtd];
    const qtdN = hist_parseNumber_(qtd0);

    if (isFinite(qtdN)) {
      const old = (typeof qtd0 === 'number') ? qtd0 : NaN;
      if (!(typeof qtd0 === 'number' && Math.abs(old - qtdN) <= 1e-12)) {
        row[idxQtd] = hist_roundSmart_(qtdN);
        changed = true; changedCells++;
      }
    } else if (qtd0 !== '' && qtd0 !== null) {
      const extracted = hist_extractFirstNumber_(qtd0);
      if (isFinite(extracted)) {
        row[idxQtd] = hist_roundSmart_(extracted);
        changed = true; changedCells++;
      } else {
        const qClean = hist_cleanTextKeepCase_(qtd0);
        if (qClean !== String(qtd0 ?? '')) {
          row[idxQtd] = qClean;
          changed = true; changedCells++;
        }
      }
    }

    // Preço
    const preco0 = row[idxPreco];
    const precoN = hist_parseNumber_(preco0);

    if (isFinite(precoN)) {
      const old = (typeof preco0 === 'number') ? preco0 : NaN;
      if (!(typeof preco0 === 'number' && Math.abs(old - precoN) <= 1e-12)) {
        row[idxPreco] = hist_roundMoney_(precoN);
        changed = true; changedCells++;
      }
    } else if (preco0 !== '' && preco0 !== null) {
      const extracted = hist_extractFirstNumber_(preco0);
      if (isFinite(extracted)) {
        row[idxPreco] = hist_roundMoney_(extracted);
        changed = true; changedCells++;
      } else {
        const pClean = hist_cleanTextKeepCase_(preco0);
        if (pClean !== String(preco0 ?? '')) {
          row[idxPreco] = pClean;
          changed = true; changedCells++;
        }
      }
    }

    // Timestamp K (fora de A:H)
    const ts0 = valsK[i][0];
    if (ts0 && !(ts0 instanceof Date)) {
      const s = String(ts0).trim();
      const parsed = Date.parse(s);
      if (!isNaN(parsed)) {
        valsK[i][0] = new Date(parsed);
        changed = true; changedCells++;
      } else {
        const tClean = hist_cleanTextKeepCase_(ts0);
        if (tClean !== String(ts0 ?? '')) {
          valsK[i][0] = tClean;
          changed = true; changedCells++;
        }
      }
    }

    if (changed) changedRows++;
  }

  // Escreve só A:H e K, preserva I:J (fórmulas)
  rangeAH.setValues(valsAH);
  rangeK.setValues(valsK);

  ss.toast(`Normalização concluída: ${changedRows} linhas ajustadas (${changedCells} células).`, 'Normalizar', 8);
}

// ====== CORE ======

/**
 * Função central de processamento para salvar os dados da Pesquisa para o Histórico.
 * Calcula inserções, atualizações e exclusões necessárias e salva a transação para o sistema de Undo.
 * @private
 * @param {boolean} deleteStale - Se verdadeiro (Sync), apaga registros da série ausentes no dia atual.
 * @returns {void}
 */
function salvarPrecosDoPeriodo_(deleteStale) {
  const ss = SpreadsheetApp.getActive();
  let sh;
  let hist;

  try {
    sh = CFG.mustGetSheet(ss, CFG.SHEETS.PESQ);
  } catch (e) {
    ss.toast(String(e.message || e), 'Histórico', 8);
    return;
  }

  hist = ss.getSheetByName(CFG.SHEETS.HIST) || ss.insertSheet(CFG.SHEETS.HIST);

  // garante cabeçalho do histórico
  const { headerRowHist, firstDataRowHist } = hist_findHeader_(hist);
  hist_ensureHeaders_(hist, headerRowHist);

  const lastRow = sh.getLastRow();
  if (lastRow < CFG.ROWS.FIRST_DATA) {
    ss.toast('Nada para salvar (sem dados).', 'Histórico', 6);
    return;
  }

  // intervalo alvo (B1, sempre texto display)
  const dr = hist_getPesquisaDateRangeFromB1_(sh);
  if (dr.error) {
    ss.toast(dr.error, 'Histórico', 8);
    return;
  }
  const { TZ, start, end, startKey, endKey, days, label } = dr;

  // timestamp desta execução
  const savedAtMs = Date.now();

  // detectar lojas via CFG
  let storeInfo;
  try {
    storeInfo = CFG.getStoreInfo(sh, CFG.ROWS.HEADER);
  } catch (e) {
    ss.toast(String(e.message || e), 'Histórico', 8);
    return;
  }

  const storeStartCol = storeInfo.storeStart;
  const storeEndCol = storeInfo.storeEnd;
  const storeCols = storeInfo.storeCols;

  if (storeCols <= 0) {
    ss.toast('Não encontrei colunas de lojas.', 'Histórico', 8);
    return;
  }

  const lojas = sh.getRange(CFG.ROWS.HEADER, storeStartCol, 1, storeCols)
    .getValues()[0].map(v => hist_cleanTextKeepCase_(v));

  const data = sh.getRange(
    CFG.ROWS.FIRST_DATA,
    1,
    lastRow - CFG.ROWS.FIRST_DATA + 1,
    storeEndCol
  ).getValues();

  // Monta obsBase 1 vez (sem data), depois replica no intervalo
  // baseKey = itemKey|marcaKey|lojaKey|tipoKey|qtdKey
  const obsBase = new Map();
  const obsSeries = new Set(); // itemKey|marcaKey|tipoKey|qtdKey

  for (const row of data) {
    const itemRaw  = row[CFG.COL.PESQ.ITEM - 1];
    const marcaRaw = row[CFG.COL.PESQ.MARCA - 1];
    const qtdRaw   = row[CFG.COL.PESQ.QTD - 1];
    const tipoRaw  = row[CFG.COL.PESQ.TIPO - 1];

    const item  = hist_cleanTextKeepCase_(itemRaw);
    const marca = hist_cleanTextKeepCase_(marcaRaw);
    const tipo  = hist_cleanTextKeepCase_(tipoRaw);

    const qtdN = hist_parseNumber_(qtdRaw);
    if (!item || !tipo || !isFinite(qtdN) || qtdN <= 0) continue;

    const itemKey  = hist_normKey_(item);
    const marcaKey = hist_normKey_(marca);
    const tipoKey  = hist_normKey_(tipo);
    const qtdKey   = hist_qtyKey_(qtdN);

    obsSeries.add(`${itemKey}|${marcaKey}|${tipoKey}|${qtdKey}`);

    for (let c = 0; c < lojas.length; c++) {
      const loja = lojas[c];
      const precoN = hist_parseNumber_(row[(storeStartCol - 1) + c]);
      if (!loja || !isFinite(precoN) || precoN <= 0) continue;

      const lojaKey = hist_normKey_(loja);
      const baseKey = `${itemKey}|${marcaKey}|${lojaKey}|${tipoKey}|${qtdKey}`;
      const existing = obsBase.get(baseKey);

      // repetido: fica com menor preço
      if (!existing || precoN < existing.preco) {
        obsBase.set(baseKey, {
          item,
          marca,
          loja,
          tipo,
          qtd: hist_roundSmart_(qtdN),
          preco: hist_roundMoney_(precoN)
        });
      }
    }
  }

  if (obsBase.size === 0) {
    ss.toast(`Nada salvo (${label}), sem preços válidos.`, 'Histórico', 8);
    return;
  }

  // Planejar tudo antes de aplicar
  const plans = [];
  let totalInsert = 0;
  let totalUpdate = 0;
  let totalDelete = 0;

  for (let d = new Date(start); d.getTime() <= end.getTime(); d.setDate(d.getDate() + 1)) {
    const dateKey = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
    const dateObj = hist_dateFromKey_(dateKey);

    const existingMap = hist_buildKeyRowMap_(hist, dateKey, false, headerRowHist, firstDataRowHist);

    const toInsert = [];
    const toUpdate = []; // {key, oldPrice, newPrice, oldSavedAtMs, newSavedAtMs}
    const toDelete = []; // {key, values}

    // inserts / updates
    for (const [baseKey, v] of obsBase.entries()) {
      const key = `${dateKey}|${baseKey}`;

      if (existingMap.has(key)) {
        const ex = existingMap.get(key);

        const oldPrice = hist_parseNumber_(ex.values[6]);
        const newPrice = hist_parseNumber_(v.preco);

        const oldSavedAtMs = ex.values[8] || null;
        const newSavedAtMs = savedAtMs;

        const priceChanged =
          (isFinite(oldPrice) && isFinite(newPrice) && Math.abs(oldPrice - newPrice) > 1e-9) ||
          (!isFinite(oldPrice) && isFinite(newPrice)) ||
          (isFinite(oldPrice) && !isFinite(newPrice));

        // timestamp sempre atualiza para linhas não travadas (map já exclui travadas)
        if (priceChanged || oldSavedAtMs !== newSavedAtMs) {
          toUpdate.push({ key, oldPrice, newPrice, oldSavedAtMs, newSavedAtMs });
        }
      } else {
        // values: [dateKey, item, marca, loja, tipo, qtd(number), preco(number), travar(false), savedAtMs]
        toInsert.push([dateKey, v.item, v.marca, v.loja, v.tipo, v.qtd, v.preco, false, savedAtMs]);
      }
    }

    // deletes (somente sync)
    if (deleteStale) {
      for (const [k, ex] of existingMap.entries()) {
        // k = date|itemKey|marcaKey|lojaKey|tipoKey|qtdKey
        const parts = String(k).split('|');
        if (parts.length < 6) continue;

        const itemKey  = parts[1];
        const marcaKey = parts[2];
        const lojaKey  = parts[3];
        const tipoKey  = parts[4];
        const qtdKey   = parts[5];

        const seriesKey = `${itemKey}|${marcaKey}|${tipoKey}|${qtdKey}`;
        const baseKey = `${itemKey}|${marcaKey}|${lojaKey}|${tipoKey}|${qtdKey}`;

        if (obsSeries.has(seriesKey) && !obsBase.has(baseKey)) {
          toDelete.push({ key: k, values: ex.values.slice() });
        }
      }
    }

    plans.push({ dateKey, dateObj, toInsert, toUpdate, toDelete });
    totalInsert += toInsert.length;
    totalUpdate += toUpdate.length;
    totalDelete += toDelete.length;
  }

  const totalOps = totalInsert + totalUpdate + totalDelete;
  if (totalOps === 0) {
    ss.toast(`Nada a alterar (${label}).`, 'Histórico', 6);
    return;
  }

  // Sempre limpa redo ao salvar algo novo
  hist_clearRedo_();

  // Undo compacto (aba _HIST_TXN)
  let undoStored = false;
  let txnMeta = null;

  if (totalOps <= HIST_STACK.maxOpsToStore) {
    const txnId = hist_newTxnId_();
    const mode = deleteStale ? 'sync' : 'incremental';

    // cria lista de ops compactas
    const ops = [];
    for (const p of plans) {
      // INS
      for (const r of p.toInsert) {
        const key = hist_makeKey_(r[0], r[1], r[2], r[3], r[4], r[5]);
        const rowValues = [r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8] || savedAtMs];
        ops.push({
          op: 'INS',
          dateKey: r[0],
          key,
          rowValues
        });
      }

      // UPD
      for (const u of p.toUpdate) {
        ops.push({
          op: 'UPD',
          dateKey: p.dateKey,
          key: u.key,
          oldPrice: isFinite(u.oldPrice) ? u.oldPrice : null,
          newPrice: isFinite(u.newPrice) ? hist_roundMoney_(u.newPrice) : null,
          oldSavedAtMs: u.oldSavedAtMs || null,
          newSavedAtMs: u.newSavedAtMs || null
        });
      }

      // DEL
      for (const d of p.toDelete) {
        ops.push({
          op: 'DEL',
          dateKey: p.dateKey,
          key: d.key,
          rowValues: d.values.slice()
        });
      }
    }

    const append = hist_writeCompactTxnOps_(ss, txnId, ops);
    if (append) {
      txnMeta = {
        v: 'compact1',
        id: txnId,
        ts: new Date().toISOString(),
        label,
        mode,
        startKey,
        endKey,
        days,
        counts: { ins: totalInsert, upd: totalUpdate, del: totalDelete },
        startRow: append.startRow,
        nOps: append.nOps
      };
      hist_pushUndo_(txnMeta);
      undoStored = true;

      // manutenção (compacta se a aba crescer muito)
      hist_txnMaintenance_(ss);
    }
  }

  // Aplicar mudanças por dia (mesma lógica)
  for (const p of plans) {
    const dateKey = p.dateKey;
    const dateObj = p.dateObj;

    // 1) deletes
    if (p.toDelete.length > 0) {
      const keyToRow = hist_buildKeyRowMap_(hist, dateKey, true, headerRowHist, firstDataRowHist);
      const rows = [];
      for (const d of p.toDelete) {
        const r = keyToRow.get(d.key)?.row;
        if (r) rows.push(r);
      }
      rows.sort((a, b) => b - a);
      for (const r of rows) hist.deleteRow(r);
    }

    // 2) updates (Preço + Timestamp)
    if (p.toUpdate.length > 0) {
      const keyToRow = hist_buildKeyRowMap_(hist, dateKey, true, headerRowHist, firstDataRowHist);
      for (const u of p.toUpdate) {
        const r = keyToRow.get(u.key)?.row;
        if (!r) continue;

        // Preço (G): só muda se mudou
        if (isFinite(u.oldPrice) && isFinite(u.newPrice) && Math.abs(u.oldPrice - u.newPrice) > 1e-9) {
          hist.getRange(r, CFG.COL.HIST.PRECO).setValue(hist_roundMoney_(u.newPrice));
        } else if (!isFinite(u.oldPrice) && isFinite(u.newPrice)) {
          hist.getRange(r, CFG.COL.HIST.PRECO).setValue(hist_roundMoney_(u.newPrice));
        } else if (isFinite(u.oldPrice) && !isFinite(u.newPrice)) {
          hist.getRange(r, CFG.COL.HIST.PRECO).setValue('');
        }

        // Timestamp (K) sempre
        hist.getRange(r, CFG.COL.HIST.TIMESTAMP).setValue(new Date(u.newSavedAtMs));
      }
    }

    // 3) inserts
    if (p.toInsert.length > 0) {
      const { firstDataRowHist: fd } = hist_findHeader_(hist);
      hist.insertRowsBefore(fd, p.toInsert.length);

      // A:H
      const out = p.toInsert.map(r => [
        dateObj, r[1], r[2], r[3], r[4], r[5], r[6], r[7]
      ]);
      hist.getRange(fd, 1, out.length, 8).setValues(out);

      // K
      const tsCol = p.toInsert.map(r => [r[8] ? new Date(r[8]) : new Date(savedAtMs)]);
      hist.getRange(fd, CFG.COL.HIST.TIMESTAMP, tsCol.length, 1).setValues(tsCol);
    }
  }

  const msg =
    `Período ${label}: +${totalInsert} | ~${totalUpdate} | -${totalDelete}` +
    (deleteStale ? ' | modo: sync' : ' | modo: incremental') +
    (undoStored ? ' | Undo OK' : ' | Undo indisponível');

  ss.toast(msg, 'Histórico', 10);
}

// ====== UNDO/REDO: compact helpers ======

/**
 * Valida se um objeto de transação possui o formato compacto esperado.
 * @private
 * @param {Object} txn - O objeto transacional.
 * @returns {boolean}
 */
function hist_isCompactTxn_(txn) {
  return !!txn && txn.v === 'compact1' && !!txn.id && !!txn.startRow && !!txn.nOps;
}

/**
 * Obtém (ou cria se não existir) a aba oculta responsável por armazenar os deltas do histórico.
 * @private
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss 
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function hist_getTxnSheet_(ss) {
  let sh = ss.getSheetByName(HIST_STACK.txnSheetName);
  if (!sh) {
    sh = ss.insertSheet(HIST_STACK.txnSheetName);
    sh.hideSheet();
    // header
    sh.getRange(1, 1, 1, HIST_STACK.txnHeaderCols).setValues([[
      'TxnId', 'Op', 'DateKey', 'Key',
      'OldPrice', 'NewPrice', 'OldTsMs', 'NewTsMs',
      'RowJson', 'Note'
    ]]);
  } else {
    try { sh.hideSheet(); } catch (_) {}
    const head = sh.getRange(1, 1, 1, HIST_STACK.txnHeaderCols).getValues()[0];
    if (!String(head[0] || '').trim()) {
      sh.getRange(1, 1, 1, HIST_STACK.txnHeaderCols).setValues([[
        'TxnId', 'Op', 'DateKey', 'Key',
        'OldPrice', 'NewPrice', 'OldTsMs', 'NewTsMs',
        'RowJson', 'Note'
      ]]);
    }
  }
  return sh;
}

/**
 * Grava na aba de transações as operações realizadas para posterior reversão (undo).
 * @private
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss 
 * @param {string} txnId 
 * @param {Array} ops 
 * @returns {Object|null}
 */
function hist_writeCompactTxnOps_(ss, txnId, ops) {
  try {
    const sh = hist_getTxnSheet_(ss);
    const startRow = Math.max(2, (sh.getLastRow() || 1) + 1);

    const rows = ops.map(op => ([
      txnId,
      op.op || '',
      op.dateKey || '',
      op.key || '',
      (op.oldPrice === null || op.oldPrice === undefined) ? '' : op.oldPrice,
      (op.newPrice === null || op.newPrice === undefined) ? '' : op.newPrice,
      (op.oldSavedAtMs === null || op.oldSavedAtMs === undefined) ? '' : op.oldSavedAtMs,
      (op.newSavedAtMs === null || op.newSavedAtMs === undefined) ? '' : op.newSavedAtMs,
      op.rowValues ? JSON.stringify(op.rowValues) : '',
      ''
    ]));

    sh.getRange(startRow, 1, rows.length, HIST_STACK.txnHeaderCols).setValues(rows);
    return { startRow, nOps: rows.length };
  } catch (e) {
    SpreadsheetApp.getActive().toast('Falha ao armazenar Undo (txn sheet).', 'Histórico', 8);
    return null;
  }
}

/**
 * Lê os dados de uma transação específica armazenada na aba oculta para reverter/refazer.
 * @private
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss 
 * @param {Object} txnMeta 
 * @returns {Array|null}
 */
function hist_readCompactTxnOps_(ss, txnMeta) {
  try {
    const sh = hist_getTxnSheet_(ss);
    const start = txnMeta.startRow;
    const n = txnMeta.nOps;

    if (!start || !n) {
      SpreadsheetApp.getActive().toast('Undo/Redo: transação inválida.', 'Histórico', 8);
      return null;
    }

    const lr = sh.getLastRow();
    if (lr < start + n - 1) {
      SpreadsheetApp.getActive().toast('Undo/Redo: dados da transação não encontrados.', 'Histórico', 8);
      return null;
    }

    const vals = sh.getRange(start, 1, n, HIST_STACK.txnHeaderCols).getValues();

    const ops = [];
    for (const r of vals) {
      const txnId = String(r[0] || '').trim();
      if (txnId !== txnMeta.id) continue;

      const op = String(r[1] || '').trim();
      const dateKey = String(r[2] || '').trim();
      const key = String(r[3] || '').trim();

      const oldPrice = (r[4] === '' || r[4] === null) ? null : hist_parseNumber_(r[4]);
      const newPrice = (r[5] === '' || r[5] === null) ? null : hist_parseNumber_(r[5]);

      const oldSavedAtMs = (r[6] === '' || r[6] === null) ? null : Number(r[6]);
      const newSavedAtMs = (r[7] === '' || r[7] === null) ? null : Number(r[7]);

      let rowValues = null;
      const rowJson = String(r[8] || '').trim();
      if (rowJson) {
        try { rowValues = JSON.parse(rowJson); } catch (_) { rowValues = null; }
      }

      ops.push({ op, dateKey, key, oldPrice, newPrice, oldSavedAtMs, newSavedAtMs, rowValues });
    }

    return ops;
  } catch (e) {
    SpreadsheetApp.getActive().toast('Undo/Redo: erro ao ler transação.', 'Histórico', 8);
    return null;
  }
}

/**
 * Faz a manutenção (limpeza) da aba oculta de transações, mantendo apenas 
 * os dados das transações mais recentes da pilha de Undo/Redo.
 * @private
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss 
 */
function hist_txnMaintenance_(ss) {
  const sh = ss.getSheetByName(HIST_STACK.txnSheetName);
  if (!sh) return;

  const lr = sh.getLastRow() || 1;
  if (lr <= HIST_STACK.txnRebuildThresholdRows) return;

  // mantém só ids presentes nos stacks
  const undo = hist_loadUndo_();
  const redo = hist_loadRedo_();

  const keepIds = new Set();
  for (const t of undo) if (hist_isCompactTxn_(t)) keepIds.add(t.id);
  for (const t of redo) if (hist_isCompactTxn_(t)) keepIds.add(t.id);

  if (keepIds.size === 0) return;

  // lê tudo e filtra
  const vals = sh.getRange(2, 1, lr - 1, HIST_STACK.txnHeaderCols).getValues();
  const kept = [];
  for (const r of vals) {
    const id = String(r[0] || '').trim();
    if (keepIds.has(id)) kept.push(r);
  }

  // reescreve compactado
  sh.clearContents();
  sh.getRange(1, 1, 1, HIST_STACK.txnHeaderCols).setValues([[
    'TxnId', 'Op', 'DateKey', 'Key',
    'OldPrice', 'NewPrice', 'OldTsMs', 'NewTsMs',
    'RowJson', 'Note'
  ]]);

  if (kept.length > 0) {
    sh.getRange(2, 1, kept.length, HIST_STACK.txnHeaderCols).setValues(kept);
  }

  // recalcula startRow e nOps por id
  const map = new Map(); // id -> {startRow, nOps}
  for (let i = 0; i < kept.length; i++) {
    const id = String(kept[i][0] || '').trim();
    const rowNum = 2 + i;
    if (!map.has(id)) map.set(id, { startRow: rowNum, nOps: 1 });
    else map.get(id).nOps++;
  }

  // atualiza pointers nos stacks
  for (const arr of [undo, redo]) {
    for (const t of arr) {
      if (!hist_isCompactTxn_(t)) continue;
      const p = map.get(t.id);
      if (p) { t.startRow = p.startRow; t.nOps = p.nOps; }
    }
  }

  hist_saveUndo_(undo);
  hist_saveRedo_(redo);
}

// ====== HELPERS ======

/**
 * Cria um ID aleatório para cada transação salva.
 * @private
 * @returns {string}
 */
function hist_newTxnId_() {
  const r = Math.floor(Math.random() * 1e9);
  return `TXN_${Date.now()}_${r}`;
}

/**
 * Encontra dinamicamente a linha de cabeçalho buscando a coluna 'data'.
 * @private
 * @param {GoogleAppsScript.Spreadsheet.Sheet} histSheet 
 * @returns {Object}
 */
function hist_findHeader_(histSheet) {
  const maxScan = Math.min(15, histSheet.getLastRow() || 15);
  if (maxScan < 1) return { headerRowHist: CFG.ROWS.HEADER, firstDataRowHist: CFG.ROWS.FIRST_DATA };

  const colA = histSheet.getRange(1, 1, maxScan, 1).getValues()
    .map(r => String(r[0] || '').trim().toLowerCase());

  let headerRow = null;
  for (let i = 0; i < colA.length; i++) {
    if (colA[i] === 'data') { headerRow = i + 1; break; }
  }
  if (!headerRow) headerRow = CFG.ROWS.HEADER;
  return { headerRowHist: headerRow, firstDataRowHist: headerRow + 1 };
}

/**
 * Cria ou restaura o cabeçalho base do histórico caso esteja corrompido.
 * @private
 * @param {GoogleAppsScript.Spreadsheet.Sheet} histSheet 
 * @param {number} headerRowHist 
 */
function hist_ensureHeaders_(histSheet, headerRowHist) {
  const headA = String(histSheet.getRange(headerRowHist, CFG.COL.HIST.DATA).getValue() || '').trim().toLowerCase();
  if (headA !== 'data') {
    histSheet.getRange(headerRowHist, 1, 1, 8)
      .setValues([['Data','Item','Marca','Loja','Tipo','Quantidade','Preço','Travar']]);
  } else {
    const travarCell = histSheet.getRange(headerRowHist, CFG.COL.HIST.TRAVAR).getValue();
    if (!String(travarCell || '').trim()) histSheet.getRange(headerRowHist, CFG.COL.HIST.TRAVAR).setValue('Travar');
  }

  const tsHead = histSheet.getRange(headerRowHist, CFG.COL.HIST.TIMESTAMP).getValue();
  if (!String(tsHead || '').trim()) histSheet.getRange(headerRowHist, CFG.COL.HIST.TIMESTAMP).setValue('Salvo em');
}

/**
 * Limpa espaços invisíveis mantendo as letras maiúsculas/minúsculas.
 * @private
 * @param {any} v 
 * @returns {string}
 */
// Texto: limpa espaços invisíveis/extra, sem mexer em maiúsculas/minúsculas
function hist_cleanTextKeepCase_(v) {
  if (v === null || v === undefined) return '';
  let s = String(v);

  s = s
    .replace(/\u00A0/g, ' ')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\s+/g, ' ')
    .trim();

  return s;
}

/**
 * Normaliza os textos de forma estrita para criação de chave (lowercase, sem acento).
 * @private
 * @param {any} v 
 * @returns {string}
 */
// Normalização para chave: lower + remove acentos + trim
function hist_normKey_(v) {
  const s = hist_cleanTextKeepCase_(v);
  return s
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}

/**
 * Converte strings contendo formatos financeiros para números flutuantes.
 * @private
 * @param {any} v 
 * @returns {number}
 */
// Parse numérico (aceita 0,4 e 0.4; aceita milhar 1.234,56)
function hist_parseNumber_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : NaN;
  if (v === null || v === undefined) return NaN;

  let s = String(v).trim();
  if (!s) return NaN;

  const m = s.match(/-?\d[\d.,]*/);
  if (!m) return NaN;
  s = m[0];

  const hasDot = s.includes('.');
  const hasComma = s.includes(',');

  if (hasDot && hasComma) {
    const lastDot = s.lastIndexOf('.');
    const lastComma = s.lastIndexOf(',');
    if (lastComma > lastDot) {
      s = s.replace(/\./g, '').replace(',', '.');
    } else {
      s = s.replace(/,/g, '');
    }
  } else if (hasComma && !hasDot) {
    s = s.replace(',', '.');
  } else if (hasDot && !hasComma) {
    if (/^\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');
  }

  const n = Number(s);
  return isFinite(n) ? n : NaN;
}

/**
 * Extrai e converte apenas a parte numérica de uma string.
 * @private
 * @param {any} v 
 * @returns {number}
 */
function hist_extractFirstNumber_(v) {
  return hist_parseNumber_(v);
}

/**
 * Arredondamento inteligente para remover ruídos de pontos flutuantes.
 * @private
 * @param {number} n 
 * @returns {number}
 */
function hist_roundSmart_(n) {
  if (!isFinite(n)) return n;
  const r = Math.round(n * 1e9) / 1e9;
  if (Math.abs(r - Math.round(r)) <= 1e-12) return Math.round(r);
  return r;
}

/**
 * Arredondamento para conversão financeira (2 casas).
 * @private
 * @param {number} n 
 * @returns {number}
 */
function hist_roundMoney_(n) {
  if (!isFinite(n)) return n;
  return Math.round(n * 100) / 100;
}

/**
 * Chave canônica da quantidade convertida para string estável.
 * @private
 * @param {any} qtdAny 
 * @returns {string}
 */
// chave canônica da quantidade (string estável)
function hist_qtyKey_(qtdAny) {
  const n = (typeof qtdAny === 'number') ? qtdAny : hist_parseNumber_(qtdAny);
  if (!isFinite(n)) return String(hist_cleanTextKeepCase_(qtdAny));
  const r = hist_roundSmart_(n);
  return String(r);
}

/**
 * Transforma uma dateKey string num objeto Data.
 * @private
 * @param {string} dateKey 
 * @returns {Date}
 */
function hist_dateFromKey_(dateKey) {
  const [y, m, d] = String(dateKey).split('-').map(Number);
  return new Date(y, (m - 1), d);
}

/**
 * Formata dateKey para string BR (DD/MM/YYYY).
 * @private
 * @param {string} dateKey 
 * @returns {string}
 */
function hist_formatBr_(dateKey) {
  const [y, m, d] = String(dateKey).split('-');
  return `${d}/${m}/${y}`;
}

/**
 * Monta a chave concatenada para indexação da linha no mapa de pesquisa do Histórico.
 * @private
 * @param {string} dateKey 
 * @param {string} item 
 * @param {string} marca 
 * @param {string} loja 
 * @param {string} tipo 
 * @param {any} qtd 
 * @returns {string}
 */
// key = date|itemKey|marcaKey|lojaKey|tipoKey|qtdKey
function hist_makeKey_(dateKey, item, marca, loja, tipo, qtd) {
  const dk = String(dateKey).trim();
  return `${dk}|${hist_normKey_(item)}|${hist_normKey_(marca)}|${hist_normKey_(loja)}|${hist_normKey_(tipo)}|${hist_qtyKey_(qtd)}`;
}

/**
 * Mapeia as linhas do histórico gerando chaves únicas para comparações com a aba pesquisa.
 * @private
 * @param {GoogleAppsScript.Spreadsheet.Sheet} histSheet 
 * @param {string} dateKey 
 * @param {boolean} includeTravar 
 * @param {number} headerRowHistOpt 
 * @param {number} firstDataRowHistOpt 
 * @returns {Map<string, Object>}
 */
// values = [dateKey, item, marca, loja, tipo, qtd(number), preco(number), travar, savedAtMs]
function hist_buildKeyRowMap_(histSheet, dateKey, includeTravar, headerRowHistOpt, firstDataRowHistOpt) {
  const TZ = Session.getScriptTimeZone();
  const lr = histSheet.getLastRow();
  const map = new Map();
  if (lr < 2) return map;

  const { headerRowHist, firstDataRowHist } = (headerRowHistOpt && firstDataRowHistOpt)
    ? { headerRowHist: headerRowHistOpt, firstDataRowHist: firstDataRowHistOpt }
    : hist_findHeader_(histSheet);

  if (lr < firstDataRowHist) return map;

  const nRows = lr - firstDataRowHist + 1;
  const vals = histSheet.getRange(firstDataRowHist, 1, nRows, CFG.COL.HIST.TIMESTAMP).getValues(); // A:K

  for (let i = 0; i < vals.length; i++) {
    const d      = vals[i][CFG.COL.HIST.DATA - 1];
    const item   = vals[i][CFG.COL.HIST.ITEM - 1];
    const marca  = vals[i][CFG.COL.HIST.MARCA - 1];
    const loja   = vals[i][CFG.COL.HIST.LOJA - 1];
    const tipo   = vals[i][CFG.COL.HIST.TIPO - 1];
    const qtd    = vals[i][CFG.COL.HIST.QTD - 1];
    const preco  = vals[i][CFG.COL.HIST.PRECO - 1];
    const travar = vals[i][CFG.COL.HIST.TRAVAR - 1];
    const ts     = vals[i][CFG.COL.HIST.TIMESTAMP - 1];

    if (!includeTravar && travar === true) continue;
    if (!d || !item || !loja || !tipo || qtd === '' || qtd === null) continue;

    const dk = Utilities.formatDate(new Date(d), TZ, 'yyyy-MM-dd');
    if (dk !== dateKey) continue;

    let tsMs = null;
    if (ts instanceof Date) tsMs = ts.getTime();
    else if (typeof ts === 'number' && isFinite(ts)) tsMs = ts;
    else {
      const s = String(ts || '').trim();
      if (s) {
        const parsed = Date.parse(s);
        if (!isNaN(parsed)) tsMs = parsed;
      }
    }

    const key = hist_makeKey_(dk, item, marca, loja, tipo, qtd);
    map.set(key, {
      row: firstDataRowHist + i,
      values: [
        dk,
        hist_cleanTextKeepCase_(item),
        hist_cleanTextKeepCase_(marca || ''),
        hist_cleanTextKeepCase_(loja),
        hist_cleanTextKeepCase_(tipo),
        (typeof qtd === 'number') ? qtd : hist_parseNumber_(qtd),
        hist_parseNumber_(preco),
        !!travar,
        tsMs
      ]
    });
  }

  return map;
}

// ====== DATA RANGE (B1 display) ======

/**
 * Retorna uma data zerando a parte de horas, minutos e segundos.
 * @private
 * @param {Date} d 
 * @returns {Date}
 */
function hist_dateOnly_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

/**
 * Processa a conversão de uma string completa (DD/MM/YYYY) para Data.
 * @private
 * @param {string} s 
 * @returns {Date|null}
 */
function hist_parseBrDateFull_(s) {
  const m = String(s || '').trim().match(/^(\d{1,2})\s*[\/.-]\s*(\d{1,2})\s*[\/.-]\s*(\d{2,4})$/);
  if (!m) return null;

  const dd = Number(m[1]);
  const mm = Number(m[2]);
  let yy = Number(m[3]);
  if (yy < 100) yy += 2000;

  const d = new Date(yy, mm - 1, dd);
  if (d.getFullYear() !== yy || d.getMonth() !== (mm - 1) || d.getDate() !== dd) return null;
  return d;
}

/**
 * Processa a conversão de uma string parcial (DD/MM) utilizando um ano fallback.
 * @private
 * @param {string} s 
 * @param {number} yearFallback 
 * @returns {Date|null}
 */
function hist_parseBrDateNoYear_(s, yearFallback) {
  const m = String(s || '').trim().match(/^(\d{1,2})\s*[\/.-]\s*(\d{1,2})$/);
  if (!m) return null;

  const dd = Number(m[1]);
  const mm = Number(m[2]);
  const yy = Number(yearFallback);

  const d = new Date(yy, mm - 1, dd);
  if (d.getFullYear() !== yy || d.getMonth() !== (mm - 1) || d.getDate() !== dd) return null;
  return d;
}

/**
 * Lê o campo B1 da pesquisa como texto e interpreta as possíveis variações de data
 * avulsa ou intervalo de dias usando Regex.
 * @private
 * @param {GoogleAppsScript.Spreadsheet.Sheet} shPesquisa 
 * @returns {Object}
 */
function hist_getPesquisaDateRangeFromB1_(shPesquisa) {
  const TZ = Session.getScriptTimeZone();
  const raw = shPesquisa.getRange(CFG.CELLS.PESQ.DATA_PESQUISA).getDisplayValue();
  let s = hist_cleanTextKeepCase_(raw).toLowerCase();

  if (!s) {
    const now = hist_dateOnly_(new Date());
    const startKey = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
    return {
      TZ,
      start: now,
      end: now,
      startKey,
      endKey: startKey,
      days: 1,
      label: hist_formatBr_(startKey)
    };
  }

  // normaliza conectores e traços
  s = s
    .replace(/[–—]/g, '-')
    .replace(/\s+(até|ate|a)\s+/g, ' - ')
    .replace(/\s+/g, ' ')
    .trim();

  // 1) atalho: 18-25/02/2026
  {
    const m = s.match(/^(\d{1,2})\s*-\s*(\d{1,2})\s*[\/.-]\s*(\d{1,2})\s*[\/.-]\s*(\d{2,4})$/);
    if (m) {
      const d1 = Number(m[1]);
      const d2 = Number(m[2]);
      const mm = Number(m[3]);
      let yy = Number(m[4]);
      if (yy < 100) yy += 2000;

      const start = new Date(yy, mm - 1, d1);
      const end = new Date(yy, mm - 1, d2);

      if (start.getFullYear() !== yy || start.getMonth() !== mm - 1 || start.getDate() !== d1) {
        return { error: 'B1: data inicial inválida.' };
      }
      if (end.getFullYear() !== yy || end.getMonth() !== mm - 1 || end.getDate() !== d2) {
        return { error: 'B1: data final inválida.' };
      }

      const st = hist_dateOnly_(start);
      const en = hist_dateOnly_(end);
      if (en.getTime() < st.getTime()) return { error: 'B1: data final menor que a inicial.' };

      const startKey = Utilities.formatDate(st, TZ, 'yyyy-MM-dd');
      const endKey = Utilities.formatDate(en, TZ, 'yyyy-MM-dd');
      const days = Math.floor((en.getTime() - st.getTime()) / 86400000) + 1;

      return {
        TZ,
        start: st,
        end: en,
        startKey,
        endKey,
        days,
        label: `${hist_formatBr_(startKey)} a ${hist_formatBr_(endKey)}`
      };
    }
  }

  // 2) duas datas completas no texto
  const fullRe = /(\d{1,2}\s*[\/.-]\s*\d{1,2}\s*[\/.-]\s*\d{2,4})/g;
  const full = [];
  for (const m of s.matchAll(fullRe)) full.push(m[1]);
  if (full.length >= 2) {
    const start0 = hist_parseBrDateFull_(full[0].replace(/\s*/g, ''));
    const end0 = hist_parseBrDateFull_(full[1].replace(/\s*/g, ''));
    if (!start0 || !end0) return { error: 'B1: não consegui interpretar as datas.' };

    const st = hist_dateOnly_(start0);
    const en = hist_dateOnly_(end0);
    if (en.getTime() < st.getTime()) return { error: 'B1: data final menor que a inicial.' };

    const startKey = Utilities.formatDate(st, TZ, 'yyyy-MM-dd');
    const endKey = Utilities.formatDate(en, TZ, 'yyyy-MM-dd');
    const days = Math.floor((en.getTime() - st.getTime()) / 86400000) + 1;

    return {
      TZ,
      start: st,
      end: en,
      startKey,
      endKey,
      days,
      label: (days === 1) ? hist_formatBr_(startKey) : `${hist_formatBr_(startKey)} a ${hist_formatBr_(endKey)}`
    };
  }

  // 3) uma data completa só
  if (full.length === 1) {
    const end0 = hist_parseBrDateFull_(full[0].replace(/\s*/g, ''));
    if (!end0) return { error: 'B1: data inválida.' };

    // tenta achar um dd/mm antes para formar intervalo do tipo "18/02 - 25/02/2026"
    const year = end0.getFullYear();
    const dmRe = /(\d{1,2}\s*[\/.-]\s*\d{1,2})(?!\s*[\/.-]\s*\d{2,4})/g;
    const dm = [];
    for (const m of s.matchAll(dmRe)) dm.push(m[1]);

    let start0 = null;
    if (dm.length >= 1) {
      // pega o primeiro dd/mm encontrado
      start0 = hist_parseBrDateNoYear_(dm[0].replace(/\s*/g, ''), year);
    }

    const st = hist_dateOnly_(start0 || end0);
    const en = hist_dateOnly_(end0);

    if (en.getTime() < st.getTime()) return { error: 'B1: data final menor que a inicial.' };

    const startKey = Utilities.formatDate(st, TZ, 'yyyy-MM-dd');
    const endKey = Utilities.formatDate(en, TZ, 'yyyy-MM-dd');
    const days = Math.floor((en.getTime() - st.getTime()) / 86400000) + 1;

    return {
      TZ,
      start: st,
      end: en,
      startKey,
      endKey,
      days,
      label: (days === 1) ? hist_formatBr_(startKey) : `${hist_formatBr_(startKey)} a ${hist_formatBr_(endKey)}`
    };
  }

  // 4) nenhuma data completa: tenta dd/mm/aaaa simples com separadores diferentes
  const single = hist_parseBrDateFull_(s.replace(/\s*/g, ''));
  if (single) {
    const st = hist_dateOnly_(single);
    const startKey = Utilities.formatDate(st, TZ, 'yyyy-MM-dd');
    return {
      TZ,
      start: st,
      end: st,
      startKey,
      endKey: startKey,
      days: 1,
      label: hist_formatBr_(startKey)
    };
  }

  // 5) dd/mm sem ano: assume ano atual
  const dm = hist_parseBrDateNoYear_(s.replace(/\s*/g, ''), (new Date()).getFullYear());
  if (dm) {
    const st = hist_dateOnly_(dm);
    const startKey = Utilities.formatDate(st, TZ, 'yyyy-MM-dd');
    return {
      TZ,
      start: st,
      end: st,
      startKey,
      endKey: startKey,
      days: 1,
      label: hist_formatBr_(startKey)
    };
  }

  return { error: 'B1: formato de data/período não reconhecido.' };
}

// ====== STACK (DocumentProperties) ======

/** @private */
function hist_loadUndo_() { return hist_loadStack_(HIST_STACK.undoKey); }
/** @private */
function hist_saveUndo_(v) { hist_saveStack_(HIST_STACK.undoKey, v); }
/** @private */
function hist_loadRedo_() { return hist_loadStack_(HIST_STACK.redoKey); }
/** @private */
function hist_saveRedo_(v) { hist_saveStack_(HIST_STACK.redoKey, v); }
/** @private */
function hist_clearRedo_() { hist_saveRedo_([]); }

/**
 * Adiciona uma transação à pilha de Undo, descartando as mais velhas.
 * @private
 * @param {Object} txn 
 */
function hist_pushUndo_(txn) {
  const s = hist_loadUndo_();
  s.push(txn);
  while (s.length > HIST_STACK.keepLast) s.shift();
  hist_saveUndo_(s);
}

/**
 * Adiciona uma transação à pilha de Redo, descartando as mais velhas.
 * @private
 * @param {Object} txn 
 */
function hist_pushRedo_(txn) {
  const s = hist_loadRedo_();
  s.push(txn);
  while (s.length > HIST_STACK.keepLast) s.shift();
  hist_saveRedo_(s);
}

/**
 * Puxa os dados persistidos das propriedades do documento baseado numa Key.
 * @private
 * @param {string} key 
 * @returns {Array}
 */
function hist_loadStack_(key) {
  const p = PropertiesService.getDocumentProperties();
  const raw = p.getProperty(key);
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}

/**
 * Salva a pilha de arrays JSONizada nas propriedades do documento.
 * @private
 * @param {string} key 
 * @param {Array} arr 
 */
function hist_saveStack_(key, arr) {
  const p = PropertiesService.getDocumentProperties();
  p.setProperty(key, JSON.stringify(arr || []));
}

/**
 * Carregar Histórico.js — unificado no Histórico.gs (v4)
 * - Lê B1 usando o MESMO parser do Histórico v5 (hist_getPesquisaDateRangeFromB1_)
 * - Se B1 for período, carrega a DATA INICIAL do período
 * - NÃO mexe na coluna E (Tipo)
 * - Limpa e escreve apenas:
 * - A:D (Check, Item, Marca, Qtd)
 * - F..storeEnd (lojas)
 * - Lê só A:G do Histórico e compara data por chave numérica
 * @returns {void}
 */
function carregarPesquisaDoHistorico() {
  const ss = SpreadsheetApp.getActive();
  const shP = ss.getSheetByName(CFG.SHEETS.PESQ);
  const shH = ss.getSheetByName(CFG.SHEETS.HIST);

  if (!shP) return ss.toast(`Não encontrei a aba "${CFG.SHEETS.PESQ}".`, 'Carregar histórico', 8);
  if (!shH) return ss.toast(`Não encontrei a aba "${CFG.SHEETS.HIST}".`, 'Carregar histórico', 8);

  // Data alvo (unificado com o parser do Histórico v5: B1 sempre texto/display)
  let dateObj, dateKey, dateKeyNum;
  let periodoInfo = null;

  if (typeof hist_getPesquisaDateRangeFromB1_ === 'function') {
    const dr = hist_getPesquisaDateRangeFromB1_(shP);
    if (dr.error) {
      return ss.toast(dr.error, 'Carregar histórico', 8);
    }

    // Carregar usa a DATA INICIAL do período
    dateObj = dr.start;
    dateKey = dr.startKey;
    dateKeyNum = LOAD2_dateToKeyNum_(dr.startKey);
    periodoInfo = dr;
  } else {
    // fallback, caso a função do histórico não exista por algum motivo
    const TZ = Session.getScriptTimeZone();
    const rawDate = shP.getRange(CFG.CELLS.PESQ.DATA_PESQUISA).getValue();
    dateObj = (rawDate instanceof Date && !isNaN(rawDate.getTime())) ? rawDate : new Date();
    dateKey = Utilities.formatDate(dateObj, TZ, 'yyyy-MM-dd');
    dateKeyNum = LOAD2_dateToKeyNum_(dateObj);
  }

  // Detecta lojas (F.. antes de Custo/Mult)
  let store;
  try {
    store = CFG.getStoreInfo(shP, CFG.ROWS.HEADER);
  } catch (err) {
    return ss.toast(`Erro detectando lojas: ${err.message}`, 'Carregar histórico', 10);
  }

  const lojas = shP.getRange(CFG.ROWS.HEADER, store.storeStart, 1, store.storeCols)
    .getDisplayValues()[0]
    .map(v => String(v || '').trim());

  // loja(normalizada) -> índice
  const lojaToIdx = new Map();
  for (let i = 0; i < lojas.length; i++) {
    const lk = LOAD2_norm_(lojas[i]);
    if (lk) lojaToIdx.set(lk, i);
  }

  // Lê Histórico A:G
  const firstDataRowHist = CFG.ROWS.FIRST_DATA;
  const lastRowH = shH.getLastRow();
  if (lastRowH < firstDataRowHist) {
    ss.toast('Histórico vazio.', 'Carregar histórico', 6);
    return;
  }

  const nRowsH = lastRowH - firstDataRowHist + 1;
  const valsH = shH.getRange(firstDataRowHist, 1, nRowsH, 7).getValues();

  // Agrupa por série (item+marca+tipo+qtd) e distribui preços por loja
  const seriesMap = new Map();

  for (let i = 0; i < valsH.length; i++) {
    const d = valsH[i][0];
    const item = String(valsH[i][1] || '').trim();
    const marca = String(valsH[i][2] || '').trim();
    const loja = String(valsH[i][3] || '').trim();
    const tipo = String(valsH[i][4] || '').trim();
    const qtd = valsH[i][5];
    const preco = valsH[i][6];

    if (!d || !item || !loja || !tipo || qtd === '' || qtd === null) continue;
    if (LOAD2_dateToKeyNum_(d) !== dateKeyNum) continue;

    const lojaIdx = lojaToIdx.get(LOAD2_norm_(loja));
    if (lojaIdx === undefined) continue;

    const p = CFG.toNum(preco);
    if (!isFinite(p) || p <= 0) continue;

    const sk = `${LOAD2_norm_(item)}|${LOAD2_norm_(marca)}|${LOAD2_norm_(tipo)}|${String(qtd).trim()}`;

    if (!seriesMap.has(sk)) {
      seriesMap.set(sk, { item, marca, qtd, lojaPrices: new Map() });
    }

    const obj = seriesMap.get(sk);
    if (!obj.lojaPrices.has(lojaIdx)) obj.lojaPrices.set(lojaIdx, []);
    obj.lojaPrices.get(lojaIdx).push(p);
  }

  if (seriesMap.size === 0) {
    ss.toast(`Nada encontrado para ${LOAD2_formatBr_(dateKey)}.`, 'Carregar histórico', 8);
    return;
  }

  // Monta saída em 2 blocos (pra não tocar a coluna E):
  // meta: A:D
  // preços: F..storeEnd
  const metaRows = [];
  const priceRows = [];
  const notesRows = [];
  const dupFlags = [];

  // Ordena: Item (A-Z) -> Quantidade (crescente) -> Marca (A-Z) -> Tipo (A-Z)
const seriesList = Array.from(seriesMap.values()).sort((a, b) => {
  const itemCmp = String(a.item || '').localeCompare(String(b.item || ''), 'pt-BR', { sensitivity: 'base' });
  if (itemCmp !== 0) return itemCmp;

  // Quantidade (numérica quando possível)
  const aq = (typeof hist_parseNumber_ === 'function') ? hist_parseNumber_(a.qtd) : CFG.toNum(a.qtd);
  const bq = (typeof hist_parseNumber_ === 'function') ? hist_parseNumber_(b.qtd) : CFG.toNum(b.qtd);

  const aHasNum = isFinite(aq);
  const bHasNum = isFinite(bq);

  if (aHasNum && bHasNum) {
    if (Math.abs(aq - bq) > 1e-12) return aq - bq; // crescente
  } else if (aHasNum && !bHasNum) {
    return -1;
  } else if (!aHasNum && bHasNum) {
    return 1;
  } else {
    const qtdCmp = String(a.qtd || '').localeCompare(String(b.qtd || ''), 'pt-BR', { sensitivity: 'base' });
    if (qtdCmp !== 0) return qtdCmp;
  }

  const marcaCmp = String(a.marca || '').localeCompare(String(b.marca || ''), 'pt-BR', { sensitivity: 'base' });
  if (marcaCmp !== 0) return marcaCmp;

  const tipoCmp = String(a.tipo || '').localeCompare(String(b.tipo || ''), 'pt-BR', { sensitivity: 'base' });
  if (tipoCmp !== 0) return tipoCmp;

  return 0;
});

for (const obj of seriesList) {
  let maxDup = 1;
  for (const arr of obj.lojaPrices.values()) maxDup = Math.max(maxDup, arr.length);

  for (let dup = 0; dup < maxDup; dup++) {
    metaRows.push([false, obj.item, obj.marca, obj.qtd]);

    const pr = new Array(store.storeCols).fill('');
    let hasDuplicate = false;

    for (const [lojaIdx, arr] of obj.lojaPrices.entries()) {
      if (arr.length > 1) hasDuplicate = true;
      pr[lojaIdx] = (dup < arr.length) ? arr[dup] : '';
    }

    priceRows.push(pr);
    notesRows.push(hasDuplicate ? `Duplicado no histórico (${LOAD2_formatBr_(dateKey)}) | revise loja(s)` : '');
    dupFlags.push(hasDuplicate);
  }
}
  // Limpa área editável SEM encostar na coluna E
  LOAD2_clearPesquisaArea_(shP, store);

  // Garante linhas suficientes
  const needRows = metaRows.length;
  const lastRowP = shP.getLastRow();
  const existingRows = Math.max(0, lastRowP - CFG.ROWS.FIRST_DATA + 1);
  if (needRows > existingRows) {
    shP.insertRowsAfter(Math.max(CFG.ROWS.FIRST_DATA, lastRowP), needRows - existingRows);
  }

  const firstData = CFG.ROWS.FIRST_DATA;

  // Escreve A:D
  shP.getRange(firstData, 1, needRows, 4).setValues(metaRows);

  // Escreve lojas F..storeEnd
  shP.getRange(firstData, store.storeStart, needRows, store.storeCols).setValues(priceRows);

  // Notas na Marca (C)
  shP.getRange(firstData, CFG.COL.PESQ.MARCA, needRows, 1)
    .setNotes(notesRows.map(n => [n]));

  // Borda vermelha duplicados (A..storeEnd)
  const toBorder = [];
  const colLetter = LOAD2_colToLetter_(store.storeEnd);
  for (let i = 0; i < dupFlags.length; i++) {
    if (!dupFlags[i]) continue;
    const r = firstData + i;
    toBorder.push(`A${r}:${colLetter}${r}`);
  }

  if (toBorder.length > 0) {
    const rl = shP.getRangeList(toBorder);
    if (typeof rl.setBorder === 'function') {
      rl.setBorder(true, true, true, true, true, true, '#d32f2f', SpreadsheetApp.BorderStyle.SOLID);
    } else {
      for (const a1 of toBorder) {
        shP.getRange(a1).setBorder(true, true, true, true, true, true, '#d32f2f', SpreadsheetApp.BorderStyle.SOLID);
      }
    }
  }

  // Reaplica gradiente e separadores (seguro)
  if (typeof adicionarBordasPesquisa === 'function') {
    try { adicionarBordasPesquisa(); } catch (e) {}
  }
  if (typeof atualizarGradienteTudo === 'function') {
    try { atualizarGradienteTudo(); } catch (e) {}
  }
  if (typeof aplicarSeparadoresItens === 'function') {
    try { aplicarSeparadoresItens(); } catch (e) {}
  }

  const extraPeriodo = (periodoInfo && periodoInfo.days > 1)
    ? ` (período detectado, usando ${LOAD2_formatBr_(dateKey)})`
    : '';

  ss.toast(`Carregado: ${needRows} linha(s) de ${LOAD2_formatBr_(dateKey)}${extraPeriodo}.`, 'Carregar histórico', 10);
}

/**
 * Limpa a área editável SEM apagar fórmulas da coluna E (Tipo).
 * Limpa conteúdo: A:D e F..storeEnd
 * Limpa notas: coluna C
 * Remove bordas (inclui duplicados)
 * @private
 * @param {GoogleAppsScript.Spreadsheet.Sheet} shP 
 * @param {Object} store 
 */
function LOAD2_clearPesquisaArea_(shP, store) {
  const firstData = CFG.ROWS.FIRST_DATA;
  const lastRow = Math.max(shP.getLastRow(), firstData);
  const n = lastRow - firstData + 1;
  if (n <= 0) return;

  // A:D
  shP.getRange(firstData, 1, n, 4).clearContent();

  // F..storeEnd (lojas)
  shP.getRange(firstData, store.storeStart, n, store.storeCols).clearContent();

  // notas em Marca (C)
  shP.getRange(firstData, CFG.COL.PESQ.MARCA, n, 1)
    .setNotes(Array.from({ length: n }, () => ['']));

  // bordas no bloco A..storeEnd
  shP.getRange(firstData, 1, n, store.storeEnd).setBorder(false, false, false, false, false, false);
}

/***************
 * Helpers locais do carregamento
 ***************/

/** @private */
function LOAD2_norm_(s) {
  return String(s || '').trim().toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

/** @private */
function LOAD2_colToLetter_(col) {
  let s = '';
  while (col > 0) {
    const m = (col - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    col = Math.floor((col - 1) / 26);
  }
  return s;
}

/** @private */
function LOAD2_formatBr_(dateKey) {
  const [y, m, d] = String(dateKey).split('-');
  return `${d}/${m}/${y}`;
}

/** @private */
function LOAD2_dateToKeyNum_(v) {
  if (!v) return NaN;

  if (v instanceof Date && !isNaN(v.getTime())) {
    return v.getFullYear() * 10000 + (v.getMonth() + 1) * 100 + v.getDate();
  }

  const s = String(v).trim();
  if (!s) return NaN;

  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return (+m[1]) * 10000 + (+m[2]) * 100 + (+m[3]);

  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return (+m[3]) * 10000 + (+m[2]) * 100 + (+m[1]);

  const d = new Date(s);
  if (d instanceof Date && !isNaN(d.getTime())) {
    return d.getFullYear() * 10000 + (d.getMonth() + 1) * 100 + d.getDate();
  }

  return NaN;
}