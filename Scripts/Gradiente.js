/**
 * @fileoverview Sistema de Formatação Condicional (Gradiente) e Checkbox.
 * Aplica um mapa de calor aos preços digitados, baseando-se no custo ajustado
 * (Preço * Multiplicador) do item. Utiliza gatilhos (onEdit) para atualizar as
 * cores apenas dos itens editados, garantindo alta performance.
 */

/***************
 * CONFIG (alinhado com sua spec)
 ***************/
const GRAD_HEADER_ROW = 2;
const GRAD_FIRST_DATA_ROW = 3;

// Colunas fixas da aba "Preços (pesquisa)"
const GRAD_COL_CHECK = 1;        // A
const GRAD_COL_ITEM  = 2;        // B
const GRAD_COL_MARCA = 3;        // C
const GRAD_COL_QTD   = 4;        // D
const GRAD_COL_TIPO  = 5;        // E
const GRAD_COL_LOJA_START = 6;   // F (primeira loja)

// Cabeçalhos auxiliares que marcam o fim das lojas
const GRAD_HDR_CUSTO = 'Custo';
const GRAD_HDR_MULT  = 'Mult';

// Gradiente (normal)
const GRAD_GREEN  = '#63BE7B';
const GRAD_YELLOW = '#FFEB84';
const GRAD_RED    = '#F8696B';
const GRAD_WHITE  = '#ffffff';
const GRAD_BLACK  = '#000000';

// Destaques
const GRAD_BEST_BG   = '#1B5E20'; // verde escuro
const GRAD_BEST_FONT = '#ffffff'; // branco
const GRAD_SECOND_BG   = '#90CAF9'; // azul claro (fora da escala)
const GRAD_SECOND_FONT = '#000000'; // preto

/**
 * Função pública para atualizar o gradiente da aba inteira.
 * Criada como alias por compatibilidade com menus legados.
 * @returns {void}
 */
// Alias (se algum menu antigo chamar esse nome)
function atualizarGradientePrecos() {
  return atualizarGradienteTudo();
}

/**
 * Identifica dinamicamente onde começam e terminam as colunas de lojas.
 * Varre o cabeçalho procurando os marcos finais "Custo" ou "Mult".
 * @private
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - A aba de pesquisa.
 * @returns {Object} Posições start, end e quantidade de colunas.
 * @throws {Error} Se os cabeçalhos não forem encontrados.
 */
/**
 * Detecta dinamicamente:
 * - col inicial de lojas (fixa em F)
 * - col final de lojas (antes de "Custo" ou "Mult")
 * - coluna "Mult"
 */
function grad_getStoreInfo_(sh) {
  const lastCol = sh.getLastColumn();
  const header = sh.getRange(GRAD_HEADER_ROW, 1, 1, lastCol).getDisplayValues()[0];

  let custoCol = -1;
  let multCol = -1;

  for (let c = 0; c < header.length; c++) {
    const h = String(header[c] || '').trim().toLowerCase();
    if (h === GRAD_HDR_CUSTO.toLowerCase()) custoCol = c + 1;
    if (h === GRAD_HDR_MULT.toLowerCase()) multCol = c + 1;
  }

  const storeStart = GRAD_COL_LOJA_START;
  let storeEnd = lastCol;

  if (custoCol > 0) storeEnd = custoCol - 1;
  else if (multCol > 0) storeEnd = multCol - 1;

  if (storeEnd < storeStart) {
    throw new Error(`Gradiente: não consegui detectar colunas de loja. storeEnd=${storeEnd}, storeStart=${storeStart}`);
  }
  if (multCol <= 0) {
    throw new Error(`Gradiente: não encontrei a coluna "${GRAD_HDR_MULT}" no cabeçalho (linha ${GRAD_HEADER_ROW}).`);
  }

  return {
    storeStart,
    storeEnd,
    storeCols: storeEnd - storeStart + 1,
    multCol
  };
}

/**
 * Gatilho principal que reage a edições manuais do usuário.
 * Dispara atualizações de separadores visuais, replicação de checkboxes e 
 * recalcula as cores de gradiente apenas para o item da linha editada.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - Objeto do evento de edição.
 * @returns {void}
 */
/***************
 * ÚNICO onEdit (checkbox + gradiente + separadores)
 ***************/
function onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== CFG.SHEETS.PESQ) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  const numRows = e.range.getNumRows();
  const numCols = e.range.getNumColumns();

  if (row < GRAD_FIRST_DATA_ROW) return;

  // ✅ Só atualiza separadores se a edição tocar a coluna do Item (B)
  const colEnd = col + numCols - 1;
  const affectsItemCol = (col <= GRAD_COL_ITEM && colEnd >= GRAD_COL_ITEM);

  let info;
  try {
    info = CFG.getStoreInfo(sh, CFG.ROWS.HEADER);
    info.multCol = info.colMult; // para manter o resto do arquivo funcionando sem refatoração grande
  } catch (err) {
    if (affectsItemCol) sep_updateAroundRange_(sh, row, numRows);
    return;
  }

  const isPriceCol = (col >= info.storeStart && col <= info.storeEnd);
  const isRelevantForGradient =
    col === GRAD_COL_CHECK ||
    col === GRAD_COL_ITEM ||
    col === GRAD_COL_QTD ||
    col === GRAD_COL_TIPO ||
    col === info.multCol ||
    isPriceCol;

  // Se não for relevante, não atualiza gradiente (e separadores só se tocou Item)
  if (!isRelevantForGradient) {
    if (affectsItemCol) sep_updateAroundRange_(sh, row, numRows);
    return;
  }

  // 1) Checkbox: marca/desmarca todas do mesmo item + atualiza gradiente do item
  if (col === GRAD_COL_CHECK) {
    const item = String(sh.getRange(row, GRAD_COL_ITEM).getValue() || '').trim();
    if (!item) return;

    const newVal = sh.getRange(row, GRAD_COL_CHECK).getValue(); // TRUE/FALSE
    grad_marcarTodasCheckboxesDoItem_(sh, item, newVal);

    // Sem flush aqui: checkbox não afeta mult/preço.
    atualizarGradientePorItem_(item, info);

    if (affectsItemCol) sep_updateAroundRange_(sh, row, numRows);
    return;
  }

  // 2) Atualiza gradiente do item da linha editada (preço/mult/qtd/tipo/item)
  const item = String(sh.getRange(row, GRAD_COL_ITEM).getValue() || '').trim();

  // Se editou algo que pode mexer em fórmulas de Mult, dá um flush pontual:
  const needsFlush = (col === GRAD_COL_QTD || col === GRAD_COL_TIPO || col === info.multCol);
  if (needsFlush) SpreadsheetApp.flush();

  if (item) atualizarGradientePorItem_(item, info);

  if (affectsItemCol) sep_updateAroundRange_(sh, row, numRows);
}

/**
 * Ao marcar o checkbox de um item, replica a marcação para todas as linhas que 
 * contenham o mesmo item na planilha inteira.
 * @private
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 
 * @param {string} item 
 * @param {boolean} newVal 
 */
/***************
 * Checkbox: marcar todas do mesmo item (mantido simples e seguro)
 ***************/
function grad_marcarTodasCheckboxesDoItem_(sheet, item, newVal) {
  const lr = sheet.getLastRow();
  if (lr < GRAD_FIRST_DATA_ROW) return;

  const nRows = lr - GRAD_FIRST_DATA_ROW + 1;
  const key = grad_norm_(item);

  const items = sheet.getRange(GRAD_FIRST_DATA_ROW, GRAD_COL_ITEM, nRows, 1).getValues();
  const checksRange = sheet.getRange(GRAD_FIRST_DATA_ROW, GRAD_COL_CHECK, nRows, 1);
  const checks = checksRange.getValues();

  for (let i = 0; i < items.length; i++) {
    const it = grad_norm_(items[i][0]);
    if (it === key) checks[i][0] = newVal;
  }
  checksRange.setValues(checks);
}

/**
 * Analisa os preços de toda a aba de pesquisa, calcula estatísticas e recolore
 * a tabela inteira baseando-se no mapa de calor.
 * @returns {void}
 */
/***************
 * Atualizar gradiente (tudo) — menu
 ***************/
function atualizarGradienteTudo() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.SHEETS.PESQ);
  if (!sh) return;

  const lr = sh.getLastRow();
  if (lr < GRAD_FIRST_DATA_ROW) return;

  let info;
  try {
    info = CFG.getStoreInfo(sh, CFG.ROWS.HEADER);
    info.multCol = info.colMult; // para manter o resto do arquivo funcionando sem refatoração grande
  } catch (err) {
    ss.toast(`Gradiente: ${err.message}`);
    return;
  }

  const nRows = lr - GRAD_FIRST_DATA_ROW + 1;

  const items = sh.getRange(GRAD_FIRST_DATA_ROW, GRAD_COL_ITEM, nRows, 1).getValues().map(r => grad_norm_(r[0]));
  const mults = sh.getRange(GRAD_FIRST_DATA_ROW, info.multCol, nRows, 1).getValues().map(r => grad_toNum_(r[0]));

  const pricesRange = sh.getRange(GRAD_FIRST_DATA_ROW, info.storeStart, nRows, info.storeCols);
  const prices = pricesRange.getValues();

  const stats = grad_buildItemStats_(items, mults, prices);

  const bgs = [];
  const fonts = [];

  for (let r = 0; r < nRows; r++) {
    const rowBg = new Array(info.storeCols).fill(GRAD_WHITE);
    const rowFont = new Array(info.storeCols).fill(GRAD_BLACK);

    const item = items[r];
    if (!item || !stats.has(item)) {
      bgs.push(rowBg);
      fonts.push(rowFont);
      continue;
    }

    const st = stats.get(item);
    const mult = grad_effectiveMult_(mults[r]); // fallback = 1 (visual)

    for (let c = 0; c < info.storeCols; c++) {
      const p = grad_toNum_(prices[r][c]);
      if (!isFinite(p) || p <= 0) {
        rowBg[c] = GRAD_WHITE;
        rowFont[c] = GRAD_BLACK;
        continue;
      }

      const cost = p * mult;

      if (grad_eq_(cost, st.best)) {
        rowBg[c] = GRAD_BEST_BG;
        rowFont[c] = GRAD_BEST_FONT;
        continue;
      }
      if (isFinite(st.second) && grad_eq_(cost, st.second)) {
        rowBg[c] = GRAD_SECOND_BG;
        rowFont[c] = GRAD_SECOND_FONT;
        continue;
      }

      const t = grad_clamp01_(st.span > 0 ? (cost - st.min) / st.span : 0.5);
      rowBg[c] = grad_triColor_(t, GRAD_GREEN, GRAD_YELLOW, GRAD_RED);
      rowFont[c] = GRAD_BLACK;
    }

    bgs.push(rowBg);
    fonts.push(rowFont);
  }

  pricesRange.setBackgrounds(bgs);
  pricesRange.setFontColors(fonts);
}

/**
 * Função otimizada (PATCH) que recolore apenas os blocos de linhas de um único item.
 * Evita o processamento da planilha inteira a cada edição do onEdit.
 * @param {string} itemRaw 
 * @param {Object} infoOpt 
 * @returns {void}
 */
/***************
 * Atualizar gradiente: só um item (PATCH OTIMIZADO)
 ***************/
function atualizarGradientePorItem_(itemRaw, infoOpt) {
  const itemText = String(itemRaw || '').trim();
  if (!itemText) return;

  const itemKey = grad_norm_(itemText);
  if (!itemKey) return;

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.SHEETS.PESQ);
  if (!sh) return;

  const lr = sh.getLastRow();
  if (lr < GRAD_FIRST_DATA_ROW) return;

  const info = infoOpt || grad_getStoreInfo_(sh);

  const nRows = lr - GRAD_FIRST_DATA_ROW + 1;
  const itemRange = sh.getRange(GRAD_FIRST_DATA_ROW, GRAD_COL_ITEM, nRows, 1);

  // Encontra todas as ocorrências do item (itens podem estar espalhados)
  const matches = itemRange
    .createTextFinder(itemText)
    .matchEntireCell(true)
    .matchCase(false)
    .findAll();

  if (!matches || matches.length === 0) return;

  const rows = matches.map(m => m.getRow()).filter(r => r >= GRAD_FIRST_DATA_ROW).sort((a, b) => a - b);
  if (rows.length === 0) return;

  const segments = grad_groupContiguousRows_(rows);

  // 1) Coletar dados só desses segmentos
  const segData = segments.map(seg => {
    const segRows = seg.end - seg.start + 1;
    const mults = sh.getRange(seg.start, info.multCol, segRows, 1).getValues().map(r => grad_toNum_(r[0]));
    const pricesRange = sh.getRange(seg.start, info.storeStart, segRows, info.storeCols);
    const prices = pricesRange.getValues();
    return { startRow: seg.start, numRows: segRows, mults, prices, pricesRange };
  });

  // 2) Calcular stats do item só com esses dados
  const st = grad_buildStatsFromSegData_(segData);
  if (!st) return;

  // 3) Aplicar gradiente só nos segmentos
  for (const seg of segData) {
    const bgs = [];
    const fonts = [];

    for (let r = 0; r < seg.numRows; r++) {
      const rowBg = new Array(info.storeCols).fill(GRAD_WHITE);
      const rowFont = new Array(info.storeCols).fill(GRAD_BLACK);

      const mult = grad_effectiveMult_(seg.mults[r]);

      for (let c = 0; c < info.storeCols; c++) {
        const p = grad_toNum_(seg.prices[r][c]);
        if (!isFinite(p) || p <= 0) {
          rowBg[c] = GRAD_WHITE;
          rowFont[c] = GRAD_BLACK;
          continue;
        }

        const cost = p * mult;

        if (grad_eq_(cost, st.best)) {
          rowBg[c] = GRAD_BEST_BG;
          rowFont[c] = GRAD_BEST_FONT;
          continue;
        }
        if (isFinite(st.second) && grad_eq_(cost, st.second)) {
          rowBg[c] = GRAD_SECOND_BG;
          rowFont[c] = GRAD_SECOND_FONT;
          continue;
        }

        const t = grad_clamp01_(st.span > 0 ? (cost - st.min) / st.span : 0.5);
        rowBg[c] = grad_triColor_(t, GRAD_GREEN, GRAD_YELLOW, GRAD_RED);
        rowFont[c] = GRAD_BLACK;
      }

      bgs.push(rowBg);
      fonts.push(rowFont);
    }

    seg.pricesRange.setBackgrounds(bgs);
    seg.pricesRange.setFontColors(fonts);
  }
}

/**
 * Constrói as estatísticas (menor preço, maior preço, amplitude) agrupando
 * o array completo da planilha por chaves de "Itens".
 * @private
 */
/***************
 * STATS
 ***************/
function grad_buildItemStats_(items, mults, prices) {
  const map = new Map();

  for (let r = 0; r < items.length; r++) {
    const item = items[r];
    if (!item) continue;

    const mult = grad_effectiveMult_(mults[r]); // fallback=1
    for (let c = 0; c < prices[r].length; c++) {
      const p = grad_toNum_(prices[r][c]);
      if (!isFinite(p) || p <= 0) continue;

      const cost = p * mult;

      if (!map.has(item)) map.set(item, { min: cost, max: cost, values: [cost] });
      else {
        const obj = map.get(item);
        if (cost < obj.min) obj.min = cost;
        if (cost > obj.max) obj.max = cost;
        obj.values.push(cost);
      }
    }
  }

  const out = new Map();
  for (const [item, obj] of map.entries()) {
    const uniq = grad_uniqueSorted_(obj.values);
    out.set(item, {
      min: obj.min,
      max: obj.max,
      span: obj.max - obj.min,
      best: uniq.length >= 1 ? uniq[0] : NaN,
      second: uniq.length >= 2 ? uniq[1] : NaN
    });
  }
  return out;
}

/**
 * Versão otimizada de cálculo de estatísticas (min, max, etc.) para trabalhar
 * apenas em cima de arrays curtos (segmentos parciais dos itens).
 * @private
 */
function grad_buildStatsFromSegData_(segData) {
  let min = Infinity;
  let max = -Infinity;
  const vals = [];

  for (const seg of segData) {
    for (let r = 0; r < seg.numRows; r++) {
      const mult = grad_effectiveMult_(seg.mults[r]);
      for (let c = 0; c < seg.prices[r].length; c++) {
        const p = grad_toNum_(seg.prices[r][c]);
        if (!isFinite(p) || p <= 0) continue;
        const cost = p * mult;
        vals.push(cost);
        if (cost < min) min = cost;
        if (cost > max) max = cost;
      }
    }
  }

  if (!isFinite(min) || !isFinite(max) || vals.length === 0) return null;

  const uniq = grad_uniqueSorted_(vals);
  return {
    min,
    max,
    span: max - min,
    best: uniq.length >= 1 ? uniq[0] : NaN,
    second: uniq.length >= 2 ? uniq[1] : NaN
  };
}

/**
 * Agrupa arrays numéricos soltos (ex: [2, 3, 4, 10, 11]) em segmentos lógicos
 * ex: [{start:2, end:4}, {start:10, end:11}] para processamento em batch.
 * @private
 */
function grad_groupContiguousRows_(sortedRows) {
  const segs = [];
  let start = sortedRows[0];
  let prev = sortedRows[0];

  for (let i = 1; i < sortedRows.length; i++) {
    const r = sortedRows[i];
    if (r === prev + 1) {
      prev = r;
      continue;
    }
    segs.push({ start, end: prev });
    start = r;
    prev = r;
  }
  segs.push({ start, end: prev });
  return segs;
}

/***************
 * HELPERS
 ***************/

/** @private */
function grad_norm_(v) {
  return String(v || '').trim().toLowerCase();
}

/** @private */
function grad_clamp01_(x) {
  return Math.max(0, Math.min(1, x));
}

/** @private */
function grad_effectiveMult_(m) {
  return (isFinite(m) && m > 0) ? m : 1; // fallback visual
}

/** @private */
function grad_toNum_(v) {
  if (typeof v === 'number') return v;
  const s = String(v || '').trim();
  if (!s) return NaN;
  const cleaned = s.replace(/\s/g, '').replace(/^R\$/i, '').replace(/\./g, '').replace(',', '.');
  const n = Number(cleaned);
  return isFinite(n) ? n : NaN;
}

/** @private */
function grad_eq_(a, b) {
  return Math.abs(a - b) < 1e-9;
}

/** @private */
function grad_uniqueSorted_(arr) {
  const sorted = arr.slice().sort((x, y) => x - y);
  const out = [];
  for (const v of sorted) {
    if (out.length === 0 || !grad_eq_(v, out[out.length - 1])) out.push(v);
  }
  return out;
}

/** @private */
function grad_hexToRgb_(hex) {
  const h = hex.replace('#', '');
  const n = parseInt(h, 16);
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255 };
}

/** @private */
function grad_rgbToHex_(r, g, b) {
  const toHex = (x) => x.toString(16).padStart(2, '0');
  return '#' + toHex(r) + toHex(g) + toHex(b);
}

/**
 * Calcula dinamicamente o código Hex correspondente a um ponto (t) em um
 * gradiente interpolado de três cores (min, mid, max).
 * @private
 */
function grad_triColor_(t, cMin, cMid, cMax) {
  const a = grad_hexToRgb_(t <= 0.5 ? cMin : cMid);
  const b = grad_hexToRgb_(t <= 0.5 ? cMid : cMax);
  const tt = (t <= 0.5) ? (t * 2) : ((t - 0.5) * 2);

  const r = Math.round(a.r + (b.r - a.r) * tt);
  const g = Math.round(a.g + (b.g - a.g) * tt);
  const bb = Math.round(a.b + (b.b - a.b) * tt);
  return grad_rgbToHex_(r, g, bb);
}