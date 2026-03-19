/**
 * @fileoverview Gerenciador de Separadores Visuais de Itens.
 * Aplica automaticamente bordas inferiores para separar visualmente grupos de itens 
 * diferentes na aba de pesquisa, facilitando a leitura e a identificação de blocos.
 */

/* ===== Separadores.gs.js — v2 (centralizado em Config.js) =====
 * - Usa CFG para nome de aba, linhas e coluna do item.
 * - Corrige o bug de bordas “presas” (limpa horizontal).
 * - Mantém a lógica otimizada (batch).
 */

/**
 * Configurações locais para os separadores visuais.
 * @constant {Object}
 */
const SEP_CFG = {
  sheetName: () => CFG.SHEETS.PESQ,
  firstDataRow: () => CFG.ROWS.FIRST_DATA,
  colItem: () => CFG.COL.PESQ.ITEM,

  color: '#9E9E9E',
  style: SpreadsheetApp.BorderStyle.SOLID_MEDIUM,

  // Quantas linhas no mínimo limpar (pra remover separadores antigos após "Limpar")
  minClearRows: 300
};

/**
 * Aplica separadores visuais em toda a aba de pesquisa.
 * Remove bordas residuais de uma grande área e depois varre todos os itens,
 * aplicando uma borda inferior sempre que o item da linha atual for diferente do próximo.
 * @returns {void}
 */
// Manual (menu) — aplica separadores na aba inteira
function aplicarSeparadoresItens() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SEP_CFG.sheetName());
  if (!sh) return;

  const firstData = SEP_CFG.firstDataRow();
  const colItem = SEP_CFG.colItem();

  const lastCol = sh.getLastColumn();
  const lastRow = sh.getLastRow();

  // Limpa um bloco "grande o suficiente" (mesmo se hoje a aba está vazia)
  const targetLastRow = Math.max(lastRow, firstData + SEP_CFG.minClearRows);
  const nClear = targetLastRow - firstData + 1;

  // Remove separadores antigos: bottom=false e horizontal=false
  sh.getRange(firstData, 1, nClear, lastCol)
    .setBorder(null, null, false, null, null, false);

  // Se não tem dados, acabou: já limpou as bordas “presas”
  if (lastRow < firstData) return;

  const nRows = lastRow - firstData + 1;

  const items = sh.getRange(firstData, colItem, nRows, 1)
    .getDisplayValues()
    .map(r => String(r[0] || '').trim().toLowerCase());

  const rowsToBorder = [];
  for (let i = 0; i < items.length - 1; i++) {
    const cur = items[i];
    const next = items[i + 1];
    if (cur && next && cur !== next) rowsToBorder.push(firstData + i);
  }

  sep_applyBottomBorders_(sh, rowsToBorder, lastCol);
}

/**
 * Atualiza os separadores apenas ao redor de um intervalo modificado (otimização).
 * Verifica uma linha antes e uma depois da edição, limpa as bordas desse pequeno 
 * bloco e as reaplica, economizando tempo de processamento em edições manuais.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Aba onde ocorreu a edição.
 * @param {number} startRow - Linha inicial da edição.
 * @param {number} numRows - Quantidade de linhas editadas.
 * @returns {void}
 */
// Automático — chamado pelo Gradiente.js quando mexer no Item
function sep_updateAroundRange_(sheet, startRow, numRows) {
  if (!sheet || sheet.getName() !== SEP_CFG.sheetName()) return;

  const firstData = SEP_CFG.firstDataRow();
  const colItem = SEP_CFG.colItem();

  const lastRow = sheet.getLastRow();
  if (lastRow < firstData) return;

  const lastCol = sheet.getLastColumn();

  // inclui uma linha antes e uma depois
  const r0 = Math.max(firstData, startRow - 1);
  const r1 = Math.min(lastRow, startRow + Math.max(1, numRows));

  // separador na linha r depende de r e r+1
  const checkFrom = r0;
  const checkTo = Math.min(lastRow - 1, r1);
  if (checkTo < checkFrom) return;

  // Lê itens de checkFrom até checkTo+1
  const n = (checkTo - checkFrom + 2);
  const items = sheet.getRange(checkFrom, colItem, n, 1)
    .getDisplayValues()
    .map(r => String(r[0] || '').trim().toLowerCase());

  // Limpa separadores antigos nesse bloco (mesmo fix de horizontal=false)
  const blockRows = (checkTo - checkFrom + 1);
  sheet.getRange(checkFrom, 1, blockRows, lastCol)
    .setBorder(null, null, false, null, null, false);

  // Reaplica só onde precisa
  const rowsToBorder = [];
  for (let i = 0; i < n - 1; i++) {
    const cur = items[i];
    const next = items[i + 1];
    if (cur && next && cur !== next) rowsToBorder.push(checkFrom + i);
  }

  sep_applyBottomBorders_(sheet, rowsToBorder, lastCol);
}

/**
 * Aplica a formatação de borda inferior em lote usando RangeList (alta performance).
 * Possui um fallback para loop individual caso RangeList falhe.
 * @private
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - A aba alvo.
 * @param {number[]} rows - Array contendo os índices das linhas que receberão borda inferior.
 * @param {number} lastCol - Número da última coluna para definir a extensão da borda.
 * @returns {void}
 */
// Aplica borda inferior em várias linhas com poucas chamadas
function sep_applyBottomBorders_(sheet, rows, lastCol) {
  if (!rows || rows.length === 0) return;

  const colLetter = sep_colToLetter_(lastCol);
  const a1s = rows.map(r => `A${r}:${colLetter}${r}`);

  const rl = sheet.getRangeList(a1s);
  if (typeof rl.setBorder === 'function') {
    rl.setBorder(null, null, true, null, null, null, SEP_CFG.color, SEP_CFG.style);
    return;
  }

  // fallback
  for (const r of rows) {
    sheet.getRange(r, 1, 1, lastCol)
      .setBorder(null, null, true, null, null, null, SEP_CFG.color, SEP_CFG.style);
  }
}

/**
 * Converte o índice numérico da coluna para a notação alfabética (ex: 1 -> A, 2 -> B).
 * @private
 * @param {number} col - O índice da coluna (1-based).
 * @returns {string} Letra correspondente à coluna.
 */
function sep_colToLetter_(col) {
  let s = '';
  while (col > 0) {
    const m = (col - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    col = Math.floor((col - 1) / 26);
  }
  return s;
}