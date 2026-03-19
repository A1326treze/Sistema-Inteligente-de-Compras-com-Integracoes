/* ===== Inventário.js (atualizado para Config.js / CFG) ===== */

/**
 * @fileoverview Organização e Ordenação do Inventário.
 * Script responsável por limpar (padronizar) os itens e reordenar as linhas 
 * do inventário de forma alfanumérica, baseando-se nas colunas de Ordem e Item.
 */

/**
 * Executa a padronização de nomenclatura e ordena o inventário inteiro.
 * A ordenação ocorre em dois níveis: 
 * 1º - Coluna de Ordem (Ascendente)
 * 2º - Coluna de Item (Ascendente)
 * Depende do objeto global `CFG` e da função `padronizarItensInventario()`.
 * @returns {void}
 */
function reordenarInventario() {
  padronizarItensInventario();

  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEETS.INV);
  if (!sh) return;

  const FIRST_DATA_ROW = CFG.ROWS.FIRST_DATA;
  const COL_ORDEM = CFG.COL.INV.ORDEM; // B
  const COL_ITEM  = CFG.COL.INV.ITEM;  // C

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < FIRST_DATA_ROW) return;

  const range = sh.getRange(FIRST_DATA_ROW, 1, lastRow - FIRST_DATA_ROW + 1, lastCol);

  range.sort([
    { column: COL_ORDEM, ascending: true },
    { column: COL_ITEM,  ascending: true }
  ]);

  SpreadsheetApp.getActive().toast('Inventário reordenado.');
}