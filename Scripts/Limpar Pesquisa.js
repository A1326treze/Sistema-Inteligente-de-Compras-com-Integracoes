/**
 * @fileoverview Utilitários de Limpeza para a aba de Pesquisa.
 * Limpa as colunas A:D e as colunas de lojas (F até antes de Custo/Mult), 
 * preservando estritamente a coluna E (Tipo).
 * Também remove notas, bordas e cores de fundo residuais.
 */

/* ===== Limpar Pesquisa.js (atualizado para Config.js / CFG) =====
 * Limpa A:D + lojas (F..antes de Custo/Mult), sem mexer na coluna E (Tipo).
 * Também remove notas/bordas/cores aplicadas por Carregar/Gradiente.
 */

/**
 * Restaura a borda preta espessa na coluna F (separador visual das lojas).
 * Sobre os parâmetros do `setBorder()`:
 * setBorder(top, left, bottom, right, vertical, horizontal, cor, estilo)
 * @returns {void}
 */
function adicionarBordasPesquisa() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEETS.PESQ);
  const lastRow = sh.getLastRow();

  // Coluna F
  sh.getRange("F3:F" + lastRow).setBorder(
    false, true, false, false, false, false,
    "#000000",
    SpreadsheetApp.BorderStyle.SOLID_THICK
  );
}

/**
 * Executa a limpeza dos dados da aba de pesquisa de forma seletiva.
 * Limpa o intervalo de A até D e o intervalo contendo os preços das lojas.
 * Após a limpeza dos dados e formatos, aciona as funções de formatação visual
 * para reconstruir o layout da planilha.
 * @returns {void}
 */
function limparPrecosPesquisa() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.SHEETS.PESQ);
  if (!sh) {
    ss.toast(`Não encontrei a aba "${CFG.SHEETS.PESQ}".`, 'Limpar', 6);
    return;
  }

  const firstData = CFG.ROWS.FIRST_DATA;
  const lr = sh.getLastRow();
  const lastRow = Math.max(lr, firstData);
  const nRows = lastRow - firstData + 1;

  const store = CFG.getStoreInfo(sh, CFG.ROWS.HEADER);

  // A:D
  const rngAD = sh.getRange(firstData, 1, nRows, 4);
  rngAD.clearContent();
  rngAD.clearNote();
  rngAD.setBorder(false, false, false, false, false, false);

  // lojas (F..storeEnd)
  const rngStores = sh.getRange(firstData, store.storeStart, nRows, store.storeCols);
  rngStores.clearContent();
  rngStores.clearNote();
  rngStores.setBackground(null);
  rngStores.setFontColor(null);
  rngStores.setBorder(false, false, false, false, false, false);

  SpreadsheetApp.flush();

  // Visual
  try { atualizarGradienteTudo(); } catch (e) {}
  try { aplicarSeparadoresItens(); } catch (e) {}
  try { adicionarBordasPesquisa(); } catch (e) {}

  ss.toast('Pesquisa limpa + visual atualizado.', 'Limpar', 4);
}