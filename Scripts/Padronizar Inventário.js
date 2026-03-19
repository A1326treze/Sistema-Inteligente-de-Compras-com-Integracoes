/**
 * @fileoverview Padronização e Formatação do Inventário.
 * Este script aplica estilos visuais (bordas) em colunas específicas para melhorar 
 * a legibilidade e padroniza a formatação de texto dos itens listados no inventário.
 * Depende do objeto global `CFG` (Config.js) para mapeamento.
 */

/**
 * Aplica estilos de borda personalizados em colunas específicas da aba de Inventário.
 * As bordas são aplicadas sempre à esquerda das colunas definidas (D, I, L, P, R),
 * a partir da linha 3 até a última linha com dados.
 * * Parâmetros do `setBorder()`: (top, left, bottom, right, vertical, horizontal, cor, estilo)
 * @returns {void}
 */
function adicionarBordas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEETS.INV);
  const lastRow = sh.getLastRow();

  if (lastRow < 3) return; // Evita erros se a planilha estiver vazia

  // Coluna D (índice 4) - Borda esquerda sólida espessa
  sh.getRange(3, 4, lastRow, 1).setBorder(
    false, true, false, false, false, false,
    "#000000",
    SpreadsheetApp.BorderStyle.SOLID_THICK
  );

  // Coluna I (índice 9) - Borda esquerda sólida média
  sh.getRange(3, 9, lastRow, 1).setBorder(
    false, true, false, false, false, false,
    "#000000",
    SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );

  // Coluna L (índice 12) - Borda esquerda sólida espessa
  sh.getRange(3, 12, lastRow, 1).setBorder(
    false, true, false, false, false, false,
    "#000000",
    SpreadsheetApp.BorderStyle.SOLID_THICK
  );

  // Coluna P - Borda esquerda dupla
  sh.getRange("P3:P" + lastRow).setBorder(
    false, true, false, false, false, false,
    "#000000",
    SpreadsheetApp.BorderStyle.DOUBLE
  );

  // Coluna R - Borda esquerda sólida espessa
  sh.getRange("R3:R" + lastRow).setBorder(
    false, true, false, false, false, false,
    "#000000",
    SpreadsheetApp.BorderStyle.SOLID_THICK
  );
}

/**
 * Limpa e padroniza a nomenclatura dos itens na aba de Inventário.
 * Identifica a coluna de itens, varre os dados removendo espaços extras e invisíveis,
 * e por fim, reescreve os dados corrigidos em lote e aplica as bordas visuais.
 * @returns {void}
 */
function padronizarItensInventario() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.SHEETS.INV);
  if (!sh) return;

  const ITEM_COL = CFG.COL.INV.ITEM;
  const lr = sh.getLastRow();
  if (lr < 2) return;

  // Busca o cabeçalho para detectar onde começam os dados
  const scan = Math.min(10, lr);
  const head = sh.getRange(1, ITEM_COL, scan, 1).getValues().map(r => INV_clean_(r[0]).toLowerCase());

  let firstData = CFG.ROWS.FIRST_DATA;
  for (let i = 0; i < head.length; i++) {
    if (head[i] === 'item' || head[i] === 'produto') { 
      firstData = i + 2; 
      break; 
    }
  }
  
  if (lr < firstData) return;

  const n = lr - firstData + 1;
  const rng = sh.getRange(firstData, ITEM_COL, n, 1);
  const vals = rng.getValues();

  let changed = 0;
  
  // Analisa e limpa cada item
  for (let i = 0; i < vals.length; i++) {
    const raw = vals[i][0];
    const cleaned = INV_clean_(raw);
    if (cleaned && cleaned !== raw) {
      vals[i][0] = cleaned;
      changed++;
    }
  }
  
  SpreadsheetApp.flush();
  
  // Tenta aplicar as bordas, ignorando erros silenciosamente caso falhe
  try { adicionarBordas(); } catch (e) { console.warn("Erro ao adicionar bordas:", e); }
  
  rng.setValues(vals);
  ss.toast(`Inventário padronizado: ${changed} itens ajustados`, 'Inventário', 6);
}

/**
 * Função utilitária para higienização de strings.
 * Remove espaços invisíveis (ex: non-breaking spaces), normaliza a codificação (NFC),
 * e remove espaços em branco duplicados ou nas extremidades.
 * @private
 * @param {any} v - O valor bruto da célula.
 * @returns {string} A string formatada e limpa.
 */
function INV_clean_(v) {
  if (v === null || v === undefined) return '';
  let s = String(v);

  // Substitui non-breaking spaces por espaços normais e remove zero-width spaces
  s = s.replace(/\u00A0/g, ' ')
       .replace(/[\u200B-\u200D\uFEFF]/g, '');

  if (s.normalize) s = s.normalize('NFC');

  s = s.replace(/\s+/g, ' ').trim();
  return s;
}
