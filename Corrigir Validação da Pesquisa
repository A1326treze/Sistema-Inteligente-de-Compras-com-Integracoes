/**
 * @fileoverview Validação e Correção de Itens da Pesquisa.
 * Compara os itens lançados na aba de "Pesquisa" com os itens válidos 
 * da aba de "Inventário", corrigindo automaticamente divergências de 
 * digitação, espaços e acentuação.
 * * Depende do objeto global `CFG` (Config.js) para mapeamento de planilhas, colunas e linhas.
 */

/**
 * Função principal que verifica e corrige os itens inválidos na aba de Pesquisa.
 * Lê a lista de itens do inventário, cria um mapa de chaves canônicas (sem acentos, 
 * em minúsculas e sem espaços) e compara com os itens da aba de pesquisa. 
 * Substitui os valores na aba de pesquisa pela nomenclatura exata do inventário.
 * * @returns {void}
 */
function corrigirItensInvalidosPesquisa() {
  const ss = SpreadsheetApp.getActive();
  const shP = ss.getSheetByName(CFG.SHEETS.PESQ);
  const shI = ss.getSheetByName(CFG.SHEETS.INV);
  if (!shP || !shI) return;

  const invLast = shI.getLastRow();
  if (invLast < 2) return;

  const itemColInv = CFG.COL.INV.ITEM;

  // Detecta onde os dados de inventário começam
  const invStart = CVP_detectarPrimeiraLinhaDados_(shI, itemColInv, ['item', 'produto'], 10) || CFG.ROWS.FIRST_DATA;
  const invVals = shI.getRange(invStart, itemColInv, invLast - invStart + 1, 1).getValues();

  // Constrói um mapa de itens válidos (Chave Canônica -> Nome Original)
  const canonMap = new Map();
  for (const [v] of invVals) {
    const s = CVP_limparTexto_(v);
    if (!s) continue;
    const k = CVP_chaveCanon_(s);
    if (!canonMap.has(k)) canonMap.set(k, s);
  }

  const firstDataRow = CFG.ROWS.FIRST_DATA;
  const lr = shP.getLastRow();
  if (lr < firstDataRow) return;

  // Busca os dados da aba de pesquisa
  const n = lr - firstDataRow + 1;
  const rng = shP.getRange(firstDataRow, CFG.COL.PESQ.ITEM, n, 1);
  const vals = rng.getValues();

  let changed = 0;
  
  // Analisa e corrige os itens da pesquisa com base no mapa canônico
  for (let i = 0; i < vals.length; i++) {
    const raw = vals[i][0];
    const s = CVP_limparTexto_(raw);
    if (!s) continue;

    const k = CVP_chaveCanon_(s);
    const canon = canonMap.get(k);
    if (canon && canon !== raw) {
      vals[i][0] = canon;
      changed++;
    }
  }

  // Aplica as correções em lote
  rng.setValues(vals);
  ss.toast(`Itens corrigidos: ${changed}`, 'Pesquisa', 6);
}

/**
 * Limpa uma string removendo espaços invisíveis, espaços duplos e caracteres de formatação.
 * Normaliza o texto para o formato NFC.
 * * @private
 * @param {any} v - O valor bruto da célula.
 * @returns {string} A string limpa e formatada.
 */
function CVP_limparTexto_(v) {
  if (v === null || v === undefined) return '';
  let s = String(v);
  s = s.replace(/\u00A0/g, ' ').replace(/[\u200B-\u200D\uFEFF]/g, '');
  if (s.normalize) s = s.normalize('NFC');
  return s.replace(/\s+/g, ' ').trim();
}

/**
 * Gera uma chave canônica para padronização de buscas.
 * Remove acentuação, caracteres especiais e espaços, convertendo tudo para letras minúsculas.
 * * @private
 * @param {string} s - O texto base.
 * @returns {string} A chave canônica (ex: "Maçã Verde" -> "macaverde").
 */
function CVP_chaveCanon_(s) {
  return String(s)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '');
}

/**
 * Varre as primeiras linhas de uma coluna para detectar a primeira linha real de dados,
 * baseando-se em palavras-chave que indicam o cabeçalho.
 * * @private
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - A aba da planilha a ser analisada.
 * @param {number} col - O índice da coluna (1-based).
 * @param {string[]} headerOptionsLower - Array de possíveis nomes de cabeçalho em minúsculas (ex: ['item', 'produto']).
 * @param {number} maxScanRows - O limite máximo de linhas a serem lidas na busca do cabeçalho.
 * @returns {number|null} O número da primeira linha de dados (linha do cabeçalho + 1) ou null se não for encontrado.
 */
function CVP_detectarPrimeiraLinhaDados_(sheet, col, headerOptionsLower, maxScanRows) {
  const scan = Math.min(maxScanRows, sheet.getLastRow());
  const vals = sheet.getRange(1, col, scan, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    const t = String(vals[i][0] || '').trim().toLowerCase();
    if (headerOptionsLower.includes(t)) return i + 2; // +1 do índice do array, +1 para a linha debaixo do cabeçalho
  }
  return null;
}
