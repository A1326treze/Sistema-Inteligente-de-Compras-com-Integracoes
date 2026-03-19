/**
 * @fileoverview Filtro Dinâmico por Setor no Inventário.
 * Permite filtrar os itens da aba de Inventário através da digitação de termos 
 * em uma célula específica (D1) e inverter a lógica do filtro via checkbox (E1).
 * Depende do objeto global `CFG` (Config.js).
 */

/**
 * Objeto de configuração local para o módulo de filtro.
 * Encapsula as chamadas ao `CFG` global para facilitar o acesso.
 * @constant
 */
const FS_CFG = {
  sheetName: () => CFG.SHEETS.INV,
  inputA1: () => CFG.CELLS.INV.FILTER_SETOR,          // Ex: D1
  invertA1: () => CFG.CELLS.INV.FILTER_SETOR_INVERT,  // Ex: E1 (checkbox)
  headerRow: () => CFG.ROWS.HEADER,
  firstDataRow: () => CFG.ROWS.FIRST_DATA,
  firstCol: 1, // A
  lastCol: () => CFG.COL.INV.LOCAL, // V (22)
  sectorCol: () => CFG.COL.INV.SETOR // E (5)
};

/**
 * Função gatilho (event handler) para capturar edições na planilha.
 * Deve ser chamada por um gatilho `onEdit(e)` principal ou configurada como gatilho instalável.
 * Só executa o filtro se a edição ocorrer na aba correta e nas células D1 ou E1.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - O objeto de evento de edição do Google Sheets.
 * @returns {void}
 */
function onEditFiltroSetor(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== FS_CFG.sheetName()) return;

    const a1 = e.range.getA1Notation();
    // Reage ao editar a célula de texto OU o checkbox
    if (a1 !== FS_CFG.inputA1() && a1 !== FS_CFG.invertA1()) return;

    aplicarFiltroPorSetor_();
  } catch (err) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Erro no script: ' + err, 'Filtro de Setor', 6);
  }
}

/** * Função utilitária para autorização inicial do script.
 * Execute 1x manualmente no editor do Apps Script para conceder as permissões necessárias.
 * @returns {void}
 */
function autorizar() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Autorização OK.', 'Filtro de Setor', 3);
  aplicarFiltroPorSetor_();
}

/**
 * Aplica a lógica principal de filtragem na aba de inventário.
 * Lê os critérios, cria ou atualiza o filtro nativo do Google Sheets e oculta/exibe
 * as linhas baseando-se no texto digitado e no estado do checkbox de inversão.
 * @private
 * @returns {void}
 */
function aplicarFiltroPorSetor_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(FS_CFG.sheetName());
  if (!sh) return;

  const inputCell = sh.getRange(FS_CFG.inputA1());
  const raw = String(inputCell.getDisplayValue() ?? '').trim();

  const invertCell = sh.getRange(FS_CFG.invertA1());
  const invert = Boolean(invertCell.getValue()); // checkbox => true/false

  // Define uma nota explicativa permanente no checkbox
  invertCell.setNote(
    'Inverter filtro de setor\n' +
    'Desmarcado: mostra apenas os setores digitados.\n' +
    'Marcado: esconde os setores digitados e mostra o resto.'
  );

  const headerRow = FS_CFG.headerRow();
  const lastRow = Math.max(sh.getLastRow(), headerRow);
  const lastCol = FS_CFG.lastCol();

  const filterRange = sh.getRange(
    headerRow,
    FS_CFG.firstCol,
    Math.max(1, lastRow - headerRow + 1),
    lastCol
  );

  let filter = sh.getFilter();
  
  // Cria o filtro se não existir ou recria se o intervalo estiver incorreto
  if (!filter) {
    filterRange.createFilter();
    filter = sh.getFilter();
  } else {
    const fr = filter.getRange();
    const mismatch =
      fr.getRow() !== headerRow ||
      fr.getColumn() !== FS_CFG.firstCol ||
      fr.getNumColumns() !== lastCol;

    if (mismatch) {
      filter.remove();
      filterRange.createFilter();
      filter = sh.getFilter();
    }
  }

  const normRaw = FS_normalize_(raw);
  
  // Palavras-chave para limpar o filtro e mostrar todos os dados
  if (!raw || normRaw === 'todos' || normRaw === 'tudo' || normRaw === 'all') {
    inputCell.setNote('');
    filter.removeColumnFilterCriteria(FS_CFG.sectorCol());
    ss.toast('Filtro limpo (mostrando tudo).', 'Filtro de Setor', 2);
    return;
  }

  // Dica amigável caso o usuário digite "e" em vez de vírgula
  if (normRaw.includes(' e ')) {
    inputCell.setNote('Dica: para filtrar mais de um setor, use vírgula.\nEx.: limpeza, mercearia');
    ss.toast('Use vírgula para separar setores (ex.: limpeza, mercearia).', 'Filtro de Setor', 5);
  }

  const tokens = FS_parseTokens_(raw);
  if (tokens.length === 0) {
    inputCell.setNote('Digite um setor (ou vários separados por vírgula).');
    ss.toast('Nada para filtrar: Célula está vazia/inválida.', 'Filtro de Setor', 4);
    return;
  }

  const firstDataRow = FS_CFG.firstDataRow();
  const numRows = Math.max(0, lastRow - firstDataRow + 1);
  if (numRows === 0) return;

  const sectorCol = FS_CFG.sectorCol();
  const colValues = sh.getRange(firstDataRow, sectorCol, numRows, 1).getDisplayValues().flat();
  const unique = new Set(colValues.map(v => String(v ?? '').trim()));

  // Conjunto de valores presentes na planilha que correspondem à pesquisa
  const matched = new Set();
  for (const val of unique) {
    const nval = FS_normalize_(val);
    for (const t of tokens) {
      if (t && nval.includes(t)) {
        matched.add(val);
        break;
      }
    }
  }

  // Comportamento caso não encontre o setor digitado
  if (matched.size === 0) {
    inputCell.setNote(
      'Setor não encontrado.\n' +
      'Verifique a digitação.\n' +
      'Para múltiplos: use vírgula (ex.: limpeza, mercearia).'
    );
    ss.toast('Nenhum setor encontrado para: ' + raw, 'Filtro de Setor', 6);
    filter.removeColumnFilterCriteria(sectorCol);
    return;
  }

  // Lógica de ocultação (Hidden Values):
  // - normal (invert=false): esconde quem NÃO bateu (mostra só match)
  // - invert (invert=true): esconde quem bateu (mostra o resto)
  const hidden = [];
  for (const val of unique) {
    const isMatch = matched.has(val);
    if (!invert && !isMatch) hidden.push(val);
    if (invert && isMatch) hidden.push(val);
  }

  const criteria = SpreadsheetApp.newFilterCriteria().setHiddenValues(hidden).build();
  filter.setColumnFilterCriteria(sectorCol, criteria);

  inputCell.setNote('');
  ss.toast(
    (invert ? 'Exceto setor: ' : 'Filtrando setor: ') + raw,
    'Filtro de Setor',
    2
  );
}

/**
 * Processa a string de entrada, separando os termos por vírgula ou quebra de linha.
 * @private
 * @param {string} raw - Texto bruto inserido pelo usuário.
 * @returns {string[]} Array de termos normalizados e sem duplicatas.
 */
function FS_parseTokens_(raw) {
  const s = String(raw ?? '');
  const parts = s
    .split(/[,|\n]+/g)
    .map(p => FS_normalize_(p.trim()))
    .filter(Boolean);
  return Array.from(new Set(parts));
}

/**
 * Normaliza uma string removendo acentos, convertendo para minúsculas e removendo espaços extras.
 * @private
 * @param {string} s - Texto a ser normalizado.
 * @returns {string} Texto limpo e formatado para comparação.
 */
function FS_normalize_(s) {
  return String(s ?? '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}
