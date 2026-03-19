/* ===== Oculta Colunas.js (atualizado para Config.js / CFG) ===== */

/**
 * @fileoverview Sistema Inteligente de Ocultação de Colunas.
 * Permite ao usuário ocultar/exibir colunas de lojas específicas baseando-se em comandos
 * de texto (nomes, letras ou exclusões com "!") digitados em uma célula específica.
 * Bloqueia a célula de input enquanto o filtro está ativo.
 */

/**
 * Configurações locais para a rotina de ocultação de colunas.
 * @constant {Object}
 */
const HIDE_CFG = {
  sheetName: () => CFG.SHEETS.PESQ,
  headerRow: () => CFG.ROWS.HEADER,
  inputA1: 'E1',
  lojasStartCol: () => CFG.COL.PESQ.LOJA_START, // F
  stateKey: 'HIDE_COLS_STATE_PESQ',
  protectDesc: 'LOCK_E1_HIDE_COLS'
};

/**
 * Lê o comando inserido na célula E1 e oculta as colunas correspondentes.
 * Aceita nomes de lojas, intervalos de letras (ex: F-X) e exclusões (ex: !Assaí).
 * Após ocultar, protege a célula E1.
 * @returns {void}
 */
function ocultarColunasPorE1() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(HIDE_CFG.sheetName());
  if (!sh) return ss.toast(`Não encontrei "${HIDE_CFG.sheetName()}".`, 'Ocultar', 6);

  const cmd = String(sh.getRange(HIDE_CFG.inputA1).getValue() || '').trim();
  if (!cmd) {
    return ss.toast(
      'E1 vazio. Exemplos: "Assaí", "J-L", "!Tauste", "!K-L", "Assaí, !Carrefour".',
      'Ocultar',
      8
    );
  }

  // Reverte ocultações anteriores antes de aplicar novo comando
  mostrarColunasOcultas_(true);

  const store = CFG.getStoreInfo(sh, HIDE_CFG.headerRow());
  const storeEndCol = store.storeEnd;

  const lojas = sh.getRange(
    HIDE_CFG.headerRow(),
    HIDE_CFG.lojasStartCol(),
    1,
    storeEndCol - HIDE_CFG.lojasStartCol() + 1
  ).getValues()[0];

  const lojaToCol = buildLojaMap_(lojas, HIDE_CFG.lojasStartCol());
  const colToName = buildColToHeaderMap_(lojas, HIDE_CFG.lojasStartCol());

  // sem ! = base visível (mostrar somente)
  // com ! = exclusões aplicadas depois
  const colsToShow = new Set();
  const colsToExclude = new Set();

  const parts = cmd.split(/[,;]+/).map(s => s.trim()).filter(Boolean);
  const notFound = [];

  for (const raw of parts) {
    const isExclude = raw.includes('!');
    const p = raw.replace(/!/g, '').trim(); // aceita !J, !K-L, !K-!L, !Assaí

    if (!p) continue;
    const targetSet = isExclude ? colsToExclude : colsToShow;

    // Intervalo por letras (ex.: F-X ou F:X)
    const mRange = p.match(/^([A-Za-z]{1,3})\s*[-:]\s*([A-Za-z]{1,3})$/);
    if (mRange) {
      const c1 = letterToCol_(mRange[1]);
      const c2 = letterToCol_(mRange[2]);
      addRangeToSet_(targetSet, Math.min(c1, c2), Math.max(c1, c2));
      continue;
    }

    // Coluna única por letra (ex.: J)
    const mSingle = p.match(/^([A-Za-z]{1,3})$/);
    if (mSingle) {
      targetSet.add(letterToCol_(mSingle[1]));
      continue;
    }

    // Intervalo por nomes de loja (ex.: Assaí-Atacadão)
    if (p.includes('-')) {
      const [a, b] = p.split('-').map(x => x.trim());
      const ca = lojaToCol.get(normKey_(a));
      const cb = lojaToCol.get(normKey_(b));
      if (!ca || !cb) {
        notFound.push(raw);
        continue;
      }
      addRangeToSet_(targetSet, Math.min(ca, cb), Math.max(ca, cb));
      continue;
    }

    // Loja única por nome
    const c = lojaToCol.get(normKey_(p));
    if (!c) notFound.push(raw);
    else targetSet.add(c);
  }

  const startCol = HIDE_CFG.lojasStartCol();

  // Todas as colunas válidas de loja
  const allStoreCols = [];
  for (let c = startCol; c <= storeEndCol; c++) allStoreCols.push(c);

  // Limita ao intervalo válido
  const showCols = Array.from(colsToShow).filter(c => c >= startCol && c <= storeEndCol);
  const excludeCols = Array.from(colsToExclude).filter(c => c >= startCol && c <= storeEndCol);

  // Se não houver itens "normais", assume todas visíveis
  const baseVisible = showCols.length ? new Set(showCols) : new Set(allStoreCols);

  // Aplica exclusões
  for (const c of excludeCols) baseVisible.delete(c);

  const visibleCols = Array.from(baseVisible).sort((a, b) => a - b);

  const onlyExcludeMode = (showCols.length === 0 && excludeCols.length > 0);
  const excludeNames = formatColsAsNames_(excludeCols.slice().sort((a, b) => a - b), colToName);
  const shownNames = formatColsAsNames_(visibleCols, colToName);
  const msgNF = notFound.length ? ` | Não encontrei: ${notFound.join(', ')}` : '';

  if (visibleCols.length === 0) {
    return ss.toast(
      `Nada visível após aplicar o comando. Exemplos: "Assaí", "J-L", "!Tauste", "!K-L".${msgNF}`,
      'Mostrar somente',
      8
    );
  }

  // Oculta tudo que não ficou visível
  const visibleSet = new Set(visibleCols);
  const colsToHide = allStoreCols.filter(c => !visibleSet.has(c));

  // Se não há nada para ocultar, não salva estado/proteção
  if (colsToHide.length === 0) {
    if (onlyExcludeMode) {
      return ss.toast(`Exceto ${excludeNames}${msgNF}`, 'Mostrar somente', 8);
    }

    if (excludeCols.length > 0) {
      return ss.toast(
        `Nenhuma coluna ocultada. Visíveis: ${shownNames} | Exceto: ${excludeNames}${msgNF}`,
        'Mostrar somente',
        8
      );
    }

    return ss.toast(`Nenhuma coluna ocultada. Visíveis: ${shownNames}${msgNF}`, 'Mostrar somente', 8);
  }

  const ranges = mergeConsecutive_(colsToHide);

  for (const r of ranges) {
    sh.hideColumns(r.start, r.len);
  }

  PropertiesService.getDocumentProperties().setProperty(
    HIDE_CFG.stateKey,
    JSON.stringify({ sheetId: sh.getSheetId(), ranges })
  );

  protectE1_(sh);

  if (onlyExcludeMode) {
    return ss.toast(`Exceto ${excludeNames}${msgNF}`, 'Mostrar somente', 8);
  }

  if (excludeCols.length > 0) {
    return ss.toast(
      `Mostrando somente ${visibleCols.length} colunas: ${shownNames} | Exceto: ${excludeNames}${msgNF}`,
      'Mostrar somente',
      8
    );
  }

  ss.toast(`Mostrando somente ${visibleCols.length} colunas: ${shownNames}${msgNF}`, 'Mostrar somente', 8);
}

/**
 * Função pública para restaurar a visibilidade de todas as colunas.
 * Desbloqueia a célula E1 e limpa seu conteúdo.
 * @returns {void}
 */
function mostrarColunasOcultas() {
  mostrarColunasOcultas_(false);
}

/**
 * Lógica interna de restauração de colunas. Pode operar em modo silencioso
 * para não disparar toasts desnecessários ao ser chamada por outras funções.
 * @private
 * @param {boolean} silent - Define se a função emitirá avisos (toasts).
 * @returns {void}
 */
function mostrarColunasOcultas_(silent) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(HIDE_CFG.sheetName());
  if (!sh) return;

  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty(HIDE_CFG.stateKey);

  const store = CFG.getStoreInfo(sh, HIDE_CFG.headerRow());
  const storeEndCol = store.storeEnd;

  const lojas = sh.getRange(
    HIDE_CFG.headerRow(),
    HIDE_CFG.lojasStartCol(),
    1,
    storeEndCol - HIDE_CFG.lojasStartCol() + 1
  ).getValues()[0];

  const colToName = buildColToHeaderMap_(lojas, HIDE_CFG.lojasStartCol());
  const startCol = HIDE_CFG.lojasStartCol();

  // Fallback: se não tiver estado salvo, tenta abrir colunas ocultas por usuário no range de lojas
  if (!raw) {
    const hiddenCols = [];
    for (let c = startCol; c <= storeEndCol; c++) {
      if (sh.isColumnHiddenByUser(c)) hiddenCols.push(c);
    }

    if (hiddenCols.length === 0) {
      unprotectE1_(sh);
      if (!silent) ss.toast('Nada para mostrar.', 'Mostrar', 4);
      return;
    }

    const ranges = mergeConsecutive_(hiddenCols);
    for (const r of ranges) {
      sh.showColumns(r.start, r.len);
    }

    unprotectE1_(sh);
    sh.getRange(HIDE_CFG.inputA1).clearContent();

    if (!silent) {
      const shownNames = formatColsAsNames_(hiddenCols, colToName);
      ss.toast(`Mostradas ${hiddenCols.length} colunas: ${shownNames}`, 'Mostrar', 8);
    }
    return;
  }

  let state;
  try {
    state = JSON.parse(raw);
  } catch (e) {
    state = null;
  }

  if (!state || !state.ranges) {
    props.deleteProperty(HIDE_CFG.stateKey);
    unprotectE1_(sh);
    if (!silent) ss.toast('Estado inválido. Removi o bloqueio do E1.', 'Mostrar', 6);
    return;
  }

  const colsShown = expandRangesToCols_(state.ranges).filter(c => c >= startCol && c <= storeEndCol);
  const shownNames = formatColsAsNames_(colsShown, colToName);

  for (const r of state.ranges) {
    sh.showColumns(r.start, r.len);
  }

  // Só apaga o estado depois de mostrar com sucesso
  props.deleteProperty(HIDE_CFG.stateKey);

  unprotectE1_(sh);
  sh.getRange(HIDE_CFG.inputA1).clearContent();

  if (!silent) ss.toast(`Mostradas ${colsShown.length} colunas: ${shownNames}`, 'Mostrar', 8);
}

/*************** HELPERS ***************/

/**
 * Cria um mapa vinculando a chave normalizada do nome da loja à sua coluna.
 * @private
 */
function buildLojaMap_(lojasRow, startCol) {
  const map = new Map();
  for (let i = 0; i < lojasRow.length; i++) {
    const name = String(lojasRow[i] || '').trim();
    if (!name) continue;
    const key = normKey_(name);
    if (!map.has(key)) map.set(key, startCol + i);
  }
  return map;
}

/**
 * Cria um mapa vinculando o índice da coluna ao nome exato do cabeçalho da loja.
 * @private
 */
function buildColToHeaderMap_(headersRow, startCol) {
  const map = new Map();
  for (let i = 0; i < headersRow.length; i++) {
    const col = startCol + i;
    const name = String(headersRow[i] || '').trim();
    if (name) map.set(col, name);
  }
  return map;
}

/**
 * Normaliza uma string de loja para criação de chaves exclusivas.
 * @private
 */
function normKey_(s) {
  return String(s || '')
    .trim()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '');
}

/**
 * Converte a notação de letra de coluna (ex: "A", "AA") para número de índice (1-based).
 * @private
 */
function letterToCol_(letters) {
  const s = String(letters || '').toUpperCase().replace(/[^A-Z]/g, '');
  let n = 0;
  for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n;
}

/**
 * Adiciona um intervalo numérico [a, b] dentro de um Set.
 * @private
 */
function addRangeToSet_(set, a, b) {
  for (let c = a; c <= b; c++) set.add(c);
}

/**
 * Agrupa arrays de números de colunas soltas em intervalos consecutivos (ranges)
 * otimizando a velocidade das funções showColumns/hideColumns.
 * @private
 */
function mergeConsecutive_(colsSorted) {
  const ranges = [];
  let start = colsSorted[0];
  let prev = colsSorted[0];

  for (let i = 1; i < colsSorted.length; i++) {
    const c = colsSorted[i];
    if (c === prev + 1) prev = c;
    else {
      ranges.push({ start, len: prev - start + 1 });
      start = prev = c;
    }
  }
  ranges.push({ start, len: prev - start + 1 });
  return ranges;
}

/**
 * Desempacota objetos de range de volta para arrays soltos com os índices das colunas.
 * @private
 */
function expandRangesToCols_(ranges) {
  const out = [];
  for (const r of ranges) for (let c = r.start; c < r.start + r.len; c++) out.push(c);
  return out;
}

/**
 * Converte um índice numérico de coluna para notação de letra (ex: 1 -> A).
 * @private
 */
function colToLetter_(n) {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/**
 * Formata um array de colunas para uma string amigável de nomes, truncando se for muito longo.
 * @private
 */
function formatColsAsNames_(cols, colToNameMap) {
  const names = cols.map(c => colToNameMap.get(c) || colToLetter_(c));
  const max = 8;
  if (names.length <= max) return names.join(', ');
  return names.slice(0, max).join(', ') + ` (+${names.length - max})`;
}

/**
 * Protege a célula E1 impedindo edições por outros usuários do domínio.
 * Utilizado para travar a célula enquanto o filtro de colunas estiver ativo.
 * @private
 */
function protectE1_(sh) {
  unprotectE1_(sh);

  const r = sh.getRange(HIDE_CFG.inputA1);
  const p = r.protect().setDescription(HIDE_CFG.protectDesc);
  p.setWarningOnly(false);

  try {
    const editors = p.getEditors();
    if (editors && editors.length) p.removeEditors(editors);
    if (p.canDomainEdit()) p.setDomainEdit(false);
  } catch (e) {}
}

/**
 * Remove a proteção da célula E1 (identificada pela descrição padrão).
 * @private
 */
function unprotectE1_(sh) {
  const prots = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (const p of prots) {
    if (p.getDescription() === HIDE_CFG.protectDesc) {
      try { p.remove(); } catch (e) {}
    }
  }
}