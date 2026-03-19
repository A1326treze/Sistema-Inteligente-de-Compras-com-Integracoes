/* ===== Verificar Duplicados (Histórico) — v2 (robusto + CFG) ===== */

/**
 * @fileoverview Auditoria de Duplicatas no Histórico.
 * Analisa a aba de Histórico em busca de registros repetidos baseados na chave
 * Data + Item + Marca + Loja + Tipo + Quantidade. Ignora linhas marcadas como "Travar".
 * Gera um relatório em uma aba separada detalhando os conflitos de preço e repetições idênticas.
 */

/**
 * Função principal que executa a varredura de duplicidades no Histórico.
 * Cria a aba "Repetidos (Histórico)" com o relatório de inconsistências, caso existam.
 * @returns {void}
 */
function verificarRepeticoesHistoricoTudo() {
  const ss = SpreadsheetApp.getActive();
  const histName = (typeof CFG !== 'undefined' && CFG.SHEETS && CFG.SHEETS.HIST) ? CFG.SHEETS.HIST : 'Histórico';

  const shH = ss.getSheetByName(histName);
  if (!shH) return ss.toast(`Não encontrei a aba "${histName}".`, 'Repetidos', 8);

  const reportName = 'Repetidos (Histórico)';
  const shReport = ss.getSheetByName(reportName);

  const { firstDataRowHist } = VDH_findHistHeader_(shH);
  const lr = shH.getLastRow();
  
  // Se não há dados, exclui relatórios antigos e encerra
  if (lr < firstDataRowHist) {
    if (shReport) ss.deleteSheet(shReport);
    return ss.toast('Histórico vazio. Nada para verificar.', 'Repetidos', 8);
  }

  // A:H (inclui Travar em H)
  const vals = shH.getRange(firstDataRowHist, 1, lr - firstDataRowHist + 1, 8).getValues();

  // key -> { dateKey, item, marca, loja, tipo, qtdKey, qtdDisplay, prices[], rows[] }
  const map = new Map();

  for (let i = 0; i < vals.length; i++) {
    const rowNum = firstDataRowHist + i;

    const d     = vals[i][0]; // A Data
    const item  = vals[i][1]; // B Item
    const marca = vals[i][2]; // C Marca
    const loja  = vals[i][3]; // D Loja
    const tipo  = vals[i][4]; // E Tipo
    const qtd   = vals[i][5]; // F Quantidade
    const preco = vals[i][6]; // G Preço
    const travar= vals[i][7]; // H Travar

    // Ignora linhas travadas
    if (travar === true) continue;

    const itemS  = VDH_cleanDisplay_(item);
    const marcaS = VDH_cleanDisplay_(marca);
    const lojaS  = VDH_cleanDisplay_(loja);
    const tipoS  = VDH_cleanDisplay_(tipo);

    if (!d || !itemS || !lojaS || !tipoS) continue;
    if (qtd === '' || qtd === null || typeof qtd === 'undefined') continue;

    // Preço: usa CFG.toNum (se existir), com fallback robusto
    let priceNum = NaN;
    if (typeof CFG !== 'undefined' && typeof CFG.toNum === 'function') {
      priceNum = CFG.toNum(preco);
    }
    if (!isFinite(priceNum)) priceNum = VDH_toNumFallback_(preco);
    if (!isFinite(priceNum) || priceNum <= 0) continue;

    const dk = VDH_dateToKey_(d);
    if (!dk) continue;

    const qtdKey = VDH_normQtdKey_(qtd);
    if (!qtdKey) continue;

    // Montagem da Chave de Identificação
    const key = [
      dk,
      VDH_normTextKey_(itemS),
      VDH_normTextKey_(marcaS),
      VDH_normTextKey_(lojaS),
      VDH_normTextKey_(tipoS),
      qtdKey
    ].join('|');

    if (!map.has(key)) {
      map.set(key, {
        dateKey: dk,
        item: itemS,
        marca: marcaS,
        loja: lojaS,
        tipo: tipoS,
        qtdKey,
        qtdDisplay: VDH_qtdDisplay_(qtdKey),
        prices: [],
        rows: []
      });
    }

    const obj = map.get(key);
    obj.prices.push(VDH_round2_(priceNum));
    obj.rows.push(rowNum);
  }

  const repeats = [];
  let conflitos = 0;
  let identicos = 0;

  // Analisa o mapa em busca de chaves com mais de uma entrada (rows.length > 1)
  for (const obj of map.values()) {
    if (obj.rows.length > 1) {
      const distinct = Array.from(new Set(obj.prices)).sort((a, b) => a - b);
      const kind = (distinct.length > 1) ? 'Conflito de preço' : 'Repetição idêntica';
      if (distinct.length > 1) conflitos++; else identicos++;
      repeats.push({ ...obj, distinct, kind });
    }
  }

  // Se estiver tudo limpo, apaga relatórios residuais e avisa
  if (repeats.length === 0) {
    if (shReport) ss.deleteSheet(shReport);
    return ss.toast('Nenhuma repetição encontrada no Histórico inteiro.', 'Repetidos', 10);
  }

  // Cria ou limpa a aba de relatório
  const shD = shReport || ss.insertSheet(reportName);
  shD.clear();

  // Cabeçalho do relatório
  shD.getRange(1, 1, 1, 11).setValues([[
    'Tipo', 'Data', 'Item', 'Marca', 'Loja', 'Tipo(un)', 'Quantidade',
    'Preços (distintos)', 'Qtd linhas', 'Linhas no Histórico', 'Obs'
  ]]);

  // Prepara os dados de saída
  const out = repeats.map(o => ([
    o.kind,
    VDH_formatBR_(o.dateKey),
    o.item,
    o.marca,
    o.loja,
    o.tipo,
    o.qtdDisplay,
    o.distinct.map(p => `R$ ${VDH_money_(p)}`).join(' | '),
    o.rows.length,
    o.rows.join(', '),
    'Se for erro: corrija/apague no Histórico e rode de novo'
  ]));

  shD.getRange(2, 1, out.length, 11).setValues(out);
  try { shD.autoResizeColumns(1, 11); } catch (e) {}

  ss.toast(
    `Pente fino: ${repeats.length} chaves repetidas — Conflitos: ${conflitos} | Idênticos: ${identicos}. Veja "${reportName}".`,
    'Repetidos',
    12
  );
}

/* ========= Helpers ========= */

/**
 * Localiza dinamicamente a linha de cabeçalho da aba Histórico buscando a palavra "data".
 * @private
 * @param {GoogleAppsScript.Spreadsheet.Sheet} shH - Aba do histórico.
 * @returns {Object} Números das linhas de cabeçalho e de dados.
 */
function VDH_findHistHeader_(shH) {
  const defaultHeader = (typeof CFG !== 'undefined' && CFG.ROWS && CFG.ROWS.HEADER) ? CFG.ROWS.HEADER : 2;
  const maxScan = Math.min(8, shH.getLastRow() || 8);
  const colA = shH.getRange(1, 1, maxScan, 1).getValues()
    .map(r => String(r[0] || '').trim().toLowerCase());

  let headerRow = null;
  for (let i = 0; i < colA.length; i++) {
    if (colA[i] === 'data') { headerRow = i + 1; break; }
  }
  if (!headerRow) headerRow = defaultHeader;
  return { headerRowHist: headerRow, firstDataRowHist: headerRow + 1 };
}

/**
 * Mantém o texto "bonito" para exibição no relatório, removendo apenas lixo invisível.
 * @private
 * @param {any} v - Valor a ser limpo.
 * @returns {string}
 */
// mantém o texto “bonito” pro relatório, só limpando lixo invisível
function VDH_cleanDisplay_(v) {
  return String(v ?? '')
    .replace(/[\u200B-\u200D\uFEFF]/g, '') // zero-width
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Normaliza o texto de forma forte para ser usado como chave de comparação (sem acentos/minúsculo).
 * @private
 * @param {any} s - Texto a ser normalizado.
 * @returns {string}
 */
// chave forte: sem invisíveis + sem acento + lower
function VDH_normTextKey_(s) {
  return String(s ?? '')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

/**
 * Arredonda um número para duas casas decimais.
 * @private
 * @param {number} n - Número alvo.
 * @returns {number}
 */
function VDH_round2_(n) { return Math.round(n * 100) / 100; }

/**
 * Converte entradas de data (Date object ou string) para uma chave no formato YYYY-MM-DD.
 * @private
 * @param {Date|string} v - Data alvo.
 * @returns {string}
 */
function VDH_dateToKey_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    const y = v.getFullYear();
    const m = String(v.getMonth() + 1).padStart(2, '0');
    const d = String(v.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  const s = String(v || '').trim();
  if (!s) return '';
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${String(+m[2]).padStart(2,'0')}-${String(+m[3]).padStart(2,'0')}`;
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return `${m[3]}-${String(+m[2]).padStart(2,'0')}-${String(+m[1]).padStart(2,'0')}`;
  const d2 = new Date(s);
  if (!isNaN(d2.getTime())) {
    const y = d2.getFullYear();
    const mm = String(d2.getMonth() + 1).padStart(2,'0');
    const dd = String(d2.getDate()).padStart(2,'0');
    return `${y}-${mm}-${dd}`;
  }
  return '';
}

/**
 * Formata uma data YYYY-MM-DD de volta para o padrão brasileiro DD/MM/YYYY.
 * @private
 * @param {string} dateKey - Chave de data.
 * @returns {string}
 */
function VDH_formatBR_(dateKey) {
  const [y, m, d] = String(dateKey).split('-');
  return `${d}/${m}/${y}`;
}

/**
 * Formata um número para o padrão monetário brasileiro.
 * @private
 * @param {number} n - Valor.
 * @returns {string}
 */
function VDH_money_(n) {
  const s = (Math.round(n * 100) / 100).toFixed(2);
  return s.replace('.', ',');
}

/**
 * Fallback para conversão de strings monetárias problemáticas em Float.
 * @private
 * @param {any} v - Valor bruto da célula de preço.
 * @returns {number}
 */
// fallback caso algum preço venha fora do padrão esperado
function VDH_toNumFallback_(v) {
  if (typeof v === 'number') return v;
  let s = String(v ?? '').trim();
  if (!s) return NaN;
  s = s
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\s/g, '')
    .replace(/^R\$/i, '');
  if (s.includes(',') && s.includes('.')) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (s.includes(',')) {
    s = s.replace(',', '.');
  }
  const n = Number(s);
  return isFinite(n) ? n : NaN;
}

/**
 * Gera uma chave de quantidade padronizada (Ex: " 100g" -> "100").
 * Suporta formatos numéricos e textos BR.
 * @private
 * @param {any} v - Valor bruto da quantidade.
 * @returns {string}
 */
// Quantidade -> chave canônica (string BR):
// - " 100", "100 " -> "100"
// - "0,4" / "0.4"  -> "0,4"
// - "100g" -> "100"
function VDH_normQtdKey_(v) {
  if (typeof v === 'number' && isFinite(v)) {
    return VDH_numToQtdKey_(v);
  }

  let s = String(v ?? '').replace(/[\u200B-\u200D\uFEFF]/g, '').trim();
  if (!s) return '';

  const m = s.match(/-?\d+(?:[.,]\d+)?/);
  if (!m) return '';

  let numStr = m[0];

  if (numStr.includes(',') && numStr.includes('.')) {
    numStr = numStr.replace(/\./g, '').replace(',', '.');
  } else if (numStr.includes(',')) {
    numStr = numStr.replace(',', '.');
  }
  const n = Number(numStr);
  if (!isFinite(n)) return '';

  return VDH_numToQtdKey_(n);
}

/**
 * Converte um número limpo para uma string de quantidade formatada (chave).
 * @private
 * @param {number} n - Número da quantidade.
 * @returns {string}
 */
function VDH_numToQtdKey_(n) {
  // inteiro
  if (Math.abs(n - Math.round(n)) < 1e-9) return String(Math.round(n));

  // decimal: até 6 casas, corta zeros, usa vírgula
  let s = n.toFixed(6);
  s = s.replace(/0+$/g, '').replace(/\.$/g, '');
  s = s.replace('.', ',');
  return s;
}

/**
 * Limpa a chave de quantidade para exibição no relatório.
 * @private
 * @param {string} qtdKey - A chave gerada.
 * @returns {string}
 */
function VDH_qtdDisplay_(qtdKey) {
  return String(qtdKey ?? '').trim();
}