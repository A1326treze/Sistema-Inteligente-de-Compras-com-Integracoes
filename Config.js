/**
 * @fileoverview Arquivo de Configuração Global (Single Source of Truth).
 * Centraliza todas as referências estruturais do projeto (nomes de abas, índices 
 * de linhas/colunas, células fixas) e funções utilitárias compartilhadas.
 * O objeto `CFG` é exposto globalmente para ser consumido pelos demais módulos.
 */

/* ===== Config.js — Fonte única de verdade (v1) =====
 * Centraliza nomes de abas, linhas, colunas e detecção dinâmica das lojas.
 * Objetivo: scripts mais escaláveis e fáceis de manter.
 * NÃO muda comportamento por si só; os outros arquivos passam a consultar este CFG.
 */

/**
 * Objeto global de configuração da planilha.
 * Construído através de uma IIFE (Immediately Invoked Function Expression).
 * @constant {Object}
 */
const CFG = (() => {
  /**
   * Mapeamento dos nomes das abas (Sheets).
   * @constant {Object}
   */
  // ===== Abas =====
  const SHEETS = {
    INV: 'Inventário',
    PESQ: 'Preços (pesquisa)',
    CALC: 'Preços (calculos)',
    HIST: 'Histórico'
  };

  /**
   * Definição das linhas padrão (Cabeçalhos e Primeira linha de dados).
   * @constant {Object}
   */
  // ===== Linhas padrão =====
  const ROWS = {
    HEADER: 2,
    FIRST_DATA: 3
  };

  /**
   * Mapeamento dos índices numéricos das colunas (1-based) para cada aba.
   * @constant {Object}
   */
  // ===== Colunas (1-based) =====
  const COL = {
    INV: {
      PRIORIDADE: 1,
      ORDEM: 2,
      ITEM: 3,
      TIPO: 4,
      SETOR: 5,
      ESTOQUE: 6,
      DESEJADO: 7,
      STATUS: 8,
      NECESSARIO: 9,
      UN: 10,
      QNTD: 11,
      UNITARIO: 12,
      TOTAL: 13,
      MELHOR_LOJA: 14,
      MELHOR_MARCA: 15,
      BRENCH_UN: 16,
      BRENCH_QNTD: 17,
      BRENCH_UNITARIO: 18,
      BRENCH_TOTAL: 19,
      BRENCH_LOJA: 20,
      BRENCH_MARCA: 21,
      LOCAL: 22
    },

    PESQ: {
      CHECK: 1,   // A
      ITEM: 2,    // B
      MARCA: 3,   // C
      QTD: 4,     // D
      TIPO: 5,    // E
      LOJA_START: 6 // F (primeira loja)
      // LOJA_END é dinâmico (antes de Custo/Mult)
    },

    HIST: {
      DATA: 1,
      ITEM: 2,
      MARCA: 3,
      LOJA: 4,
      TIPO: 5,
      QTD: 6,
      PRECO: 7,
      TRAVAR: 8,
      TIMESTAMP: 11 // K ("Salvo em")
    }
  };

  /**
   * Mapeamento de referências exatas de células (Notação A1).
   * @constant {Object}
   */
  // ===== Células fixas =====
  const CELLS = {
    INV: {
      FILTER_SETOR: 'D1',
      FILTER_SETOR_INVERT: 'E1' // checkbox
    },
    PESQ: { DATA_PESQUISA: 'B1' }
  };


  /**
   * Mapeamento de palavras-chave usadas para localizar colunas limitadoras.
   * @constant {Object}
   */
  // ===== Cabeçalhos auxiliares (fim das lojas) =====
  const HDR = {
    CUSTO: 'custo',
    MULT: 'mult'
  };

  // ===== Util =====
  
  /**
   * Remove espaços nas extremidades de uma string.
   * @private
   * @param {any} s - Valor a ser normalizado.
   * @returns {string} String limpa.
   */
  function norm_(s) {
    return String(s || '').trim();
  }

  /**
   * Normaliza textos de cabeçalho para comparação (minúsculo, sem acento, sem espaços extras).
   * @private
   * @param {any} s - Texto do cabeçalho.
   * @returns {string} Texto padronizado.
   */
  function normHeader_(s) {
    // minúsculo + remove acentos + tira espaços duplicados
    return norm_(s)
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/\s+/g, ' ');
  }

  /**
   * Retorna os valores da linha de cabeçalho de uma planilha.
   * @private
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - Aba alvo.
   * @param {number} [headerRow=ROWS.HEADER] - Linha do cabeçalho.
   * @returns {Array} Array contendo os valores do cabeçalho.
   */
  function getHeaderRowValues_(sh, headerRow = ROWS.HEADER) {
    const lastCol = sh.getLastColumn();
    if (lastCol < 1) return [];
    return sh.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];
  }

  /**
   * Busca e retorna o índice numérico (1-based) de uma coluna com base no nome do cabeçalho.
   * @private
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sh - Aba alvo.
   * @param {string} headerName - Nome da coluna procurada.
   * @param {number} [headerRow=ROWS.HEADER] - Linha do cabeçalho.
   * @returns {number} O índice numérico da coluna, ou -1 se não encontrar.
   */
  function findHeaderCol_(sh, headerName, headerRow = ROWS.HEADER) {
    const target = normHeader_(headerName);
    const headers = getHeaderRowValues_(sh, headerRow).map(normHeader_);
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] === target) return i + 1; // 1-based
    }
    return -1;
  }

  /**
   * Detecta dinamicamente as colunas de lojas na aba "Preços (pesquisa)".
   * Regra: lojas começam em F (6) e terminam na coluna imediatamente antes 
   * do primeiro cabeçalho "Custo" ou "Mult" (tolerante a acentos/maiúsculas).
   * @private
   * @param {GoogleAppsScript.Spreadsheet.Sheet} shPesquisa - Aba de pesquisa.
   * @param {number} [headerRow=ROWS.HEADER] - Linha de referência.
   * @returns {Object} { storeStart, storeEnd, storeCols, colCusto, colMult }
   * @throws {Error} Se não existirem colunas suficientes.
   */
  /**
   * Detecta dinamicamente colunas de lojas na aba "Preços (pesquisa)".
   * Regra: lojas começam em F (6) e terminam na coluna imediatamente antes do primeiro cabeçalho
   * "Custo" ou "Mult" (tolerante a acentos/maiúsculas).
   *
   * Retorna:
   * { storeStart, storeEnd, storeCols, colCusto, colMult }
   */
  function getStoreInfo_(shPesquisa, headerRow = ROWS.HEADER) {
    const lastCol = shPesquisa.getLastColumn();
    if (lastCol < COL.PESQ.LOJA_START) {
      throw new Error('CFG.getStoreInfo: não há colunas suficientes para lojas.');
    }

    const headers = getHeaderRowValues_(shPesquisa, headerRow).map(normHeader_);

    let colCusto = -1;
    let colMult = -1;

    for (let i = 0; i < headers.length; i++) {
      if (headers[i] === HDR.CUSTO) colCusto = i + 1;
      if (headers[i] === HDR.MULT) colMult = i + 1;
    }

    const storeStart = COL.PESQ.LOJA_START;

    // fim das lojas: antes do primeiro (Custo ou Mult) que apareça depois do start
    let storeEnd = lastCol;

    const candidates = [colCusto, colMult].filter(c => c > storeStart);
    if (candidates.length > 0) storeEnd = Math.min(...candidates) - 1;

    if (storeEnd < storeStart) {
      throw new Error(`CFG.getStoreInfo: storeEnd < storeStart (${storeEnd} < ${storeStart}).`);
    }

    return {
      storeStart,
      storeEnd,
      storeCols: storeEnd - storeStart + 1,
      colCusto,
      colMult
    };
  }

  /**
   * Retorna o objeto Sheet pelo nome fornecido ou lança um erro para interromper execuções falhas.
   * @private
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - Objeto Spreadsheet ativo.
   * @param {string} sheetName - Nome exato da aba.
   * @returns {GoogleAppsScript.Spreadsheet.Sheet}
   * @throws {Error} Se a aba não for encontrada.
   */
  /**
   * Retorna o sheet pelo nome do CFG ou lança erro.
   */
  function mustGetSheet_(ss, sheetName) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) throw new Error(`Não encontrei a aba "${sheetName}".`);
    return sh;
  }

  /**
   * Converte texto numérico monetário (ex.: "R$ 5,75") em número Float.
   * @private
   * @param {any} v - Valor bruto a ser convertido.
   * @returns {number} Valor em float ou NaN.
   */
  /**
   * Converte texto numérico BRL (ex.: "R$ 5,75") em número.
   */
  function toNum_(v) {
    if (typeof v === 'number') return v;
    const s = norm_(v);
    if (!s) return NaN;
    const cleaned = s.replace(/\s/g, '').replace(/^R\$/i, '').replace(/\./g, '').replace(',', '.');
    const n = Number(cleaned);
    return isFinite(n) ? n : NaN;
  }

  // Objeto CFG retornado e acessível globalmente
  return {
    SHEETS,
    ROWS,
    COL,
    CELLS,
    HDR,

    norm: norm_,
    normHeader: normHeader_,
    getHeaderRowValues: getHeaderRowValues_,
    findHeaderCol: findHeaderCol_,
    getStoreInfo: getStoreInfo_,
    mustGetSheet: mustGetSheet_,
    toNum: toNum_
  };
})();