/* ===== topbar.js (atualizado para Config.js / CFG) ===== */

/**
 * @fileoverview Gerenciador da Barra de Navegação Flutuante (Topbar).
 * Controla a exibição, preferências de abertura automática e navegação entre abas
 * através de uma interface HTML (Modeless Dialog) integrada ao Google Sheets.
 */

/**
 * Instância do PropertiesService para armazenar configurações a nível de documento.
 * @constant
 */
const TOPBAR_PROPS = PropertiesService.getDocumentProperties();

/** @constant {string} Chave para salvar a preferência de abertura automática. */
const TOPBAR_AUTO_KEY = 'TOPBAR_AUTO_OPEN';
/** @constant {string} Chave para salvar a largura customizada da topbar. */
const TOPBAR_W_KEY = 'TOPBAR_WIDTH';
/** @constant {string} Chave para salvar a altura customizada da topbar. */
const TOPBAR_H_KEY = 'TOPBAR_HEIGHT';

/**
 * Abre a interface HTML da Topbar em uma janela flutuante (Modeless Dialog).
 * O arquivo HTML deve se chamar 'Topbar.html'.
 * @returns {void}
 */
function abrirTopbar() {
  const html = HtmlService.createHtmlOutputFromFile('Topbar')
    .setWidth(480)
    .setHeight(45);

  SpreadsheetApp.getUi().showModelessDialog(html, ' ');
}

/**
 * Define explicitamente a preferência de abertura automática da Topbar.
 * @param {boolean} isOn - Verdadeiro para ativar, falso para desativar.
 * @returns {void}
 */
function setTopbarAutoOpen(isOn) {
  TOPBAR_PROPS.setProperty(TOPBAR_AUTO_KEY, String(!!isOn));
}

/**
 * Alterna (liga/desliga) a configuração de abertura automática da Topbar
 * e exibe uma notificação de confirmação ao usuário.
 * @returns {void}
 */
function toggleTopbarAutoOpen() {
  const cur = TOPBAR_PROPS.getProperty(TOPBAR_AUTO_KEY) === 'true';
  const next = !cur;
  TOPBAR_PROPS.setProperty(TOPBAR_AUTO_KEY, String(next));
  SpreadsheetApp.getActive().toast(next ? 'Topbar: auto-abrir ATIVADO' : 'Topbar: auto-abrir DESATIVADO');
  if (next) abrirTopbar();
}

/**
 * Verifica se a abertura automática está ativada e, se estiver, abre a Topbar.
 * Geralmente chamado por um gatilho onOpen().
 * @private
 * @returns {void}
 */
function maybeAutoOpenTopbar_() {
  const auto = TOPBAR_PROPS.getProperty(TOPBAR_AUTO_KEY) === 'true';
  if (auto) abrirTopbar();
}

/**
 * Recupera as configurações salvas da Topbar (estado automático e dimensões).
 * Caso as dimensões não estejam salvas, aplica valores padrão.
 * @returns {Object} Objeto contendo {autoOpen, width, height}.
 */
function getTopbarSettings() {
  return {
    autoOpen: TOPBAR_PROPS.getProperty(TOPBAR_AUTO_KEY) === 'true',
    width: Number(TOPBAR_PROPS.getProperty(TOPBAR_W_KEY)) || 50,
    height: Number(TOPBAR_PROPS.getProperty(TOPBAR_H_KEY)) || 54
  };
}

/**
 * Mapeia e retorna as abas disponíveis para exibição na Topbar e identifica a aba atual.
 * Baseia-se nas definições do objeto global `CFG` para encontrar as abas principais.
 * @returns {Object} Objeto contendo o array de abas filtradas e o nome da aba ativa.
 */
function getTopbarData() {
  const ss = SpreadsheetApp.getActive();
  const active = ss.getActiveSheet().getName();

  const desired = [
    { name: CFG.SHEETS.INV,  label: '🏠 Inventário' },
    { name: CFG.SHEETS.PESQ, label: '🛒 Preços' },
    { name: 'Gráfico',       label: '📊 Gráfico' },
    { name: CFG.SHEETS.HIST, label: '📈 Histórico' },
  ];

  const existing = new Set(ss.getSheets().map(s => s.getName()));
  let tabs = desired.filter(t => existing.has(t.name));

  // Fallback: se não encontrar nenhuma das abas desejadas, lista todas as abas existentes
  if (tabs.length === 0) {
    tabs = ss.getSheets().map(s => ({ name: s.getName(), label: s.getName() }));
  }

  return { tabs, active };
}

/**
 * Retorna o nome da aba (planilha) que o usuário está visualizando no momento.
 * @returns {string} Nome da aba ativa.
 */
function getActiveSheetName() {
  return SpreadsheetApp.getActive().getActiveSheet().getName();
}

/**
 * Navega o usuário para a aba especificada, tornando-a a aba ativa.
 * @param {string} name - O nome da aba de destino.
 * @returns {void}
 */
function goToSheet(name) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(name);
  if (sh) ss.setActiveSheet(sh);
}