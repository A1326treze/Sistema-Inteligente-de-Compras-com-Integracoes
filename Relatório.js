/**
 * @fileoverview Gerenciador de Atalho para Dashboard Externo.
 * Permite que o usuário salve a URL de um relatório (ex: Looker Studio) nas propriedades 
 * do documento e o abra em uma nova guia diretamente pela interface do Google Sheets.
 */

// ===== Abrir Dashboard (Looker) em nova guia =====

/**
 * Chave utilizada para armazenar a URL nas propriedades do documento.
 * @constant {string}
 */
const DASH_URL_KEY = 'LOOKER_DASH_URL';

/**
 * Exibe um prompt de interface (UI) solicitando ao usuário que insira a URL do Dashboard.
 * Salva a URL informada no PropertiesService (nível de documento).
 * @returns {void}
 */
function definirDashboardUrl() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();
  const atual = props.getProperty(DASH_URL_KEY) || '';

  const res = ui.prompt(
    'Link do Dashboard (Looker Studio)',
    'Cole a URL completa do seu relatório.\n' +
    'Atual: ' + (atual ? atual : '(não definido)'),
    ui.ButtonSet.OK_CANCEL
  );

  if (res.getSelectedButton() !== ui.Button.OK) return;

  const url = (res.getResponseText() || '').trim();
  if (!url) return;

  props.setProperty(DASH_URL_KEY, url);
  SpreadsheetApp.getActive().toast('Link do Dashboard salvo.', 'Dashboard', 5);
}

/**
 * Resgata a URL salva nas propriedades e a abre em uma nova guia do navegador.
 * Utiliza um dialog HTML temporário e sem foco (modeless) para contornar bloqueios 
 * nativos de pop-up e acionar a função window.open do lado do cliente (browser).
 * @returns {void}
 */
function abrirDashboard() {
  const props = PropertiesService.getDocumentProperties();
  const url = (props.getProperty(DASH_URL_KEY) || '').trim();

  if (!url) {
    SpreadsheetApp.getUi().alert(
      'Ainda não tem link salvo.\n' +
      'Use: Planilha → Definir link do Dashboard'
    );
    return;
  }

  // Abre em nova guia e fecha o dialog
  const html = HtmlService.createHtmlOutput(
    `<script>
      window.open(${JSON.stringify(url)}, "_blank");
      google.script.host.close();
    </script>`
  ).setWidth(120).setHeight(60);

  SpreadsheetApp.getUi().showModelessDialog(html, 'Abrindo…');
}