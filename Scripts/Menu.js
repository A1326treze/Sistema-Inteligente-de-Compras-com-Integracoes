/**
 * @fileoverview Criação de Menus Customizados e Gatilhos de Abertura.
 * Este script é executado automaticamente ao abrir a planilha e monta
 * a interface de botões na barra de menus superior do Google Sheets.
 */

/**
 * Gatilho simples (Simple Trigger) acionado nativamente ao abrir o documento.
 * Constrói os menus dropdown "Planilha", "Pesquisa" e "Histórico", 
 * mapeando os opções de interface para as suas respectivas funções.
 * @returns {void}
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Planilha')
    .addItem('Abrir Topbar', 'abrirTopbar')
    .addItem('Topbar: Auto-abrir (alternar)', 'toggleTopbarAutoOpen')
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi()
      .createMenu('Gráfico')
        .addItem('Abrir Dashboard (Looker)', 'abrirDashboard')
        .addItem('Definir link do Dashboard', 'definirDashboardUrl')
    )
    .addToUi();
  
  ui.createMenu('Pesquisa')
    .addItem('Atualizar gradiente (tudo)', 'atualizarGradienteTudo')
    .addItem('Aplicar separadores por Item (manual)', 'aplicarSeparadoresItens')
    .addSeparator()
    .addItem('Sincronizar AppSheet (App ⇄ PC)', 'syncPcEApp')
    .addToUi();

  ui.createMenu('Histórico')
    .addItem('Sincronizar preços do dia (limpar ausentes)', 'sincronizarPrecosDoDia')
    .addItem('Refazer última ação', 'refazerUltimaAcaoHistorico')
    .addSeparator()
    .addItem('Verificar duplicados', 'verificarRepeticoesHistoricoTudo')
    .addToUi();
}

/**
 * Gatilho instalável (Installable Trigger) acionado ao abrir a planilha.
 * Utilizado para rodar funções que carregam interfaces HTML (como a Topbar) 
 * automaticamente logo após o carregamento da planilha.
 * @param {GoogleAppsScript.Events.SheetsOnOpen} e - Objeto do evento de abertura.
 * @returns {void}
 */
function onOpenInstalavel(e) {
  // só abre se o auto-abrir estiver ligado
  maybeAutoOpenTopbar_();
}