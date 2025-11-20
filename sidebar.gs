function UpdateSideBar(msg) {
  const ui = SpreadsheetApp.getUi();
  ui.showSidebar(HtmlService.createHtmlOutput(`<p>${msg}</p>`).setTitle('Status'));
}