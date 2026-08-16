function doGet(e) {
  const htmlServ= HtmlService.createTemplateFromFile("Planning");
  const html = htmlServ.evaluate();
//  html.setWidth(1200).setHeight(800);
//  const ui = SpreadsheetApp.getUi();
//  ui.showModalDialog(html, "Hotel Manager v0");
  return html;
}

function loadPartialHTML (partial) {
  const htmlServ= HtmlService.createTemplateFromFile(partial);
  return html = htmlServ.evaluate().getContent();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}