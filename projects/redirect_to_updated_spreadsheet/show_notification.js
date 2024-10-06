// Example of how to use redirectToUpdatedSpreadsheet:
//
// function onOpenRedirect() {
//   var newSpreadsheetId = '1CIgSDfo3n9DJXP7rKMc-4R9xfDc6s4dS22zmGYp8Epg';
//   var title = "Outdated BIA Template";
//   var message = "This BIA assessment is outdated.";
//   RedirectToUpdatedSpreadsheet.showNotification(newSpreadsheetId, title, message);
// }

function showNotification(newSpreadsheetId, title, message) {
  var newSpreadsheetUrl = 'https://docs.google.com/spreadsheets/d/' + newSpreadsheetId;
  var body = `
    <style>body { font-family: 'Yahoo Sans', sans-serif; }</style>
    <p>${message}</p>
    <p>Use <a href="${newSpreadsheetUrl}" target="_blank">the new version</a> instead.</p>
  `;
  var ui = SpreadsheetApp.getUi();
  var htmlOutput = HtmlService.createHtmlOutput(body).setWidth(300).setHeight(75);
  ui.showModalDialog(htmlOutput, title);
}
