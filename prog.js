const SHEET_NAME = "estoque";

function doGet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  const headers = data[0];
  const rows = data.slice(1);

  const result = rows.map(row => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = JSON.parse(e.postData.contents);

  const rows = sheet.getDataRange().getValues();

  // verifica se posição já existe
  for (let i = 1; i < rows.length; i++) {
    if (
      rows[i][0] == data.rua &&
      rows[i][1] == data.coluna &&
      rows[i][2] == data.altura
    ) {
      sheet.getRange(i + 1, 4).setValue(data.item);
      return ContentService.createTextOutput("Atualizado");
    }
  }

  // se não existir, adiciona
  sheet.appendRow([data.rua, data.coluna, data.altura, data.item]);

  return ContentService.createTextOutput("Adicionado");
}