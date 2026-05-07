const SHEET_NAME = "Página1";

const sheet =
  SpreadsheetApp.openById(
    "1EqNHODFZ_XIDOahT1clXIhfs1ZzMw_OagdIm36YGwRM"
  );

function doGet(e) {

  const aba = sheet.getSheetByName(SHEET_NAME);

  if (e.parameter.rua) {

    const rua = e.parameter.rua;
    const coluna = e.parameter.coluna;
    const altura = e.parameter.altura;
    const item = e.parameter.item;

    const rows = aba.getDataRange().getValues();

    for (let i = 1; i < rows.length; i++) {

      if (
        rows[i][0] == rua &&
        rows[i][1] == coluna &&
        rows[i][2] == altura
      ) {

        aba.getRange(i + 1, 4).setValue(item);

        return ContentService
  .createTextOutput(
    e.parameter.callback + '("Atualizado")'
  )
  .setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
    }

    aba.appendRow([
      rua,
      coluna,
      altura,
      item
    ]);

    return ContentService
  .createTextOutput(
    e.parameter.callback + '("Adicionado")'
  )
  .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  const data = aba.getDataRange().getValues();

  const headers = data[0];
  const rows = data.slice(1);

  const result = rows.map(row => {

    let obj = {};

    headers.forEach((h, i) => {
      obj[h] = row[i];
    });

    return obj;
  });

  const callback = e.parameter.callback;

return ContentService
  .createTextOutput(
    callback + "(" + JSON.stringify(result) + ")"
  )
  .setMimeType(ContentService.MimeType.JAVASCRIPT);
}