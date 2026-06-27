// Removemos o .setHeader pois não existe no Apps Script
function createResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    result.push(obj);
  }
  return createResponse(result);
}

function doPost(e) {
  // Nota: O Google Apps Script trata POST como uma requisição direta.
  // Se o seu navegador ainda bloquear, é por causa da falta de resposta "OPTIONS".
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var params = JSON.parse(e.postData.contents);
  var data = sheet.getDataRange().getValues();
  
  if (params.action === 'bulk_import') {
    sheet.getRange(2, 1, sheet.getLastRow(), 12).clearContent();
    var rows = params.data.map(d => [d.ID, d.Regiao, d.Nome, d.Categoria, d.Meses_Ideais, d.Esforco, d.Janela, d.Custo, d.Hub, d.Lat, d.Lng, d.Familias]);
    sheet.getRange(2, 1, rows.length, 12).setValues(rows);
    return createResponse({status: "success"});
  }
  
  if (params.action === 'delete_row') {
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() == params.id.toString()) {
        sheet.deleteRow(i + 1);
        return createResponse({status: "success"});
      }
    }
  }

  if (params.action === 'save_row') {
    var p = params.data;
    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() == p.ID.toString()) {
        sheet.getRange(i + 1, 1, 1, 12).setValues([[p.ID, p.Regiao, p.Nome, p.Categoria, p.Meses_Ideais, p.Esforco, p.Janela, p.Custo, p.Hub, p.Lat, p.Lng, p.Familias]]);
        found = true; break;
      }
    }
    if (!found) sheet.appendRow([p.ID, p.Regiao, p.Nome, p.Categoria, p.Meses_Ideais, p.Esforco, p.Janela, p.Custo, p.Hub, p.Lat, p.Lng, p.Familias]);
    return createResponse({status: "success"});
  }
  
  return createResponse({status: "error"});
}