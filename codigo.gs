function doGet() {
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
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var params = JSON.parse(e.postData.contents);
  var data = sheet.getDataRange().getValues();
  
  if (params.action === 'save_row') {
    var p = params.data;
    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == p.ID) {
        sheet.getRange(i + 1, 1, 1, 12).setValues([[p.ID, p.Regiao, p.Nome, p.Categoria, p.Meses_Ideais, p.Esforco, p.Janela, p.Custo, p.Hub, p.Lat, p.Lng, p.Familias]]);
        found = true; break;
      }
    }
    if (!found) sheet.appendRow([p.ID, p.Regiao, p.Nome, p.Categoria, p.Meses_Ideais, p.Esforco, p.Janela, p.Custo, p.Hub, p.Lat, p.Lng, p.Familias]);
  } else if (params.action === 'delete_row') {
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == params.id) { sheet.deleteRow(i + 1); break; }
    }
  }
  return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
}