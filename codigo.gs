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
  
  // --- ADICIONE ESTE BLOCO AQUI ---
  if (params.action === 'bulk_import') {
    var dataToImport = params.data;
    // Limpa a planilha (do cabeçalho para baixo)
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).clearContent();
    }
    // Prepara as linhas
    var rows = dataToImport.map(d => [d.ID, d.Regiao, d.Nome, d.Categoria, d.Meses_Ideais, d.Esforco, d.Janela, d.Custo, d.Hub, d.Lat, d.Lng, d.Familias]);
    // Salva na planilha
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 12).setValues(rows);
    }
    return ContentService.createTextOutput(JSON.stringify({status: "success"}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // --- FIM DO BLOCO DE IMPORTAÇÃO --
  
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