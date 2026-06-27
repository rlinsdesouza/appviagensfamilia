function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var result = [];
    
    for (var i = 1; i < data.length; i++) {
      var obj = {};
      for (var j = 0; j < headers.length; j++) {
        var val = data[i][j];
        if (headers[j] == "Lat" || headers[j] == "Lng") {
          obj[headers[j]] = val ? parseFloat(val.toString().replace(',', '.')) : 0;
        } else {
          obj[headers[j]] = val;
        }
      }
      result.push(obj);
    }
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({"error": err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  // A trava de segurança do Google: quem não for editor da planilha receberá um erro de permissão do próprio Google antes mesmo de rodar isso.
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var params = JSON.parse(e.postData.contents);
  
  // AÇÃO 1: UPLOAD DE BASE JSON (NOVO)
  if (params.action === 'bulk_import') {
    var dadosInjetados = params.data;
    if (!dadosInjetados || dadosInjetados.length === 0) return ContentService.createTextOutput(JSON.stringify({"status": "error"})).setMimeType(ContentService.MimeType.JSON);
    
    // Apaga os dados antigos (mantendo a linha 1 do cabeçalho)
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
    
    var headers = ["ID", "Regiao", "Nome", "Categoria", "Meses_Ideais", "Esforco", "Janela", "Custo", "Hub", "Lat", "Lng", "Familias"];
    var rows = [];
    for (var i = 0; i < dadosInjetados.length; i++) {
      var row = [];
      for (var h = 0; h < headers.length; h++) {
        row.push(dadosInjetados[i][headers[h]] || "");
      }
      rows.push(row);
    }
    
    // Grava tudo de uma vez de forma super rápida
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    return ContentService.createTextOutput(JSON.stringify({"status": "success", "msg": rows.length + " destinos injetados!"})).setMimeType(ContentService.MimeType.JSON);
  }
  
  // AÇÃO 2: ATUALIZAR STATUS DE FAMÍLIA (O que você já usava)
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == params.id) {
      sheet.getRange(i + 1, 12).setValue(params.familias);
      return ContentService.createTextOutput(JSON.stringify({"status": "success"})).setMimeType(ContentService.MimeType.JSON);
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify({"status": "error"})).setMimeType(ContentService.MimeType.JSON);
}