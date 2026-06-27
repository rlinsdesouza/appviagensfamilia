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
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var params = JSON.parse(e.postData.contents);
  var data = sheet.getDataRange().getValues();
  
  // 1. AÇÃO: SALVAR OU EDITAR (Upsert)
  if (params.action === 'save_row') {
    var p = params.data;
    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == p.ID) {
        // Atualiza linha existente (colunas 1 a 12)
        sheet.getRange(i + 1, 1, 1, 12).setValues([[p.ID, p.Regiao, p.Nome, p.Categoria, p.Meses_Ideais, p.Esforco, p.Janela, p.Custo, p.Hub, p.Lat, p.Lng, p.Familias]]);
        found = true; break;
      }
    }
    if (!found) { // Cria nova linha
      sheet.appendRow([p.ID, p.Regiao, p.Nome, p.Categoria, p.Meses_Ideais, p.Esforco, p.Janela, p.Custo, p.Hub, p.Lat, p.Lng, p.Familias]);
    }
    return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
  }

 // AÇÃO: EXCLUIR
  else if (params.action === 'delete_row') {
    var idParaDeletar = params.id.toString(); // Forçamos para texto
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() == idParaDeletar) {
        sheet.deleteRow(i + 1);
        return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status: "error", msg: "ID não encontrado"})).setMimeType(ContentService.MimeType.JSON);
  }
  // 3. AÇÃO PADRÃO: ATUALIZAR FAMÍLIAS (Sua lógica original preservada)
  else {
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == params.id) {
        sheet.getRange(i + 1, 12).setValue(params.familias);
        return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
      }
    }
  }
}