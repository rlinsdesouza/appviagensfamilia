// Função de cálculo de distância
function getDistancia(lat1, lon1, lat2, lon2) {
  var R = 6371000;
  var dLat = (lat2 - lat1) * Math.PI / 180;
  var dLon = (lon2 - lon1) * Math.PI / 180;
  var a = Math.sin(dLat/2) * Math.sin(dLat/2) + Math.cos(lat1 * Math.PI/180) * Math.cos(lat2 * Math.PI/180) * Math.sin(dLon/2) * Math.sin(dLon/2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
}

// VALIDAÇÃO: Sem o sinal de '=' nos parâmetros para evitar SyntaxError
function validarDestino(novoP, dataExistente, idAtual) {
  for (var i = 1; i < dataExistente.length; i++) {
    var row = dataExistente[i];
    var idExistente = row[0] ? row[0].toString() : "";
    var latExistente = row[9];
    var lngExistente = row[10];
    
    // Checa ID duplicado
    if (idExistente === novoP.ID.toString() && idExistente !== idAtual.toString()) return "ID já cadastrado!";
    
    // Checa Proximidade (200m)
    if (idExistente !== novoP.ID.toString()) {
      var dist = getDistancia(novoP.Lat, novoP.Lng, latExistente, lngExistente);
      if (dist < 200) return "Local duplicado! Já existe um destino a menos de 200m.";
    }
  }
  return null;
}

function doGet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) { obj[headers[j]] = data[i][j]; }
    result.push(obj);
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  try {
    var params = JSON.parse(e.postData.contents);
    var data = sheet.getDataRange().getValues();
    
    // IMPORTAÇÃO
    if (params.action === 'append_import') {
      var rowsToAppend = [];
      var debugLog = []; // Vamos registrar o que acontece aqui
      
      params.data.forEach(function(d) {
        var erro = validarDestino(d, data, "");
        if (!erro) {
          rowsToAppend.push([d.ID, d.Regiao, d.Nome, d.Categoria, d.Meses_Ideais, d.Esforco, d.Janela, d.Custo, d.Hub, d.Lat, d.Lng, d.Familias, d.Historico || ""]);
          data.push([d.ID, "", "", "", "", "", "", "", "", d.Lat, d.Lng]); 
        } else {
          debugLog.push("ID " + d.ID + " rejeitado: " + erro);
        }
      });
      
      // Registra no Log do Google para você ver
      Logger.log("Importação: Tentou " + params.data.length + " itens.");
      Logger.log("Rejeitados: " + debugLog.join(" | "));
      
      if (rowsToAppend.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 13).setValues(rowsToAppend);
      }
      
      return ContentService.createTextOutput(JSON.stringify({
        status: "success", 
        count: rowsToAppend.length, 
        rejeitados: debugLog // Isso vai aparecer no seu console se houver erro
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // SALVAR
    if (params.action === 'save_row') {
      var p = params.data;
      var erro = validarDestino(p, data, p.ID);
      if (erro) return ContentService.createTextOutput(JSON.stringify({status: "error", msg: erro})).setMimeType(ContentService.MimeType.JSON);
      
      var found = false;
      for (var i = 1; i < data.length; i++) {
        if (data[i][0].toString() == p.ID.toString()) {
          sheet.getRange(i + 1, 1, 1, 13).setValues([[p.ID, p.Regiao, p.Nome, p.Categoria, p.Meses_Ideais, p.Esforco, p.Janela, p.Custo, p.Hub, p.Lat, p.Lng, p.Familias, p.Historico]]);
          found = true; break;
        }
      }
      if (!found) sheet.appendRow([p.ID, p.Regiao, p.Nome, p.Categoria, p.Meses_Ideais, p.Esforco, p.Janela, p.Custo, p.Hub, p.Lat, p.Lng, p.Familias, p.Historico]);
      return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({status: "error", msg: "Ação inválida"})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", msg: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}