// 1. Função de cálculo de distância (Haversine)
function getDistancia(lat1, lon1, lat2, lon2) {
  var R = 6371000; // Raio da Terra em metros
  var dLat = (lat2 - lat1) * Math.PI / 180;
  var dLon = (lon2 - lon1) * Math.PI / 180;
  var a = Math.sin(dLat/2) * Math.sin(dLat/2) + 
          Math.cos(lat1 * Math.PI/180) * Math.cos(lat2 * Math.PI/180) * Math.sin(dLon/2) * Math.sin(dLon/2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
}

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
	var isNew = true; // Assumimos que é novo
	
	// Validação de segurança: ID não pode ser vazio
if (!p.ID || p.ID.toString().trim() === "") {
    return ContentService.createTextOutput(JSON.stringify({status: "error", msg: "O ID é obrigatório!"})).setMimeType(ContentService.MimeType.JSON);
}
	
	// Verifica se ID já existe
for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() == p.ID.toString()) {
        // Se encontramos o ID, não é novo, é uma edição
        isNew = false; 
        break;
    }
}

// Se tentou criar um ID que já existe (e não é uma edição), bloqueia
if (isNew && data.some(row => row[0].toString() == p.ID.toString())) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", msg: "ID já cadastrado!"})).setMimeType(ContentService.MimeType.JSON);
}
	
	// Check de Integridade: Verifica se existe outro lugar num raio de 200m
    for (var i = 1; i < data.length; i++) {
        // Pula a checagem se for o mesmo ID (estamos editando, não criando)
        if (data[i][0].toString() == p.ID.toString()) continue; 
        
        var dist = getDistancia(p.Lat, p.Lng, data[i][9], data[i][10]); // lat=col 10, lng=col 11
        if (dist < 200) {
            return ContentService.createTextOutput(JSON.stringify({status: "error", msg: "Local já cadastrado a menos de 200m!"})).setMimeType(ContentService.MimeType.JSON);
        }
    }
	
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == p.ID) {
        sheet.getRange(i + 1, 1, 1, 13).setValues([[
    p.ID, p.Regiao, p.Nome, p.Categoria, p.Meses_Ideais, 
    p.Esforco, p.Janela, p.Custo, p.Hub, p.Lat, 
    p.Lng, p.Familias, p.Historico]]);
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