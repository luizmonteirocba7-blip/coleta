function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({status:"erro", mensagem:"Requisição sem conteúdo JSON."});
    }

    var data = JSON.parse(e.postData.contents);
    var equipe = String(data.equipe || "").trim();

    var equipeParaAba = {
      "SSPAN-Pantanal 01 (pocone)": "SSPAN-Pantanal 01 (pocone)",
      "SSPAN-Pantanal 02": "SSPAN-Pantanal 02",
      // Compatibilidade com a versão anterior do app
      "SSPAN-Pocone": "SSPAN-Pantanal 01 (pocone)",
      "SSPAN-Barao": "SSPAN-Pantanal 02"
    };

    if (!equipeParaAba[equipe]) {
      return jsonResponse({
        status:"erro",
        mensagem:"Acesso negado. Selecione SSPAN-Pantanal 01 (pocone) ou SSPAN-Pantanal 02."
      });
    }

    var abaNome = equipeParaAba[equipe];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(abaNome);

    if (!sheet) {
      return jsonResponse({status:"erro", mensagem:"Aba não encontrada: " + abaNome});
    }

    // A:R — mesma ordem da planilha:
    // A Latitude | B Longitude | C Período | D Data | E Hora
    // F T ar IRT | G T ar ARGOS | H UR IRT | I UR ARGOS
    // J T_Orvalho IRT | K T_Orvalho ARGOS
    // L Pressão IRT | M Pressão ARGOS
    // N Vento IRT | O Direção Vento IRT
    // P Vento ARGOS | Q Direção Vento ARGOS
    // R Cenário

    var rowData = [[
      data.latitude_irt || data.latitude_argos || "",
      data.longitude_irt || data.longitude_argos || "",
      data.periodo_irt || data.periodo_argos || "",
      data.data_irt || data.data_argos || "",
      data.hora_irt || data.hora_argos || "",
      data.t_ar_irt || "",
      data.t_ar_argos || "",
      data.ur_irt || "",
      data.ur_argos || "",
      data.t_orvalho_irt || "",
      data.t_orvalho_argos || "",
      data.pressao_irt || "",
      data.pressao_argos || "",
      data.vento_irt || "",
      data.direcao_irt || "",
      data.vento_argos || "",
      data.direcao_argos || "",
      data.cenario_irt || ""
    ]];

    var nextRow = Math.max(7, sheet.getLastRow() + 1);
    sheet.getRange(nextRow, 1, 1, 18).setValues(rowData);

    return jsonResponse({
      status:"sucesso",
      equipe:equipe,
      aba:abaNome,
      linha:nextRow,
      mensagem:"Dados salvos com sucesso em " + abaNome + "."
    });

  } catch (error) {
    return jsonResponse({status:"erro", mensagem:"Erro interno: " + error.message});
  }
}

function doGet() {
  return ContentService
    .createTextOutput("API do App de Campo SSPAN ativa.")
    .setMimeType(ContentService.MimeType.TEXT);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
