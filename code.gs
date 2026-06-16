/**
 * GOOGLE APPS SCRIPT - API INTEGRATION FOR CONSTRUMAIS / SYSTEMA DE GESTÃO DE OBRAS
 * 
 * Este arquivo (code.gs) deve ser colado no Editor de Apps Script do seu Google Sheets.
 * Ele gerencia as planilhas: "Pedidos", "Baixas", "Obras" e "Usuarios".
 * Ele expõe uma API REST funcional para leitura, escrita, atualização e exclusão via doGet e doPost.
 */

// Define as constantes de nomes das Abas
var SHEET_PEDIDOS = "Pedidos";
var SHEET_BAIXAS = "Baixas";
var SHEET_OBRAS = "Obras";
var SHEET_USUARIOS = "Usuarios";

/**
 * Função de Inicialização automática das Planilhas se não existirem
 */
function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Aba Pedidos
  var sheetPedidos = ss.getSheetByName(SHEET_PEDIDOS);
  if (!sheetPedidos) {
    sheetPedidos = ss.insertSheet(SHEET_PEDIDOS);
    sheetPedidos.appendRow(["ID", "Insumo", "Código", "Obra", "Qtd Solicitada", "Qtd Recebida", "Unidade", "Status", "Data Pedido"]);
    sheetPedidos.getRange("A1:I1").setFontWeight("bold").setBackground("#e2e8f0");
  }
  
  // 2. Aba Baixas
  var sheetBaixas = ss.getSheetByName(SHEET_BAIXAS);
  if (!sheetBaixas) {
    sheetBaixas = ss.insertSheet(SHEET_BAIXAS);
    sheetBaixas.appendRow(["ID", "Insumo", "Código", "Obra", "Quantidade", "Unidade", "Colaborador", "Data"]);
    sheetBaixas.getRange("A1:H1").setFontWeight("bold").setBackground("#e2e8f0");
  }
  
  // 3. Aba Obras
  var sheetObras = ss.getSheetByName(SHEET_OBRAS);
  if (!sheetObras) {
    sheetObras = ss.insertSheet(SHEET_OBRAS);
    sheetObras.appendRow(["Nome da Obra"]);
    sheetObras.getRange("A1").setFontWeight("bold").setBackground("#e2e8f0");
    // Obras default iniciais
    sheetObras.appendRow(["Residencial Alvorada"]);
    sheetObras.appendRow(["Torre Infinito"]);
    sheetObras.appendRow(["Complexo Hospitalar"]);
  }
  
  // 4. Aba Usuarios
  var sheetUsuarios = ss.getSheetByName(SHEET_USUARIOS);
  if (!sheetUsuarios) {
    sheetUsuarios = ss.insertSheet(SHEET_USUARIOS);
    sheetUsuarios.appendRow(["ID", "Nome", "Role", "Login", "Senha", "Pode Criar Pedido", "Pode Dar Baixa", "Pode Receber Mercadoria", "Pode Excluir Pedido", "Pode Excluir Obra"]);
    sheetUsuarios.getRange("A1:J1").setFontWeight("bold").setBackground("#e2e8f0");
    
    // Usuários default iniciais (Carlos e Mestre José)
    sheetUsuarios.appendRow(["usr-1", "Carlos Roberto", "Administrador", "carlos@empresa.com", "admin123", "TRUE", "TRUE", "TRUE", "TRUE", "TRUE"]);
    sheetUsuarios.appendRow(["usr-2", "Mestre José", "Colaborador", "jose@empresa.com", "mestre123", "TRUE", "TRUE", "TRUE", "FALSE", "FALSE"]);
  }
}

/**
 * Trata requisições HTTP GET (Retorna os dados cadastrados em formato JSON)
 * Chamada exemplo: https://script.google.com/.../exec?action=getData
 */
function doGet(e) {
  try {
    var action = e.parameter.action || "getData";
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    if (action === "getData") {
      var data = {
        pedidos: readSheetData(ss.getSheetByName(SHEET_PEDIDOS)),
        baixas: readSheetData(ss.getSheetByName(SHEET_BAIXAS)),
        obras: readObrasList(ss.getSheetByName(SHEET_OBRAS)),
        usuarios: readUsuariosList(ss.getSheetByName(SHEET_USUARIOS))
      };
      return responseJson(data);
    }
    
    return responseJson({ error: "Ação descrita não suportada no GET" }, 400);
  } catch (err) {
    return responseJson({ error: err.toString() }, 500);
  }
}

/**
 * Trata requisições HTTP POST (Cadastros, Atualizações e Remoções)
 */
function doPost(e) {
  try {
    var rawBody = e.postData.contents;
    var payload = JSON.parse(rawBody);
    var action = payload.action;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    if (!action) {
      return responseJson({ error: "Ação de payload 'action' não especificada!" }, 400);
    }

    setupDatabase(); // Garante estrutura antes de processar

    switch (action) {
      // --- PEDIDOS ---
      case "savePedido":
        return saveOrUpdatePedido(ss.getSheetByName(SHEET_PEDIDOS), payload.data);
        
      case "deletePedido":
        return deleteRowById(ss.getSheetByName(SHEET_PEDIDOS), payload.id);

      // --- BAIXAS ---
      case "saveBaixa":
        return saveBaixa(ss.getSheetByName(SHEET_BAIXAS), payload.data);

      // --- OBRAS ---
      case "createObra":
        return createObra(ss.getSheetByName(SHEET_OBRAS), payload.name);
        
      case "renameObra":
        return renameObra(ss, payload.oldName, payload.newName);
        
      case "deleteObra":
        return deleteObra(ss, payload.name);

      // --- USUARIOS ---
      case "saveUsuario":
        return saveOrUpdateUsuario(ss.getSheetByName(SHEET_USUARIOS), payload.data);
        
      case "deleteUsuario":
        return deleteRowById(ss.getSheetByName(SHEET_USUARIOS), payload.id);

      default:
        return responseJson({ error: "Ação POST '" + action + "' não reconhecida." }, 400);
    }
  } catch (err) {
    return responseJson({ error: err.toString() }, 500);
  }
}

/**
 * Utilitário de leitura genérica de dados da tabela (excluindo headers), retornando objetos chave-valor mapeados
 */
function readSheetData(sheet) {
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  
  var headers = rows[0];
  var list = [];
  
  for (var r = 1; r < rows.length; r++) {
    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      var key = toCamelCase(headers[c]);
      var val = rows[r][c];
      
      // Converte booleans e formatos convenientes
      if (val === "TRUE" || val === true) val = true;
      if (val === "FALSE" || val === false) val = false;
      
      obj[key] = val;
    }
    list.push(obj);
  }
  return list;
}

/**
 * Lê lista de obras específica (uma única coluna)
 */
function readObrasList(sheet) {
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  
  var list = [];
  for (var r = 1; r < rows.length; r++) {
    var val = rows[r][0].toString().trim();
    if (val) list.push(val);
  }
  return list;
}

/**
 * Lê lista de usuários e reconverte booleanos de string para tipo lógico boolean
 */
function readUsuariosList(sheet) {
  var list = readSheetData(sheet);
  return list.map(function(u) {
    return {
      id: String(u.id),
      nome: String(u.nome),
      role: String(u.role),
      login: String(u.login),
      senha: String(u.senha),
      podeCriarPedido: String(u.podeCriarPedido).toLowerCase() === "true" || u.podeCriarPedido === true,
      podeDarBaixa: String(u.podeDarBaixa).toLowerCase() === "true" || u.podeDarBaixa === true,
      podeReceberMercadoria: String(u.podeReceberMercadoria).toLowerCase() === "true" || u.podeReceberMercadoria === true,
      podeExcluirPedido: String(u.podeExcluirPedido).toLowerCase() === "true" || u.podeExcluirPedido === true,
      podeExcluirObra: String(u.podeExcluirObra).toLowerCase() === "true" || u.podeExcluirObra === true
    };
  });
}

/**
 * Insere ou Atualiza Pedido de compras com base no ID
 */
function saveOrUpdatePedido(sheet, item) {
  if (!sheet) return responseJson({ error: "Planilha de pedidos não encontrada" }, 404);
  var data = sheet.getDataRange().getValues();
  var rowIdx = -1;
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === item.id.toString().trim()) {
      rowIdx = i + 1; // 1-index + headers
      break;
    }
  }
  
  var record = [
    item.id,
    item.insumo,
    item.codigo || "",
    item.obra,
    Number(item.qtdSolicitada),
    Number(item.qtdRecebida),
    item.unidade,
    item.status,
    item.dataPedido
  ];
  
  if (rowIdx !== -1) {
    sheet.getRange(rowIdx, 1, 1, record.length).setValues([record]);
    return responseJson({ success: true, message: "Pedido atualizado!", item: item });
  } else {
    sheet.appendRow(record);
    return responseJson({ success: true, message: "Pedido criado com sucesso!", item: item });
  }
}

/**
 * Salva um novo diário de baixa estoque físico
 */
function saveBaixa(sheet, item) {
  if (!sheet) return responseJson({ error: "Planilha de baixas não encontrada" }, 404);
  var record = [
    item.id,
    item.insumo,
    item.codigo || "",
    item.obra,
    Number(item.quantidade),
    item.unidade,
    item.colaborador,
    item.data
  ];
  sheet.appendRow(record);
  return responseJson({ success: true, message: "Registro de diário de baixa lançado com sucesso!", item: item });
}

/**
 * Cria nova Obra
 */
function createObra(sheet, name) {
  if (!sheet || !name) return responseJson({ error: "Nome ou planilha inválida" }, 400);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase().trim() === name.toLowerCase().trim()) {
      return responseJson({ error: "Empreendimento/Obra já cadastrado no sistema!" }, 400);
    }
  }
  sheet.appendRow([name.trim()]);
  return responseJson({ success: true, message: "Obra cadastrada com sucesso!" });
}

/**
 * Renomeia Obra nas listas globais e atualiza todos os pedidos e baixas vinculados
 */
function renameObra(ss, oldName, newName) {
  var sheetObras = ss.getSheetByName(SHEET_OBRAS);
  if (!sheetObras) return responseJson({ error: " planilha Obras indisponível" }, 404);
  
  var rows = sheetObras.getDataRange().getValues();
  var renamed = false;
  
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0].toString().trim() === oldName.trim()) {
      sheetObras.getRange(i + 1, 1).setValue(newName.trim());
      renamed = true;
      break;
    }
  }
  
  if (renamed) {
    // Cascade update nos pedidos
    var sheetPedidos = ss.getSheetByName(SHEET_PEDIDOS);
    if (sheetPedidos) {
      var pData = sheetPedidos.getDataRange().getValues();
      for (var p = 1; p < pData.length; p++) {
        if (pData[p][3].toString().trim() === oldName.trim()) {
          sheetPedidos.getRange(p + 1, 4).setValue(newName.trim());
        }
      }
    }
    
    // Cascade update nas baixas
    var sheetBaixas = ss.getSheetByName(SHEET_BAIXAS);
    if (sheetBaixas) {
      var bData = sheetBaixas.getDataRange().getValues();
      for (var b = 1; b < bData.length; b++) {
        if (bData[b][3].toString().trim() === oldName.trim()) {
          sheetBaixas.getRange(b + 1, 4).setValue(newName.trim());
        }
      }
    }
    return responseJson({ success: true, message: "Obra alterada com sucesso!" });
  }
  
  return responseJson({ error: "Obra de origem não encontrada" }, 404);
}

/**
 * Remove a obra do cadastro geral
 */
function deleteObra(ss, name) {
  var sheet = ss.getSheetByName(SHEET_OBRAS);
  if (!sheet) return responseJson({ error: "Planilha Obras não encontrada" }, 404);
  
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0].toString().trim() === name.trim()) {
      sheet.deleteRow(i + 1);
      return responseJson({ success: true, message: "Obra excluída com sucesso dos registros cadastrais!" });
    }
  }
  return responseJson({ error: "Obra não encontrada para o termo solicitado" }, 404);
}

/**
 * Salva ou Atualiza colaboradores / perfis no painel administrativo
 */
function saveOrUpdateUsuario(sheet, item) {
  if (!sheet) return responseJson({ error: "Planilha de usuários ausente" }, 404);
  var data = sheet.getDataRange().getValues();
  var rowIdx = -1;
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === item.id.toString().trim()) {
      rowIdx = i + 1;
      break;
    }
  }
  
  var record = [
    item.id,
    item.nome,
    item.role,
    item.login,
    item.senha,
    String(item.podeCriarPedido).toUpperCase(),
    String(item.podeDarBaixa).toUpperCase(),
    String(item.podeReceberMercadoria).toUpperCase(),
    String(item.podeExcluirPedido).toUpperCase(),
    String(item.podeExcluirObra).toUpperCase()
  ];
  
  if (rowIdx !== -1) {
    sheet.getRange(rowIdx, 1, 1, record.length).setValues([record]);
    return responseJson({ success: true, message: "Dados do colaborador atualizados!", item: item });
  } else {
    sheet.appendRow(record);
    return responseJson({ success: true, message: "Novo colaborador cadastrado com êxito!", item: item });
  }
}

/**
 * Exclui uma linha por ID em qualquer aba que utilize o ID como primeira coluna
 */
function deleteRowById(sheet, id) {
  if (!sheet) return responseJson({ error: "Planilha alvo inexistente" }, 404);
  var rows = sheet.getDataRange().getValues();
  for (var r = 1; r < rows.length; r++) {
    if (rows[r][0].toString().trim() === id.toString().trim()) {
      sheet.deleteRow(r + 1);
      return responseJson({ success: true, message: "Registro excluído com êxito de código " + id });
    }
  }
  return responseJson({ error: "Registro não encotrado com ID " + id, id: id }, 404);
}

/**
 * Retorna uma resposta amoldada em JSON seguro
 */
function responseJson(object, statusCode) {
  var status = statusCode || 200;
  var output = ContentService.createTextOutput(JSON.stringify(object))
                 .setMimeType(ContentService.MimeType.JSON);
  return output;
}

/**
 * Converte headers de texto para camelCase (ex: "Qtd Solicitada" -> "qtdSolicitada")
 */
function toCamelCase(str) {
  var clean = str.toString().toLowerCase().trim()
    .replace(/[áàãâä]/g, "a")
    .replace(/[éèêë]/g, "e")
    .replace(/[íìîï]/g, "i")
    .replace(/[óòõôö]/g, "o")
    .replace(/[úùûü]/g, "u")
    .replace(/ç/g, "c")
    .replace(/[^a-z0-9\s]/g, ""); // remove caracteres especiais
    
  var words = clean.split(/\s+/);
  if (words.length === 0) return "";
  
  var result = words[0];
  for (var i = 1; i < words.length; i++) {
    result += words[i].charAt(0).toUpperCase() + words[i].slice(1);
  }
  return result;
}
