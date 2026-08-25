/**
 * ALFA PRECATÓRIOS — Backend de Propostas (Google Sheets)
 * ------------------------------------------------------------
 * Cole este código em Extensões > Apps Script da sua planilha.
 * Depois: Implantar > Nova implantação > Tipo "App da Web".
 *   - Executar como: Eu mesmo (sua conta)
 *   - Quem tem acesso: Qualquer pessoa
 * Copie a URL gerada (termina em /exec) e informe ao sistema.
 */

var ABA = 'Propostas';

var COLUNAS = [
  'id','numero','savedAt','nome','cpf','rg','estcivil','profissao','email','telefone',
  'endereco','cidade_cli','estado','cep','processo','precatorio','tribunal',
  'vara','comarca','valorprec','valorprop','data','cidade','obs'
];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ABA);
  if (!sh) {
    sh = ss.insertSheet(ABA);
    sh.appendRow(COLUNAS);
    sh.getRange(1, 1, 1, COLUNAS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** LEITURA e GRAVAÇÃO via GET
 *  - Sem parâmetros: retorna todas as propostas.
 *  - Com ?payload=...: executa salvar/excluir (contorna o bloqueio de CORS do POST). */
function doGet(e) {
  try {
    // Se veio um payload na URL, trata como gravação/exclusão
    if (e && e.parameter && e.parameter.payload) {
      var body = JSON.parse(e.parameter.payload);
      return processar_(body);
    }
    // Caso contrário, leitura normal
    var sh = getSheet_();
    var vals = sh.getDataRange().getValues();
    var out = [];
    for (var i = 1; i < vals.length; i++) {
      var row = vals[i];
      if (!row[0]) continue;
      var obj = {};
      for (var c = 0; c < COLUNAS.length; c++) obj[COLUNAS[c]] = row[c];
      out.push(obj);
    }
    return json_({ ok: true, data: out });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** GRAVAÇÃO / EXCLUSÃO — também aceita POST */
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    return processar_(body);
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** Lógica central de salvar/excluir, usada por doGet e doPost */
function processar_(body) {
  var sh = getSheet_();

  if (body.action === 'delete') {
    _deleteById(sh, body.id);
    return json_({ ok: true });
  }

  var p = body.proposta || body;
  if (!p.id) p.id = Date.now();
  p.savedAt = new Date().toISOString();

  var rowIndex = _findRow(sh, p.id);
  var linha = COLUNAS.map(function (k) { return p[k] != null ? p[k] : ''; });

  if (rowIndex > 0) {
    sh.getRange(rowIndex, 1, 1, COLUNAS.length).setValues([linha]);
  } else {
    sh.appendRow(linha);
  }
  return json_({ ok: true, id: p.id });
}

function _findRow(sh, id) {
  var ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 0), 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function _deleteById(sh, id) {
  var r = _findRow(sh, id);
  if (r > 0) sh.deleteRow(r);
}
