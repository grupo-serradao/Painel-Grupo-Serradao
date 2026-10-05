// Painel Grupo Serradão — Apps Script v2.1
// ────────────────────────────────────────────────────────────────
// v2.1: nova ação 'update' — altera só as células pedidas, conferindo
// o valor antigo de cada uma antes. Se alguma não bater, não grava nada.

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, status: 'online', versao: '2.1' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var body      = JSON.parse(e.postData.contents);
    var action    = body.action;
    var sheetName = body.sheet;
    var ss        = SpreadsheetApp.getActiveSpreadsheet();

    // ── ALTERAR CÉLULAS ──────────────────────────────────────────
    // body.cells = [{ row, col, expected, value }]  (row/col começam em 1)
    if (action === 'update') {
      var aba = ss.getSheetByName(sheetName);
      if (!aba) return err('Aba não encontrada: ' + sheetName);
      var cells = body.cells || [];
      if (!cells.length) return err('Nenhuma célula informada');
      if (cells.length > 200) return err('Máximo de 200 células por vez');

      var lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try {
        var divergentes = [];
        cells.forEach(function(c) {
          var atual = celula(aba.getRange(c.row, c.col).getValue());
          if (!mesmoValor(atual, c.expected)) {
            divergentes.push('linha ' + c.row + ' col ' + c.col + ': esperado "' + c.expected + '", está "' + atual + '"');
          }
        });
        if (divergentes.length) {
          return err('Nada foi alterado — valores diferentes do esperado (a planilha mudou?): ' + divergentes.slice(0, 5).join('; '));
        }
        cells.forEach(function(c) { aba.getRange(c.row, c.col).setValue(c.value); });
        SpreadsheetApp.flush();
      } finally {
        lock.releaseLock();
      }
      return ok({ updated: cells.length });
    }

    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);  // cria aba se não existir

    // ── LER ──────────────────────────────────────────────────────
    if (action === 'read') {
      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      if (lastRow === 0 || lastCol === 0) return ok({ values: [] });

      var raw = sheet.getRange(1, 1, lastRow, lastCol).getValues();
      var values = raw.map(function(row) { return row.map(celula); });
      return ok({ values: values });
    }

    // ── ADICIONAR LINHAS ─────────────────────────────────────────
    if (action === 'append') {
      var rows = body.values;
      if (rows && rows.length > 0) {
        var startRow = sheet.getLastRow() + 1;
        sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
      }
      return ok({});
    }

    // ── LIMPAR ABA ───────────────────────────────────────────────
    // IMPORTANTE: clear() apaga conteúdo E formatação
    // clearContents() (versão anterior) deixava formatação de data
    // nas células, corrompendo números salvos depois
    if (action === 'clear') {
      sheet.clear();
      return ok({});
    }

    return err('Ação desconhecida: ' + action);

  } catch (e) {
    return err(e.message);
  }
}

// ── Helpers ──────────────────────────────────────────────────────
// Mesmo formato que o 'read' devolve (datas em ISO UTC)
function celula(cell) {
  if (cell instanceof Date) {
    return Utilities.formatDate(cell, 'UTC', "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'");
  }
  return cell;
}

function mesmoValor(a, b) {
  if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) < 1e-9;
  return String(a) === String(b);
}

function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify(Object.assign({ ok: true }, data)))
    .setMimeType(ContentService.MimeType.JSON);
}

function err(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}
