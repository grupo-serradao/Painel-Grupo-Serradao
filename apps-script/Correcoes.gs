// Correções de dados — rodar pelo editor (selecione "corrigirDados" e clique em ▶ Executar).
// Não precisa implantar. Cada correção confere o valor atual antes de gravar;
// se já estiver corrigido ou diferente do esperado, não mexe e avisa no registro.

// Preencha se souber o valor real (deixe null para não alterar):
var KG_TOTAL_BOV_10_01 = null;   // kg total de bovinos de 10/01/2026 (hoje: 20782)
var PARADA_MESA_10_01_MIN = null; // minutos da parada "Mesa de vísceras parou" de 10/01/2026 (hoje: 300)

function corrigirDados() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  var log = [];

  // 1) FÍGADO 14/01/2026: 19665,2 → 1966,52 (vírgula no lugar errado)
  var mi = ss.getSheetByName('Miudos');
  var linhas = acharLinhas(mi, tz, 1, '2026-01-14', function(r) { return norm(r[1]) === 'FIGADO'; });
  log.push(trocarValor(mi, linhas, 4, 19665.2, 1966.52, 'FÍGADO 14/01 (Miudos col D)', true));

  // 2) Pesos 02/02/2026: colunas C e D trocadas (10241 / 23793)
  var pe = ss.getSheetByName('Pesos');
  linhas = acharLinhas(pe, tz, 2, '2026-02-02');
  if (linhas.length !== 1) {
    log.push('⚠ Pesos 02/02: ' + linhas.length + ' linhas encontradas — não alterado');
  } else {
    var c = pe.getRange(linhas[0], 3), d = pe.getRange(linhas[0], 4);
    if (igual(c.getValue(), 10241) && igual(d.getValue(), 23793)) {
      c.setValue(23793); d.setValue(10241);
      log.push('✓ Pesos 02/02: colunas C e D trocadas (linha ' + linhas[0] + ')');
    } else if (igual(c.getValue(), 23793) && igual(d.getValue(), 10241)) {
      log.push('= Pesos 02/02: já estava corrigido');
    } else {
      log.push('⚠ Pesos 02/02: valores diferentes do esperado (C=' + c.getValue() + ', D=' + d.getValue() + ') — não alterado');
    }
  }

  // 3) kg total de bovinos 10/01/2026 (só se preenchido lá em cima)
  if (KG_TOTAL_BOV_10_01 !== null) {
    var ta = ss.getSheetByName('TempoAbate');
    linhas = acharLinhas(ta, tz, 1, '2026-01-10', function(r) { return norm(r[1]) === 'BOVINO'; });
    log.push(trocarValor(ta, linhas, 4, 20782, KG_TOTAL_BOV_10_01, 'kg 10/01 (TempoAbate col D)', false));
    linhas = acharLinhas(mi, tz, 1, '2026-01-10', function(r) { return norm(r[1]) !== 'VERGALHO'; });
    var alterados = 0, outros = 0;
    linhas.forEach(function(l) {
      var e = mi.getRange(l, 5);
      if (igual(e.getValue(), 20782)) { e.setValue(KG_TOTAL_BOV_10_01); recalcPct(mi, l); alterados++; }
      else if (!igual(e.getValue(), KG_TOTAL_BOV_10_01)) outros++;
    });
    log.push('✓ kg 10/01 (Miudos col E): ' + alterados + ' linha(s) alterada(s)' + (outros ? ', ⚠ ' + outros + ' com valor inesperado (não alteradas)' : ''));
  }

  // 4) Parada "Mesa de vísceras parou" 10/01/2026 (só se preenchido lá em cima)
  if (PARADA_MESA_10_01_MIN !== null) {
    var pa = ss.getSheetByName('Paradas');
    linhas = acharLinhas(pa, tz, 1, '2026-01-10', function(r) { return norm(r[4]).indexOf('MESA DE VISCERAS PAROU') >= 0; });
    log.push(trocarValor(pa, linhas, 6, 300, PARADA_MESA_10_01_MIN, 'Parada mesa 10/01 (Paradas col F)', false));
  }

  SpreadsheetApp.flush();
  var texto = log.join('\n');
  console.log(texto);
  return texto;
}

// ── Helpers ──────────────────────────────────────────────────────
function acharLinhas(aba, tz, colData, data, filtro) {
  if (!aba) return [];
  var vals = aba.getDataRange().getValues(), out = [];
  for (var i = 1; i < vals.length; i++) {
    var v = vals[i][colData - 1];
    var dt = v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v).slice(0, 10);
    if (dt === data && (!filtro || filtro(vals[i]))) out.push(i + 1);
  }
  return out;
}

function trocarValor(aba, linhas, col, antigo, novo, nome, pctMiudos) {
  if (linhas.length !== 1) return '⚠ ' + nome + ': ' + linhas.length + ' linhas encontradas — não alterado';
  var cel = aba.getRange(linhas[0], col), atual = cel.getValue();
  if (igual(atual, novo)) return '= ' + nome + ': já estava corrigido';
  if (!igual(atual, antigo)) return '⚠ ' + nome + ': valor atual ' + atual + ' (esperado ' + antigo + ') — não alterado';
  cel.setValue(novo);
  if (pctMiudos) recalcPct(aba, linhas[0]);
  return '✓ ' + nome + ': ' + antigo + ' → ' + novo + ' (linha ' + linhas[0] + ')';
}

// Miudos col F (PctReal) = PesoKg (D) / KgTotalAbate (E) × 100
function recalcPct(aba, linha) {
  var p = Number(aba.getRange(linha, 4).getValue()), k = Number(aba.getRange(linha, 5).getValue());
  if (k > 0) aba.getRange(linha, 6).setValue(Math.round(p / k * 100 * 10000) / 10000);
}

function igual(a, b) { return Math.abs(Number(a) - Number(b)) < 1e-6; }

function norm(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().trim();
}
