/**
 * Baixa de IMEI vendido na Base do Feishu — VERSÃO INDEPENDENTE.
 * Pode ficar num PROJETO SEPARADO do Apps Script (não depende do projeto do painel de fotos).
 * Cole este arquivo sozinho num projeto novo.
 *
 * O que faz: lê a planilha de VENDAS do Google (IMEI na coluna E, data da venda na coluna G)
 * e, para cada IMEI, escreve "Vendido <data>" na coluna VENDIDO (已售) da Base do Feishu,
 * procurando em todas as abas da Base que tenham coluna de IMEI e de VENDIDO (Sheet2 e Transferência).
 * Os IMEIs que não existirem na Base vão para a aba "Não encontrados Feishu" da planilha de vendas.
 *
 * Propriedades do script (⚙️ Configurações do projeto → Propriedades do script):
 *   VENDAS_PLANILHA   = link da planilha de vendas do Google. VÁRIOS MESES: separe os links por vírgula.
 *   VENDAS_ABA        = nome da aba das vendas (a mesma em todas, ex.: dados). Em branco = a primeira aba de cada.
 *   FEISHU_ESTOQUE    = link da Base do Feishu (começa com .../wiki/... ou .../base/...).
 *   FEISHU_APP_ID     = App ID do app no Feishu (ex.: cli_xxx).      ← em 凭证与基础信息
 *   FEISHU_APP_SECRET = App Secret do mesmo app.                     ← em 凭证与基础信息 (botão de revelar)
 *
 * Permissões do app no Feishu (verdes/免审): bitable:app e wiki:wiki:readonly.
 * E o app precisa estar como EDITOR na Base (abrir a Base → Compartilhar → adicionar o app).
 *
 * 1) marcarVendidosTeste() → NÃO grava nada: mostra quantos marcaria e quantos não achou.
 * 2) marcarVendidos()      → grava de verdade.
 * 3) marcarVendidosLigar() → roda sozinho a cada hora. marcarVendidosDesligar() para parar.
 */
var VENDAS_LIN_INI = 4;    // primeira linha de dados na planilha de vendas
var VENDAS_COL_IMEI = 5;   // coluna E
var VENDAS_COL_DATA = 7;   // coluna G
var VD_FEISHU = 'https://open.feishu.cn/open-apis';
var VD_TK = null;          // token guardado durante a execução

// ---------- conexão com o Feishu ----------
function vdToken_() {
  if (VD_TK) return VD_TK;
  var p = PropertiesService.getScriptProperties();
  var id = p.getProperty('FEISHU_APP_ID'), sec = p.getProperty('FEISHU_APP_SECRET');
  if (!id || !sec) throw new Error('Faltam as propriedades FEISHU_APP_ID e FEISHU_APP_SECRET.');
  var res = UrlFetchApp.fetch(VD_FEISHU + '/auth/v3/tenant_access_token/internal', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ app_id: id, app_secret: sec }),
  });
  var j = JSON.parse(res.getContentText() || '{}');
  if (!j.tenant_access_token) throw new Error('Não consegui o token do Feishu: ' + String(j.msg || res.getContentText()).slice(0, 150));
  VD_TK = j.tenant_access_token;
  return VD_TK;
}
function vdFeishu_(method, path, body) {
  var opt = { method: String(method).toLowerCase(), muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + vdToken_() } };
  if (body !== undefined) { opt.contentType = 'application/json'; opt.payload = JSON.stringify(body); }
  var res = UrlFetchApp.fetch(VD_FEISHU + path, opt);
  var j = JSON.parse(res.getContentText() || '{}');
  if (j.code && j.code !== 0) throw new Error('Feishu ' + path + ' → ' + j.code + ' ' + (j.msg || ''));
  return j.data || j;
}
function vdLista_(path, grande) {
  var out = [], page = '';
  do {
    var r = vdFeishu_('GET', path + (path.indexOf('?') >= 0 ? '&' : '?') + 'page_size=' + (grande ? 500 : 100) + (page ? '&page_token=' + page : ''));
    (r.items || []).forEach(function (x) { out.push(x); });
    page = r.has_more ? r.page_token : '';
  } while (page);
  return out;
}

// ---------- utilidades ----------
function soDigitos_(v) {
  if (v === null || v === undefined || v instanceof Date) return '';
  var s = String(v).replace(/\D/g, '');
  return s.replace(/^0+/, '') || s;
}
function campoTexto_(val) {
  if (val === null || val === undefined) return '';
  if (typeof val === 'string' || typeof val === 'number') return String(val);
  if (Array.isArray(val)) return val.map(function (x) { return (x && (x.text || x.name)) || ''; }).join('');
  if (typeof val === 'object') return String(val.text || val.name || '');
  return String(val);
}
function formataData_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'America/Sao_Paulo', 'dd/MM/yyyy');
  return String(v || '').trim();
}
function achaCampo_(campos, re) { for (var i = 0; i < campos.length; i++) { if (re.test(campos[i].field_name)) return campos[i].field_name; } return ''; }

// ---------- planilhas de vendas (Google) — uma ou várias (uma por mês), separadas por vírgula ----------
function vendasFontes_() {
  var p = PropertiesService.getScriptProperties();
  var links = String(p.getProperty('VENDAS_PLANILHA') || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
  if (!links.length) throw new Error('Falta a propriedade VENDAS_PLANILHA (link da planilha; separe vários meses por vírgula).');
  var nome = String(p.getProperty('VENDAS_ABA') || '').trim();
  var fontes = [];
  links.forEach(function (link) {
    var m = link.match(/\/d\/([A-Za-z0-9_-]+)/);
    var ss = SpreadsheetApp.openById(m ? m[1] : link);
    var a = nome ? ss.getSheetByName(nome) : ss.getSheets()[0];
    if (a) fontes.push({ ss: ss, aba: a });
    else console.warn('Aba "' + nome + '" não existe em "' + ss.getName() + '" — pulei essa planilha.');
  });
  if (!fontes.length) throw new Error('Nenhuma aba encontrada. Confira VENDAS_ABA (nome exato) e os links em VENDAS_PLANILHA.');
  return fontes;
}
function escreveNaoEncontrados_(ss, linhas) {
  var a = ss.getSheetByName('Não encontrados Feishu') || ss.insertSheet('Não encontrados Feishu');
  a.clear();
  a.getRange(1, 1, 1, 3).setValues([['IMEI', 'Data da venda', 'Conferido em']]);
  a.getRange(2, 1, Math.max(linhas.length, 1), 1).setNumberFormat('@');
  if (linhas.length) {
    var agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');
    a.getRange(2, 1, linhas.length, 3).setValues(linhas.map(function (l) { return [l[0], l[1], agora]; }));
  }
}

// ---------- Base do Feishu (Bitable) ----------
function estoqueBase_() {
  var link = String(PropertiesService.getScriptProperties().getProperty('FEISHU_ESTOQUE') || '').trim();
  if (!link) throw new Error('Falta a propriedade FEISHU_ESTOQUE com o link da Base do Feishu.');
  var m = link.match(/\/(wiki|base|app)\/([A-Za-z0-9]+)/);
  var token = m ? m[2] : (/^[A-Za-z0-9]{10,}$/.test(link) ? link : '');
  if (!token) throw new Error('Não entendi o link da Base no FEISHU_ESTOQUE.');
  if (m && m[1] === 'wiki') {
    var r = vdFeishu_('GET', '/wiki/v2/spaces/get_node?obj_type=wiki&token=' + token);
    token = r.node && r.node.obj_token;
    if (!token) throw new Error('Não consegui abrir a Base do Wiki. Confira a permissão wiki:wiki:readonly e se o app está na Base.');
  }
  return token;
}

function marcarVendidosTeste() { marcarVendidos_(true); }
function marcarVendidos() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) throw new Error('Outra rotina está rodando agora. Tente de novo em 1 minuto.');
  try { marcarVendidos_(false); } finally { lock.releaseLock(); }
}
function marcarVendidosDiario() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try { marcarVendidos_(false); } finally { lock.releaseLock(); }
}
function marcarVendidosLigar() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'marcarVendidosDiario') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('marcarVendidosDiario').timeBased().everyHours(1).create();
  console.log('Baixa de vendidos no Feishu ligada (a cada hora).');
}
function marcarVendidosDesligar() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'marcarVendidosDiario') ScriptApp.deleteTrigger(t); });
  console.log('Baixa de vendidos no Feishu desligada.');
}

function marcarVendidos_(teste) {
  var inicio = Date.now(), limite = 5 * 60 * 1000;
  var fontes = vendasFontes_();
  var app = estoqueBase_();

  var tabelas = vdLista_('/bitable/v1/apps/' + app + '/tables', false).map(function (t) {
    var campos = vdLista_('/bitable/v1/apps/' + app + '/tables/' + t.table_id + '/fields', false);
    return { id: t.table_id, nome: t.name, imei: achaCampo_(campos, /imei/i), vendido: achaCampo_(campos, /vendido|已售|已卖/i) };
  }).filter(function (t) { return t.imei && t.vendido; });
  if (!tabelas.length) throw new Error('Nenhuma aba da Base tem as colunas IMEI e VENDIDO. Confira os nomes das colunas.');

  var idx = {};
  tabelas.forEach(function (t) {
    vdLista_('/bitable/v1/apps/' + app + '/tables/' + t.id + '/records', true).forEach(function (rec) {
      var d = soDigitos_(campoTexto_(rec.fields[t.imei]));
      if (d.length >= 10) { (idx[d] = idx[d] || []).push({ tid: t.id, campo: t.vendido, rid: rec.record_id, ja: campoTexto_(rec.fields[t.vendido]) }); }
    });
  });

  var naoAchou = [], marcou = 0, jaTinha = 0, lin = 0, parou = false;
  for (var s = 0; s < fontes.length && !parou; s++) {
    var aba = fontes[s].aba, ultima = aba.getLastRow();
    if (ultima < VENDAS_LIN_INI) continue;
    var vals = aba.getRange(VENDAS_LIN_INI, 1, ultima - VENDAS_LIN_INI + 1, Math.max(VENDAS_COL_DATA, VENDAS_COL_IMEI)).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (Date.now() - inicio > limite) { parou = true; break; }
      var imei = soDigitos_(vals[i][VENDAS_COL_IMEI - 1]);
      if (imei.length < 10) continue;
      lin++;
      var data = formataData_(vals[i][VENDAS_COL_DATA - 1]);
      var txt = 'Vendido' + (data ? ' ' + data : '');
      var achados = idx[imei];
      if (!achados || !achados.length) { naoAchou.push([imei, data || '']); continue; }
      achados.forEach(function (a) {
        if (a.ja && /vendido/i.test(a.ja)) { jaTinha++; return; }
        if (!teste) { var f = {}; f[a.campo] = txt; vdFeishu_('PUT', '/bitable/v1/apps/' + app + '/tables/' + a.tid + '/records/' + a.rid, { fields: f }); }
        a.ja = txt; marcou++;
      });
    }
  }

  if (!teste) escreveNaoEncontrados_(fontes[0].ss, naoAchou);

  console.log((teste ? 'TESTE (nada gravado): ' : '') + lin + ' vendas lidas em ' + fontes.length + ' planilha(s) · ' +
    marcou + (teste ? ' seriam marcadas' : ' marcadas') + ' como Vendido na Base (' + tabelas.map(function (t) { return t.nome; }).join(', ') + ') · ' +
    jaTinha + ' já estavam marcadas · ' + naoAchou.length + ' não encontradas na Base' +
    (!teste && naoAchou.length ? ' (veja a aba "Não encontrados Feishu")' : '') +
    (parou ? ' — parou no limite de tempo: rode de novo.' : '') + '.');
}
