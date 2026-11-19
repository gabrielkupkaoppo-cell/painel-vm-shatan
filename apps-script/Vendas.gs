/**
 * Painel VM Shatan — marca os IMEIs vendidos na Base do Feishu.
 * Oitavo arquivo do mesmo Apps Script (usa feishuToken_/feishu_ do Feishu.gs).
 *
 * Lê a planilha de VENDAS do Google (IMEI na coluna E, data da venda na coluna G) e,
 * para cada IMEI, escreve "Vendido <data>" na coluna VENDIDO (已售) da Base do Feishu,
 * tanto no Sheet2 (estoque do escritório) quanto na Transferência (foi para a loja).
 * Os IMEIs que não existirem na Base vão para a aba "Não encontrados Feishu" da planilha.
 *
 * O script procura sozinho, em todas as abas da Base, as que têm uma coluna de IMEI e
 * uma coluna de VENDIDO — por isso não precisa configurar nomes de aba nem de coluna.
 *
 * Propriedades do script (⚙️ Configurações do projeto → Propriedades do script):
 *   VENDAS_PLANILHA = link (ou ID) da planilha de vendas do Google.
 *   VENDAS_ABA      = (opcional) nome da aba das vendas. Em branco = a primeira aba.
 *   FEISHU_ESTOQUE  = link da Base do Feishu (o link do navegador, começa com .../wiki/... ou .../base/...).
 *
 * Permissões do app no Feishu (as duas são verdes/免审, entram sozinhas):
 *   bitable:app            (ver, editar e gerenciar Bases)
 *   wiki:wiki:readonly     (a Base está dentro de um Wiki; serve para abrir a Base)
 * E o app "Painel VM Shatan" precisa estar como EDITOR na Base (abrir a Base → Compartilhar → adicionar o app).
 *
 * 1) marcarVendidosTeste() → NÃO grava nada: mostra quantos marcaria e quantos não achou.
 * 2) marcarVendidos()      → grava de verdade.
 * 3) marcarVendidosLigar() → roda sozinho a cada hora. marcarVendidosDesligar() para parar.
 */
var VENDAS_LIN_INI = 4;    // primeira linha de dados na planilha de vendas (as de cima são cabeçalho)
var VENDAS_COL_IMEI = 5;   // coluna E
var VENDAS_COL_DATA = 7;   // coluna G

function marcarVendidosTeste() { marcarVendidos_(true); }
function marcarVendidos() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) throw new Error('Outra rotina está rodando agora. Tente de novo em 1 minuto.');
  try { marcarVendidos_(false); } finally { lock.releaseLock(); }
}
function marcarVendidosDiario() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try { marcarVendidos_(false); } finally { lock.releaseLock(); }
}
function marcarVendidosLigar() {
  ScriptApp.getProjectTriggers().forEach((t) => { if (t.getHandlerFunction() === 'marcarVendidosDiario') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('marcarVendidosDiario').timeBased().everyHours(1).create();
  console.log('Baixa de vendidos no Feishu ligada (a cada hora).');
}
function marcarVendidosDesligar() {
  ScriptApp.getProjectTriggers().forEach((t) => { if (t.getHandlerFunction() === 'marcarVendidosDiario') ScriptApp.deleteTrigger(t); });
  console.log('Baixa de vendidos no Feishu desligada.');
}

// só os dígitos (serve para comparar IMEI escrito como número ou como texto); tira zeros à esquerda
function soDigitos_(v) {
  if (v === null || v === undefined || v instanceof Date) return '';
  const s = String(v).replace(/\D/g, '');
  return s.replace(/^0+/, '') || s;
}
// valor de um campo do Feishu vira texto limpo
function campoTexto_(val) {
  if (val === null || val === undefined) return '';
  if (typeof val === 'string' || typeof val === 'number') return String(val);
  if (Array.isArray(val)) return val.map((x) => (x && (x.text || x.name)) || '').join('');
  if (typeof val === 'object') return String(val.text || val.name || '');
  return String(val);
}
function formataData_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'America/Sao_Paulo', 'dd/MM/yyyy');
  return String(v || '').trim();
}

// ---------- planilha de vendas (Google) ----------
function vendasPlanilha_() {
  const link = String(PropertiesService.getScriptProperties().getProperty('VENDAS_PLANILHA') || '').trim();
  if (!link) throw new Error('Falta a propriedade VENDAS_PLANILHA (link da planilha de vendas do Google).');
  const m = link.match(/\/d\/([A-Za-z0-9_-]+)/);
  return SpreadsheetApp.openById(m ? m[1] : link);
}
function vendasAba_(ss) {
  const nome = String(PropertiesService.getScriptProperties().getProperty('VENDAS_ABA') || '').trim();
  if (nome) { const a = ss.getSheetByName(nome); if (!a) throw new Error('Não achei a aba "' + nome + '" na planilha de vendas.'); return a; }
  return ss.getSheets()[0];
}
function escreveNaoEncontrados_(ss, linhas) {
  const a = ss.getSheetByName('Não encontrados Feishu') || ss.insertSheet('Não encontrados Feishu');
  a.clear();
  a.getRange(1, 1, 1, 3).setValues([['IMEI', 'Data da venda', 'Conferido em']]);
  a.getRange(2, 1, Math.max(linhas.length, 1), 1).setNumberFormat('@');   // IMEI como texto
  if (linhas.length) {
    const agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');
    a.getRange(2, 1, linhas.length, 3).setValues(linhas.map((l) => [l[0], l[1], agora]));
  }
}

// ---------- Base do Feishu (Bitable) ----------
function estoqueBase_(tk) {
  const link = String(PropertiesService.getScriptProperties().getProperty('FEISHU_ESTOQUE') || '').trim();
  if (!link) throw new Error('Falta a propriedade FEISHU_ESTOQUE com o link da Base do Feishu.');
  const m = link.match(/\/(wiki|base|app)\/([A-Za-z0-9]+)/);
  let token = m ? m[2] : (/^[A-Za-z0-9]{10,}$/.test(link) ? link : '');
  if (!token) throw new Error('Não entendi o link da Base no FEISHU_ESTOQUE.');
  if (m && m[1] === 'wiki') {
    const r = feishu_(tk, 'GET', '/wiki/v2/spaces/get_node?obj_type=wiki&token=' + token);
    token = r.node && r.node.obj_token;
    if (!token) throw new Error('Não consegui abrir a Base do Wiki. Confira a permissão wiki:wiki:readonly e se o app está na Base.');
  }
  return token;   // app_token da Base
}
function listaFeishu_(tk, path, chave) {
  let out = [], page = '';
  do {
    const r = feishu_(tk, 'GET', path + (path.indexOf('?') >= 0 ? '&' : '?') + 'page_size=' + (chave === 'records' ? 500 : 100) + (page ? '&page_token=' + page : ''));
    (r.items || []).forEach((x) => out.push(x));
    page = r.has_more ? r.page_token : '';
  } while (page);
  return out;
}
function achaCampo_(campos, re) { for (let i = 0; i < campos.length; i++) { if (re.test(campos[i].field_name)) return campos[i].field_name; } return ''; }

function marcarVendidos_(teste) {
  const inicio = Date.now(), limite = 5 * 60 * 1000;
  const ss = vendasPlanilha_(), aba = vendasAba_(ss);
  const tk = feishuToken_(), app = estoqueBase_(tk);

  // abas da Base que têm IMEI e VENDIDO
  const tabelas = listaFeishu_(tk, '/bitable/v1/apps/' + app + '/tables', 'tables').map((t) => {
    const campos = listaFeishu_(tk, '/bitable/v1/apps/' + app + '/tables/' + t.table_id + '/fields', 'fields');
    return { id: t.table_id, nome: t.name, imei: achaCampo_(campos, /imei/i), vendido: achaCampo_(campos, /vendido|已售|已卖/i) };
  }).filter((t) => t.imei && t.vendido);
  if (!tabelas.length) throw new Error('Nenhuma aba da Base tem as colunas IMEI e VENDIDO. Confira os nomes das colunas.');

  // índice IMEI -> ocorrências na Base
  const idx = {};
  tabelas.forEach((t) => {
    listaFeishu_(tk, '/bitable/v1/apps/' + app + '/tables/' + t.id + '/records', 'records').forEach((rec) => {
      const d = soDigitos_(campoTexto_(rec.fields[t.imei]));
      if (d.length >= 10) (idx[d] = idx[d] || []).push({ tid: t.id, campo: t.vendido, rid: rec.record_id, ja: campoTexto_(rec.fields[t.vendido]) });
    });
  });

  // lê as vendas
  const ultima = aba.getLastRow();
  const naoAchou = []; let marcou = 0, jaTinha = 0, lin = 0, parou = false;
  if (ultima >= VENDAS_LIN_INI) {
    const vals = aba.getRange(VENDAS_LIN_INI, 1, ultima - VENDAS_LIN_INI + 1, Math.max(VENDAS_COL_DATA, VENDAS_COL_IMEI)).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (Date.now() - inicio > limite) { parou = true; break; }
      const imei = soDigitos_(vals[i][VENDAS_COL_IMEI - 1]);
      if (imei.length < 10) continue;
      lin++;
      const data = formataData_(vals[i][VENDAS_COL_DATA - 1]);
      const txt = 'Vendido' + (data ? ' ' + data : '');
      const achados = idx[imei];
      if (!achados || !achados.length) { naoAchou.push([imei, data || '']); continue; }
      achados.forEach((a) => {
        if (a.ja && /vendido/i.test(a.ja)) { jaTinha++; return; }
        if (!teste) { const f = {}; f[a.campo] = txt; feishu_(tk, 'PUT', '/bitable/v1/apps/' + app + '/tables/' + a.tid + '/records/' + a.rid, { fields: f }); }
        a.ja = txt; marcou++;   // marca no índice para não repetir se o IMEI aparecer 2x na planilha
      });
    }
  }

  if (!teste) escreveNaoEncontrados_(ss, naoAchou);

  console.log((teste ? 'TESTE (nada gravado): ' : '') + lin + ' vendas lidas · ' +
    marcou + (teste ? ' seriam marcadas' : ' marcadas') + ' como Vendido na Base (' + tabelas.map((t) => t.nome).join(', ') + ') · ' +
    jaTinha + ' já estavam marcadas · ' + naoAchou.length + ' não encontradas na Base' +
    (!teste && naoAchou.length ? ' (veja a aba "Não encontrados Feishu")' : '') +
    (parou ? ' — parou no limite de tempo: rode de novo.' : '') + '.');
}
