/**
 * Painel VM Shatan — leva as fotos APROVADAS para a planilha do Feishu.
 * Arquivo extra no mesmo Apps Script do Sincronizar.gs (usa as funções supa_ de lá).
 * Propriedades do script: FEISHU_APP_ID, FEISHU_APP_SECRET, FEISHU_PLANILHA (código da planilha).
 *
 * 1) Rode feishuTeste()  → manda UMA loja aprovada e mostra no registro o que fez.
 * 2) Conferiu no Feishu? Rode feishuLigar() → passa a mandar sozinho a cada 10 min.
 */
const FEISHU_API = 'https://open.feishu.cn/open-apis';
const FEISHU_LOTE = 15; // aprovações por rodada

function feishuTeste() { return enviarFeishu_(1, true); }
function feishuLigar() {
  ScriptApp.getProjectTriggers().forEach((t) => { if (t.getHandlerFunction() === 'sincronizarFeishu') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sincronizarFeishu').timeBased().everyMinutes(10).create();
  console.log('Envio automático para o Feishu ligado (a cada 10 min).');
}
function feishuDesligar() {
  ScriptApp.getProjectTriggers().forEach((t) => { if (t.getHandlerFunction() === 'sincronizarFeishu') ScriptApp.deleteTrigger(t); });
  console.log('Envio automático para o Feishu desligado.');
}
// Depois do teste, antes de apontar para a planilha oficial: faz o script reenviar tudo do zero.
function feishuRecomecar() {
  supa_('PATCH', '/rest/v1/revisoes?feishu_em=not.is.null', { feishu_em: null, feishu_info: null }, { Prefer: 'return=minimal' });
  // as aprovações em lote ("Validada antes no painel") são de lojas cujas fotos já estão coladas na planilha: não reenviar
  supa_('PATCH', '/rest/v1/revisoes?obs=eq.' + encodeURIComponent('Validada antes no painel (planilha Excel)'), { feishu_em: new Date().toISOString(), feishu_info: 'já estava na planilha (aprovada em lote)' }, { Prefer: 'return=minimal' });
  console.log('Pronto: as aprovações feitas uma a uma serão enviadas para a planilha em FEISHU_PLANILHA.');
}
function sincronizarFeishu() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try { enviarFeishu_(FEISHU_LOTE, false); } finally { lock.releaseLock(); }
}

function enviarFeishu_(limite, detalhado) {
  const props = PropertiesService.getScriptProperties();
  const planilha = String(props.getProperty('FEISHU_PLANILHA') || '').trim();
  if (!planilha) throw new Error('Falta a propriedade FEISHU_PLANILHA.');
  const tk = feishuToken_();

  // abas da planilha do Feishu
  const abas = feishu_(tk, 'GET', '/sheets/v3/spreadsheets/' + planilha + '/sheets/query').sheets || [];
  if (detalhado) abas.forEach((a) => console.log('Aba no Feishu: ' + a.title));
  const cacheAba = {};

  // aprovações ainda não enviadas, com fotos já copiadas
  const pend = supa_('GET', '/rest/v1/revisoes?select=key,status,canal_corrigido,loja_corrigida&status=eq.Aprovado&feishu_em=is.null&order=revisado_em.desc&limit=100');
  if (!pend.length) { console.log('Nada novo para mandar ao Feishu.'); return; }
  const envioDe = (key) => (supa_('GET', '/rest/v1/envios?select=key,form,canal,loja_raw,fotos,produtos,mesa&key=eq.' + encodeURIComponent(key)) || [])[0];

  let feitos = 0;
  for (let i = 0; i < pend.length && feitos < limite; i++) {
    const key = pend[i].key, e = envioDe(key);
    if (e && pend[i].canal_corrigido) e.canal = pend[i].canal_corrigido;   // código da loja corrigido no site
    if (e && pend[i].loja_corrigida) e.loja_raw = pend[i].loja_corrigida;
    let info = '';
    try {
      if (!e) { info = 'envio não encontrado'; marcar_(key, info); continue; }
      const fotos = (e.fotos || []).filter((p) => p.path);
      if (!fotos.length) { if ((e.fotos || []).length) continue; info = 'sem fotos'; marcar_(key, info); continue; } // ainda copiando do Drive
      const canal = canalFeishu_(e.canal);
      const aba = abas.find((a) => semAcento_(a.title) === semAcento_(canal));
      if (!aba) { info = 'aba "' + canal + '" não existe no Feishu'; marcar_(key, info); if (detalhado) console.log(key + ': ' + info); continue; }
      const A = cacheAba[aba.sheet_id] || (cacheAba[aba.sheet_id] = lerAba_(tk, planilha, aba));
      const linha = acharLinha_(A, e.loja_raw);
      if (!linha) { info = 'loja "' + e.loja_raw + '" não encontrada na aba ' + aba.title; marcar_(key, info); if (detalhado) console.log(key + ': ' + info); continue; }
      const destino = distribuir_(fotos, A.fotoCols);
      const feitas = [], falhas = [];
      destino.forEach((d) => {
        const blob = fotoParaFeishu_(d.foto);
        const bytes = Array.from(blob.getBytes(), (b) => b & 0xff);
        const cel = aba.sheet_id + '!' + letra_(d.col) + linha + ':' + letra_(d.col) + linha;
        const corpo = { range: cel, image: bytes, name: d.foto.path };
        try {
          try { feishu_(tk, 'POST', '/sheets/v2/spreadsheets/' + planilha + '/values_image', corpo); }
          catch (e1) { Utilities.sleep(3000); feishu_(tk, 'POST', '/sheets/v2/spreadsheets/' + planilha + '/values_image', corpo); } // 2ª tentativa
          feitas.push(letra_(d.col) + linha + ' ← ' + d.foto.label);
        } catch (e2) { falhas.push(letra_(d.col) + linha + ' (' + d.foto.label.split(' · ')[0] + ')'); console.warn('Célula ' + letra_(d.col) + linha + ' da aba ' + aba.title + ' recusou a foto: ' + e2); }
      });
      let caixasTxt = '';
      try { caixasTxt = marcarCaixas_(tk, planilha, aba, A, linha, e); } catch (e3) { console.warn('Caixas de seleção da linha ' + linha + ': ' + e3); caixasTxt = ' · caixas: erro (' + String(e3).replace(/^Error: /, '').slice(0, 120) + ')'; }
      if (!feitas.length && falhas.length) throw new Error('o Feishu recusou as fotos em ' + aba.title + ' ' + falhas.join(', ') + ' — confira se essas células não estão mescladas ou protegidas');
      info = aba.title + ' linha ' + linha + ': ' + (feitas.join('; ') || 'nenhuma coluna de foto') + (falhas.length ? ' · NÃO ENTROU: ' + falhas.join(', ') : '') + caixasTxt;
      marcar_(key, info, true);
      feitos++;
      console.log(key + ' → ' + info);
    } catch (err) {
      console.warn(key + ': ' + err);
      if (detalhado) throw err;
      // conta tentativas; depois de 3 rodadas com erro, desiste e registra o motivo
      try {
        const atual = (supa_('GET', '/rest/v1/revisoes?select=feishu_info&key=eq.' + encodeURIComponent(key)) || [])[0] || {};
        const n = ((String(atual.feishu_info || '').match(/^tentativa (\d+)/) || [])[1] | 0) + 1;
        if (n >= 3) marcar_(key, 'erro do Feishu depois de 3 tentativas: ' + String(err).slice(0, 150));
        else supa_('PATCH', '/rest/v1/revisoes?key=eq.' + encodeURIComponent(key), { feishu_info: 'tentativa ' + n + ': ' + String(err).slice(0, 150) }, { Prefer: 'return=minimal' });
      } catch (e2) { /* segue */ }
    }
  }
  console.log('Enviadas ao Feishu nesta rodada: ' + feitos);
}

// ---------- leitura da aba ----------
function lerAba_(tk, planilha, aba) {
  const nCol = Math.min((aba.grid_properties && aba.grid_properties.column_count) || 40, 60);
  const nLin = (aba.grid_properties && aba.grid_properties.row_count) || 500;
  const fim = letra_(nCol);
  const cab = valores_(tk, planilha, aba.sheet_id + '!A1:' + fim + '3');
  const H = {}; // coluna -> textos do cabeçalho
  cab.forEach((row) => (row || []).forEach((v, c) => { const s = texto_(v); if (s) (H[c + 1] = H[c + 1] || []).push(s); }));
  const achar = (fn) => { const cs = Object.keys(H).map(Number).sort((a, b) => a - b); for (const c of cs) if (H[c].some(fn)) return c; return null; };
  const colCod = achar((v) => /store code|código|codigo|门店编码|门店代码/i.test(v));
  const colNome = achar((v) => /store name/i.test(v));
  const fotoCols = Object.keys(H).map(Number).filter((c) => H[c].some((v) => v.indexOf('照片') >= 0)).sort((a, b) => a - b)
    .map((c) => ({ col: c, h: H[c].join(' ') }));
  const cols = [colCod, colNome].filter(Boolean);
  const dados = {};
  cols.forEach((c) => { const vs = valores_(tk, planilha, aba.sheet_id + '!' + letra_(c) + '1:' + letra_(c) + nLin); dados[c] = vs.map((r) => texto_(r && r[0])); });
  console.log('Aba ' + aba.title + ': código na coluna ' + (colCod ? letra_(colCod) : '—') + ', nome na ' + (colNome ? letra_(colNome) : '—') +
    ', fotos em ' + fotoCols.map((f) => letra_(f.col)).join(', '));
  const caixas = {}; // nome normalizado do cabeçalho -> coluna
  Object.keys(H).forEach((c) => H[c].forEach((v) => { const k = normItem_(v); if (k && !caixas[k]) caixas[k] = +c; }));
  return { colCod, colNome, fotoCols, dados, nLin, caixas };
}
function acharLinha_(A, codigoForms) {
  const alvo = chaves_(codigoForms); if (!alvo.length) return null;
  const linhas = A.dados[A.colCod] || A.dados[A.colNome] || [];
  // 1) código igual (completo); 2) só os números, se for único na aba
  const full = [], dig = [];
  for (let r = 1; r < linhas.length; r++) {
    const k = chaves_(linhas[r]); const nomeK = A.colNome && A.dados[A.colNome] ? chaves_(String(A.dados[A.colNome][r] || '').split(/[\s-]/)[0]) : [];
    if (k[0] && k[0] === alvo[0]) full.push(r + 1);
    if ((k[1] && k[1] === alvo[1]) || (nomeK[1] && nomeK[1] === alvo[1])) dig.push(r + 1);
  }
  if (full.length === 1) return full[0];
  if (dig.length === 1) return dig[0];
  return null;
}
function chaves_(raw) {
  const u = String(raw || '').toUpperCase().trim(); if (!u) return [];
  const m = u.match(/\d+/g);
  return [u.replace(/[^A-Z0-9]/g, '').replace(/^0+/, ''), m ? (m[m.length - 1].replace(/^0+/, '') || '0') : null];
}
// qual foto do Forms vai em qual coluna de foto da planilha
function distribuir_(fotos, cols) {
  const tipo = (lbl) => { const u = semAcento_(lbl);
    if (/entrada/.test(u)) return 'fachada'; if (/frente/.test(u)) return 'frente';
    if (/lado 2|outro lado/.test(u)) return 'tras'; if (/lado 1|mesa completa|aparelhos/.test(u)) return 'dentro'; return null; };
  const tipoCol = (h) => { const u = semAcento_(h);
    if (/frente|前面/.test(u)) return 'frente'; if (/atras|后面/.test(u)) return 'tras'; if (/da loja|店面/.test(u) && !/dentro/.test(u)) return 'fachada';
    if (/dentro|店内/.test(u)) return 'dentro'; return 'outra'; };
  const livres = cols.map((c) => ({ col: c.col, t: tipoCol(c.h), usada: false }));
  const out = [];
  fotos.forEach((f) => { const t = tipo(f.label); if (!t) return; const c = livres.find((x) => !x.usada && x.t === t); if (c) { c.usada = true; out.push({ col: c.col, foto: f }); } });
  fotos.forEach((f) => { if (out.some((o) => o.foto === f) || !tipo(f.label)) return; const c = livres.find((x) => !x.usada); if (c) { c.usada = true; out.push({ col: c.col, foto: f }); } });
  return out;
}

// ---------- caixas de seleção (aparelhos, mesa, displays) ----------
function normItem_(v) { return String(v || '').toLowerCase().replace(/renho/g, 'reno').replace(/[\s\-_.·()（）]/g, ''); }
// o que o envio diz que existe na loja -> nomes de cabeçalho possíveis no Feishu
function itensDoEnvio_(e) {
  const out = new Set();
  const add = (...ks) => ks.forEach((k) => out.add(normItem_(k)));
  String(e.produtos || '').split(',').forEach((x) => { const k = normItem_(x); if (k) out.add(k); });
  const m = semAcento_(e.mesa);
  if (/1[.,]6/.test(m)) add('1.6M 体验台'); if (/1[.,]2/.test(m)) add('1.2M 体验台'); if (/0[.,]8/.test(m)) add('0.8M 体验台');
  if (/display find/.test(m)) add('find display');
  if (/display reno/.test(m)) add('reno display', '桌面reno发光字');
  if (/a series/.test(m)) add('A seris display', 'A series display');
  if (/display iot/.test(m)) add('IOT display');
  if (/a4 display/.test(m)) add('A4立牌');
  if (/luminoso/.test(m)) add('Logo发光字');
  (e.fotos || []).forEach((f) => { const l = semAcento_(f.label);
    if (/totem/.test(l)) add('人行立牌');
    if (/display a4/.test(l)) add('A4立牌');
    if (/display reno/.test(l)) add('桌面reno发光字', 'reno display');
    if (/display oppo logo/.test(l)) add('Logo发光字');
    if (/^logo de mesa/.test(l)) add('Logo 发光灯柱'); });
  return out;
}
const CAIXAS_GERENCIADAS_ = ['findx9pro', 'reno16', 'reno16f', 'reno14', 'reno14f', 'a6pro', 'a6x5g', 'a55g', 'a5pro', '1.6m体验台', '1.2m体验台', '0.8m体验台',
  '人行立牌', '桌面reno发光字', 'logo发光字', 'a4立牌', 'logo发光灯柱', 'finddisplay', 'renodisplay', 'aserisdisplay', 'aseriesdisplay', 'iotdisplay'].map(normItem_);
function marcarCaixas_(tk, planilha, aba, A, linha, e) {
  const tem = itensDoEnvio_(e);
  const desmarcar = String(PropertiesService.getScriptProperties().getProperty('FEISHU_DESMARCAR') || '').trim().toLowerCase() === 'sim';
  const ranges = [], ligadas = [];
  CAIXAS_GERENCIADAS_.forEach((k) => {
    const col = A.caixas[k]; if (!col) return;
    const sim = tem.has(k);
    if (!sim && !desmarcar) return;
    ranges.push({ range: aba.sheet_id + '!' + letra_(col) + linha + ':' + letra_(col) + linha, values: [[sim ? 1 : 0]] }); // no Feishu a caixa marcada vale 1 e a desmarcada 0
    if (sim) ligadas.push(letra_(col));
  });
  if (!ranges.length) return '';
  feishu_(tk, 'POST', '/sheets/v2/spreadsheets/' + planilha + '/values_batch_update', { valueRanges: ranges });
  return ' · caixas marcadas: ' + (ligadas.join(', ') || 'nenhuma') + (desmarcar ? ' (demais desmarcadas)' : '');
}

// ---------- utilidades ----------
function feishuToken_() {
  const p = PropertiesService.getScriptProperties();
  const id = String(p.getProperty('FEISHU_APP_ID') || '').trim(), sec = String(p.getProperty('FEISHU_APP_SECRET') || '').trim();
  if (!id || !sec) throw new Error('Faltam FEISHU_APP_ID e/ou FEISHU_APP_SECRET nas Propriedades do script.');
  const r = JSON.parse(UrlFetchApp.fetch(FEISHU_API + '/auth/v3/tenant_access_token/internal', { method: 'post', contentType: 'application/json; charset=utf-8', payload: JSON.stringify({ app_id: id, app_secret: sec }), muteHttpExceptions: true }).getContentText());
  if (r.code !== 0) throw new Error('Feishu recusou o login do app: ' + r.msg + ' (App ID lido: "' + id.slice(0, 6) + '…", ' + id.length + ' caracteres; App Secret com ' + sec.length + ' caracteres)');
  return r.tenant_access_token;
}
function feishu_(tk, method, path, body) {
  const res = UrlFetchApp.fetch(FEISHU_API + path, { method: method.toLowerCase(), contentType: 'application/json; charset=utf-8', headers: { Authorization: 'Bearer ' + tk },
    payload: body === undefined ? undefined : JSON.stringify(body), muteHttpExceptions: true });
  let j = {}; try { j = JSON.parse(res.getContentText()); } catch (e) { /* */ }
  if (res.getResponseCode() >= 300 || j.code !== 0) {
    const dica = j.code === 91403 || j.code === 1310213 || res.getResponseCode() === 403 ? ' (o app precisa estar adicionado na planilha com permissão de edição)' : '';
    throw new Error('Feishu ' + path.split('?')[0] + ' → ' + (j.code || res.getResponseCode()) + ' ' + (j.msg || res.getContentText().slice(0, 200)) + dica);
  }
  return j.data || {};
}
function valores_(tk, planilha, range) {
  const d = feishu_(tk, 'GET', '/sheets/v2/spreadsheets/' + planilha + '/values/' + encodeURIComponent(range) + '?valueRenderOption=ToString');
  return (d.valueRange && d.valueRange.values) || [];
}
function texto_(v) {
  if (v == null) return '';
  if (Array.isArray(v)) return v.map((x) => (x && (x.text || x.link)) || (typeof x === 'string' ? x : '')).join('').trim();
  if (typeof v === 'object') return String(v.text || '').trim();
  return String(v).trim();
}
// versão reduzida (800 px) direto do Drive, para não pesar no Feishu; se não der, usa a cópia do Supabase
function fotoParaFeishu_(foto) {
  try {
    const meta = JSON.parse(UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + foto.id + '?fields=thumbnailLink',
      { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getContentText());
    if (meta.thumbnailLink) {
      const r = UrlFetchApp.fetch(meta.thumbnailLink.replace(/=s\d+$/, '=s800'), { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
      if (r.getResponseCode() === 200 && r.getBlob().getBytes().length > 5000) return r.getBlob();
    }
  } catch (e) { /* usa o Supabase */ }
  return fotoSupabase_(foto.path);
}
function fotoSupabase_(path) {
  const res = UrlFetchApp.fetch(SUPABASE_URL + '/storage/v1/object/fotos/' + encodeURIComponent(path), { headers: cabecalhos_({}), muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('foto ' + path + ' não encontrada no Supabase (' + res.getResponseCode() + ')');
  return res.getBlob();
}
function marcar_(key, info, ok) {
  supa_('PATCH', '/rest/v1/revisoes?key=eq.' + encodeURIComponent(key), { feishu_em: new Date().toISOString(), feishu_info: (ok ? '' : 'NÃO ENVIADA: ') + info }, { Prefer: 'return=minimal' });
}
function canalFeishu_(c) {
  const t = semAcento_(c);
  const M = { fastshop: 'Fast', fast: 'Fast', 'magazine luiza': 'Magalu', magalu: 'Magalu', paraiba: 'Paraiba', americanas: 'Americanas', claro: 'Claro', vivo: 'Vivo', tim: 'Tim', bemol: 'Bemol', martinello: 'Martinello', lider: 'Lider' };
  return M[t] || String(c || '').trim();
}
function semAcento_(s) { return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim(); }
function letra_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

// Diagnóstico: mostra como o Feishu devolve as caixas de seleção (para acertar a gravação)
function feishuVerCaixas() {
  const planilha = String(PropertiesService.getScriptProperties().getProperty('FEISHU_PLANILHA') || '').trim();
  const tk = feishuToken_();
  const abas = feishu_(tk, 'GET', '/sheets/v3/spreadsheets/' + planilha + '/sheets/query').sheets || [];
  const aba = abas.find((a) => semAcento_(a.title) === 'magalu') || abas[0];
  ['ToString', 'FormattedValue', 'UnformattedValue', 'Formula'].forEach((modo) => {
    try {
      const d = feishu_(tk, 'GET', '/sheets/v2/spreadsheets/' + planilha + '/values/' + encodeURIComponent(aba.sheet_id + '!K3:AC6') + '?valueRenderOption=' + modo);
      console.log(modo + ': ' + JSON.stringify((d.valueRange || {}).values).slice(0, 900));
    } catch (e) { console.log(modo + ': erro ' + e); }
  });
}
