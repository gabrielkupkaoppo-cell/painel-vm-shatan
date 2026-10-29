/**
 * Painel VM Shatan — sincroniza os 2 Google Forms com o Supabase.
 * Cole este arquivo no Apps Script da planilha "Controle de Mesas Shatan"
 * (Extensões → Apps Script), preencha SUPABASE_KEY nas Propriedades do script
 * e rode a função instalar() uma vez.
 */
const SUPABASE_URL = 'https://crhkkumzfgayysgcnlhe.supabase.co';
const FORMS = [
  { nome: 'Controle de Mesas', id: '1_SYs6plPs8T-8vsG_ZEOb3O_niz6GkHjL9eoNyK5oXU', aba: 'Respostas ao formulário 1', prefixo: '', notas: [27, 28] },
  { nome: 'Checklist Fast/Operadoras', id: '1JesHMataw8iB-abJpNBCK_MuKXlqwOs8Sn2a6NVv-3E', aba: 'Respostas ao formulário 1', prefixo: 'F2|', notas: [] },
];
const ABA_REVISAO = 'Revisão';
const LIMITE_MS = 4.5 * 60 * 1000; // para antes do limite de 6 min do Apps Script

function instalar() {
  ScriptApp.getProjectTriggers().forEach((t) => { if (t.getHandlerFunction() === 'sincronizar') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sincronizar').timeBased().everyMinutes(10).create();
  sincronizar();
}

function sincronizar() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  const inicio = Date.now();
  try {
    const existentes = {};
    const antigosEnvios = supa_('GET', '/rest/v1/envios?select=key,form,carimbo,nome,loja_raw,fotos&limit=20000');
    const fotoPorId = {}; // o mesmo arquivo do Drive guarda o que já foi feito (cópia, impressão digital, data) mesmo se a linha mudar
    antigosEnvios.forEach((e) => { existentes[e.key] = e.fotos || []; (e.fotos || []).forEach((p) => { if (p && p.id && (p.path || p.md5 || p.arquivada)) fotoPorId[p.id] = p; }); });

    // 1) envios dos 2 Forms
    const envios = [];
    FORMS.forEach((F) => {
      const sh = SpreadsheetApp.openById(F.id).getSheetByName(F.aba);
      if (!sh) return;
      const v = sh.getDataRange().getDisplayValues();
      const header = v[0] || [];
      for (let r = 1; r < v.length; r++) {
        const row = v[r];
        const carimbo = String(row[0] || '').trim();
        const nome = String(row[1] || '').trim(), loja = String(row[5] || '').trim();
        const data = parseData_(carimbo);
        if (!data || (!nome && !loja)) continue;
        const key = F.prefixo + carimbo + '#' + (r + 1);
        const antigas = {}; (existentes[key] || []).forEach((p) => { antigas[p.id] = p; });
        Object.keys(fotoPorId).length && driveIds_(row.slice(9).join(' ')).forEach((id) => { if (!antigas[id] && fotoPorId[id]) antigas[id] = fotoPorId[id]; });
        const fotos = [];
        for (let c = 9; c < row.length; c++) {
          const ids = driveIds_(row[c]);
          ids.forEach((id, k) => fotos.push({ id: id, label: rotulo_(header[c] || 'Foto') + (ids.length > 1 ? ' (' + (k + 1) + ')' : ''), path: (antigas[id] || {}).path || null, arquivada: !!(antigas[id] || {}).arquivada, px: (antigas[id] || {}).px || null, md5: (antigas[id] || {}).md5 || null, tirada: (antigas[id] || {}).tirada != null ? antigas[id].tirada : null }));
        }
        envios.push({
          key: key, form: F.nome, linha: r + 1, carimbo: carimbo, data: data.toISOString(),
          nome: nome, cargo: String(row[2] || '').trim(), tel: String(row[3] || '').trim(),
          canal: String(row[4] || '').trim(), loja_raw: loja, loja: codigo_(loja), uf: String(row[6] || '').trim().toUpperCase(),
          produtos: String(row[7] || '').trim(), mesa: String(row[8] || '').trim(),
          notas: F.notas.map((i) => String(row[i] || '').trim()).filter(String).join(' · '),
          fotos: fotos, atualizado_em: new Date().toISOString(),
        });
      }
    });
    for (let i = 0; i < envios.length; i += 200) {
      supa_('POST', '/rest/v1/envios?on_conflict=key', envios.slice(i, i + 200), { Prefer: 'resolution=merge-duplicates,return=minimal' });
    }

    // 1b) linhas que mudaram de lugar ou foram apagadas na planilha de respostas:
    //     a revisão passa para a chave nova e o envio antigo é removido (evita envio duplicado e falsa "foto repetida")
    const atuais = {}; envios.forEach((e) => { atuais[e.key] = e; });
    const velhos = antigosEnvios.filter((e) => !atuais[e.key]);
    if (velhos.length && envios.length >= 0.9 * antigosEnvios.length) {
      const porAssinatura = {}; envios.forEach((e) => { porAssinatura[[e.form, e.carimbo, e.nome, e.loja_raw].join('|')] = e.key; });
      const revKeys = {}; supa_('GET', '/rest/v1/revisoes?select=key,revisado_em&limit=20000').forEach((r) => { revKeys[r.key] = r; });
      velhos.forEach((v) => {
        const novo = porAssinatura[[v.form, v.carimbo, v.nome, v.loja_raw].join('|')];
        const rv = revKeys[v.key];
        if (rv) {
          if (novo && !revKeys[novo]) { supa_('PATCH', '/rest/v1/revisoes?key=eq.' + encodeURIComponent(v.key), { key: novo }, { Prefer: 'return=minimal' }); revKeys[novo] = rv; }
          else supa_('DELETE', '/rest/v1/revisoes?key=eq.' + encodeURIComponent(v.key), undefined, { Prefer: 'return=minimal' });
        }
        supa_('DELETE', '/rest/v1/envios?key=eq.' + encodeURIComponent(v.key), undefined, { Prefer: 'return=minimal' });
      });
      console.log('Envios que mudaram de linha ou saíram da planilha: ' + velhos.length + ' (revisões levadas para a linha nova).');
    }

    // 2) decisões antigas da aba "Revisão" (Portal do Claude) → Supabase
    const aba = SpreadsheetApp.getActive().getSheetByName(ABA_REVISAO);
    if (aba && aba.getLastRow() > 1) {
      const linhas = aba.getRange(2, 1, aba.getLastRow() - 1, 12).getDisplayValues();
      const antigas = linhas.filter((l) => l[0] && atuais[l[0]] && (l[7] === 'Aprovado' || l[7] === 'Reprovado')).map((l) => ({
        key: l[0], status: l[7], motivos: l[8] || '', obs: l[9] || '', revisado_em: (parseData_(l[10]) || new Date()).toISOString(),
        revisado_por_email: l[11] || 'portal do Claude', origem: 'planilha',
      }));
      for (let i = 0; i < antigas.length; i += 200) {
        supa_('POST', '/rest/v1/revisoes?on_conflict=key', antigas.slice(i, i + 200), { Prefer: 'resolution=ignore-duplicates,return=minimal' });
      }
    }

    // 3) espelha todas as decisões do site na aba "Revisão"
    if (aba) {
      const porKey = {}; envios.forEach((e) => { porKey[e.key] = e; });
      const rev = supa_('GET', '/rest/v1/revisoes?select=*&order=revisado_em.asc&limit=20000');
      const out = rev.map((x) => {
        const e = porKey[x.key] || {};
        return [x.key, e.carimbo || '', e.linha || '', e.loja || '', e.canal || '', e.uf || '', e.nome || '', x.status, x.motivos || '', x.obs || '',
          Utilities.formatDate(new Date(x.revisado_em), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm'), x.revisado_por_email || ''];
      });
      aba.getRange(1, 1, 1, 12).setValues([['Chave do envio', 'Carimbo do envio', 'Linha na planilha', 'Loja (código)', 'Canal', 'Estado', 'Promotor', 'Status', 'Motivo', 'Observação', 'Revisado em', 'Revisado por']]);
      if (aba.getLastRow() > 1) aba.getRange(2, 1, aba.getLastRow() - 1, 12).clearContent();
      if (out.length) aba.getRange(2, 1, out.length, 12).setValues(out);
    }

    // 4) copia as fotos novas do Drive para o Supabase (aos poucos, a cada 10 min)
    const mudou = {};
    const ate = (min) => Date.now() - inicio < min * 60000;
    for (let i = 0; i < envios.length && ate(2.5); i++) {
      const e = envios[envios.length - 1 - i]; // mais recentes primeiro
      e.fotos.forEach((p) => {
        if (p.path || p.arquivada || !ate(2.5)) return;
        try {
          const blob = imagem_(p.id);
          const path = p.id + '.jpg';
          supaRaw_('POST', '/storage/v1/object/fotos/' + path, blob.getBytes(), { 'Content-Type': blob.getContentType() || 'image/jpeg', 'x-upsert': 'true' });
          p.path = path; p.px = 800; mudou[e.key] = e;
        } catch (err) { console.warn('foto ' + p.id + ': ' + err); }
      });
    }
    // 4b) "impressão digital" (md5) e data em que a foto foi tirada, direto do Drive (para achar foto repetida ou antiga)
    for (let i = 0; i < envios.length && ate(3.3); i++) {
      const e = envios[envios.length - 1 - i];
      e.fotos.forEach((p) => {
        if ((p.md5 && p.tirada != null) || !ate(3.3)) return;
        try {
          const m = JSON.parse(UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + p.id + '?fields=md5Checksum,imageMediaMetadata(time)', { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true }).getContentText());
          if (m.md5Checksum) { p.md5 = m.md5Checksum; mudou[e.key] = e; }
          // data em que a câmera tirou a foto ("2026:08:15 10:22:11"); vazio = foto sem essa informação (print, WhatsApp…)
          const t = m.imageMediaMetadata && m.imageMediaMetadata.time;
          const mt = t && String(t).match(/^(\d{4})[:-](\d{2})[:-](\d{2})[ T](\d{2}):(\d{2})/);
          if (m.md5Checksum || m.imageMediaMetadata) { p.tirada = mt ? mt[1] + '-' + mt[2] + '-' + mt[3] + ' ' + mt[4] + ':' + mt[5] : ''; mudou[e.key] = e; }
        } catch (err) { /* tenta na próxima rodada */ }
      });
    }
    const lista = Object.keys(mudou).map((k) => ({ key: k, fotos: mudou[k].fotos }));
    for (let i = 0; i < lista.length; i += 100) {
      lista.slice(i, i + 100).forEach((x) => supa_('PATCH', '/rest/v1/envios?key=eq.' + encodeURIComponent(x.key), { fotos: x.fotos }, { Prefer: 'return=minimal' }));
    }
    console.log('envios: ' + envios.length + ' · fotos copiadas agora: ' + lista.length + ' envios');
    if (ate(3.8)) limparFotos_(envios, inicio);
  } finally {
    lock.releaseLock();
  }
}

// ---------- utilidades ----------
function chave_() {
  const k = PropertiesService.getScriptProperties().getProperty('SUPABASE_KEY');
  if (!k) throw new Error('Falta a propriedade SUPABASE_KEY (Configurações do projeto → Propriedades do script).');
  return k;
}
function cabecalhos_(extra) {
  const k = chave_();
  const h = { apikey: k };
  if (k.indexOf('eyJ') === 0) h.Authorization = 'Bearer ' + k;
  Object.keys(extra || {}).forEach((n) => { h[n] = extra[n]; });
  return h;
}
function supaRaw_(method, path, payload, extra) {
  const res = UrlFetchApp.fetch(SUPABASE_URL + path, { method: method.toLowerCase(), headers: cabecalhos_(extra), payload: payload, muteHttpExceptions: true });
  const code = res.getResponseCode();
  if (code >= 300) throw new Error(method + ' ' + path.split('?')[0] + ' → ' + code + ' ' + res.getContentText().slice(0, 300));
  return res.getContentText();
}
function supa_(method, path, body, extra) {
  const h = Object.assign({ 'Content-Type': 'application/json' }, extra || {});
  const txt = supaRaw_(method, path, body === undefined ? undefined : JSON.stringify(body), h);
  return txt ? JSON.parse(txt) : null;
}
function imagem_(id) {
  // tenta uma versão reduzida (1280 px) pela miniatura do Drive; se não der, usa o arquivo original
  try {
    const meta = JSON.parse(UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id + '?fields=thumbnailLink,size',
      { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getContentText());
    if (meta.thumbnailLink) {
      const r = UrlFetchApp.fetch(meta.thumbnailLink.replace(/=s\d+$/, '=s800'), { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
      if (r.getResponseCode() === 200) return r.getBlob();
    }
  } catch (e) { /* segue para o original */ }
  return DriveApp.getFileById(id).getBlob();
}
function parseData_(s) {
  const m = String(s || '').match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})(?::(\d{2}))?/);
  return m ? new Date(+m[3], +m[2] - 1, +m[1], +m[4], +m[5], +(m[6] || 0)) : null;
}
function driveIds_(v) {
  const out = [];
  String(v || '').split(/[,\s]+/).forEach((p) => {
    const m = p.match(/[?&]id=([\w-]{10,})/) || p.match(/\/d\/([\w-]{10,})/);
    if (m) out.push(m[1]);
  });
  return out;
}
function codigo_(raw) {
  const m = String(raw || '').trim().match(/(\d[\d.]*)/g);
  return m ? m[m.length - 1].replace(/\./g, '') : String(raw || '').trim();
}
function rotulo_(h) {
  const u = String(h).toUpperCase();
  if (u.indexOf('ENTRADA') >= 0) return 'Entrada da loja · 门店入口';
  if (u.indexOf('FRENTE') >= 0) return 'Mesa de frente · 体验台正面';
  if (u.indexOf('OUTRO LADO') >= 0) return 'Mesa — lado 2 · 侧面2';
  if (u.indexOf('UM LADO') >= 0) return 'Mesa — lado 1 · 侧面1';
  if (u.indexOf('MESA COMPLETA') >= 0) return 'Mesa completa · 整台';
  if (u.indexOf('APARELHOS') >= 0) return 'Mesa com aparelhos · 样机';
  if (u.indexOf('TOTEM') >= 0) return 'Totem Marina Sena · 人形立牌';
  if (u.indexOf('RENO') >= 0) return 'Display Reno16 · Reno16 展示';
  if (u.indexOf('OPPO LOGO') >= 0) return 'Display OPPO logo · 台面标识';
  if (u.indexOf('LOGO') >= 0) return 'Logo de mesa · 台面Logo';
  if (u.indexOf('A4') >= 0) return 'Display A4 · A4 展示';
  return String(h).trim();
}


// ---------- economia de espaço no Supabase ----------
// Fica guardado: a última foto APROVADA de cada loja (referência para comparar), os envios aguardando revisão
// (até 30 dias), os aprovados que ainda vão para o Feishu e qualquer envio revisado nos últimos 15 dias.
// O resto é "arquivado": sai do Supabase, mas continua no Google Drive (botão "abrir no Drive" no site).
// Também reduz para 800 px as fotos antigas que foram copiadas maiores.
const DIAS_AGUARDANDO = 30, DIAS_CARENCIA = 15;
function limparFotos() { // pode rodar à mão para ver o resumo
  const existentes = supa_('GET', '/rest/v1/envios?select=key,canal,loja_raw,data,fotos&limit=20000');
  limparFotos_(existentes, Date.now(), true);
}
function lojaChave_(canal, raw) {
  const c = String(canal || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim();
  const M = { fastshop: 'fast', 'magazine luiza': 'magalu' };
  return (M[c] || c) + '|' + String(raw || '').toUpperCase().replace(/[^A-Z0-9]/g, '').replace(/^0+/, '');
}
function limparFotos_(envios, inicio, detalhado) {
  const rev = {}; supa_('GET', '/rest/v1/revisoes?select=key,status,revisado_em,feishu_em&limit=20000').forEach((r) => { rev[r.key] = r; });
  const agora = Date.now(), dia = 864e5;
  // referência: último envio aprovado por loja
  const ref = {};
  envios.forEach((e) => { const r = rev[e.key]; if (!r || r.status !== 'Aprovado') return; const k = lojaChave_(e.canal, e.loja_raw);
    if (!/\|.+/.test(k)) return; if (!ref[k] || new Date(e.data) > new Date(ref[k].data)) ref[k] = e; });
  const manter = {}; Object.keys(ref).forEach((k) => { manter[ref[k].key] = true; });
  envios.forEach((e) => {
    const r = rev[e.key];
    if (!r && agora - new Date(e.data).getTime() < DIAS_AGUARDANDO * dia) manter[e.key] = true;          // aguardando revisão
    if (r && r.status === 'Aprovado' && !r.feishu_em) manter[e.key] = true;                              // ainda vai para o Feishu
    if (r && agora - new Date(r.revisado_em).getTime() < DIAS_CARENCIA * dia) manter[e.key] = true;       // revisado há pouco
  });
  const apagar = [], mudou = {};
  envios.forEach((e) => {
    if (manter[e.key]) return;
    (e.fotos || []).forEach((p) => { if (p.path) { apagar.push(p.path); p.path = null; p.arquivada = true; mudou[e.key] = e; } });
  });
  for (let i = 0; i < apagar.length; i += 100) {
    UrlFetchApp.fetch(SUPABASE_URL + '/storage/v1/object/fotos', { method: 'delete', contentType: 'application/json', headers: cabecalhos_({}), payload: JSON.stringify({ prefixes: apagar.slice(i, i + 100) }), muteHttpExceptions: true });
  }
  // reduz para 800 px as que ficaram (aos poucos)
  let reduzidas = 0;
  const fim = inicio + 4.6 * 60000;
  Object.keys(mudou).forEach((k) => supa_('PATCH', '/rest/v1/envios?key=eq.' + encodeURIComponent(k), { fotos: mudou[k].fotos }, { Prefer: 'return=minimal' }));
  for (const k in mudou) delete mudou[k];
  for (let i = 0; i < envios.length && Date.now() < fim; i++) {
    const e = envios[i]; if (!manter[e.key]) continue;
    let mexeu = false;
    (e.fotos || []).forEach((p) => {
      if (!p.path || p.px === 800 || Date.now() >= fim) return;
      try {
        const meta = JSON.parse(UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + p.id + '?fields=thumbnailLink', { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getContentText());
        if (!meta.thumbnailLink) return;
        const r = UrlFetchApp.fetch(meta.thumbnailLink.replace(/=s\d+$/, '=s800'), { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
        if (r.getResponseCode() !== 200 || r.getBlob().getBytes().length < 5000) return;
        supaRaw_('POST', '/storage/v1/object/fotos/' + p.path, r.getBlob().getBytes(), { 'Content-Type': 'image/jpeg', 'x-upsert': 'true' });
        p.px = 800; mexeu = true; reduzidas++;
      } catch (err) { /* tenta na próxima rodada */ }
    });
    if (mexeu) supa_('PATCH', '/rest/v1/envios?key=eq.' + encodeURIComponent(e.key), { fotos: e.fotos }, { Prefer: 'return=minimal' }); // grava na hora
  }
  Object.keys(mudou).forEach((k) => supa_('PATCH', '/rest/v1/envios?key=eq.' + encodeURIComponent(k), { fotos: mudou[k].fotos }, { Prefer: 'return=minimal' }));
  const msg = 'Espaço: ' + Object.keys(ref).length + ' lojas com foto de referência · ' + apagar.length + ' fotos arquivadas (seguem no Drive) · ' + reduzidas + ' reduzidas para 800 px';
  if (detalhado || apagar.length || reduzidas) console.log(msg);
}
