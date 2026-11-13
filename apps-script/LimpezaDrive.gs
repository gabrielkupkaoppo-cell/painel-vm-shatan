/**
 * Painel VM Shatan — guarda as fotos do Forms numa PLANILHA DE ARQUIVO no Feishu e libera o Google Drive.
 * Sexto arquivo do mesmo Apps Script (usa supa_ do Sincronizar.gs e feishuToken_/feishu_ do Feishu.gs).
 * Usa só a permissão de planilha que o app já tem (não precisa de aprovação do admin).
 *
 * O Forms sempre salva a foto no seu Drive (não dá para mudar). Depois que a foto foi decidida, este script:
 *   1) escreve uma linha na planilha de arquivo do Feishu (uma aba por mês: "2026-11"):
 *      Data do envio | Canal | Loja | UF | Promotor | Foto | Status | Imagem (a foto dentro da célula, até 2000 px)
 *   2) manda o original do Google Drive para a LIXEIRA (só depois de arquivado no Feishu).
 * O site mostra a miniatura de sempre e o link "abrir original no Feishu" (vai direto para a linha da foto).
 *
 * Nunca vai para a lixeira:
 *   - foto aguardando aprovação
 *   - foto aprovada que ainda não foi para a planilha das lojas no Feishu
 *   - foto que não foi arquivada (se a planilha recusar, ela fica no Drive)
 *
 * Antes da 1ª vez:
 *   - Crie a planilha no Feishu, adicione o app "Painel VM Shatan" com permissão de EDIÇÃO.
 *   - O link da planilha já está em FEISHU_ARQUIVO_PADRAO (a propriedade FEISHU_ARQUIVO, se existir, tem prioridade).
 *
 * 1) limparDriveTeste()   → NÃO muda nada: mostra quantas fotos seriam arquivadas e quantos GB sairiam do Drive.
 * 2) limparDrive()        → faz o trabalho (em partes de ~5 min; rode de novo se faltar).
 * 3) limparDriveLigar()   → roda sozinho a cada hora. limparDriveDesligar() para parar.
 *
 * A lixeira do Google guarda os arquivos por 30 dias. O espaço só é liberado quando ela esvazia
 * (sozinha em 30 dias, ou na hora em drive.google.com → Lixeira → Esvaziar).
 */
var DIAS_DRIVE = 1;
var FEISHU_ARQUIVO_PADRAO = 'https://vcn8bhpny9gy.feishu.cn/sheets/Y39SsskmrhLvRhtVewhctb1lnpg'; // planilha Arquivo Fotos VM
var ARQ_CAB = ['Data do envio · 提交日期', 'Canal · 渠道', 'Loja · 门店', 'UF · 州', 'Promotor · 提交人', 'Foto · 照片类型', 'Status · 审核结果', 'Imagem · 图片'];

function limparDriveTeste() { limparDrive_(true); }
function limparDrive() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) throw new Error('Outra rotina está rodando agora. Tente de novo em 1 minuto.');
  try { limparDrive_(false); } finally { lock.releaseLock(); }
}
function limparDriveLigar() {
  ScriptApp.getProjectTriggers().forEach((t) => { if (t.getHandlerFunction() === 'limparDriveDiario') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('limparDriveDiario').timeBased().everyHours(1).create();
  console.log('Arquivo de fotos no Feishu + limpeza do Drive ligados (a cada hora).');
}
function limparDriveDesligar() {
  ScriptApp.getProjectTriggers().forEach((t) => { if (t.getHandlerFunction() === 'limparDriveDiario') ScriptApp.deleteTrigger(t); });
  console.log('Limpeza do Drive desligada.');
}
function limparDriveDiario() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try { limparDrive_(false); } finally { lock.releaseLock(); }
}

// ---------- planilha de arquivo ----------
function arquivoCfg_() {
  const url = String(PropertiesService.getScriptProperties().getProperty('FEISHU_ARQUIVO') || FEISHU_ARQUIVO_PADRAO || '').trim();
  const m = url.match(/sheets\/([A-Za-z0-9]+)/) || url.match(/^([A-Za-z0-9]{15,})$/);
  if (!m) throw new Error('Falta a propriedade FEISHU_ARQUIVO com o link da planilha de arquivo do Feishu.');
  const host = (url.match(/^https:\/\/[^/]+/) || ['https://www.feishu.cn'])[0];
  return { token: m[1], host: host };
}
function arquivoEstado_() {
  const p = PropertiesService.getScriptProperties();
  let st = {}; try { st = JSON.parse(p.getProperty('FEISHU_ARQUIVO_ABAS') || '{}'); } catch (e) { st = {}; }
  st.abas = st.abas || {};
  return { st: st, salvar: () => p.setProperty('FEISHU_ARQUIVO_ABAS', JSON.stringify(st)) };
}
// aba do mês: acha ou cria (com cabeçalho) e descobre a próxima linha livre
function arquivoAba_(tk, cfg, E, titulo) {
  const cache = E.st.abas[titulo];
  if (cache && cache.token === cfg.token) return cache;
  const abas = feishu_(tk, 'GET', '/sheets/v3/spreadsheets/' + cfg.token + '/sheets/query').sheets || [];
  let a = abas.find((x) => x.title === titulo), id;
  if (a) id = a.sheet_id;
  else {
    const r = feishu_(tk, 'POST', '/sheets/v2/spreadsheets/' + cfg.token + '/sheets_batch_update', { requests: [{ addSheet: { properties: { title: titulo, index: 0 } } }] });
    id = r.replies[0].addSheet.properties.sheetId;
  }
  const col = feishu_(tk, 'GET', '/sheets/v2/spreadsheets/' + cfg.token + '/values/' + encodeURIComponent(id + '!A1:A20000') + '?valueRenderOption=ToString');
  const vals = (col.valueRange && col.valueRange.values) || [];
  let n = vals.length; while (n > 0 && (vals[n - 1][0] == null || vals[n - 1][0] === '')) n--;
  if (n === 0) {
    feishu_(tk, 'PUT', '/sheets/v2/spreadsheets/' + cfg.token + '/values', { valueRange: { range: id + '!A1:H1', values: [ARQ_CAB] } });
    n = 1;
  }
  E.st.abas[titulo] = { token: cfg.token, id: id, prox: n + 1 }; E.salvar();
  return E.st.abas[titulo];
}
// foto em até 2000 px (qualidade boa e leve); se não der, o original
function fotoArquivo_(id) {
  const tk = ScriptApp.getOAuthToken();
  try {
    const meta = JSON.parse(UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id + '?fields=thumbnailLink', { headers: { Authorization: 'Bearer ' + tk } }).getContentText());
    if (meta.thumbnailLink) {
      const r = UrlFetchApp.fetch(meta.thumbnailLink.replace(/=s\d+$/, '=s2000'), { headers: { Authorization: 'Bearer ' + tk }, muteHttpExceptions: true });
      if (r.getResponseCode() === 200 && r.getBlob().getBytes().length > 5000) return r.getBlob();
    }
  } catch (e) { /* usa o original */ }
  return DriveApp.getFileById(id).getBlob();
}
function arquivarFoto_(tk, cfg, E, e, p, status) {
  const d = new Date(e.data);
  const aba = arquivoAba_(tk, cfg, E, Utilities.formatDate(d, 'America/Sao_Paulo', 'yyyy-MM'));
  const n = aba.prox;
  feishu_(tk, 'PUT', '/sheets/v2/spreadsheets/' + cfg.token + '/values', { valueRange: { range: aba.id + '!A' + n + ':G' + n, values: [[
    Utilities.formatDate(d, 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm'), String(e.canal || ''), String(e.loja_raw || ''), String(e.uf || ''), String(e.nome || ''),
    String(p.label || 'Foto').split(' · ')[0], status]] } });
  aba.prox = n + 1; E.salvar();   // a linha já está ocupada mesmo se a imagem falhar
  const blob = fotoArquivo_(p.id);
  const corpo = { range: aba.id + '!H' + n + ':H' + n, image: Array.from(blob.getBytes(), (b) => b & 0xff), name: (p.id + '.jpg') };
  try { feishu_(tk, 'POST', '/sheets/v2/spreadsheets/' + cfg.token + '/values_image', corpo); }
  catch (e1) { Utilities.sleep(3000); feishu_(tk, 'POST', '/sheets/v2/spreadsheets/' + cfg.token + '/values_image', corpo); }
  return { bytes: blob.getBytes().length, url: cfg.host + '/sheets/' + cfg.token + '?sheet=' + aba.id + '&range=H' + n };
}

function limparDrive_(teste) {
  const inicio = Date.now(), limite = 4.6 * 60 * 1000;
  const dia = 864e5, agora = Date.now();
  const rev = {};
  (supa_('GET', '/rest/v1/revisoes?select=key,status,revisado_em,feishu_em&limit=20000') || []).forEach((r) => { rev[r.key] = r; });
  const envios = (supa_('GET', '/rest/v1/envios?select=key,data,canal,loja_raw,uf,nome,fotos&limit=20000') || [])
    .sort((a, b) => new Date(a.data) - new Date(b.data));
  const gtk = ScriptApp.getOAuthToken();
  const cfg = teste ? null : arquivoCfg_();
  const tk = teste ? null : feishuToken_();
  const E = arquivoEstado_();
  let nArq = 0, bArq = 0, nLixo = 0, bLixo = 0, erros = 0, parou = false, ultErro = '';

  for (let i = 0; i < envios.length; i++) {
    if (Date.now() - inicio > limite) { parou = true; break; }
    const e = envios[i], r = rev[e.key];
    if (!r || !r.revisado_em) continue;                                          // aguardando: fica no Drive
    if (agora - new Date(r.revisado_em).getTime() < DIAS_DRIVE * dia) continue;    // carência
    const ok = (r.status === 'Aprovado' && r.feishu_em) || r.status === 'Reprovado' || r.status === 'Substituído' || r.status === 'Descartado';
    if (!ok) continue;
    const fotos = e.fotos || [];
    let mudou = false;
    for (const p of fotos) {
      if (!p || !p.id || p.drive_lixeira) continue;
      if (Date.now() - inicio > limite) { parou = true; break; }
      try {
        const res = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + p.id + '?fields=size,trashed,ownedByMe', { headers: { Authorization: 'Bearer ' + gtk }, muteHttpExceptions: true });
        if (res.getResponseCode() === 404) { p.drive_lixeira = true; mudou = true; continue; }
        const m = JSON.parse(res.getContentText());
        if (!m.ownedByMe) continue;
        if (m.trashed) { p.drive_lixeira = true; mudou = true; continue; }
        if (teste) { nLixo++; bLixo += +m.size || 0; continue; }
        // 1) arquiva na planilha do Feishu
        if (!p.feishu_url) {
          const a = arquivarFoto_(tk, cfg, E, e, p, r.status);
          p.feishu_url = a.url; mudou = true; nArq++; bArq += a.bytes;
        }
        // 2) original → lixeira do Drive
        const t = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + p.id, { method: 'patch', contentType: 'application/json', payload: JSON.stringify({ trashed: true }), headers: { Authorization: 'Bearer ' + gtk }, muteHttpExceptions: true });
        if (t.getResponseCode() >= 300) { erros++; continue; }
        p.drive_lixeira = true; mudou = true; nLixo++; bLixo += +m.size || 0;
      } catch (err) {
        erros++; ultErro = String(err && err.message || err);
        if (mudou) supa_('PATCH', '/rest/v1/envios?key=eq.' + encodeURIComponent(e.key), { fotos: fotos }, { Prefer: 'return=minimal' });
        if (/FEISHU_ARQUIVO|91403|1310213|403|permiss/i.test(ultErro)) throw new Error('Parei: ' + ultErro + ' — confira se o app "Painel VM Shatan" está na planilha de arquivo com permissão de edição.');
        mudou = false;
      }
    }
    if (mudou) supa_('PATCH', '/rest/v1/envios?key=eq.' + encodeURIComponent(e.key), { fotos: fotos }, { Prefer: 'return=minimal' });
  }

  const gb = (b) => (b / 1073741824).toFixed(2) + ' GB';
  if (teste) console.log('TESTE (nada mudou): ' + nLixo + ' fotos já decididas seriam arquivadas no Feishu e sairiam do Drive (~' + gb(bLixo) + '). As que estão aguardando saem depois de decididas.' + (parou ? ' — a contagem parou no limite de tempo; o total real é maior.' : ''));
  else console.log(nArq + ' fotos arquivadas na planilha do Feishu · ' + nLixo + ' originais foram para a lixeira do Drive (~' + gb(bLixo) + ').' + (erros ? ' ' + erros + ' erro(s). Último: ' + ultErro : '') + (parou ? ' Parou no limite de tempo: rode de novo (ou deixe ligado a cada hora).' : ' Concluído.'));
}
