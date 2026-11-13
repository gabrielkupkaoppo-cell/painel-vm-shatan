/**
 * Painel VM Shatan — guarda as fotos do Forms no FEISHU e libera o Google Drive.
 * Sexto arquivo do mesmo Apps Script (usa supa_ do Sincronizar.gs e feishuToken_/feishu_ do Feishu.gs).
 *
 * O Forms sempre salva a foto no seu Drive (não dá para mudar). Este script:
 *   1) copia cada foto ORIGINAL (qualidade cheia) para uma pasta no Feishu Drive
 *      "Fotos VM Mesas (Forms)", com subpastas por mês e canal (ex.: "2026-11 Bemol");
 *   2) depois que a foto foi decidida e já está guardada no Feishu, manda o original do Google Drive
 *      para a LIXEIRA (DIAS_DRIVE dias após a decisão).
 * O site mostra a miniatura de sempre e o link "abrir original no Feishu".
 *
 * Nunca vai para a lixeira:
 *   - foto que ainda NÃO foi copiada para o Feishu
 *   - foto aguardando aprovação
 *   - foto aprovada que ainda não foi para a planilha do Feishu
 *
 * Antes da 1ª vez:
 *   - No app do Feishu (open.feishu.cn → seu app → Permissões), ligue a permissão de Drive
 *     "查看、评论、编辑和管理云空间中所有文件" (drive:drive) e publique a nova versão do app.
 *   - Propriedade do script FEISHU_FOTOS_DONO = seu e-mail do Feishu (a pasta é compartilhada com você).
 *
 * 1) limparDriveTeste()   → NÃO muda nada: mostra quantas fotos iriam para o Feishu e quantos GB sairiam do Drive.
 * 2) limparDrive()        → faz o trabalho (em partes de ~5 min; rode de novo se faltar).
 * 3) limparDriveLigar()   → roda sozinho a cada hora. limparDriveDesligar() para parar.
 *
 * A lixeira do Google guarda os arquivos por 30 dias. O espaço só é liberado quando ela esvazia
 * (sozinha em 30 dias, ou na hora em drive.google.com → Lixeira → Esvaziar).
 */
var DIAS_DRIVE = 1;

function limparDriveTeste() { limparDrive_(true); }
function limparDrive() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) throw new Error('Outra rotina está rodando agora. Tente de novo em 1 minuto.');
  try { limparDrive_(false); } finally { lock.releaseLock(); }
}
function limparDriveLigar() {
  ScriptApp.getProjectTriggers().forEach((t) => { if (t.getHandlerFunction() === 'limparDriveDiario') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('limparDriveDiario').timeBased().everyHours(1).create();
  console.log('Fotos → Feishu e limpeza do Drive ligadas (a cada hora).');
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

// ---------- pastas no Feishu ----------
function fotosPastas_() {
  const p = PropertiesService.getScriptProperties();
  let st = {}; try { st = JSON.parse(p.getProperty('FEISHU_FOTOS_PASTAS') || '{}'); } catch (e) { st = {}; }
  st.subs = st.subs || {};
  return { st: st, salvar: () => p.setProperty('FEISHU_FOTOS_PASTAS', JSON.stringify(st)) };
}
function fotosRaiz_(tk, P) {
  if (P.st.raiz) return P.st.raiz;
  const dono = String(PropertiesService.getScriptProperties().getProperty('FEISHU_FOTOS_DONO') || '').trim();
  if (!dono) throw new Error('Falta a propriedade FEISHU_FOTOS_DONO (seu e-mail do Feishu), para a pasta das fotos ser compartilhada com você.');
  const root = feishu_(tk, 'GET', '/drive/explorer/v2/root_folder/meta');
  const f = feishu_(tk, 'POST', '/drive/v1/files/create_folder', { name: 'Fotos VM Mesas (Forms)', folder_token: root.token });
  feishu_(tk, 'POST', '/drive/v1/permissions/' + f.token + '/members?type=folder&need_notification=true', { member_type: 'email', member_id: dono, perm: 'full_access' });
  P.st.raiz = f.token; P.st.url = f.url || ''; P.salvar();
  console.log('Pasta criada no Feishu e compartilhada com ' + dono + ': ' + (f.url || f.token));
  return f.token;
}
function fotosSub_(tk, P, nome) {
  if (P.st.subs[nome]) return P.st.subs[nome];
  const f = feishu_(tk, 'POST', '/drive/v1/files/create_folder', { name: nome, folder_token: fotosRaiz_(tk, P) });
  P.st.subs[nome] = f.token; P.salvar();
  return f.token;
}
function fotoUrlFeishu_(P, token) {
  const m = String(P.st.url || '').match(/^https:\/\/[^/]+/);
  return (m ? m[0] : 'https://www.feishu.cn') + '/file/' + token;
}
function enviarFotoFeishu_(tk, P, e, p) {
  const blob = DriveApp.getFileById(p.id).getBlob();
  const bytes = blob.getBytes();
  if (bytes.length > 20 * 1048576) throw new Error('arquivo maior que 20 MB');
  const d = new Date(e.data);
  const canal = String(e.canal || 'Sem canal').replace(/[\\/:*?"<>|]/g, ' ').trim();
  const sub = fotosSub_(tk, P, Utilities.formatDate(d, 'America/Sao_Paulo', 'yyyy-MM') + ' ' + canal);
  const ext = (String(blob.getName() || '').match(/\.[A-Za-z0-9]{2,5}$/) || ['.jpg'])[0];
  const nome = (Utilities.formatDate(d, 'America/Sao_Paulo', 'yyyy-MM-dd HH.mm') + ' ' + canal + ' ' + String(e.loja_raw || '').trim() + ' ' + String(p.label || 'Foto').split(' · ')[0])
    .replace(/[\\/:*?"<>|]/g, ' ').replace(/\s+/g, ' ').slice(0, 180) + ext;
  const res = UrlFetchApp.fetch(FEISHU_API + '/drive/v1/files/upload_all', {
    method: 'post', headers: { Authorization: 'Bearer ' + tk }, muteHttpExceptions: true,
    payload: { file_name: nome, parent_type: 'explorer', parent_node: sub, size: String(bytes.length), file: blob.setName(nome) },
  });
  let j = {}; try { j = JSON.parse(res.getContentText()); } catch (x) { /* */ }
  if (j.code !== 0 || !j.data || !j.data.file_token) throw new Error('Feishu upload → ' + (j.code || res.getResponseCode()) + ' ' + (j.msg || res.getContentText().slice(0, 150)));
  return { bytes: bytes.length, token: j.data.file_token };
}

function limparDrive_(teste) {
  const inicio = Date.now(), limite = 4.6 * 60 * 1000;
  const dia = 864e5, agora = Date.now();
  const rev = {};
  (supa_('GET', '/rest/v1/revisoes?select=key,status,revisado_em,feishu_em&limit=20000') || []).forEach((r) => { rev[r.key] = r; });
  const envios = (supa_('GET', '/rest/v1/envios?select=key,data,canal,loja_raw,fotos&limit=20000') || [])
    .sort((a, b) => new Date(a.data) - new Date(b.data));
  const gtk = ScriptApp.getOAuthToken();
  const tk = teste ? null : feishuToken_();
  const P = fotosPastas_();
  let nFeishu = 0, bFeishu = 0, nLixo = 0, bLixo = 0, faltaFeishu = 0, erros = 0, parou = false, ultErro = '';

  for (let i = 0; i < envios.length; i++) {
    if (Date.now() - inicio > limite) { parou = true; break; }
    const e = envios[i], r = rev[e.key];
    const fotos = e.fotos || [];
    let mudou = false;
    const decidida = r && r.revisado_em && agora - new Date(r.revisado_em).getTime() >= DIAS_DRIVE * dia &&
      ((r.status === 'Aprovado' && r.feishu_em) || r.status === 'Reprovado' || r.status === 'Substituído' || r.status === 'Descartado');

    fotos.forEach((p) => {
      if (!p || !p.id || p.drive_lixeira) return;
      if (Date.now() - inicio > limite) { parou = true; return; }
      try {
        // 1) guardar o original no Feishu
        if (!p.feishu_token) {
          if (teste) {
            const m = JSON.parse(UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + p.id + '?fields=size,trashed,ownedByMe', { headers: { Authorization: 'Bearer ' + gtk }, muteHttpExceptions: true }).getContentText());
            if (m.size && !m.trashed && m.ownedByMe) { faltaFeishu++; bFeishu += +m.size || 0; if (decidida) { nLixo++; bLixo += +m.size || 0; } }
            return;
          }
          const up = enviarFotoFeishu_(tk, P, e, p);
          p.feishu_token = up.token; p.feishu_url = fotoUrlFeishu_(P, up.token); mudou = true;
          nFeishu++; bFeishu += up.bytes;
        }
        // 2) original já está no Feishu e a foto foi decidida → lixeira do Drive
        if (!decidida || teste) return;
        const res = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + p.id + '?fields=size,trashed,ownedByMe', { headers: { Authorization: 'Bearer ' + gtk }, muteHttpExceptions: true });
        if (res.getResponseCode() === 404) { p.drive_lixeira = true; mudou = true; return; }
        const m = JSON.parse(res.getContentText());
        if (!m.ownedByMe) return;
        if (m.trashed) { p.drive_lixeira = true; mudou = true; return; }
        const t = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + p.id, { method: 'patch', contentType: 'application/json', payload: JSON.stringify({ trashed: true }), headers: { Authorization: 'Bearer ' + gtk }, muteHttpExceptions: true });
        if (t.getResponseCode() >= 300) { erros++; return; }
        p.drive_lixeira = true; mudou = true; nLixo++; bLixo += +m.size || 0;
      } catch (err) {
        erros++; ultErro = String(err && err.message || err);
        if (/Falta a propriedade|99991672|99991679|permission|scope/i.test(ultErro)) throw err; // problema de configuração: para já
      }
    });
    if (mudou) supa_('PATCH', '/rest/v1/envios?key=eq.' + encodeURIComponent(e.key), { fotos: fotos }, { Prefer: 'return=minimal' });
  }

  const gb = (b) => (b / 1073741824).toFixed(2) + ' GB';
  if (teste) console.log('TESTE (nada mudou): ' + faltaFeishu + ' fotos iriam para o Feishu (' + gb(bFeishu) + '). Dessas, ' + nLixo + ' já poderiam sair do Drive agora (~' + gb(bLixo) + '); as outras saem depois de decididas.' + (parou ? ' — a contagem parou no limite de tempo; o total real é maior.' : ''));
  else console.log(nFeishu + ' fotos copiadas para o Feishu (' + gb(bFeishu) + ') · ' + nLixo + ' originais foram para a lixeira do Drive (~' + gb(bLixo) + ').' + (erros ? ' ' + erros + ' erro(s). Último: ' + ultErro : '') + (parou ? ' Parou no limite de tempo: rode de novo (ou deixe ligado a cada hora).' : ' Concluído.') + (P.st.url ? ' Pasta no Feishu: ' + P.st.url : ''));
}
