/**
 * Painel VM Shatan — libera espaço no Google Drive.
 * Sexto arquivo do mesmo Apps Script (usa supa_ do Sincronizar.gs).
 *
 * O Forms guarda cada foto no seu Drive. Depois que a foto foi decidida, o original no Drive
 * não é mais necessário: a foto aprovada já está na planilha do Feishu e o site guarda a cópia
 * reduzida da última foto de cada loja. Este script manda esses originais para a LIXEIRA do Drive.
 *
 * Vai para a lixeira (depois de DIAS_DRIVE dias da decisão):
 *   - foto APROVADA que já foi para o Feishu
 *   - foto REPROVADA, SUBSTITUÍDA ou DESCARTADA
 * Nunca mexe em:
 *   - foto aguardando aprovação
 *   - foto aprovada que ainda não foi para o Feishu
 *   - arquivos que não são das respostas do Forms
 *
 * 1) limparDriveTeste()   → NÃO apaga nada: mostra quantas fotos e quantos GB sairiam.
 * 2) limparDrive()        → manda para a lixeira (roda em partes de ~5 min; rode de novo se faltar).
 * 3) limparDriveLigar()   → roda sozinho todo dia de madrugada. limparDriveDesligar() para parar.
 *
 * A lixeira do Google guarda os arquivos por 30 dias (dá para recuperar). O espaço só é liberado
 * quando a lixeira esvazia: sozinha em 30 dias, ou na hora em drive.google.com → Lixeira → Esvaziar.
 */
var DIAS_DRIVE = 7;

function limparDriveTeste() { limparDrive_(true); }
function limparDrive() { limparDrive_(false); }
function limparDriveLigar() {
  ScriptApp.getProjectTriggers().forEach((t) => { if (t.getHandlerFunction() === 'limparDriveDiario') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('limparDriveDiario').timeBased().everyDays(1).atHour(3).inTimezone('America/Sao_Paulo').create();
  console.log('Limpeza do Drive ligada (todo dia ~3h).');
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

function limparDrive_(teste) {
  const inicio = Date.now(), limite = 4.6 * 60 * 1000;
  const dia = 864e5, agora = Date.now();
  const rev = {};
  (supa_('GET', '/rest/v1/revisoes?select=key,status,revisado_em,feishu_em&limit=20000') || []).forEach((r) => { rev[r.key] = r; });
  const envios = supa_('GET', '/rest/v1/envios?select=key,fotos&limit=20000') || [];
  const tk = ScriptApp.getOAuthToken();
  let nFotos = 0, bytes = 0, nEnvios = 0, pulos = 0, erros = 0, parou = false;
  const porStatus = {};

  for (let i = 0; i < envios.length; i++) {
    if (Date.now() - inicio > limite) { parou = true; break; }
    const e = envios[i], r = rev[e.key];
    if (!r || !r.revisado_em) continue;                                       // aguardando: nunca
    if (agora - new Date(r.revisado_em).getTime() < DIAS_DRIVE * dia) continue; // ainda no prazo de carência
    const ok = (r.status === 'Aprovado' && r.feishu_em) || r.status === 'Reprovado' || r.status === 'Substituído' || r.status === 'Descartado';
    if (!ok) continue;
    const fotos = e.fotos || [];
    const alvo = fotos.filter((p) => p && p.id && !p.drive_lixeira);
    if (!alvo.length) continue;
    let mudou = false;
    alvo.forEach((p) => {
      if (Date.now() - inicio > limite) { parou = true; return; }
      try {
        const res = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + p.id + '?fields=size,trashed,ownedByMe', { headers: { Authorization: 'Bearer ' + tk }, muteHttpExceptions: true });
        if (res.getResponseCode() === 404) { p.drive_lixeira = true; mudou = true; return; }   // já não existe
        const m = JSON.parse(res.getContentText());
        if (!m.ownedByMe) { pulos++; return; }                                                  // não é seu: não mexe
        if (m.trashed) { p.drive_lixeira = true; mudou = true; return; }
        nFotos++; bytes += +m.size || 0; porStatus[r.status] = (porStatus[r.status] || 0) + 1;
        if (teste) return;
        const up = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + p.id, { method: 'patch', contentType: 'application/json', payload: JSON.stringify({ trashed: true }), headers: { Authorization: 'Bearer ' + tk }, muteHttpExceptions: true });
        if (up.getResponseCode() >= 300) { erros++; return; }
        p.drive_lixeira = true; mudou = true;
      } catch (err) { erros++; }
    });
    if (mudou) { supa_('PATCH', '/rest/v1/envios?key=eq.' + encodeURIComponent(e.key), { fotos: fotos }, { Prefer: 'return=minimal' }); nEnvios++; }
  }

  const gb = (bytes / 1073741824).toFixed(2);
  const det = Object.keys(porStatus).map((s) => s + ': ' + porStatus[s]).join(', ');
  if (teste) console.log('TESTE (nada foi apagado): ' + nFotos + ' fotos iriam para a lixeira, liberando ~' + gb + ' GB. (' + det + ')' + (pulos ? ' ' + pulos + ' arquivo(s) de outra pessoa seriam ignorados.' : '') + (parou ? ' — a contagem parou no limite de tempo; o total real é maior.' : ''));
  else console.log(nFotos + ' fotos foram para a lixeira (~' + gb + ' GB). (' + det + ')' + (erros ? ' ' + erros + ' erro(s).' : '') + (parou ? ' Parou no limite de tempo: rode de novo para continuar.' : ' Concluído.') + ' O espaço é liberado quando a lixeira do Drive esvaziar.');
}
