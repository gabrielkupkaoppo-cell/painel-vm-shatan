/**
 * Painel VM Shatan — relatório do espaço do Google Drive (NÃO apaga nada).
 * Sétimo arquivo do mesmo Apps Script.
 *
 * Rode relatorioDrive() e veja o resultado em "Registro de execução":
 *   - quanto o Drive ocupa no total e quanto está na lixeira
 *   - espaço por tipo de arquivo (fotos, vídeos, PDFs...)
 *   - as 15 pastas que mais ocupam
 *   - os 25 maiores arquivos (com link)
 */
function relatorioDrive() {
  DriveApp.getRootFolder(); // garante a permissão do Drive
  const tk = ScriptApp.getOAuthToken(), inicio = Date.now();
  const listar = (q, cb) => {
    let page = '';
    do {
      const url = 'https://www.googleapis.com/drive/v3/files?pageSize=1000&orderBy=quotaBytesUsed%20desc' +
        '&q=' + encodeURIComponent(q) +
        '&fields=' + encodeURIComponent('nextPageToken,files(id,name,mimeType,quotaBytesUsed,parents,webViewLink)') +
        (page ? '&pageToken=' + encodeURIComponent(page) : '');
      const r = JSON.parse(UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + tk } }).getContentText());
      (r.files || []).forEach(cb);
      page = r.nextPageToken;
    } while (page && Date.now() - inicio < 5 * 60 * 1000);
    return !!page;
  };
  const tipoDe = (m) => /^image\//.test(m) ? 'Fotos/imagens' : /^video\//.test(m) ? 'Vídeos' : m === 'application/pdf' ? 'PDFs' :
    /zip|rar|7z|tar|gzip/.test(m) ? 'Compactados (zip/rar)' : /^audio\//.test(m) ? 'Áudios' : /google-apps/.test(m) ? 'Docs/Planilhas Google' : 'Outros';
  const gb = (b) => (b / 1073741824).toFixed(2) + ' GB', mb = (b) => (b / 1048576).toFixed(1) + ' MB';

  let total = 0, n = 0; const porTipo = {}, porPasta = {}, top = [];
  const cortou = listar("'me' in owners and trashed=false", (f) => {
    const b = +f.quotaBytesUsed || 0; total += b; n++;
    const t = tipoDe(f.mimeType); porTipo[t] = porTipo[t] || { b: 0, n: 0 }; porTipo[t].b += b; porTipo[t].n++;
    const p = (f.parents || ['(sem pasta)'])[0]; porPasta[p] = porPasta[p] || { b: 0, n: 0 }; porPasta[p].b += b; porPasta[p].n++;
    if (top.length < 25) top.push(f);
  });
  let lixo = 0, nLixo = 0;
  listar("'me' in owners and trashed=true", (f) => { lixo += +f.quotaBytesUsed || 0; nLixo++; });

  const nomePasta = (id) => {
    if (id === '(sem pasta)') return id;
    try {
      const r = JSON.parse(UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id + '?fields=name,parents', { headers: { Authorization: 'Bearer ' + tk } }).getContentText());
      let nome = r.name;
      if (r.parents && r.parents[0]) {
        const pai = JSON.parse(UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + r.parents[0] + '?fields=name', { headers: { Authorization: 'Bearer ' + tk }, muteHttpExceptions: true }).getContentText());
        if (pai.name) nome = pai.name + ' / ' + nome;
      }
      return nome;
    } catch (e) { return id; }
  };

  const L = [];
  L.push('=== ESPAÇO DO GOOGLE DRIVE ===');
  L.push('Seus arquivos: ' + gb(total) + ' em ' + n + ' arquivos' + (cortou ? ' (parou no limite de tempo; o total real é maior)' : ''));
  L.push('Na lixeira (ainda ocupando espaço): ' + gb(lixo) + ' em ' + nLixo + ' arquivos' + (lixo > 0 ? '  → esvazie em drive.google.com → Lixeira' : ''));
  L.push('');
  L.push('--- Por tipo ---');
  Object.keys(porTipo).sort((a, b) => porTipo[b].b - porTipo[a].b).forEach((t) => L.push(t + ': ' + gb(porTipo[t].b) + ' (' + porTipo[t].n + ' arquivos)'));
  L.push('');
  L.push('--- Pastas que mais ocupam ---');
  Object.keys(porPasta).sort((a, b) => porPasta[b].b - porPasta[a].b).slice(0, 15)
    .forEach((p) => L.push(gb(porPasta[p].b) + ' · ' + porPasta[p].n + ' arq. · ' + nomePasta(p)));
  L.push('');
  L.push('--- 25 maiores arquivos ---');
  top.forEach((f) => L.push(mb(+f.quotaBytesUsed || 0) + ' · ' + f.name + ' · ' + f.webViewLink));
  console.log(L.join('\n'));
}
