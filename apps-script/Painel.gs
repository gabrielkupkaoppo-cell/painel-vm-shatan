/**
 * Painel VM Shatan — atualiza o DASH a partir da planilha do Feishu (a mesma de FEISHU_PLANILHA).
 * Terceiro arquivo do mesmo Apps Script (usa funções do Sincronizar.gs e do Feishu.gs).
 *
 * 1) Rode painelTeste()  → lê o Feishu, mostra no registro o resumo e grava no dash.
 * 2) Deu certo? Rode painelLigar() → passa a atualizar sozinho a cada 10 min.
 */
var PAINEL_CANAIS = ['Americanas', 'Magalu', 'Lider', 'Bemol', 'Martinello', 'Paraiba', 'Fast', 'Claro', 'Vivo', 'Tim'];
var PAINEL_REGIAO = { DF: 'DF/GO', GO: 'DF/GO', CE: 'CE/PI', PI: 'CE/PI', PA: 'PA/AP', AP: 'PA/AP', MT: 'MT/MS', MS: 'MT/MS', AM: 'AM/AC/RR/RO', AC: 'AM/AC/RR/RO', RR: 'AM/AC/RR/RO', RO: 'AM/AC/RR/RO', MA: 'MA/TO', TO: 'MA/TO' };

function painelTeste() { atualizarPainel_(true); }
function painelLigar() {
  ScriptApp.getProjectTriggers().forEach((t) => { if (t.getHandlerFunction() === 'atualizarPainel') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('atualizarPainel').timeBased().everyMinutes(10).create();
  console.log('Dash ligado ao Feishu: atualiza a cada 10 min.');
}
function painelDesligar() {
  ScriptApp.getProjectTriggers().forEach((t) => { if (t.getHandlerFunction() === 'atualizarPainel') ScriptApp.deleteTrigger(t); });
  console.log('Atualização automática do dash desligada.');
}
function atualizarPainel() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) return;
  try { atualizarPainel_(false); } finally { lock.releaseLock(); }
}

function atualizarPainel_(detalhado) {
  const planilha = String(PropertiesService.getScriptProperties().getProperty('FEISHU_PLANILHA') || '').trim();
  if (!planilha) throw new Error('Falta a propriedade FEISHU_PLANILHA.');
  const tk = feishuToken_();
  const abas = feishu_(tk, 'GET', '/sheets/v3/spreadsheets/' + planilha + '/sheets/query').sheets || [];
  const hoje = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd');

  // data da última foto aprovada por linha (vem do envio ao Feishu: "Magalu linha 75: ...")
  // data do último envio APROVADO de cada loja, pelo canal + código da loja (não depende da linha do Feishu)
  const ultima = {};
  const aprov = {}; supa_('GET', '/rest/v1/revisoes?select=key&status=eq.Aprovado&limit=20000').forEach((r) => { aprov[r.key] = 1; });
  supa_('GET', '/rest/v1/envios?select=key,canal,loja_raw,data&limit=20000').forEach((e) => {
    if (!aprov[e.key] || !e.data) return;
    const c = semAcento_(canalFeishu_(e.canal)); const ks = chaves_(e.loja_raw);
    const iso = Utilities.formatDate(new Date(e.data), 'America/Sao_Paulo', 'yyyy-MM-dd');
    [ks[0] && c + '|' + ks[0], ks[1] && c + '|#' + ks[1]].filter(Boolean).forEach((k) => { if (!ultima[k] || ultima[k] < iso) ultima[k] = iso; });
  });
  const ultimaDe = (canal, codigo, nome) => {
    const c = semAcento_(canal);
    for (const raw of [codigo, String(nome || '').split(/[\s-]/)[0]]) {
      const ks = chaves_(raw); if (!ks.length) continue;
      if (ks[0] && ultima[c + '|' + ks[0]]) return ultima[c + '|' + ks[0]];
      if (ks[1] && ultima[c + '|#' + ks[1]]) return ultima[c + '|#' + ks[1]];
    }
    return '';
  };

  const lojas = []; const resumo = [];
  PAINEL_CANAIS.forEach((ch) => {
    const aba = abas.find((a) => semAcento_(a.title) === semAcento_(ch));
    if (!aba) { resumo.push(ch + ': aba não encontrada'); return; }
    const nCol = Math.min((aba.grid_properties && aba.grid_properties.column_count) || 40, 60);
    const nLin = Math.min((aba.grid_properties && aba.grid_properties.row_count) || 500, 3000);
    const V = valoresBrutos_(tk, planilha, aba.sheet_id + '!A1:' + letra_(nCol) + nLin);
    const cel = (r, c) => { const row = V[r - 1]; return row ? row[c - 1] : null; };
    const txt = (r, c) => texto_(cel(r, c));
    // cabeçalho (mesma regra do importador)
    let hdr = null;
    for (let r = 1; r < 6 && !hdr; r++) for (let c = 1; c <= nCol; c++) if (txt(r, c).indexOf('Store Name') >= 0) { hdr = r; break; }
    if (!hdr) hdr = 2;
    let hrows = []; for (let r = 1; r <= hdr; r++) hrows.push(r); if (ch === 'Claro') hrows.push(3);
    if (ch === 'Tim') { hdr = 3; hrows = [1, 2, 3]; }
    const H = {};
    hrows.forEach((r) => { for (let c = 1; c <= nCol; c++) { const v = txt(r, c); if (v) (H[c] = H[c] || []).push(v); } });
    const cols = Object.keys(H).map(Number).sort((a, b) => a - b);
    const find = (fn) => { for (const c of cols) if (H[c].some(fn)) return c; return null; };
    const C = { code: find((v) => v.indexOf('Store Code') >= 0), name: find((v) => v.indexOf('Store Name') >= 0), city: find((v) => v.indexOf('City') >= 0),
      state: find((v) => (v.indexOf('State') >= 0 || v === 'STATE') && v.indexOf('Currently') < 0), cls: find((v) => v.indexOf('Classification') >= 0),
      promo: find((v) => v.indexOf('promotor exclusivo') >= 0), t16: find((v) => v.indexOf('1.6M') >= 0), t12: find((v) => v.indexOf('1.2M') >= 0), t08: find((v) => v.indexOf('0.8M') >= 0),
      note: find((v) => v.indexOf('备注anota') >= 0), ok: find((v) => v.indexOf('Ok/Pendente') >= 0), date: find((v) => v.indexOf('veicula') >= 0) };
    const fotoCols = cols.filter((c) => H[c].some((v) => v.indexOf('照片') >= 0));
    const dispCols = cols.filter((c) => H[c].some((v) => v.toLowerCase().indexOf('display') >= 0 && v.indexOf('Store') < 0));
    if (!C.name) { resumo.push(ch + ': sem coluna Store Name'); return; }
    let n = 0;
    for (let r = hdr + 1; r <= V.length; r++) {
      const nome = txt(r, C.name); if (!nome || nome.indexOf('Store') >= 0) continue;
      const g = (k) => (C[k] ? txt(r, C[k]) : '');
      const num = (k) => { const v = g(k); if (!v) return 0; const f = parseFloat(v); if (!isNaN(f) && /^[-+]?\d*\.?\d+$/.test(v)) return Math.trunc(f); return ['YES', 'SIM', 'X', 'TRUE'].indexOf(v.toUpperCase()) >= 0 ? 1 : 0; };
      const okp = g('ok').toUpperCase() === 'OK' ? 'OK' : (g('ok') || 'Pendente');
      let fotos = fotoCols.filter((c) => { const v = cel(r, c); return v != null && v !== '' && !(Array.isArray(v) && !v.length); }).length;
      let obs = g('note').replace(/\s+/g, ' ').trim(); if (obs === 'xcxvxv' || obs === 'Esse funciona') obs = '';
      let uf = g('state').toUpperCase(); let cidade = titulo_(g('city'));
      if (ch === 'Tim' && !uf) { uf = 'PA'; cidade = 'Belém'; }
      let cls = g('cls').toUpperCase().slice(0, 2); if (['S', 'A', 'B', 'C'].indexOf(cls) < 0) cls = '-';
      const prRaw = g('promo').toLowerCase(), p = semAcento_(prRaw).replace(/[^\x00-\x7f]/g, '');
      const prom = (p.indexOf('exclu') >= 0 || p.indexOf('exlcu') >= 0 || prRaw.indexOf('专') >= 0) ? 'Exclusivo' : (p.indexOf('compart') >= 0 || prRaw.indexOf('共') >= 0) ? 'Compartilhado' : (p.indexOf('sem') >= 0 || prRaw.indexOf('无') >= 0) ? 'Sem promotor' : '-';
      const tg = etiquetas_(obs);
      const sit = (fotos === 0) ? (okp === 'OK' ? 'OK' : 'Pendente') : (tg.length ? 'OK c/ apontamento' : 'OK');
      if (fotos === 0 && okp === 'OK') fotos = 1; // a fórmula do Feishu já viu a foto
      let nd = 0; dispCols.forEach((c) => { const v = txt(r, c); if (!(v === '' || v === 'NO' || v === '0')) nd++; });
      const t16 = num('t16'), t12 = num('t12'), t08 = num('t08');
      let veic = g('date'); if (/^\d{5}(\.\d+)?$/.test(veic)) veic = Utilities.formatDate(new Date(Date.UTC(1899, 11, 30) + Math.round(+veic) * 864e5), 'UTC', 'yyyy-MM-dd');
      const ult = ultimaDe(ch, C.code ? g('code') : '', nome);
      const codigo = C.code ? g('code').replace(/\.0$/, '') : '';
      lojas.push({ id: ch + '-' + r, canal: ch, codigo: codigo === '0' ? '' : codigo, loja: nome, cidade, uf, regiao: PAINEL_REGIAO[uf] || 'Outra', classe: cls, promotor: prom,
        mesa_16: t16, mesa_12: t12, mesa_08: t08, mesas: t16 + t12 + t08, displays: nd, fotos, situacao: sit, ok_planilha: okp, apontamento: tg.join('; '), observacao: obs,
        veiculacao: veic, ultima_foto: ult, fotos_hash: '', atualizado_em: hoje, dias_sem_foto: ult ? Math.round((Date.parse(hoje) - Date.parse(ult)) / 864e5) : null });
      n++;
    }
    resumo.push(ch + ': ' + n + ' lojas, ' + lojas.filter((l) => l.canal === ch && l.situacao !== 'Pendente').length + ' com foto');
  });
  // lojas dispensadas de foto no site (sem mesa etc.) não entram na base nem no histórico
  const disp = {}; (supa_('GET', '/rest/v1/lojas_ajustes?select=chave&tipo=eq.dispensa') || []).forEach((a) => { disp[a.chave] = 1; });
  const chaveL = (l) => l.canal + '|' + String(l.codigo || l.loja || '').trim().toUpperCase();
  const nAntes = lojas.length;
  for (let i = lojas.length - 1; i >= 0; i--) if (disp[chaveL(lojas[i])]) lojas.splice(i, 1);
  if (nAntes !== lojas.length) resumo.push('Dispensadas de foto (fora da conta): ' + (nAntes - lojas.length));
  if (detalhado) resumo.forEach((s) => console.log(s));
  if (lojas.length < 50) throw new Error('Li só ' + lojas.length + ' lojas no Feishu — não atualizei o dash. ' + resumo.join(' · '));

  // mantém histórico e prioridades que já estão no dash
  const atual = (supa_('GET', '/rest/v1/painel_dados?id=eq.atual&select=dados') || [])[0];
  const base = (atual && atual.dados) || {};
  const grp = {};
  lojas.forEach((l) => { const k = l.canal + '|' + l.regiao; const x = grp[k] || (grp[k] = { data: hoje, canal: l.canal, regiao: l.regiao, lojas: 0, com_foto: 0, ok_apont: 0, pendente: 0 });
    x.lojas++; if (l.situacao !== 'Pendente') x.com_foto++; if (l.situacao === 'OK c/ apontamento') x.ok_apont++; if (l.situacao === 'Pendente') x.pendente++; });
  const historico = (base.historico || []).filter((h) => h.data !== hoje).concat(Object.keys(grp).map((k) => grp[k]));
  const dados = { atualizado: hoje, lojas, historico, prioridades: base.prioridades || [] };
  supa_('POST', '/rest/v1/painel_dados?on_conflict=id', [{ id: 'atual', dados, atualizado_em: new Date().toISOString(), atualizado_por: 'Feishu (automático)' }], { Prefer: 'resolution=merge-duplicates,return=minimal' });
  const com = lojas.filter((l) => l.situacao !== 'Pendente').length;
  console.log('Dash atualizado pelo Feishu: ' + lojas.length + ' lojas, ' + com + ' com foto (' + Math.round(100 * com / lojas.length) + '%).');
}

// ---------- utilidades ----------
function valoresBrutos_(tk, planilha, range) {
  const d = feishu_(tk, 'GET', '/sheets/v2/spreadsheets/' + planilha + '/values/' + encodeURIComponent(range) + '?valueRenderOption=FormattedValue&dateTimeRenderOption=FormattedString');
  return (d.valueRange && d.valueRange.values) || [];
}
function titulo_(s) { return String(s || '').toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (a, b, c) => b + c.toUpperCase()).trim(); }
function etiquetas_(o) {
  const n = semAcento_(o).replace(/[^\x00-\x7f]/g, ''), t = [];
  if (/qualida|foto|escuro|irregular|nao mostra/.test(n)) t.push('Foto fora do padrão');
  if (n.indexOf('spec') >= 0) t.push('Sem spec card');
  if (n.indexOf('demo') >= 0) t.push('Sem modo demo');
  if (/mostruario|estoque/.test(n)) t.push('Mostruário / estoque');
  if (/mesa|bancada|atuacao/.test(n)) t.push('Sem mesa / sem espaço');
  if (/canva|material|chegou|inventario|transito/.test(n)) t.push('Material / canvas');
  return t;
}

// ---------- alerta por e-mail: lojas com a última foto aprovada há 30+ dias ----------
var PAINEL_SITE = 'https://painel-vm-shatan.gabrielkupkaoppo.workers.dev';
var PAINEL_DIAS_ALERTA = 30;
function alertaLigar() {
  ScriptApp.getProjectTriggers().forEach((t) => { if (t.getHandlerFunction() === 'alertaFotos30') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('alertaFotos30').timeBased().everyDays(1).atHour(8).inTimezone('America/Sao_Paulo').create();
  console.log('Alerta diário ligado (todo dia ~8h). Mandando um agora para testar…');
  alertaEnviar_(true);
}
function alertaDesligar() {
  ScriptApp.getProjectTriggers().forEach((t) => { if (t.getHandlerFunction() === 'alertaFotos30') ScriptApp.deleteTrigger(t); });
  console.log('Alerta diário desligado.');
}
function alertaFotos30() { alertaEnviar_(false); }
function alertaEnviar_(forcar) {
  const props = PropertiesService.getScriptProperties();
  const atual = (supa_('GET', '/rest/v1/painel_dados?id=eq.atual&select=dados') || [])[0];
  const lojas = ((atual && atual.dados && atual.dados.lojas) || []).filter((l) => l.dias_sem_foto != null && l.dias_sem_foto >= PAINEL_DIAS_ALERTA)
    .sort((a, b) => b.dias_sem_foto - a.dias_sem_foto);
  const chaves = lojas.map((l) => l.canal + '|' + (l.codigo || l.loja)).sort();
  const anteriores = JSON.parse(props.getProperty('ALERTA30') || '[]');
  const novas = chaves.filter((k) => anteriores.indexOf(k) < 0);
  props.setProperty('ALERTA30', JSON.stringify(chaves));
  if (!lojas.length) { console.log('Nenhuma loja com foto de ' + PAINEL_DIAS_ALERTA + '+ dias.'); if (!forcar) return; }
  if (!novas.length && !forcar) { console.log('Sem lojas novas passando de ' + PAINEL_DIAS_ALERTA + ' dias — não mandei e-mail.'); return; }
  const admins = (supa_('GET', '/rest/v1/perfis?select=email&papel=eq.admin&ativo=is.true') || []).map((p) => p.email).filter(Boolean);
  const para = admins.length ? admins.join(',') : Session.getEffectiveUser().getEmail();
  const esc = (t) => String(t == null ? '' : t).replace(/&/g, '&amp;').replace(/</g, '&lt;');
  const linhas = lojas.map((l) => '<tr' + (novas.indexOf(l.canal + '|' + (l.codigo || l.loja)) >= 0 ? ' style="background:#fff4e0"' : '') + '><td>' + esc(l.codigo) + '</td><td>' + esc(l.loja) + '</td><td>' + esc(l.canal) +
    '</td><td>' + esc(l.uf) + '</td><td style="text-align:right">' + l.dias_sem_foto + '</td><td>' + esc(l.ultima_foto) + '</td></tr>').join('');
  const html = '<p>Olá! Estas lojas estão com a <b>última foto aprovada há ' + PAINEL_DIAS_ALERTA + ' dias ou mais</b> e precisam mandar foto nova pelo Forms.' +
    (novas.length ? ' Em destaque, as <b>' + novas.length + '</b> que passaram do prazo desde o último aviso.' : '') + '</p>' +
    '<table cellpadding="6" style="border-collapse:collapse;font-family:Arial;font-size:13px" border="1"><tr style="background:#eef1ef"><th>Código</th><th>Loja</th><th>Canal</th><th>UF</th><th>Dias</th><th>Última foto</th></tr>' +
    (linhas || '<tr><td colspan="6">Nenhuma loja no momento.</td></tr>') + '</table>' +
    '<p>Ver no painel: <a href="' + PAINEL_SITE + '">' + PAINEL_SITE + '</a> (filtro "Foto com 30+ dias").</p><p style="color:#888">Aviso automático do Painel VM Shatan · 体验台照片超过30天提醒</p>';
  MailApp.sendEmail({ to: para, subject: '⚠ ' + lojas.length + ' lojas com foto de ' + PAINEL_DIAS_ALERTA + '+ dias · Painel VM Shatan', htmlBody: html });
  console.log('E-mail enviado para ' + para + ': ' + lojas.length + ' lojas (' + novas.length + ' novas).');
}
