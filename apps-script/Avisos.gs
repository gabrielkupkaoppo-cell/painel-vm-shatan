/**
 * Painel VM Shatan — avisos no Feishu (mensagens do app "Painel VM Shatan").
 * Quarto arquivo do mesmo Apps Script (usa funções do Sincronizar.gs, Feishu.gs e Painel.gs).
 *
 * A) AVISO DE APROVAÇÃO/REPROVAÇÃO → responsável do estado da loja
 *    Os responsáveis são cadastrados no site: aba Usuários → "Responsáveis por estado"
 *    (campo Feishu = e-mail da pessoa no Feishu ou celular com +55).
 *    1) Rode avisosTeste()  → manda agora os avisos pendentes e mostra no registro.
 *    2) Deu certo? Rode avisosLigar() → passa a mandar sozinho a cada 10 min. avisosDesligar() para parar.
 *
 * B) RESUMO DIÁRIO (opcional)
 *    Propriedade FEISHU_RESUMO_PARA = e-mail, celular com +55 ou id de grupo (oc_...). Vários: separe com vírgula.
 *    resumoTeste() manda agora; resumoLigar() manda todo dia ~8h; resumoDesligar() para.
 */

// ===================== A) avisos para o responsável do estado =====================
function avisosTeste() { avisarResponsaveis_(true); }
function avisarResponsaveis() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try { avisarResponsaveis_(false); } finally { lock.releaseLock(); }
}
function avisosLigar() {
  ScriptApp.getProjectTriggers().forEach((t) => { if (t.getHandlerFunction() === 'avisarResponsaveis') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('avisarResponsaveis').timeBased().everyMinutes(10).create();
  console.log('Avisos para os responsáveis ligados (a cada 10 min).');
}
function avisosDesligar() {
  ScriptApp.getProjectTriggers().forEach((t) => { if (t.getHandlerFunction() === 'avisarResponsaveis') ScriptApp.deleteTrigger(t); });
  console.log('Avisos para os responsáveis desligados.');
}

function avisarResponsaveis_(detalhado) {
  const pend = supa_('GET', '/rest/v1/revisoes?select=key,status,motivos,obs,revisado_em,aviso_info&aviso_em=is.null&order=revisado_em.asc&limit=300') || [];
  if (!pend.length) { if (detalhado) console.log('Nenhuma aprovação/reprovação nova para avisar.'); return; }
  const marcar = (key, info) => supa_('PATCH', '/rest/v1/revisoes?key=eq.' + encodeURIComponent(key), { aviso_em: new Date().toISOString(), aviso_info: info }, { Prefer: 'return=minimal' });
  const resp = {}; (supa_('GET', '/rest/v1/responsaveis?select=*') || []).forEach((x) => { resp[String(x.uf || '').toUpperCase()] = x; });
  const porUf = {};
  pend.forEach((r) => {
    if (r.status !== 'Aprovado' && r.status !== 'Reprovado') { marcar(r.key, 'sem aviso (' + r.status + ')'); return; }
    if (/^Validada antes/.test(r.obs || '')) { marcar(r.key, 'sem aviso (baixa em lote)'); return; }
    if (Date.now() - new Date(r.revisado_em) > 3 * 864e5) { marcar(r.key, 'sem aviso (decisão antiga)'); return; }
    const e = (supa_('GET', '/rest/v1/envios?select=key,canal,loja_raw,loja,uf,nome,data&key=eq.' + encodeURIComponent(r.key)) || [])[0];
    if (!e) { marcar(r.key, 'sem aviso (envio não encontrado)'); return; }
    const uf = String(e.uf || '').toUpperCase();
    (porUf[uf] = porUf[uf] || []).push({ r: r, e: e });
  });
  const ufs = Object.keys(porUf);
  if (!ufs.length) return;
  const tk = feishuToken_();
  ufs.forEach((uf) => {
    const itens = porUf[uf], sup = resp[uf];
    if (!sup || !String(sup.feishu || '').trim()) {
      itens.forEach((x) => marcar(x.r.key, 'sem responsável no Feishu para ' + (uf || 'estado vazio')));
      console.log(uf + ': sem responsável com Feishu cadastrado — ' + itens.length + ' aviso(s) não enviados.');
      return;
    }
    try {
      const alvo = resumoDestino_(tk, String(sup.feishu).trim(), sup.whatsapp);
      for (let i = 0; i < itens.length; i += 15) {
        const lote = itens.slice(i, i + 15);
        feishu_(tk, 'POST', '/im/v1/messages?receive_id_type=' + alvo.tipo, { receive_id: alvo.id, msg_type: 'interactive', content: JSON.stringify(avisoCartao_(uf, sup, lote)) });
        lote.forEach((x) => marcar(x.r.key, 'avisado no Feishu: ' + (sup.nome || sup.feishu)));
      }
      console.log(uf + ': ' + itens.length + ' aviso(s) enviados para ' + (sup.nome || sup.feishu) + '.');
    } catch (err) {
      const msg = String(err.message || err).slice(0, 300);
      itens.forEach((x) => {
        const n = +((String(x.r.aviso_info || '').match(/tentativa (\d+)/) || [])[1] || 0) + 1;
        if (n >= 3) marcar(x.r.key, 'NÃO ENVIADO após 3 tentativas: ' + msg);
        else supa_('PATCH', '/rest/v1/revisoes?key=eq.' + encodeURIComponent(x.r.key), { aviso_info: 'erro (tentativa ' + n + '): ' + msg }, { Prefer: 'return=minimal' });
      });
      console.log(uf + ': erro ao avisar ' + (sup.nome || sup.feishu) + ' — ' + msg);
    }
  });
}

function avisoCartao_(uf, sup, itens) {
  const tz = 'America/Sao_Paulo';
  const ap = itens.filter((x) => x.r.status === 'Aprovado').length, rp = itens.length - ap;
  const linhas = itens.map((x) => {
    const e = x.e, r = x.r;
    const loja = canalFeishu_(e.canal) + ' ' + (e.loja || e.loja_raw || '');
    const quando = e.data ? Utilities.formatDate(new Date(e.data), tz, 'dd/MM HH:mm') : '';
    const cab = (r.status === 'Aprovado' ? '✅ **APROVADA**' : '❌ **REPROVADA**') + ' · ' + loja + ' (' + uf + ')';
    const det = 'Enviada por ' + (String(e.nome || '—').trim()) + (quando ? ' em ' + quando : '');
    const extra = (r.status === 'Reprovado' && r.motivos ? '\nMotivo · 原因: ' + r.motivos : '') + (r.obs ? '\nObs. · 备注: ' + r.obs : '');
    return cab + '\n' + det + extra;
  });
  const nome = String(sup.nome || '').trim().split(/\s+/)[0];
  const intro = (nome ? 'Olá ' + nome + '! ' : '') + 'Fotos das mesas das lojas de ' + uf + ' revisadas agora:';
  const fim = rp ? '\n\n⚠️ Para as reprovadas: oriente a equipe a corrigir e enviar de novo pelo Forms.' : '';
  return {
    config: { wide_screen_mode: true },
    header: { template: rp ? 'orange' : 'green', title: { tag: 'plain_text', content: 'Fotos das mesas · 体验台照片审核 — ' + uf + ' (' + (ap ? ap + ' aprovada' + (ap > 1 ? 's' : '') : '') + (ap && rp ? ', ' : '') + (rp ? rp + ' reprovada' + (rp > 1 ? 's' : '') : '') + ')' } },
    elements: [
      { tag: 'div', text: { tag: 'lark_md', content: intro } },
      { tag: 'hr' },
      { tag: 'div', text: { tag: 'lark_md', content: linhas.join('\n\n') + fim } },
      { tag: 'action', actions: [{ tag: 'button', type: 'primary', text: { tag: 'plain_text', content: 'Abrir o painel · 打开看板' }, url: PAINEL_SITE }] },
    ],
  };
}

// ===================== B) resumo diário =====================
function resumoTeste() { resumoEnviar_(); }
function resumoDiario() { resumoEnviar_(); }
function resumoLigar() {
  ScriptApp.getProjectTriggers().forEach((t) => { if (t.getHandlerFunction() === 'resumoDiario') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('resumoDiario').timeBased().everyDays(1).atHour(8).inTimezone('America/Sao_Paulo').create();
  console.log('Resumo diário no Feishu ligado (todo dia ~8h).');
}
function resumoDesligar() {
  ScriptApp.getProjectTriggers().forEach((t) => { if (t.getHandlerFunction() === 'resumoDiario') ScriptApp.deleteTrigger(t); });
  console.log('Resumo diário no Feishu desligado.');
}

function resumoEnviar_() {
  const para = String(PropertiesService.getScriptProperties().getProperty('FEISHU_RESUMO_PARA') || '').split(',').map((s) => s.trim()).filter(Boolean);
  if (!para.length) throw new Error('Falta a propriedade FEISHU_RESUMO_PARA (e-mail, celular com +55 ou id do grupo).');
  const card = resumoCartao_();
  const tk = feishuToken_();
  para.forEach((dest) => {
    const alvo = resumoDestino_(tk, dest);
    feishu_(tk, 'POST', '/im/v1/messages?receive_id_type=' + alvo.tipo, { receive_id: alvo.id, msg_type: 'interactive', content: JSON.stringify(card) });
    console.log('Resumo enviado para ' + dest + '.');
  });
}

// e-mail, celular ou grupo → id que o Feishu entende (procura a pessoa no Feishu da empresa)
// "reserva" = outro número para tentar (ex.: o WhatsApp do responsável), se o e-mail não for achado
function resumoDestino_(tk, dest, reserva) {
  dest = String(dest || '').trim();
  if (/^oc_/.test(dest)) return { tipo: 'chat_id', id: dest };
  if (/^ou_/.test(dest)) return { tipo: 'open_id', id: dest };
  const tentativas = [];
  const fone_ = (v) => { let f = String(v || '').replace(/[^\d+]/g, ''); if (!f) return ''; if (f.charAt(0) !== '+') f = '+' + (f.replace(/^0+/, '').length <= 11 ? '55' : '') + f.replace(/^0+/, ''); return f; };
  if (dest.indexOf('@') > 0) tentativas.push({ emails: [dest.toLowerCase()] }); else if (dest) tentativas.push({ mobiles: [fone_(dest)] });
  const r = fone_(reserva); if (r && !(tentativas[0] && tentativas[0].mobiles && tentativas[0].mobiles[0] === r)) tentativas.push({ mobiles: [r] });
  for (let i = 0; i < tentativas.length; i++) {
    const d = feishu_(tk, 'POST', '/contact/v3/users/batch_get_id?user_id_type=open_id', Object.assign({ include_resigned: false }, tentativas[i]));
    const u = (d.user_list || []).find((x) => x.user_id);
    if (u) return { tipo: 'open_id', id: u.user_id };
  }
  throw new Error('Não achei "' + dest + '"' + (r ? ' nem o celular ' + r : '') + ' no Feishu da OPPO. Use o e-mail/celular que a pessoa usa para entrar no Feishu, e confira se ela está no "可用范围" (disponibilidade) do app.');
}

function resumoCartao_() {
  const tz = 'America/Sao_Paulo';
  const hoje0 = new Date(Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'00:00:00") + '-03:00');
  const ontem0 = new Date(hoje0.getTime() - 864e5);
  const envios = supa_('GET', '/rest/v1/envios?select=key,form,data,nome,canal,loja_raw,fotos&limit=20000') || [];
  const revs = {}; (supa_('GET', '/rest/v1/revisoes?select=key,status,revisado_em&limit=20000') || []).forEach((r) => { revs[r.key] = r; });
  const grupo = (e) => { const m = String(e.loja_raw || '').match(/\d+/g); return m ? e.form + '|' + canalFeishu_(e.canal) + '|' + (m[m.length - 1].replace(/^0+/, '') || '0') : 'k:' + e.key; };
  const entre = (d, a, b) => d >= a && d < b;

  // ontem
  const recebidos = envios.filter((e) => entre(new Date(e.data), ontem0, hoje0)).length;
  let aprov = 0, reprov = 0;
  Object.keys(revs).forEach((k) => { const r = revs[k]; if (!entre(new Date(r.revisado_em), ontem0, hoje0)) return; if (r.status === 'Aprovado') aprov++; else if (r.status === 'Reprovado') reprov++; });

  // fila: um envio por loja (o mais recente), igual ao site
  const decididoMaisNovo = {};
  envios.forEach((e) => { const r = revs[e.key]; if (!r || r.status === 'Substituído' || !(e.fotos || []).length) return; const g = grupo(e), d = new Date(e.data); if (!decididoMaisNovo[g] || d > decididoMaisNovo[g]) decididoMaisNovo[g] = d; });
  const fila = {};
  envios.forEach((e) => {
    if (revs[e.key]) return;
    const g = grupo(e), d = new Date(e.data);
    if (decididoMaisNovo[g] && decididoMaisNovo[g] > d) return;
    if (!fila[g] || d > new Date(fila[g].data)) fila[g] = e;
  });
  const pend = Object.keys(fila).map((g) => fila[g]);
  const maisAntigo = pend.reduce((m, e) => Math.max(m, Math.floor((Date.now() - new Date(e.data)) / 864e5)), 0);
  const porPromotor = {}; pend.forEach((e) => { const n = String(e.nome || '—').trim().split(/\s+/).slice(0, 2).join(' '); porPromotor[n] = (porPromotor[n] || 0) + 1; });
  const topProm = Object.keys(porPromotor).sort((a, b) => porPromotor[b] - porPromotor[a]).slice(0, 5);

  // 30+ dias (do painel)
  const atual = (supa_('GET', '/rest/v1/painel_dados?id=eq.atual&select=dados') || [])[0];
  const lojas = (atual && atual.dados && atual.dados.lojas) || [];
  const vencidas = lojas.filter((l) => l.dias_sem_foto != null && l.dias_sem_foto >= PAINEL_DIAS_ALERTA).sort((a, b) => b.dias_sem_foto - a.dias_sem_foto);
  const semFoto = lojas.filter((l) => l.situacao === 'Pendente').length;
  const ok = lojas.length - semFoto;

  const md = [];
  md.push('**Ontem · 昨天**  (' + Utilities.formatDate(ontem0, tz, 'dd/MM') + ')');
  md.push('📥 Envios recebidos · 收到: **' + recebidos + '**');
  md.push('✅ Aprovados · 通过: **' + aprov + '**    ❌ Reprovados · 不通过: **' + reprov + '**');
  md.push('');
  md.push('**Fila de aprovação · 待审核**');
  md.push('⏳ Lojas aguardando · 待审核门店: **' + pend.length + '**' + (pend.length ? '  (mais antiga há ' + maisAntigo + ' dias)' : ''));
  if (topProm.length) md.push(topProm.map((n) => '• ' + n + ': ' + porPromotor[n]).join('\n'));
  md.push('');
  md.push('**Painel · 看板**');
  if (lojas.length) md.push('🏬 Lojas com foto OK · 已完成: **' + ok + '** de ' + lojas.length + ' (' + Math.round(ok * 100 / lojas.length) + '%)');
  md.push('⚠️ Foto com ' + PAINEL_DIAS_ALERTA + '+ dias · 超过30天: **' + vencidas.length + '**');
  if (vencidas.length) md.push(vencidas.slice(0, 8).map((l) => '• ' + l.canal + ' ' + (l.codigo || '') + ' ' + titulo_(l.loja || '') + ' (' + (l.uf || '') + ') — ' + l.dias_sem_foto + ' dias').join('\n') + (vencidas.length > 8 ? '\n• … e mais ' + (vencidas.length - 8) : ''));

  return {
    config: { wide_screen_mode: true },
    header: { template: 'green', title: { tag: 'plain_text', content: 'Resumo VM Mesas · 体验台日报 — ' + Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy') } },
    elements: [
      { tag: 'div', text: { tag: 'lark_md', content: md.join('\n') } },
      { tag: 'hr' },
      { tag: 'action', actions: [{ tag: 'button', type: 'primary', text: { tag: 'plain_text', content: 'Abrir o painel · 打开看板' }, url: PAINEL_SITE }] },
      { tag: 'note', elements: [{ tag: 'plain_text', content: 'Aviso automático do Painel VM Shatan' }] },
    ],
  };
}
