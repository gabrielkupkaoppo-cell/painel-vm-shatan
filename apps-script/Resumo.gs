/**
 * Painel VM Shatan — resumo diário no Feishu (mensagem do app "Painel VM Shatan").
 * Quarto arquivo do mesmo Apps Script (usa funções do Sincronizar.gs, Feishu.gs e Painel.gs).
 *
 * Propriedade do script:
 *   FEISHU_RESUMO_PARA = para quem mandar. Pode ser:
 *      - e-mail da pessoa no Feishu (ex.: fulano@oppo.com)
 *      - celular com DDI (ex.: +5591999998888)
 *      - id de um grupo (começa com oc_)
 *     Vários destinos: separe com vírgula.
 *
 * 1) Rode resumoTeste()  → manda o resumo agora.
 * 2) Gostou? Rode resumoLigar() → manda sozinho todo dia ~8h. resumoDesligar() para parar.
 */
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

// e-mail, celular ou grupo → id que o Feishu entende
function resumoDestino_(tk, dest) {
  if (/^oc_/.test(dest)) return { tipo: 'chat_id', id: dest };
  if (/^ou_/.test(dest)) return { tipo: 'open_id', id: dest };
  if (dest.indexOf('@') > 0) return { tipo: 'email', id: dest };
  let fone = dest.replace(/[^\d+]/g, ''); if (fone.charAt(0) !== '+') fone = '+' + (fone.length <= 11 ? '55' : '') + fone;
  const d = feishu_(tk, 'POST', '/contact/v3/users/batch_get_id?user_id_type=open_id', { mobiles: [fone], include_resigned: false });
  const u = (d.user_list || [])[0];
  if (!u || !u.user_id) throw new Error('Não achei ninguém no Feishu com o celular ' + fone + '. Confira o número ou use o e-mail da pessoa.');
  return { tipo: 'open_id', id: u.user_id };
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
