/**
 * Painel VM Shatan — cria uma aba de canal no Feishu com a MESMA estrutura de outra aba.
 * Arquivo do mesmo Apps Script (usa funções do Feishu.gs). Trabalha na planilha de FEISHU_PLANILHA.
 *
 * criarAbaCasasBahia() → copia a aba "Magalu": cabeçalho (português/chinês), colunas, caixas de seleção,
 * fórmula do Ok/Pendente e colunas de foto. As lojas da Magalu NÃO são copiadas: ficam 20 linhas em branco
 * já formatadas. As lojas da Casas Bahia entram pela aba Instalações do site (ou à mão).
 *
 * Se já existir uma aba "Casas Bahia" VAZIA (só criada, sem nada escrito), ela é substituída.
 * Se ela já tiver qualquer conteúdo, o script para e não mexe em nada.
 */
function criarAbaCasasBahia() { copiarEstruturaAba_('Casas Bahia', 'Magalu', 20); }

function copiarEstruturaAba_(nova, modelo, linhasVazias) {
  const planilha = String(PropertiesService.getScriptProperties().getProperty('FEISHU_PLANILHA') || '').trim();
  if (!planilha) throw new Error('Falta a propriedade FEISHU_PLANILHA.');
  const tk = feishuToken_();
  const base = '/sheets/v2/spreadsheets/' + planilha;
  const listar = () => feishu_(tk, 'GET', '/sheets/v3/spreadsheets/' + planilha + '/sheets/query').sheets || [];
  let abas = listar();
  const mod = abas.find((a) => semAcento_(a.title) === semAcento_(modelo));
  if (!mod) throw new Error('Não achei a aba modelo "' + modelo + '".');

  // aba com o mesmo nome: só substitui se estiver vazia
  const ja = abas.find((a) => semAcento_(a.title) === semAcento_(nova));
  if (ja) {
    const v = valores_(tk, planilha, ja.sheet_id + '!A1:Z60');
    const temAlgo = v.some((row) => (row || []).some((c) => texto_(c)));
    if (temAlgo) throw new Error('A aba "' + ja.title + '" já tem conteúdo. Para não perder nada, apague ou renomeie essa aba no Feishu e rode de novo.');
    feishu_(tk, 'POST', base + '/sheets_batch_update', { requests: [{ deleteSheet: { sheetId: ja.sheet_id } }] });
    console.log('Aba "' + ja.title + '" vazia removida para ser recriada com a estrutura da ' + mod.title + '.');
  }

  // 1) copia a aba modelo
  const cp = feishu_(tk, 'POST', base + '/sheets_batch_update', { requests: [{ copySheet: { source: { sheetId: mod.sheet_id }, destination: { title: nova } } }] });
  const sid = cp.replies[0].copySheet.properties.sheetId;
  abas = listar();
  const ab = abas.find((a) => a.sheet_id === sid) || {};
  const nCol = Math.min((ab.grid_properties && ab.grid_properties.column_count) || 40, 60), L = letra_(nCol);
  const nLin = (ab.grid_properties && ab.grid_properties.row_count) || 500;

  // 2) cabeçalho e linhas de loja do modelo
  const cab = valores_(tk, planilha, sid + '!A1:' + L + '6');
  let hdr = 0; for (let r = 0; r < cab.length && !hdr; r++) if ((cab[r] || []).some((v) => texto_(v).indexOf('Store Name') >= 0)) hdr = r + 1;
  if (!hdr) hdr = 2;
  let colNome = 2; (cab[hdr - 1] || []).forEach((v, c) => { if (texto_(v).indexOf('Store Name') >= 0) colNome = c + 1; });
  const nomes = valores_(tk, planilha, sid + '!' + letra_(colNome) + '1:' + letra_(colNome) + nLin).map((r) => texto_(r && r[0]));
  let fim = hdr; for (let r = hdr; r < nomes.length; r++) if (nomes[r]) fim = r + 1;
  const M = fim - hdr, K = Math.min(M, linhasVazias);
  if (!M) throw new Error('A aba modelo não tem linhas de loja para copiar o formato.');

  // 3) apaga as lojas que sobram (fica só o número de linhas em branco pedido)
  if (M > K) feishu_(tk, 'DELETE', base + '/dimension_range', { dimension: { sheetId: sid, majorDimension: 'ROWS', startIndex: hdr + K + 1, endIndex: fim } });

  // 4) limpa o conteúdo das linhas que ficaram, mantendo formato, caixas e fórmulas
  const de = hdr + 1, ate = hdr + K;
  const form = (feishu_(tk, 'GET', base + '/values/' + encodeURIComponent(sid + '!A' + de + ':' + L + ate) + '?valueRenderOption=Formula').valueRange || {}).values || [];
  const ehFormula = (v) => (v && typeof v === 'object' && v.type === 'formula') || (typeof v === 'string' && v.charAt(0) === '=');
  const linhas = [];
  for (let i = 0; i < K; i++) {
    const row = form[i] || [], out = [];
    for (let c = 0; c < nCol; c++) out.push(ehFormula(row[c]) ? row[c] : null);   // null = célula vazia; fórmula fica
    linhas.push(out.map((v) => (v && typeof v === 'string' && v.charAt(0) === '=' ? { type: 'formula', text: v } : v)));
  }
  feishu_(tk, 'POST', base + '/values_batch_update', { valueRanges: [{ range: sid + '!A' + de + ':' + L + ate, values: linhas }] });

  // 5) confere: as linhas devem ficar sem nome de loja
  const sobra = valores_(tk, planilha, sid + '!' + letra_(colNome) + de + ':' + letra_(colNome) + ate).filter((r) => texto_(r && r[0])).length;
  console.log('Aba "' + nova + '" criada com a estrutura da ' + mod.title + ': cabeçalho na linha ' + hdr + ', ' + K + ' linhas em branco formatadas (' + de + ' a ' + ate + ').' +
    (sobra ? ' ⚠ ' + sobra + ' linha(s) ainda com texto da ' + mod.title + ': selecione as linhas ' + de + ' a ' + ate + ' e aperte Delete.' : ''));
  console.log('Confira no Feishu se ficou alguma foto da ' + mod.title + ' nas linhas ' + de + ' a ' + ate + ' e, se ficou, selecione as colunas de foto dessas linhas e aperte Delete.');
}
