/**
 * Painel VM Shatan — cria a aba "Gazin" na planilha do Feishu (a de FEISHU_PLANILHA).
 * Quinto arquivo do mesmo Apps Script (usa funções do Sincronizar.gs e do Feishu.gs).
 *
 * Rode criarAbaGazin() UMA vez. Ele:
 *   1) copia a aba "Bemol" (mesmo cabeçalho, caixas de seleção, fórmula do Ok/Pendente, colunas de foto);
 *   2) apaga as lojas da Bemol da cópia e coloca uma linha para cada loja Gazin que já apareceu no Forms;
 *   3) libera as aprovações da Gazin que não tinham ido para o Feishu (por falta da aba) para irem agora.
 * Se já existir uma aba "Gazin", ele não faz nada (apague/renomeie a antiga antes, se quiser recriar).
 *
 * Teste primeiro na planilha de teste (testecloude). Depois, se quiser na oficial, troque FEISHU_PLANILHA e rode de novo.
 */
var GAZIN_MODELO = 'Bemol';          // aba copiada como modelo (tem mesas OPPO 1.2M/1.6M/0.8M, para o futuro)
var GAZIN_MULTIMARCAS = true;        // cria a coluna "Multimarcas" (caixa marcada quando o promotor escolhe Mesa Multimarcas)

function criarAbaGazin() { criarAbaCanal_('Gazin', GAZIN_MODELO); }

function criarAbaCanal_(canal, modelo) {
  const planilha = String(PropertiesService.getScriptProperties().getProperty('FEISHU_PLANILHA') || '').trim();
  if (!planilha) throw new Error('Falta a propriedade FEISHU_PLANILHA.');
  const tk = feishuToken_();
  const base = '/sheets/v2/spreadsheets/' + planilha;
  let abas = feishu_(tk, 'GET', '/sheets/v3/spreadsheets/' + planilha + '/sheets/query').sheets || [];
  if (abas.find((a) => semAcento_(a.title) === semAcento_(canal))) throw new Error('Já existe uma aba "' + canal + '" nesta planilha. Apague ou renomeie a antiga se quiser criar de novo.');
  const mod = abas.find((a) => semAcento_(a.title) === semAcento_(modelo));
  if (!mod) throw new Error('Não achei a aba modelo "' + modelo + '".');

  // 1) lojas da Gazin que já apareceram no Forms
  const lojas = lojasDoForms_(canal);
  if (!lojas.length) throw new Error('Nenhuma loja ' + canal + ' com código encontrada nos envios.');
  console.log('Lojas ' + canal + ' encontradas no Forms: ' + lojas.length);

  // 2) copia a aba modelo
  const cp = feishu_(tk, 'POST', base + '/sheets_batch_update', { requests: [{ copySheet: { source: { sheetId: mod.sheet_id }, destination: { title: canal } } }] });
  const sid = cp.replies[0].copySheet.properties.sheetId;
  abas = feishu_(tk, 'GET', '/sheets/v3/spreadsheets/' + planilha + '/sheets/query').sheets || [];
  const nova = abas.find((a) => a.sheet_id === sid) || {};
  let nCol = Math.min((nova.grid_properties && nova.grid_properties.column_count) || 40, 60);
  // 2b) coluna "Multimarcas" logo depois das colunas de tamanho de mesa (herda a caixa de seleção da vizinha)
  if (GAZIN_MULTIMARCAS) {
    const cab0 = valores_(tk, planilha, sid + '!A1:' + letra_(nCol) + '6');
    let hdr0 = 0; for (let r = 0; r < cab0.length && !hdr0; r++) if ((cab0[r] || []).some((v) => texto_(v).indexOf('Store Name') >= 0)) hdr0 = r + 1;
    if (!hdr0) hdr0 = 2;
    let ultMesa = 0, jaTem = false;
    for (let r = 0; r < hdr0; r++) (cab0[r] || []).forEach((v, c) => { const t = texto_(v); if (/[018][.,][268]\s*M/i.test(t) && t.indexOf('体验台') >= 0 || /^[018][.,][268]\s*M/i.test(t)) ultMesa = Math.max(ultMesa, c + 1); if (/multimarca/i.test(t)) jaTem = true; });
    if (!jaTem) {
      const depois = ultMesa || nCol;
      feishu_(tk, 'POST', base + '/insert_dimension_range', { dimension: { sheetId: sid, majorDimension: 'COLUMNS', startIndex: depois, endIndex: depois + 1 }, inheritStyle: 'BEFORE' });
      nCol = Math.min(nCol + 1, 60);
      feishu_(tk, 'POST', base + '/values_batch_update', { valueRanges: [{ range: sid + '!' + letra_(depois + 1) + hdr0 + ':' + letra_(depois + 1) + hdr0, values: [['Multimarcas']] }] });
      console.log('Coluna "Multimarcas" criada na coluna ' + letra_(depois + 1) + (ultMesa ? ' (depois das mesas 1.2M/1.6M/0.8M).' : ' (no final).'));
    }
  }
  const nLin = (nova.grid_properties && nova.grid_properties.row_count) || 500;
  const L = letra_(nCol);
  console.log('Aba "' + canal + '" criada como cópia de "' + modelo + '" (' + nCol + ' colunas).');

  // 3) cabeçalho e colunas
  const cab = valores_(tk, planilha, sid + '!A1:' + L + '6');
  let hdr = 0; for (let r = 0; r < cab.length && !hdr; r++) if ((cab[r] || []).some((v) => texto_(v).indexOf('Store Name') >= 0)) hdr = r + 1;
  if (!hdr) hdr = 2;
  const H = {}; for (let r = 0; r < hdr; r++) (cab[r] || []).forEach((v, c) => { const t = texto_(v); if (t) H[c + 1] = (H[c + 1] ? H[c + 1] + ' ' : '') + t; });
  const achar = (re) => { for (let c = 1; c <= nCol; c++) if (H[c] && re.test(H[c])) return c; return null; };
  const C = { code: achar(/Store Code/i), name: achar(/Store Name/i), city: achar(/City/i), state: (() => { for (let c = 1; c <= nCol; c++) if (H[c] && /\bstate\b|estado/i.test(H[c]) && !/currently/i.test(H[c])) return c; return null; })(), note: achar(/备注anota|anota/i), promo: achar(/promotor exclusivo/i) };
  const gerenciadas = CAIXAS_GERENCIADAS_;
  console.log('Cabeçalho na linha ' + hdr + '. Código: ' + (C.code ? letra_(C.code) : '—') + ', Nome: ' + (C.name ? letra_(C.name) : '—') + ', Estado: ' + (C.state ? letra_(C.state) : '—') + ', Obs.: ' + (C.note ? letra_(C.note) : '—'));

  // linhas de dados do modelo (até a última com nome de loja)
  const nomes = valores_(tk, planilha, sid + '!' + letra_(C.name || 2) + '1:' + letra_(C.name || 2) + nLin).map((r) => texto_(r && r[0]));
  let fim = hdr; for (let r = hdr; r < nomes.length; r++) if (nomes[r]) fim = r + 1;
  const M = fim - hdr, N = lojas.length, linha1 = hdr + 1;
  if (!M) throw new Error('A aba modelo não tem linhas de loja para copiar o formato.');

  // linha modelo: fórmulas e valores
  const tplF = (feishu_(tk, 'GET', base + '/values/' + encodeURIComponent(sid + '!A' + linha1 + ':' + L + linha1) + '?valueRenderOption=Formula').valueRange || {}).values || [[]];
  const tplV = (feishu_(tk, 'GET', base + '/values/' + encodeURIComponent(sid + '!A' + linha1 + ':' + L + linha1) + '?valueRenderOption=ToString').valueRange || {}).values || [[]];
  const formulaDe = (v) => { if (v && typeof v === 'object' && v.type === 'formula') return v.text; if (typeof v === 'string' && v.charAt(0) === '=') return v; return null; };

  // 4) acerta a quantidade de linhas: mais linhas herdando o formato, ou apaga as que sobram
  if (N > M) {
    feishu_(tk, 'POST', base + '/insert_dimension_range', { dimension: { sheetId: sid, majorDimension: 'ROWS', startIndex: fim, endIndex: fim + (N - M) }, inheritStyle: 'BEFORE' });
  } else if (N < M) {
    feishu_(tk, 'DELETE', base + '/dimension_range', { dimension: { sheetId: sid, majorDimension: 'ROWS', startIndex: hdr + N + 1, endIndex: fim } });
  }

  // 5) escreve as lojas
  const ajusta = (f, linha) => f.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/g, (m, d1, col, d2, num) => (d2 === '' && +num === linha1 ? d1 + col + linha : m));
  const linhas = lojas.map((lj, i) => {
    const linha = linha1 + i, out = [];
    for (let c = 1; c <= nCol; c++) {
      const f = formulaDe((tplF[0] || [])[c - 1]);
      const v = texto_((tplV[0] || [])[c - 1]);
      const h = H[c] || '';
      if (c === C.code) out.push(lj.codigo);
      else if (c === C.name) out.push(canal.toUpperCase() + ' FILIAL ' + lj.codigo);
      else if (c === C.state) out.push(lj.uf);
      else if (c === C.city) out.push('');
      else if (c === C.note) out.push(lj.nota);
      else if (f) out.push({ type: 'formula', text: ajusta(f, linha) });
      else if (h.indexOf('照片') >= 0) out.push('');
      else if (gerenciadas.indexOf(normItem_(h.split(' ').pop())) >= 0 || gerenciadas.some((g) => normItem_(h).indexOf(g) >= 0) || /^(true|false)$/i.test(v)) out.push(0);
      else out.push('');
    }
    return out;
  });
  // escreve em blocos de colunas, pulando as colunas de foto (elas precisam ficar VAZIAS de verdade para o Ok/Pendente funcionar)
  const fotoCol = {}; for (let c = 1; c <= nCol; c++) if ((H[c] || '').indexOf('照片') >= 0) fotoCol[c] = true;
  const blocos = []; let ini = null;
  for (let c = 1; c <= nCol + 1; c++) { if (c <= nCol && !fotoCol[c]) { if (ini === null) ini = c; } else if (ini !== null) { blocos.push([ini, c - 1]); ini = null; } }
  for (let i = 0; i < linhas.length; i += 40) {
    const de = linha1 + i, ate = de + Math.min(40, linhas.length - i) - 1;
    const vr = blocos.map(([a, b]) => ({ range: sid + '!' + letra_(a) + de + ':' + letra_(b) + ate, values: linhas.slice(i, i + 40).map((row) => row.slice(a - 1, b)) }));
    feishu_(tk, 'POST', base + '/values_batch_update', { valueRanges: vr });
  }
  const fc = Object.keys(fotoCol).map(Number);
  if (fc.length) console.log('Colunas de foto (' + fc.map(letra_).join(', ') + '): se aparecer foto da Bemol nas linhas ' + linha1 + ' a ' + (linha1 + Math.min(N, M) - 1) + ', selecione e aperte Delete.');
  console.log(N + ' lojas ' + canal + ' escritas nas linhas ' + linha1 + ' a ' + (linha1 + N - 1) + '.');

  // 6) aprovações da Gazin que não foram ao Feishu por falta da aba → mandar de novo
  supa_('PATCH', '/rest/v1/revisoes?feishu_info=like.' + encodeURIComponent('*aba "' + canal + '" não existe*'), { feishu_em: null, feishu_info: null }, { Prefer: 'return=minimal' });
  console.log('Pronto! Confira a aba "' + canal + '" no Feishu. As aprovações da ' + canal + ' passam a ir para ela no próximo ciclo (até 10 min).');
  console.log('Se alguma foto da Bemol tiver ficado na primeira linha, apague à mão (a cópia pode manter imagens).');
}

// uma linha por loja que já mandou o Forms (código = último número digitado)
function lojasDoForms_(canal) {
  const envios = supa_('GET', '/rest/v1/envios?select=key,form,canal,loja_raw,uf,nome,mesa,data&canal=ilike.*' + encodeURIComponent(canal) + '*&limit=20000') || [];
  const rev = {}; (supa_('GET', '/rest/v1/revisoes?select=key,canal_corrigido,loja_corrigida&loja_corrigida=not.is.null&limit=20000') || []).forEach((r) => { rev[r.key] = r; });
  const por = {};
  let semCod = 0;
  envios.forEach((e) => {
    if (e.form !== FORMS[0].nome) return;
    const r = rev[e.key]; const raw = r && r.loja_corrigida ? r.loja_corrigida : e.loja_raw;
    if (r && r.canal_corrigido && semAcento_(r.canal_corrigido) !== semAcento_(canal)) return;
    const m = String(raw || '').match(/\d+/g); if (!m) { semCod++; return; }
    let cod = m[m.length - 1].replace(/^0+/, '') || '0'; if (cod.length < 3) cod = ('000' + cod).slice(-3);
    const x = por[cod] || (por[cod] = { codigo: cod, ufs: {}, n: 0, ultima: null, nome: '', mesas: {} });
    x.n++; const uf = String(e.uf || '').toUpperCase(); if (/^[A-Z]{2}$/.test(uf)) x.ufs[uf] = (x.ufs[uf] || 0) + 1;
    const d = new Date(e.data); if (!x.ultima || d > x.ultima) { x.ultima = d; x.nome = String(e.nome || '').trim(); }
    const me = String(e.mesa || '').trim(); if (me) x.mesas[me] = 1;
  });
  if (semCod) console.log(semCod + ' envio(s) ' + canal + ' sem número de loja ficaram de fora (corrija o código no site).');
  return Object.keys(por).map((k) => {
    const x = por[k];
    const uf = Object.keys(x.ufs).sort((a, b) => x.ufs[b] - x.ufs[a])[0] || '';
    const nota = 'Lista inicial pelo Forms: ' + x.n + ' envio(s), último em ' + Utilities.formatDate(x.ultima, 'America/Sao_Paulo', 'dd/MM') + ' por ' + x.nome +
      (Object.keys(x.mesas).length ? ' · mesa: ' + Object.keys(x.mesas).join(' / ') : '') + (x.codigo.length > 3 ? ' · ⚠ conferir código' : '');
    return { codigo: x.codigo, uf: uf, nota: nota };
  }).sort((a, b) => a.uf.localeCompare(b.uf) || a.codigo.localeCompare(b.codigo, undefined, { numeric: true }));
}
