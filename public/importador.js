/* Importador do Excel de VM no navegador.
 * Lê o .xlsx (zip) por partes, sem carregar o arquivo inteiro na memória:
 * planilhas por canal, fotos coladas (desenhos), datas EXIF e miniaturas.
 * Reproduz a mesma lógica usada até aqui (extract.py + clean.py + photos.py).
 * Uso: VMIMPORT.importar(file, { hoje: 'AAAA-MM-DD', historico: [], prioridades: [], onProgress(txt, frac) })
 */
window.VMIMPORT = (function () {
  'use strict';
  const CH = ['Americanas', 'Magalu', 'Lider', 'Bemol', 'Martinello', 'Paraiba', 'Fast', 'Claro', 'Vivo', 'Tim'];
  const PH = { Americanas: [30, 31, 32, 33], Magalu: [30, 31, 32, 33], Lider: [30, 31, 32, 33], Bemol: [30, 31, 32, 33], Martinello: [30, 31, 32, 33],
    Paraiba: [31, 32, 33], Fast: [26, 27, 28, 30, 31, 32], Claro: [30, 31, 32, 33], Vivo: [30, 31, 32, 33], Tim: [21] };
  const SLOT = { 30: 'Dentro da loja', 31: 'Mesa (frente)', 32: 'Mesa (trás)', 33: 'Fachada', 21: 'Loja', 26: 'Dentro da loja', 27: 'Mesa', 28: 'Mesa' };
  const SLOTS = { Paraiba: { 31: 'Dentro da loja', 32: 'Mesa', 33: 'Mesa' }, Fast: { 26: 'Dentro da loja', 27: 'Mesa', 28: 'Mesa', 30: 'Mesa (frente)', 31: 'Mesa (trás)', 32: 'Fachada' } };
  const REG = { DF: 'DF/GO', GO: 'DF/GO', CE: 'CE/PI', PI: 'CE/PI', PA: 'PA/AP', AP: 'PA/AP', MT: 'MT/MS', MS: 'MT/MS', AM: 'AM/AC/RR/RO', AC: 'AM/AC/RR/RO', RR: 'AM/AC/RR/RO', RO: 'AM/AC/RR/RO', MA: 'MA/TO', TO: 'MA/TO' };

  // ---------- zip ----------
  async function openZip(file) {
    const tailLen = Math.min(file.size, 66000);
    const tail = new DataView(await file.slice(file.size - tailLen).arrayBuffer());
    let eocd = -1;
    for (let i = tailLen - 22; i >= 0; i--) { if (tail.getUint32(i, true) === 0x06054b50) { eocd = i; break; } }
    if (eocd < 0) throw new Error('O arquivo não parece ser um Excel (.xlsx).');
    let cdSize = tail.getUint32(eocd + 12, true), cdOff = tail.getUint32(eocd + 16, true);
    if (cdOff === 0xffffffff) { // zip64
      const loc = tail.getUint32(eocd - 20, true) === 0x07064b50 ? Number(tail.getBigUint64(eocd - 20 + 8, true)) : -1;
      const z = new DataView(await file.slice(loc, loc + 56).arrayBuffer());
      cdSize = Number(z.getBigUint64(40, true)); cdOff = Number(z.getBigUint64(48, true));
    }
    const cd = new DataView(await file.slice(cdOff, cdOff + cdSize).arrayBuffer());
    const dec = new TextDecoder();
    const entries = new Map();
    for (let p = 0; p + 46 <= cd.byteLength && cd.getUint32(p, true) === 0x02014b50;) {
      const method = cd.getUint16(p + 10, true);
      let csize = cd.getUint32(p + 20, true), usize = cd.getUint32(p + 24, true);
      const nlen = cd.getUint16(p + 28, true), xlen = cd.getUint16(p + 30, true), clen = cd.getUint16(p + 32, true);
      let off = cd.getUint32(p + 42, true);
      const name = dec.decode(new Uint8Array(cd.buffer, p + 46, nlen));
      if (csize === 0xffffffff || usize === 0xffffffff || off === 0xffffffff) { // zip64 extra
        let q = p + 46 + nlen; const end = q + xlen;
        while (q + 4 <= end) {
          const id = cd.getUint16(q, true), sz = cd.getUint16(q + 2, true);
          if (id === 1) { let r = q + 4; if (usize === 0xffffffff) { usize = Number(cd.getBigUint64(r, true)); r += 8; } if (csize === 0xffffffff) { csize = Number(cd.getBigUint64(r, true)); r += 8; } if (off === 0xffffffff) { off = Number(cd.getBigUint64(r, true)); } }
          q += 4 + sz;
        }
      }
      entries.set(name, { method, csize, off });
      p += 46 + nlen + xlen + clen;
    }
    async function bytes(name) {
      const e = entries.get(name); if (!e) return null;
      const h = new DataView(await file.slice(e.off, e.off + 30).arrayBuffer());
      const start = e.off + 30 + h.getUint16(26, true) + h.getUint16(28, true);
      const blob = file.slice(start, start + e.csize);
      if (e.method === 0) return new Uint8Array(await blob.arrayBuffer());
      if (e.method !== 8) throw new Error('Compressão não suportada em ' + name);
      return new Uint8Array(await new Response(blob.stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer());
    }
    async function text(name) { const b = await bytes(name); return b ? new TextDecoder().decode(b) : null; }
    return { entries, bytes, text };
  }

  // ---------- xml helpers ----------
  const parseXml = (t) => new DOMParser().parseFromString(t, 'application/xml');
  const kids = (n, local) => Array.from(n.getElementsByTagNameNS('*', local));
  const resolve = (base, target) => { // base 'xl/worksheets/sheet1.xml', target '../drawings/drawing1.xml'
    if (target.startsWith('/')) return target.slice(1);
    const parts = base.split('/'); parts.pop();
    target.split('/').forEach((s) => { if (s === '..') parts.pop(); else if (s !== '.') parts.push(s); });
    return parts.join('/');
  };
  async function rels(zip, partPath) {
    const parts = partPath.split('/'); const fn = parts.pop();
    const t = await zip.text(parts.join('/') + '/_rels/' + fn + '.rels');
    const m = new Map(); if (!t) return m;
    kids(parseXml(t), 'Relationship').forEach((r) => m.set(r.getAttribute('Id'), { target: resolve(partPath, r.getAttribute('Target')), type: r.getAttribute('Type') || '' }));
    return m;
  }
  const colNum = (letters) => { let n = 0; for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64); return n; };
  const colLetters = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  function shiftFormula(f, dr, dc) { // traduz fórmula compartilhada
    return f.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/g, (all, c$, col, r$, row) => (c$ ? '$' + col : colLetters(colNum(col) + dc)) + (r$ ? '$' + row : String(+row + dr)));
  }

  // ---------- planilha → células ----------
  function readSheet(xml, sst) {
    const doc = parseXml(xml);
    const cells = new Map(); // 'r,c' -> {v (python-like), f, raw}
    let maxRow = 0, maxCol = 0;
    const shared = {};
    kids(doc, 'c').forEach((c) => {
      const ref = c.getAttribute('r'); const m = ref && ref.match(/^([A-Z]+)(\d+)$/); if (!m) return;
      const col = colNum(m[1]), row = +m[2];
      const t = c.getAttribute('t') || 'n';
      const vEl = kids(c, 'v')[0], fEl = kids(c, 'f')[0];
      let v = null;
      const vt = vEl ? vEl.textContent : null;
      if (t === 's') v = vt != null ? sst[+vt] : null;
      else if (t === 'inlineStr') v = kids(c, 't').map((x) => x.textContent).join('');
      else if (t === 'str' || t === 'e') v = vt;
      else if (t === 'b') v = vt == null ? null : vt === '1';
      else if (vt != null && vt !== '') v = { num: Number(vt), txt: vt };
      let f = null;
      if (fEl) {
        const si = fEl.getAttribute('si');
        if (fEl.textContent) { f = fEl.textContent; if (si != null) shared[si] = { f, row, col }; }
        else if (si != null && shared[si]) f = shiftFormula(shared[si].f, row - shared[si].row, col - shared[si].col);
        else f = '';
      }
      cells.set(row + ',' + col, { v, f });
      if (row > maxRow) maxRow = row; if (col > maxCol) maxCol = col;
    });
    return { cells, maxRow, maxCol };
  }
  // valor como o Python veria (data_only) e str() dele
  const val = (S, r, c) => { const x = S.cells.get(r + ',' + c); return x ? x.v : null; };
  const isEmptyPy = (v) => v === null || v === undefined || v === '' || v === false || (typeof v === 'object' && v.num === 0);
  function pyStr(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'boolean') return v ? 'True' : 'False';
    if (typeof v === 'object') { const isFloat = /[.eE]/.test(v.txt); return isFloat ? (Number.isInteger(v.num) ? v.num + '.0' : String(v.num)) : String(v.num); }
    return String(v);
  }
  const orStr = (v) => (isEmptyPy(v) ? '' : pyStr(v));
  const title = (s) => s.toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (a, b, c) => b + c.toUpperCase());

  // ---------- EXIF (data da foto) ----------
  function exifDate(u8) {
    try {
      if (u8[0] !== 0xff || u8[1] !== 0xd8) return '';
      let p = 2;
      while (p + 4 < u8.length) {
        if (u8[p] !== 0xff) return '';
        const mk = u8[p + 1], len = (u8[p + 2] << 8) | u8[p + 3];
        if (mk === 0xe1 && u8[p + 4] === 0x45 && u8[p + 5] === 0x78 && u8[p + 6] === 0x69 && u8[p + 7] === 0x66) {
          const t0 = p + 10, dv = new DataView(u8.buffer, u8.byteOffset);
          const le = dv.getUint16(t0) === 0x4949;
          const u16 = (o) => dv.getUint16(t0 + o, le), u32 = (o) => dv.getUint32(t0 + o, le);
          const str = (o, n) => { let s = ''; for (let i = 0; i < n; i++) { const ch = u8[t0 + o + i]; if (!ch) break; s += String.fromCharCode(ch); } return s; };
          const ifd = (o) => { const n = u16(o), out = {}; for (let i = 0; i < n; i++) { const e = o + 2 + i * 12; out[u16(e)] = { type: u16(e + 2), count: u32(e + 4), voff: u32(e + 8), at: e + 8 }; } return out; };
          const d0 = ifd(u32(4));
          const asStr = (en) => en && en.type === 2 ? str(en.count > 4 ? en.voff : en.at, en.count) : '';
          let d = '';
          if (d0[0x8769]) d = asStr(ifd(d0[0x8769].voff)[0x9003]);
          if (!d) d = asStr(d0[0x0132]);
          return d ? d.slice(0, 10).replace(/:/g, '-') : '';
        }
        if (mk === 0xda) return '';
        p += 2 + len;
      }
    } catch (e) { /* sem EXIF */ }
    return '';
  }
  async function thumb(u8, mime) {
    const bmp = await createImageBitmap(new Blob([u8], { type: mime }), { imageOrientation: 'from-image' });
    const k = Math.min(1, 300 / Math.max(bmp.width, bmp.height));
    const w = Math.max(1, Math.round(bmp.width * k)), h = Math.max(1, Math.round(bmp.height * k));
    const cv = document.createElement('canvas'); cv.width = w; cv.height = h;
    const ctx = cv.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, w, h); ctx.drawImage(bmp, 0, 0, w, h); bmp.close && bmp.close();
    return cv.toDataURL('image/jpeg', 0.5);
  }
  async function sha1hex(u8) { const d = new Uint8Array(await crypto.subtle.digest('SHA-1', u8)); return Array.from(d, (b) => b.toString(16).padStart(2, '0')).join(''); }

  // ---------- regras (iguais às do Python) ----------
  const norm = (s) => s.toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^\x00-\x7f]/g, '');
  function tags(o) {
    const n = norm(o), t = [];
    if (/qualida|foto|escuro|irregular|nao mostra/.test(n)) t.push('Foto fora do padrão');
    if (n.includes('spec')) t.push('Sem spec card');
    if (n.includes('demo')) t.push('Sem modo demo');
    if (/mostruario|estoque/.test(n)) t.push('Mostruário / estoque');
    if (/mesa|bancada|atuacao/.test(n)) t.push('Sem mesa / sem espaço');
    if (/canva|material|chegou|inventario|transito/.test(n)) t.push('Material / canvas');
    return t;
  }

  async function importar(file, opt) {
    opt = opt || {};
    const prog = opt.onProgress || (() => {});
    prog('Abrindo o arquivo…', 0);
    const zip = await openZip(file);
    const wbPath = 'xl/workbook.xml';
    const wbXml = await zip.text(wbPath); if (!wbXml) throw new Error('Não achei as abas do Excel.');
    const wbRels = await rels(zip, wbPath);
    const sheets = {};
    kids(parseXml(wbXml), 'sheet').forEach((s) => { const rid = s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || s.getAttribute('r:id'); const r = wbRels.get(rid); if (r) sheets[s.getAttribute('name')] = r.target; });
    const faltam = CH.filter((c) => !sheets[c]);
    if (faltam.length) throw new Error('Faltam abas no Excel: ' + faltam.join(', ') + '.');
    const sstXml = await zip.text('xl/sharedStrings.xml');
    const sst = sstXml ? kids(parseXml(sstXml), 'si').map((si) => kids(si, 't').map((t) => t.textContent).join('')) : [];

    const raw = []; const imgsBySheet = {};
    for (let ci = 0; ci < CH.length; ci++) {
      const ch = CH[ci]; prog('Lendo a aba ' + ch + '…', 0.02 + 0.08 * ci / CH.length);
      const path = sheets[ch];
      const S = readSheet(await zip.text(path), sst);
      // cabeçalho
      let hdr = null;
      for (let r = 1; r < 6; r++) { for (let c = 1; c <= S.maxCol; c++) { if (pyStr(val(S, r, c)).includes('Store Name')) { hdr = r; break; } } if (hdr) break; }
      if (hdr == null) hdr = 2;
      let hrows = []; for (let r = 1; r <= hdr; r++) hrows.push(r); if (ch === 'Claro') hrows.push(3);
      if (ch === 'Tim') { hdr = 3; hrows = [1, 2, 3]; }
      const H = new Map();
      hrows.forEach((r) => { for (let c = 1; c <= S.maxCol; c++) { const v = pyStr(val(S, r, c)); if (v) { if (!H.has(c)) H.set(c, []); H.get(c).push(v); } } });
      const cols = [...H.keys()].sort((a, b) => a - b);
      const find = (pred) => { for (const c of cols) if (H.get(c).some(pred)) return c; return null; };
      const C = {
        code: find((v) => v.includes('Store Code')), name: find((v) => v.includes('Store Name')), city: find((v) => v.includes('City')),
        state: find((v) => (v.includes('State') || v === 'STATE') && !v.includes('Currently')), cls: find((v) => v.includes('Classification')),
        promo: find((v) => v.includes('promotor exclusivo')), t16: find((v) => v.includes('1.6M')), t12: find((v) => v.includes('1.2M')), t08: find((v) => v.includes('0.8M')),
        note: find((v) => v.includes('备注anota')), ok: find((v) => v.includes('Ok/Pendente')), date: find((v) => v.includes('veicula')),
      };
      if (!C.name) throw new Error('Não achei a coluna "Store Name" na aba ' + ch + '.');
      const photocols = cols.filter((c) => H.get(c).some((v) => v.includes('照片')));
      const dispcols = cols.filter((c) => H.get(c).some((v) => v.toLowerCase().includes('display') && !v.includes('Store')));
      // fotos coladas (desenhos)
      const imgs = new Map(); const pics = [];
      const sRels = await rels(zip, path);
      for (const [, r] of sRels) {
        if (!/\/drawing$/.test(r.type)) continue;
        const dXml = await zip.text(r.target); if (!dXml) continue;
        const dRels = await rels(zip, r.target);
        const ddoc = parseXml(dXml);
        ['twoCellAnchor', 'oneCellAnchor'].forEach((tag) => kids(ddoc, tag).forEach((a) => {
          const from = kids(a, 'from')[0]; const pic = kids(a, 'pic')[0]; if (!from || !pic) return;
          const blip = kids(pic, 'blip')[0]; if (!blip) return;
          const rid = blip.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed') || blip.getAttribute('r:embed');
          const media = dRels.get(rid); if (!media) return;
          const row = +kids(from, 'row')[0].textContent + 1, col = +kids(from, 'col')[0].textContent + 1;
          if (row > hdr) { if (!imgs.has(row)) imgs.set(row, new Set()); imgs.get(row).add(col); }
          pics.push({ row, col, media: media.target });
        }));
      }
      imgsBySheet[ch] = pics;
      const okval = (r) => {
        if (!C.ok) return '';
        const x = S.cells.get(r + ',' + C.ok);
        if (x && x.f != null && x.f !== '') {
          const refs = [...x.f.split('"OK"')[0].matchAll(/([A-Z]{1,2})(\d+)/g)];
          return refs.some((m) => { const y = S.cells.get(+m[2] + ',' + colNum(m[1])); return y && ((y.f != null && y.f !== '') || (y.v !== null && y.v !== '')); }) ? 'OK' : 'Pendente';
        }
        const v = pyStr(x ? x.v : null).trim();
        return v.toUpperCase() === 'OK' ? 'OK' : v;
      };
      for (let r = hdr + 1; r <= S.maxRow; r++) {
        const g = (k) => (C[k] ? val(S, r, C[k]) : null);
        const name = g('name');
        if (name === null || pyStr(name).trim() === '') continue;
        if (pyStr(name).includes('Store')) continue;
        const pcount = [...(imgs.get(r) || [])].filter((c) => photocols.includes(c)).length;
        let nd = 0;
        dispcols.forEach((c) => { const v = val(S, r, c); const s = typeof v === 'object' && v ? v.num : v; if (!(v === null || s === '' || s === 'NO' || s === '0' || s === 0 || s === '\n')) nd++; });
        const tb = (k) => { const v = g(k); if (v === null || v === '' || v === '\n') return 0; if (typeof v === 'object') return Math.trunc(v.num); const n = parseFloat(v); if (!isNaN(n) && /^\s*[-+]?\d*\.?\d+(e[-+]?\d+)?\s*$/i.test(v)) return Math.trunc(n); return ['YES', 'SIM', 'X', 'TRUE'].includes(String(v).trim().toUpperCase()) ? 1 : 0; };
        const d = g('date');
        let veic = '';
        if (d && typeof d === 'object' && d.num > 20000 && d.num < 80000) { const dt = new Date(Date.UTC(1899, 11, 30) + Math.round(d.num) * 86400000); veic = dt.toISOString().slice(0, 10); }
        else if (!isEmptyPy(d)) veic = pyStr(d);
        raw.push({ canal: ch, codigo: orStr(g('code')), loja: pyStr(name).trim(), cidade: title(orStr(g('city')).trim()), uf: orStr(g('state')).trim().toUpperCase(),
          classe: orStr(g('cls')).trim().toUpperCase().slice(0, 2) || '-', promotor: orStr(g('promo')).trim().toLowerCase(),
          t16: tb('t16'), t12: tb('t12'), t08: tb('t08'), fotos: pcount, obs: orStr(g('note')).trim(), displays: nd, ok_planilha: okval(r), veiculacao: veic, linha: r });
      }
    }

    // limpeza (clean.py)
    const lojas = raw.map((s) => {
      if (s.canal === 'Claro' && s.linha === 3) Object.assign(s, { classe: '-', promotor: '', obs: '', veiculacao: '' });
      if (s.canal === 'Tim' && !s.uf) { s.uf = 'PA'; s.cidade = 'Belém'; }
      let obs = s.obs.replace(/\s+/g, ' ').trim(); if (obs === 'xcxvxv' || obs === 'Esse funciona') obs = '';
      const cl = ['S', 'A', 'B', 'C'].includes(s.classe) ? s.classe : '-';
      const p = norm(s.promotor);
      const prom = (p.includes('exclu') || p.includes('exlcu') || s.promotor.includes('专')) ? 'Exclusivo' : (p.includes('compart') || s.promotor.includes('共')) ? 'Compartilhado' : (p.includes('sem') || s.promotor.includes('无')) ? 'Sem promotor' : '-';
      const tg = tags(obs);
      const okp = s.ok_planilha.toUpperCase() === 'OK';
      const sit = s.fotos === 0 ? (okp ? 'OK' : 'Pendente') : (tg.length ? 'OK c/ apontamento' : 'OK');
      return { id: s.canal + '-' + s.linha, canal: s.canal, codigo: s.codigo, loja: s.loja, cidade: s.cidade, uf: s.uf, regiao: REG[s.uf] || 'Outra', classe: cl, promotor: prom,
        mesa_16: s.t16, mesa_12: s.t12, mesa_08: s.t08, mesas: s.t16 + s.t12 + s.t08, displays: s.displays, fotos: s.fotos, situacao: sit, ok_planilha: s.ok_planilha,
        apontamento: tg.join('; '), observacao: obs, veiculacao: s.veiculacao };
    });
    const byId = new Map(lojas.map((l) => [l.id, l]));

    // fotos (photos.py)
    const todas = [];
    CH.forEach((ch) => imgsBySheet[ch].forEach((p) => { const lid = ch + '-' + p.row; if (PH[ch].includes(p.col) && byId.has(lid)) todas.push({ ch, lid, ...p }); }));
    const fotos = []; const per = new Map();
    for (let i = 0; i < todas.length; i++) {
      const p = todas[i];
      if (i % 5 === 0) prog('Preparando as fotos… ' + i + ' de ' + todas.length, 0.12 + 0.83 * i / todas.length);
      try {
        const u8 = await zip.bytes(p.media); if (!u8) continue;
        const h = (await sha1hex(u8)).slice(0, 12);
        const ext = p.media.split('.').pop().toLowerCase();
        const mime = ext === 'png' ? 'image/png' : ext === 'gif' ? 'image/gif' : ext === 'webp' ? 'image/webp' : 'image/jpeg';
        const ex = exifDate(u8);
        const img = await thumb(u8, mime);
        const slot = (SLOTS[p.ch] || SLOT)[p.col] || 'Foto';
        fotos.push({ foto_id: p.lid + '-' + p.col + '-' + h.slice(0, 6), loja_id: p.lid, posicao: slot, coluna: p.col, data_foto: ex, hash: h, img });
        if (!per.has(p.lid)) per.set(p.lid, []); per.get(p.lid).push({ h, ex });
      } catch (e) { console.warn('foto ignorada', p, e); }
    }
    const hoje = opt.hoje || new Date().toISOString().slice(0, 10);
    const dia = (s) => Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10));
    lojas.forEach((l) => {
      const lst = per.get(l.id) || [];
      const ds = lst.map((x) => x.ex).filter(Boolean).sort();
      l.ultima_foto = ds.length ? ds[ds.length - 1] : '';
      l.fotos_hash = lst.map((x) => x.h).sort().join('|');
      l.atualizado_em = hoje;
      l.dias_sem_foto = /^\d{4}-\d{2}-\d{2}$/.test(l.ultima_foto) ? Math.round((dia(hoje) - dia(l.ultima_foto)) / 86400000) : null;
    });
    fotos.sort((a, b) => (a.loja_id < b.loja_id ? -1 : a.loja_id > b.loja_id ? 1 : a.coluna - b.coluna));

    // histórico: uma linha por canal × regional por dia de importação
    const grp = new Map();
    lojas.forEach((l) => { const k = l.canal + '|' + l.regiao; if (!grp.has(k)) grp.set(k, { data: hoje, canal: l.canal, regiao: l.regiao, lojas: 0, com_foto: 0, ok_apont: 0, pendente: 0 });
      const g = grp.get(k); g.lojas++; if (l.situacao !== 'Pendente') g.com_foto++; if (l.situacao === 'OK c/ apontamento') g.ok_apont++; if (l.situacao === 'Pendente') g.pendente++; });
    const historico = (opt.historico || []).filter((h) => h.data !== hoje).concat([...grp.values()]);
    prog('Pronto para enviar.', 0.96);
    return { dados: { atualizado: hoje, lojas, historico, prioridades: opt.prioridades || [] }, fotos };
  }
  return { importar };
})();
