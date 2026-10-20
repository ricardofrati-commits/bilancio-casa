'use strict';
/* =====================================================================
   BILANCIO CASA  -  scontrini -> spese -> Excel "FAMILIARE"
   Fase 1: foto/PDF -> OCR offline -> conferma -> archivio
   Fase 2: Excel nel formato FAMILIARE (un foglio per mese, Ricardo | Silvana)
   ===================================================================== */

// ------------------------------------------------------------ utilità
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const PERSONE = ['Ricardo', 'Silvana'];
const MESI = ['GENNAIO', 'FEBBRAIO', 'MARZO', 'APRILE', 'MAGGIO', 'GIUGNO', 'LUGLIO', 'AGOSTO', 'SETTEMBRE', 'OTTOBRE', 'NOVEMBRE', 'DICEMBRE'];
const MESI_IT = MESI.map(m => m[0] + m.slice(1).toLowerCase());
const MESI_BREVI = ['gen', 'feb', 'mar', 'apr', 'mag', 'giu', 'lug', 'ago', 'set', 'ott', 'nov', 'dic'];
const TIPI = { IMPORTO: 'Ordinaria', EXTRA: 'Extra', FISSO: 'Fissa' };
const LS_KEY = 'bilancio-casa.v1';
const EURF = new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR' });
const eur = n => EURF.format(n || 0);
const pad = n => String(n).padStart(2, '0');
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const todayISO = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
const nowYM = () => todayISO().slice(0, 7);
const norm = s => String(s || '').trim().toUpperCase().replace(/\s+/g, ' ');
const round2 = n => Math.round((n + Number.EPSILON) * 100) / 100;
const absURL = p => new URL(p, document.baseURI).href;
const ymOf = d => d.slice(0, 7);
function parseNum(s) {                       // "1.234,56" | "23,5" | "23.50" -> numero
  s = String(s ?? '').replace(/[€\s]/g, '');
  if (!s) return NaN;
  if (s.includes(',') && s.includes('.')) s = s.replace(/\./g, '').replace(',', '.');
  else s = s.replace(',', '.');
  return parseFloat(s);
}
function toast(msg) {
  const t = $('#toast'); t.textContent = msg; t.classList.add('on');
  clearTimeout(toast.h); toast.h = setTimeout(() => t.classList.remove('on'), 2600);
}

// ------------------------------------------------------------ dati
const DEFAULT_VOCI = ['SPESA', 'FARMACIA', 'GASOLIO', 'BAR', 'RISTORANTE', 'PIZZA', 'GELATO', 'AMAZON', 'ONLINE', 'SHOPPING', 'ALTRO'];
const DEFAULT_TIPO = { SPESA: 'IMPORTO', FARMACIA: 'IMPORTO', GASOLIO: 'IMPORTO', BAR: 'IMPORTO', PIZZA: 'IMPORTO', GELATO: 'IMPORTO', RISTORANTE: 'EXTRA', AMAZON: 'EXTRA', ONLINE: 'EXTRA', SHOPPING: 'EXTRA', ALTRO: 'EXTRA' };
// Voci fisse ricavate dal file FAMILIARE 2026 (modificabili dalla schermata Excel)
const seedRecurring = () => [
  { id: 'seed-1', u: 0, persona: 'Ricardo', voce: 'ASS. METLIFE', importo: 17.31, tipo: 'FISSO', giorno: 15, da: '2026-08', a: '' },
  { id: 'seed-2', u: 0, persona: 'Ricardo', voce: 'NETFLIX', importo: 13.99, tipo: 'FISSO', giorno: 15, da: '2026-08', a: '' },
  { id: 'seed-3', u: 0, persona: 'Ricardo', voce: 'ASS. ALLEANZA', importo: 31, tipo: 'FISSO', giorno: 20, da: '2026-08', a: '' },
  { id: 'seed-4', u: 0, persona: 'Ricardo', voce: 'NIMS - CAFFE', importo: 40.83, tipo: 'EXTRA', giorno: 1, da: '2026-08', a: '' },
  { id: 'seed-5', u: 0, persona: 'Silvana', voce: 'ASS. ALLEANZA', importo: 43, tipo: 'FISSO', giorno: 20, da: '2026-08', a: '' },
  { id: 'seed-6', u: 0, persona: 'Silvana', voce: 'TELEFONO', importo: 7, tipo: 'FISSO', giorno: 0, da: '2026-08', a: '' },
];
const defaults = () => ({ v: 1, expenses: [], del: {}, recurring: seedRecurring(), voci: [...DEFAULT_VOCI], tipoPerVoce: { ...DEFAULT_TIPO }, settings: { persona: 'Ricardo', periodo: '' } });
let S = (() => {
  try { const j = JSON.parse(localStorage.getItem(LS_KEY)); if (j && Array.isArray(j.expenses)) return Object.assign(defaults(), j); } catch (e) { /* nessun dato salvato */ }
  return defaults();
})();
function save(noSync) {
  try { localStorage.setItem(LS_KEY, JSON.stringify(S)); }
  catch (e) { toast('Memoria piena: esporta una copia di sicurezza'); }
  if (!noSync && typeof schedSync === 'function') schedSync();
}
const tipoDi = voce => S.tipoPerVoce[norm(voce)] || 'EXTRA';
const tutteLeVoci = () => [...new Set([...S.voci, ...Object.keys(S.tipoPerVoce)])].sort();

// ------------------------------------------------------------ lettura scontrino (testo -> importo / data / voce)
const KW_TOTALE = ['totale complessivo', 'importo pagato', 'totale euro', 'totale eur', 'da pagare', 'totale', 'pagamento elettronico', 'pagato', 'importo'];
const KW_ESCLUDI = ['subtotale', 'sub totale', 'resto', 'iva', 'imponibile', 'imposta', 'sconto', 'contanti', 'non riscosso', 'buoni', 'punti'];
const AMOUNT_RE = /(?:^|[^\d,.])(\d{1,3}(?:\.\d{3})*,\d{2}|\d+[.,]\d{2})(?!\d)/g;
const DATE_RE = /(?:^|\D)(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4}|\d{2})(?!\d)/g;
const DATE_ISO_RE = /(?:^|\D)(\d{4})-(\d{2})-(\d{2})(?!\d)/g;
const amountsIn = line => [...line.matchAll(AMOUNT_RE)].map(m => parseNum(m[1]));
function findAmount(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  const cands = [];
  lines.forEach((line, i) => {
    const low = line.toLowerCase();
    if (KW_ESCLUDI.some(k => low.includes(k))) return;
    for (let p = 0; p < KW_TOTALE.length; p++) {
      if (low.includes(KW_TOTALE[p])) {
        let f = amountsIn(line);
        if (!f.length && lines[i + 1]) f = amountsIn(lines[i + 1]);
        if (f.length) cands.push([p, f[f.length - 1]]);
        break;
      }
    }
  });
  if (cands.length) {
    const best = Math.min(...cands.map(c => c[0]));
    return Math.max(...cands.filter(c => c[0] === best).map(c => c[1]));
  }
  const all = amountsIn(text).filter(a => a > 0 && a < 10000);
  return all.length ? Math.max(...all) : null;
}
function findDate(text) {
  const today = new Date(); const found = [];
  for (const m of text.matchAll(DATE_ISO_RE)) found.push([m.index, +m[3], +m[2], +m[1]]);
  for (const m of text.matchAll(DATE_RE)) { let y = +m[3]; if (y < 100) y += 2000; found.push([m.index, +m[1], +m[2], y]); }
  found.sort((a, b) => a[0] - b[0]);
  for (const [, d, mo, y] of found) {
    const dt = new Date(y, mo - 1, d);
    if (dt.getFullYear() !== y || dt.getMonth() !== mo - 1 || dt.getDate() !== d) continue;
    if (y >= 2000 && dt <= new Date(today.getTime() + 86400000)) return `${y}-${pad(mo)}-${pad(d)}`;
  }
  return null;
}
const KW_VOCI = [
  ['FARMACIA', ['farmacia', 'parafarmacia']],
  ['GASOLIO', ['carburante', 'gasolio', 'benzina', 'self service', 'distributore', ' eni ', ' q8', 'tamoil', ' erg ']],
  ['RISTORANTE', ['ristorante', 'trattoria', 'osteria', 'coperto', 'sushi', 'braceria']],
  ['PIZZA', ['pizzeria', 'pizza']],
  ['GELATO', ['gelateria', 'gelato']],
  ['BAR', [' bar ', 'caffe', 'caffè', 'caffetteria', 'pasticceria', 'tavola calda', 'autogrill']],
  ['AMAZON', ['amazon']],
  ['SPESA', ['supermerc', 'coop', 'conad', 'esselunga', 'carrefour', 'lidl', 'eurospin', 'despar', 'famila', 'iper ', 'aldi', 'alimentari', 'macelleria', 'panificio', 'ortofrutta', ' pam ', ' md ', 'penny', 'todis', 'simply', 'sigma', 'unes']],
];
function suggestVoce(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  const head = ' ' + lines.slice(0, 8).join(' \n ').toLowerCase() + ' ';
  const all = ' ' + text.toLowerCase() + ' ';
  for (const hay of [head, all]) for (const [voce, kws] of KW_VOCI) if (kws.some(k => hay.includes(k))) return voce;
  return '';
}
const merchantOf = text => (text.split('\n').map(l => l.trim()).find(l => /[A-Za-zÀ-ú]{3,}/.test(l)) || '').slice(0, 40);

// ------------------------------------------------------------ OCR (Tesseract.js, tutto sul telefono)
let ocrWorker = null, ocrProgress = null;
async function getWorker() {
  if (!ocrWorker) {
    ocrWorker = await Tesseract.createWorker('ita', 1, {
      workerPath: absURL('vendor/tesseract/worker.min.js'),
      corePath: absURL('vendor/tesseract/core'),
      langPath: absURL('vendor/tesseract/lang'),
      gzip: true,
      logger: m => { if (m.status === 'recognizing text' && ocrProgress) ocrProgress(m.progress); }
    });
    await ocrWorker.setParameters({ tessedit_pageseg_mode: '6', preserve_interword_spaces: '1' });
  }
  return ocrWorker;
}
async function loadBitmap(file) {
  try { return await createImageBitmap(file); }
  catch (e) {
    return await new Promise((res, rej) => { const im = new Image(); im.onload = () => res(im); im.onerror = rej; im.src = URL.createObjectURL(file); });
  }
}
function prepCanvas(src, w0, h0) {
  const long = Math.max(w0, h0);
  const k = long > 2400 ? 2400 / long : long < 1500 ? Math.min(1500 / long, 2.5) : 1;
  const w = Math.round(w0 * k), h = Math.round(h0 * k);
  const c = document.createElement('canvas'); c.width = w; c.height = h;
  const ctx = c.getContext('2d', { willReadFrequently: true });
  ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, w, h); ctx.drawImage(src, 0, 0, w, h);
  const id = ctx.getImageData(0, 0, w, h), px = id.data, hist = new Uint32Array(256);
  for (let i = 0; i < px.length; i += 4) { const g = (px[i] * 0.299 + px[i + 1] * 0.587 + px[i + 2] * 0.114) | 0; px[i] = g; hist[g]++; }
  const tot = w * h; let acc = 0, lo = 0, hi = 255;
  for (let v = 0; v < 256; v++) { acc += hist[v]; if (acc > tot * 0.02) { lo = v; break; } }
  acc = 0; for (let v = 255; v >= 0; v--) { acc += hist[v]; if (acc > tot * 0.02) { hi = v; break; } }
  const span = Math.max(hi - lo, 40);
  for (let i = 0; i < px.length; i += 4) { const g = Math.max(0, Math.min(255, ((px[i] - lo) * 255) / span)); px[i] = px[i + 1] = px[i + 2] = g; px[i + 3] = 255; }
  ctx.putImageData(id, 0, 0);
  return c;
}
async function ocrCanvas(canvas) { const w = await getWorker(); const { data } = await w.recognize(canvas); return data.text || ''; }
function pdfLines(items) {
  const rows = new Map();
  items.forEach(it => { const y = Math.round(it.transform[5] / 3); (rows.get(y) || rows.set(y, []).get(y)).push(it); });
  return [...rows.entries()].sort((a, b) => b[0] - a[0]).map(([, r]) => r.sort((a, b) => a.transform[4] - b.transform[4]).map(i => i.str).join(' ')).join('\n');
}
async function readFileText(file) {
  if (/pdf$/i.test(file.type) || /\.pdf$/i.test(file.name)) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = absURL('vendor/pdf.worker.min.js');
    const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
    const n = Math.min(pdf.numPages, 3); let text = '';
    for (let p = 1; p <= n; p++) { const pg = await pdf.getPage(p); text += pdfLines((await pg.getTextContent()).items) + '\n'; }
    if (text.replace(/\s/g, '').length >= 30) return text;
    text = '';                                           // PDF scansionato: serve l'OCR
    for (let p = 1; p <= n; p++) {
      const pg = await pdf.getPage(p), vp = pg.getViewport({ scale: 2.5 });
      const c = document.createElement('canvas'); c.width = vp.width; c.height = vp.height;
      await pg.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
      text += await ocrCanvas(prepCanvas(c, c.width, c.height)) + '\n';
    }
    return text;
  }
  const bmp = await loadBitmap(file);
  return ocrCanvas(prepCanvas(bmp, bmp.width || bmp.naturalWidth, bmp.height || bmp.naturalHeight));
}
async function sha1(file) {
  try { const b = await crypto.subtle.digest('SHA-1', await file.arrayBuffer()); return [...new Uint8Array(b)].map(x => x.toString(16).padStart(2, '0')).join(''); }
  catch (e) { return null; }
}

// ------------------------------------------------------------ pannello "scontrino"
let queue = [], form = null, thumbURL = null;
const sheet = $('#sheet');
function newForm(over = {}) {
  return Object.assign({ id: null, importo: '', d: todayISO(), persona: S.settings.persona || 'Ricardo', voce: '', tipo: 'IMPORTO', nota: '', dirty: {}, src: null, hash: null, dup: false, status: '', ocr: null, raw: '' }, over);
}
function openForm(f) { form = f; renderSheet(); sheet.hidden = false; }
function closeForm() {
  sheet.hidden = true; sheet.innerHTML = ''; form = null;
  if (thumbURL) { URL.revokeObjectURL(thumbURL); thumbURL = null; }
}
function renderSheet() {
  if (!form) return;
  const f = form, isEdit = !!f.id;
  const voci = S.voci, inChips = voci.includes(f.voce);
  const prevScroll = $('.panel', sheet)?.scrollTop || 0;
  let ocrHTML = '';
  if (f.src) {
    const thumb = f.src.isPdf ? '<div class="pdf">PDF</div>' : `<img alt="Anteprima scontrino" src="${esc(f.src.url)}">`;
    const warn = f.status === 'fail' || f.dup;
    const msg = f.status === 'run' ? 'Leggo lo scontrino…' : f.status === 'ok' ? 'Letto. Controlla importo e data.' : f.status === 'fail' ? 'Lettura non riuscita: compila a mano.' : '';
    ocrHTML = `<div class="ocr ${warn ? 'warn' : ''}">${thumb}<div class="tx"><strong id="ocrMsg">${msg}</strong>${f.dup ? '<div>Questo file risulta già salvato.</div>' : ''}
      ${queue.length ? `<div class="hint">Altri ${queue.length} in coda</div>` : ''}<div class="pb"><i id="ocrBar" style="width:${f.status === 'run' ? '4%' : '100%'}"></i></div></div></div>`;
  }
  sheet.innerHTML = `<div class="scrim" data-act="close"></div>
  <section class="panel" role="dialog" aria-modal="true" aria-label="${isEdit ? 'Modifica spesa' : 'Nuova spesa'}">
    <div class="grab"></div>${ocrHTML}
    <div class="stub">
      <div class="amount"><span>€</span><input id="fImp" inputmode="decimal" autocomplete="off" placeholder="0,00" aria-label="Importo in euro" value="${esc(f.importo)}"></div>
      <label class="row"><span>Data</span><input type="date" id="fData" value="${esc(f.d)}"></label>
      <div class="row"><span>Chi</span><div class="seg">${PERSONE.map(p => `<button type="button" data-act="persona" data-p="${p}" aria-pressed="${f.persona === p}">${p}</button>`).join('')}</div></div>
      <div class="row col"><span>Voce</span>
        <div class="chips">${voci.map(v => `<button type="button" class="chip" data-act="voce" data-v="${esc(v)}" aria-pressed="${f.voce === v}">${esc(v)}</button>`).join('')}</div>
        <input id="fVoce" list="dlVoci" placeholder="Altra voce (es. AMAZON, DECATHLON)" value="${inChips ? '' : esc(f.voce)}" autocomplete="off"></div>
      <div class="row"><span>Tipo</span><div class="seg">${Object.entries(TIPI).map(([k, t]) => `<button type="button" data-act="tipo" data-t="${k}" aria-pressed="${f.tipo === k}">${t}</button>`).join('')}</div></div>
      <label class="row"><span>Note</span><input type="text" id="fNota" placeholder="facoltativo" value="${esc(f.nota)}"></label>
    </div>
    <div class="actions">
      ${isEdit ? '<button class="btn danger" data-act="elimina">Elimina</button>' : `<button class="btn" data-act="${f.src ? 'salta' : 'close'}">${f.src ? 'Salta' : 'Annulla'}</button>`}
      <button class="btn primary" data-act="salva">${isEdit ? 'Salva modifiche' : queue.length ? 'Salva e prossimo' : 'Salva'}</button>
    </div>
    ${f.raw ? `<details class="raw"><summary>Testo letto dallo scontrino</summary><pre>${esc(f.raw)}</pre></details>` : ''}
  </section>`;
  $('.panel', sheet).scrollTop = prevScroll;
}
function setOcrProgress(p) { const b = $('#ocrBar'); if (b) b.style.width = Math.max(4, Math.round(p * 100)) + '%'; }

async function processNext() {
  if (!queue.length) { closeForm(); return; }
  const file = queue.shift();
  if (thumbURL) URL.revokeObjectURL(thumbURL);
  const isPdf = /pdf$/i.test(file.type) || /\.pdf$/i.test(file.name);
  thumbURL = isPdf ? null : URL.createObjectURL(file);
  const f = newForm({ src: { name: file.name, isPdf, url: thumbURL }, status: 'run' });
  openForm(f);
  f.hash = await sha1(file);
  f.dup = !!(f.hash && S.expenses.some(e => e.hash === f.hash));
  renderSheet();
  ocrProgress = setOcrProgress;
  try {
    const text = await readFileText(file);
    if (form !== f) return;                              // l'utente ha chiuso nel frattempo
    f.raw = text;
    const amt = findAmount(text), dt = findDate(text), v = suggestVoce(text);
    if (amt != null && !f.dirty.importo) f.importo = String(amt.toFixed(2)).replace('.', ',');
    if (dt && !f.dirty.d) f.d = dt;
    if (v && !f.dirty.voce) { f.voce = v; if (!f.dirty.tipo) f.tipo = tipoDi(v); }
    if (!f.nota) f.nota = merchantOf(text);
    f.status = (amt != null || dt) ? 'ok' : 'fail';
  } catch (e) {
    console.error(e); f.status = 'fail';
  }
  if (form === f) renderSheet();
}
async function addFiles(files) {
  const ok = [...files].filter(f => /^image\/(jpeg|png)/.test(f.type) || /pdf$/i.test(f.type) || /\.(jpe?g|png|pdf)$/i.test(f.name));
  if (!ok.length) { toast('Scegli una foto JPG/PNG oppure un PDF'); return; }
  queue.push(...ok);
  if (sheet.hidden) processNext(); else renderSheet();
}
function salvaForm() {
  const f = form;
  const importo = round2(parseNum(f.importo)), d = f.d;
  if (!(importo > 0)) { toast('Inserisci un importo valido (es. 23,50)'); $('#fImp')?.focus(); return; }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) { toast('Inserisci la data'); return; }
  const voce = norm(f.voce);
  if (!voce) { toast('Scegli una voce (SPESA, BAR, FARMACIA…)'); return; }
  if (f.dup && !f.id && !confirm('Questo scontrino è già stato salvato. Salvarlo di nuovo?')) return;
  const rec = { id: f.id || uid(), d, persona: f.persona, voce, tipo: f.tipo, importo, nota: f.nota.trim(), hash: f.hash || undefined, nd: false, u: Date.now() };
  if (f.id) { const i = S.expenses.findIndex(e => e.id === f.id); if (i >= 0) { rec.hash = S.expenses[i].hash; S.expenses[i] = rec; } }
  else S.expenses.push(rec);
  S.tipoPerVoce[voce] = f.tipo; S.settings.persona = f.persona;
  if (!S.voci.includes(voce) && !(voce in DEFAULT_TIPO)) { /* voce nuova: resta nell'elenco suggerimenti */ }
  save(); toast(f.id ? 'Modifica salvata' : 'Spesa salvata');
  ui.mese = ymOf(d);
  if (f.src && queue.length) processNext(); else { closeForm(); render(); }
}

// ------------------------------------------------------------ schermate
const ui = { tab: 'scan', mese: '', persona: '', mesePick: '', sel: null };
const sumBy = (arr, fn) => round2(arr.reduce((a, e) => a + fn(e), 0));
const speseDelMese = ym => S.expenses.filter(e => ymOf(e.d) === ym);
function mesiConDati() { return [...new Set(S.expenses.map(e => ymOf(e.d)))].sort().reverse(); }
function meseCorrente() { const l = mesiConDati(); return ui.mese && /^\d{4}-\d{2}$/.test(ui.mese) ? ui.mese : (l.includes(nowYM()) ? nowYM() : l[0] || nowYM()); }
const labelMese = ym => `${MESI_IT[+ym.slice(5) - 1]} ${ym.slice(0, 4)}`;
const iconCam = '<svg viewBox="0 0 24 24"><path d="M4 8h3l2-3h6l2 3h3v11H4zM12 17a4 4 0 1 0 0-8 4 4 0 0 0 0 8z"/></svg>';
const iconFile = '<svg viewBox="0 0 24 24"><path d="M5 3h9l5 5v13H5zM14 3v5h5"/></svg>';

function viewScan() {
  const ultime = [...S.expenses].sort((a, b) => (b.d + b.id > a.d + a.id ? 1 : -1)).slice(0, 5);
  return `<section class="card hero">
      <h2>Aggiungi uno scontrino</h2>
      <p>Fotografa lo scontrino o scegli un file: leggo importo e data, tu scegli la voce.</p>
      <div class="stack">
        <button class="btn primary big" data-act="camera">${iconCam} Scatta una foto</button>
        <button class="btn big" data-act="files">${iconFile} Scegli foto o PDF</button>
        <button class="link" data-act="manuale">Inserisci a mano, senza scontrino</button>
      </div>
    </section>
    <h3 class="month-h"><span>Ultime spese</span></h3>
    ${ultime.length ? `<div>${ultime.map(itemHTML).join('')}</div>` : '<div class="card empty">Ancora nessuna spesa. Scatta la prima foto qui sopra.</div>'}`;
}
function itemHTML(e) {
  const dd = `${+e.d.slice(8)} ${MESI_BREVI[+e.d.slice(5, 7) - 1]}`;
  return `<button class="item" data-act="edit" data-id="${e.id}"><span class="dot" style="background:var(--${e.persona.toLowerCase()})" title="${e.persona}"></span>
    <span class="t">${esc(e.voce)}${e.tipo !== 'IMPORTO' ? `<span class="tag">${TIPI[e.tipo]}</span>` : ''}</span>
    <span class="s">${e.nd ? labelMese(ymOf(e.d)) : dd} · ${e.persona}${e.nota ? ' · ' + esc(e.nota) : ''}</span><span class="a">${eur(e.importo)}</span></button>`;
}
function viewList() {
  const mesi = mesiConDati();
  const sel = ui.mesePick;
  let arr = S.expenses.filter(e => (!sel || ymOf(e.d) === sel) && (!ui.persona || e.persona === ui.persona));
  arr.sort((a, b) => (b.d + b.id > a.d + a.id ? 1 : -1));
  const groups = new Map(); arr.forEach(e => { const k = ymOf(e.d); (groups.get(k) || groups.set(k, []).get(k)).push(e); });
  return `<div class="filters">
      <select class="sel" id="selMese" aria-label="Mese" style="flex:1"><option value="">Tutti i mesi</option>${mesi.map(m => `<option value="${m}" ${m === sel ? 'selected' : ''}>${labelMese(m)}</option>`).join('')}</select>
    </div>
    <div class="filters">${['', ...PERSONE].map(p => `<button class="chip" data-act="fPersona" data-p="${p}" aria-pressed="${ui.persona === p}">${p || 'Tutti'}</button>`).join('')}
      <span style="flex:1"></span><button class="btn small" data-act="manuale">+ Aggiungi</button></div>
    ${arr.length ? [...groups.entries()].map(([k, list]) => `<div class="month-h"><span>${labelMese(k)}</span><span>${eur(sumBy(list, e => e.importo))}</span></div><div>${list.map(itemHTML).join('')}</div>`).join('') : '<div class="card empty">Nessuna spesa con questi filtri.</div>'}`;
}

// --- grafici
function seriesMensili(endYM, n = 12) {
  const out = []; let y = +endYM.slice(0, 4), m = +endYM.slice(5);
  for (let i = 0; i < n; i++) { out.unshift(`${y}-${pad(m)}`); m--; if (m === 0) { m = 12; y--; } }
  return out.map(ym => { const l = speseDelMese(ym); return { ym, R: sumBy(l.filter(e => e.persona === 'Ricardo'), e => e.importo), S: sumBy(l.filter(e => e.persona === 'Silvana'), e => e.importo) }; });
}
const niceMax = v => { if (v <= 0) return 100; const p = Math.pow(10, Math.floor(Math.log10(v))); const f = v / p; return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 5 ? 5 : 10) * p; };
function stackedSVG(serie, selYM) {
  const W = 340, H = 190, L = 34, R = 6, T = 8, B = 22, iw = W - L - R, ih = H - T - B;
  const max = niceMax(Math.max(...serie.map(s => s.R + s.S), 1)), bw = iw / serie.length;
  const y = v => T + ih - (v / max) * ih, fmt = v => v >= 1000 ? (v / 1000).toString().replace('.', ',') + 'k' : v;
  let g = [0, .5, 1].map(f => `<line class="grid" x1="${L}" x2="${W - R}" y1="${y(max * f)}" y2="${y(max * f)}"/><text x="${L - 5}" y="${y(max * f) + 3}" text-anchor="end">${fmt(max * f)}</text>`).join('');
  serie.forEach((s, i) => {
    const x = L + i * bw + bw * .16, w = bw * .68, hR = (s.R / max) * ih, hS = (s.S / max) * ih;
    const on = s.ym === selYM;
    if (s.R > 0) g += `<rect x="${x}" y="${y(s.R)}" width="${w}" height="${Math.max(hR - 1, 0)}" rx="3" fill="var(--ricardo)" ${on ? '' : 'opacity=".9"'}/>`;
    if (s.S > 0) g += `<rect x="${x}" y="${y(s.R + s.S)}" width="${w}" height="${Math.max(hS - 1, 0)}" rx="3" fill="var(--silvana)"/>`;
    g += `<text x="${x + w / 2}" y="${H - 7}" text-anchor="middle" ${on ? 'style="fill:var(--ink);font-weight:700"' : ''}>${MESI_BREVI[+s.ym.slice(5) - 1][0].toUpperCase()}</text>`;
    g += `<rect data-act="barra" data-ym="${s.ym}" x="${L + i * bw}" y="${T}" width="${bw}" height="${ih + B}" fill="transparent" style="cursor:pointer"/>`;
  });
  return `<svg class="chart" viewBox="0 0 ${W} ${H}" role="img" aria-label="Spese degli ultimi 12 mesi, per persona"><line class="axis" x1="${L}" x2="${W - R}" y1="${T + ih}" y2="${T + ih}"/>${g}</svg>`;
}
function readoutHTML(s) { return s ? `<b>${labelMese(s.ym)}</b> · <span class="sw" style="background:var(--ricardo)"></span>Ricardo <b>${eur(s.R)}</b> · <span class="sw" style="background:var(--silvana)"></span>Silvana <b>${eur(s.S)}</b> · totale <b>${eur(s.R + s.S)}</b>` : ''; }
function viewChart() {
  const ym = meseCorrente(), mesi = [...new Set([...mesiConDati(), nowYM()])].sort().reverse();
  const l = speseDelMese(ym), tot = sumBy(l, e => e.importo);
  const R = sumBy(l.filter(e => e.persona === 'Ricardo'), e => e.importo), Sv = sumBy(l.filter(e => e.persona === 'Silvana'), e => e.importo);
  const per = new Map(); l.forEach(e => per.set(e.voce, (per.get(e.voce) || 0) + e.importo));
  let voci = [...per.entries()].sort((a, b) => b[1] - a[1]);
  if (voci.length > 8) { const rest = voci.slice(7).reduce((a, v) => a + v[1], 0); voci = [...voci.slice(0, 7), ['ALTRE', rest]]; }
  const mx = voci.length ? voci[0][1] : 1, serie = seriesMensili(ym), selS = serie.find(s => s.ym === (ui.sel || ym));
  const ex = sumBy(l.filter(e => e.tipo === 'EXTRA'), e => e.importo), fx = sumBy(l.filter(e => e.tipo === 'FISSO'), e => e.importo);
  return `<div class="filters"><select class="sel" id="selMeseG" aria-label="Mese">${mesi.map(m => `<option value="${m}" ${m === ym ? 'selected' : ''}>${labelMese(m)}</option>`).join('')}</select></div>
    <section class="card"><div class="kpi"><div><div class="lbl">Spesa del mese</div><div class="big">${eur(tot)}</div></div></div>
      <div class="pp"><span><span class="sw" style="background:var(--ricardo)"></span>Ricardo <b>${eur(R)}</b></span><span><span class="sw" style="background:var(--silvana)"></span>Silvana <b>${eur(Sv)}</b></span></div>
      <div class="pp hint"><span>Fisse <b>${eur(fx)}</b></span><span>Extra <b>${eur(ex)}</b></span><span>Ordinarie <b>${eur(round2(tot - ex - fx))}</b></span></div></section>
    <section class="card"><h2>Dove vanno i soldi</h2>
      ${voci.length ? `<div class="bars">${voci.map(([n, v]) => `<div class="bar"><span class="n" title="${esc(n)}">${esc(n)}</span><span class="tr"><i style="width:${Math.max(3, v / mx * 100)}%"></i></span><span class="v">${eur(v)}</span></div>`).join('')}</div>` : '<p>Nessuna spesa in questo mese.</p>'}</section>
    <section class="card"><h2>Ultimi 12 mesi</h2>
      <div class="legend"><span><span class="sw" style="background:var(--ricardo)"></span>Ricardo</span><span><span class="sw" style="background:var(--silvana)"></span>Silvana</span></div>
      ${stackedSVG(serie, selS?.ym)}<div class="readout" id="readout">${readoutHTML(selS)}</div></section>`;
}

// --- Excel / impostazioni
function periodOptions() {
  const ys = S.expenses.map(e => +e.d.slice(0, 4)), cy = new Date().getFullYear();
  const min = Math.min(cy, ...ys), max = Math.max(cy + 9, ...ys), o = [];
  for (let y = min; y <= max; y++) o.push({ v: `1-${y}`, t: `Gennaio – Dicembre ${y}` });
  return o;
}
const periodoCorrente = () => /^1-\d{4}$/.test(S.settings.periodo || '') ? S.settings.periodo : `1-${new Date().getFullYear()}`;
function viewExcel() {
  const po = periodOptions(), pc = periodoCorrente();
  return `<section class="card"><h2>Excel FAMILIARE</h2>
      <p>Un foglio per mese, Ricardo a sinistra e Silvana a destra, come il tuo file. Le spese fisse si aggiungono da sole.</p>
      <div class="field"><label for="selPer">Anno</label><select class="sel" id="selPer">${po.map(p => `<option value="${p.v}" ${p.v === pc ? 'selected' : ''}>${p.t}</option>`).join('')}</select></div>
      <div class="stack" style="margin-top:12px">
        <button class="btn primary big" data-act="export" id="btnExp">Crea l'Excel</button>
        <button class="btn big" data-act="importXlsx">Carica un Excel FAMILIARE</button>
        <button class="btn" data-act="share" id="btnShare" hidden>Condividi il file</button></div>
      <p class="hint" style="margin-top:10px">Il file caricato porta qui le sue spese, anche di anni diversi (2025, 2026, 2027…). Si può ripetere senza creare doppioni.</p>
      <div id="impReport"></div></section>
    <section class="card"><h2>Stesso bilancio su telefono e PC</h2>
      ${syncOn() ? `<p>Collegato a <b>${esc(sync.repo)}</b>.</p><p id="syncStat" class="hint">${esc(syncStatusText())}</p><div class="two" style="margin-top:10px"><button class="btn primary" data-act="syncNow">Sincronizza ora</button><button class="btn" data-act="syncOff">Scollega</button></div>` : `<p>Collega un archivio privato su GitHub: spese e voci fisse restano uguali su telefono e computer, e l'Excel esce identico da entrambi.</p><button class="btn primary" style="margin-top:10px" data-act="syncSetup">Collega</button>`}</section>
    <section class="card"><h2>Spese fisse e ricorrenti</h2><p>Compaiono ogni mese nell'Excel, fino al mese in corso, se non hai già inserito la stessa voce.</p>
      <div style="margin-top:6px">${S.recurring.length ? S.recurring.map(r => `<button class="rec" data-act="editRec" data-id="${r.id}"><span><b style="font-weight:650">${esc(r.voce)}</b><small>${r.persona} · ${TIPI[r.tipo]}${r.giorno ? ' · giorno ' + r.giorno : ''} · da ${r.da ? labelMese(r.da) : 'sempre'}${r.a ? ' a ' + labelMese(r.a) : ''}</small></span><b>${eur(r.importo)}</b></button>`).join('') : '<p class="hint">Nessuna voce.</p>'}</div>
      <button class="btn small" style="margin-top:10px" data-act="newRec">+ Aggiungi voce fissa</button></section>
    <section class="card"><h2>Copia di sicurezza</h2><p>I dati stanno solo su questo telefono. Salva una copia ogni tanto; serve anche per portarli su un altro dispositivo.</p>
      <div class="two" style="margin-top:10px"><button class="btn" data-act="backup">Salva copia</button><button class="btn" data-act="restore">Ripristina</button></div>
      <p class="hint" style="margin-top:10px">${S.expenses.length} spese salvate.</p></section>`;
}

let _ver = 0;
function render() {
  $$('#nav button').forEach(b => b.setAttribute('aria-current', b.dataset.tab === ui.tab ? 'page' : 'false'));
  const v = { scan: viewScan, list: viewList, chart: viewChart, excel: viewExcel }[ui.tab]();
  $('#view').innerHTML = v;
  const ym = meseCorrente(), tot = sumBy(speseDelMese(ym), e => e.importo);
  $('#topTot').textContent = S.expenses.length ? `${MESI_IT[+ym.slice(5) - 1]}: ${eur(tot)}` : '';
  $('#dlVoci').innerHTML = tutteLeVoci().map(v => `<option value="${esc(v)}">`).join('');
}

// ------------------------------------------------------------ editor voci fisse
function openRec(r) {
  const x = r || { id: '', persona: 'Ricardo', voce: '', importo: '', tipo: 'FISSO', giorno: '', da: nowYM(), a: '' };
  sheet.hidden = false; form = null;
  sheet.innerHTML = `<div class="scrim" data-act="close"></div><section class="panel" role="dialog" aria-modal="true" aria-label="Voce fissa">
    <div class="grab"></div><div class="card" style="box-shadow:none">
    <h2>${r ? 'Modifica voce fissa' : 'Nuova voce fissa'}</h2>
    <div class="field"><label>Chi</label><select class="sel" id="rPers">${PERSONE.map(p => `<option ${p === x.persona ? 'selected' : ''}>${p}</option>`).join('')}</select></div>
    <div class="field"><label>Voce (come nell'Excel)</label><input class="inp" id="rVoce" value="${esc(x.voce)}" placeholder="es. NETFLIX"></div>
    <div class="field"><label>Importo (€)</label><input class="inp" id="rImp" inputmode="decimal" value="${x.importo === '' ? '' : String(x.importo).replace('.', ',')}"></div>
    <div class="field"><label>Colonna nell'Excel</label><select class="sel" id="rTipo">${Object.entries(TIPI).map(([k, t]) => `<option value="${k}" ${k === x.tipo ? 'selected' : ''}>${t}${k === 'FISSO' ? ' (Importo + Fisso mensile)' : k === 'EXTRA' ? ' (colonna Extra)' : ' (colonna Importo)'}</option>`).join('')}</select></div>
    <div class="field"><label>Giorno del mese (vuoto = senza data)</label><input class="inp" id="rGiorno" inputmode="numeric" value="${x.giorno || ''}"></div>
    <div class="two"><div class="field"><label>Da (mese)</label><input class="inp" type="month" id="rDa" value="${x.da || ''}"></div><div class="field"><label>A (vuoto = sempre)</label><input class="inp" type="month" id="rA" value="${x.a || ''}"></div></div>
    <div class="actions" style="margin-top:16px">${r ? '<button class="btn danger" data-act="delRec">Elimina</button>' : '<button class="btn" data-act="close">Annulla</button>'}<button class="btn primary" data-act="saveRec" data-id="${x.id}">Salva</button></div></div></section>`;
}
function saveRec(id) {
  const imp = round2(parseNum($('#rImp').value)), voce = norm($('#rVoce').value);
  if (!voce || !(imp > 0)) { toast('Servono voce e importo'); return; }
  const rec = { id: id || uid(), persona: $('#rPers').value, voce, importo: imp, tipo: $('#rTipo').value, giorno: Math.min(31, parseInt($('#rGiorno').value) || 0), da: $('#rDa').value, a: $('#rA').value, u: Date.now() };
  const i = S.recurring.findIndex(r => r.id === id);
  if (i >= 0) S.recurring[i] = rec; else S.recurring.push(rec);
  S.tipoPerVoce[voce] = S.tipoPerVoce[voce] || (rec.tipo === 'EXTRA' ? 'EXTRA' : 'IMPORTO');
  save(); closeForm(); render(); toast('Voce fissa salvata');
}

// ------------------------------------------------------------ FASE 2: costruzione dell'Excel nel formato FAMILIARE
const FMT_EUR = '_-* #,##0.00\\ "€"_-;\\-* #,##0.00\\ "€"_-;_-* "-"??\\ "€"_-;_-@_-';
const COL = { verde: 'FFE2EFDA', blu: 'FFBDD7EE', arancio: 'FFFBE5D6', ambra: 'FFFFC000', tot: 'FF00B050', giallo: 'FFFFFF00', rosa: 'FFFF99CC' };
const thin = { style: 'thin' }, BORD = { top: thin, left: thin, bottom: thin, right: thin };

function righeMese(persona, y, m) {
  const ym = `${y}-${pad(m)}`, rows = [];
  const exist = S.expenses.filter(e => e.persona === persona && ymOf(e.d) === ym);
  const labels = new Set(exist.map(e => norm(e.voce)));
  S.recurring.filter(r => r.persona === persona).forEach((r, i) => {
    if (r.da && ym < r.da) return; if (r.a && ym > r.a) return; if (ym > nowYM()) return;
    if (labels.has(norm(r.voce))) return;
    const dim = new Date(y, m, 0).getDate(), dd = r.giorno ? Math.min(r.giorno, dim) : 0;
    rows.push({ rank: 0, date: dd ? Date.UTC(y, m - 1, dd) : null, key: dd ? `${ym}-${pad(dd)}` : '9999', ord: i, voce: r.voce, extra: r.tipo === 'EXTRA' ? r.importo : 0, importo: r.tipo === 'EXTRA' ? 0 : r.importo, fisso: r.tipo === 'FISSO' ? r.importo : 0 });
  });
  exist.forEach((e, i) => {
    const dd = +e.d.slice(8, 10);
    rows.push({ rank: e.tipo === 'FISSO' ? 0 : 1, date: e.nd ? null : Date.UTC(y, m - 1, dd), key: e.nd ? '9999' : e.d, ord: 1000 + i, voce: e.voce, extra: e.tipo === 'EXTRA' ? e.importo : 0, importo: e.tipo === 'EXTRA' ? 0 : e.importo, fisso: e.tipo === 'FISSO' ? e.importo : 0 });
  });
  rows.sort((a, b) => a.rank - b.rank || (a.key < b.key ? -1 : a.key > b.key ? 1 : 0) || a.ord - b.ord);
  return rows;
}
function styleCell(c, o = {}) {
  if (o.border !== false) c.border = BORD;
  if (o.fill) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: o.fill } };
  c.font = { name: 'Calibri', size: 11, bold: !!o.bold };
  c.alignment = { horizontal: o.center === false ? undefined : 'center', vertical: 'middle' };
  if (o.fmt) c.numFmt = o.fmt;
}
const colL = n => { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };

function buildMonthSheet(wb, m, year, rowsP) {
  const ws = wb.addWorksheet(MESI[m - 1]);
  const widths = [8.43, 10.66, 8.43, 17.55, 8.43, 10.89, 8.43, 11, 14.33, 14.55, 11, 10.66, 8.43, 23.11, 8.43, 10.66, 8.43, 11, 14.33];
  widths.forEach((w, i) => ws.getColumn(i + 1).width = w);
  const N = Math.max(35, 3 + Math.max(rowsP[0].length, rowsP[1].length));
  const T = N + 1, TT = N + 7, SM = N + 9;                 // riga totali, riga TOTALE, riga SPESA MENSILE
  const out = { N, T, TT, SM, sums: [] };
  PERSONE.forEach((p, b) => {
    const o = b * 10, c = k => colL(k + o);                 // colonne del blocco: A..I | K..S
    ws.mergeCells(`${c(4)}1:${c(6)}1`);
    const t = ws.getCell(`${c(4)}1`); t.value = `SPESE - ${p.toUpperCase()}`; styleCell(t, { fill: b ? COL.arancio : COL.blu, bold: true });
    ['E', 'F'].forEach((_, k) => styleCell(ws.getCell(`${c(5 + k)}1`), { fill: b ? COL.arancio : COL.blu, bold: true }));
    styleCell(ws.getCell(`${c(1)}3`), { center: false });
    ['DATA', null, 'RAGIONE SOCIALE', null, 'EXTRA', null, 'IMPORTO', 'FISSO MENSILE'].forEach((h, k) => { const cell = ws.getCell(`${c(2 + k)}3`); if (h) cell.value = h; styleCell(cell, { fill: COL.verde, bold: true }); });
    let sF = 0, sH = 0, sI = 0;
    for (let i = 0; i < N - 3; i++) {
      const r = 4 + i, row = rowsP[b][i];
      styleCell(ws.getCell(`${c(1)}${r}`), {}); ws.getCell(`${c(1)}${r}`).value = i + 1;
      const d = ws.getCell(`${c(2)}${r}`); styleCell(d, { fmt: 'mm-dd-yy' }); if (row && row.date != null) d.value = new Date(row.date);
      styleCell(ws.getCell(`${c(3)}${r}`), {});
      const l = ws.getCell(`${c(4)}${r}`); styleCell(l, { fill: COL.verde, bold: true }); if (row) l.value = row.voce;
      styleCell(ws.getCell(`${c(5)}${r}`), {});
      const F = ws.getCell(`${c(6)}${r}`), H = ws.getCell(`${c(8)}${r}`), I = ws.getCell(`${c(9)}${r}`);
      styleCell(F, { fmt: FMT_EUR }); styleCell(ws.getCell(`${c(7)}${r}`), {}); styleCell(H, { fmt: FMT_EUR });
      styleCell(I, { fill: COL.ambra, bold: true, fmt: FMT_EUR, center: false });
      F.value = row ? row.extra : 0; H.value = row ? row.importo : 0; I.value = row ? row.fisso : 0;
      if (row) { sF += row.extra; sH += row.importo; sI += row.fisso; }
    }
    sF = round2(sF); sH = round2(sH); sI = round2(sI);
    const tc = (k, f, res, o2) => { const cell = ws.getCell(`${c(k)}${T}`); cell.value = { formula: f, result: res }; styleCell(cell, o2); };
    tc(6, `SUM(${c(6)}4:${c(6)}${N})`, sF, { bold: true, fmt: FMT_EUR });
    tc(8, `SUM(${c(8)}4:${c(8)}${N})`, sH, { bold: true, fmt: FMT_EUR, center: false });
    tc(9, `SUM($${c(9)}$4:${c(9)}${N})`, sI, { fill: COL.tot, bold: true, fmt: FMT_EUR, center: false });
    const L = ws.getCell(`${c(4)}${TT}`); L.value = 'TOTALE'; styleCell(L, { bold: true, center: false }); styleCell(ws.getCell(`${c(5)}${TT}`), {});
    const f1 = ws.getCell(`${c(6)}${TT}`); f1.value = { formula: `SUM($${c(6)}$${T}:${c(6)}${TT - 1})`, result: sF }; styleCell(f1, { fill: COL.giallo, bold: true, fmt: FMT_EUR, center: false });
    styleCell(ws.getCell(`${c(7)}${TT}`), { bold: true, fmt: FMT_EUR, center: false });
    const h1 = ws.getCell(`${c(8)}${TT}`); h1.value = { formula: `SUM($${c(6)}$${T}:${c(8)}${TT - 1})`, result: round2(sF + sH) }; styleCell(h1, { fill: COL.tot, bold: true, fmt: FMT_EUR, center: false });
    out.sums.push({ extra: sF, importo: sH, fisso: sI, totale: round2(sF + sH) });
  });
  const j = ws.getCell(`J${SM}`); j.value = 'SPESA MENSILE'; styleCell(j, { fill: COL.rosa, center: false });
  const k = ws.getCell(`K${SM}`); k.value = { formula: `SUM(J${SM + 1},H${TT},R${TT})`, result: round2(out.sums[0].totale + out.sums[1].totale) }; styleCell(k, { fill: COL.rosa, fmt: FMT_EUR, center: false });
  return out;
}

// grafici disegnati su canvas e inseriti come immagini nel foglio RIEPILOGO
function canvasStack(mesi) {
  const W = 760, H = 340, c = document.createElement('canvas'); c.width = W * 2; c.height = H * 2;
  const x = c.getContext('2d'); x.scale(2, 2); x.fillStyle = '#fff'; x.fillRect(0, 0, W, H);
  x.fillStyle = '#0f1b24'; x.font = '600 16px sans-serif'; x.fillText('Spese mensili: Ricardo e Silvana', 20, 28);
  const L = 56, R = 16, T = 64, B = 34, iw = W - L - R, ih = H - T - B, max = niceMax(Math.max(...mesi.map(s => s.R + s.S), 1)), bw = iw / 12;
  x.font = '12px sans-serif'; x.textAlign = 'right'; x.fillStyle = '#7a8894';
  [0, .5, 1].forEach(f => { const yy = T + ih - ih * f; x.strokeStyle = f ? '#e6ebef' : '#c9d2d9'; x.beginPath(); x.moveTo(L, yy); x.lineTo(W - R, yy); x.stroke(); x.fillText(Math.round(max * f).toLocaleString('it-IT'), L - 8, yy + 4); });
  mesi.forEach((s, i) => {
    const bx = L + i * bw + bw * .18, w = bw * .64, hR = s.R / max * ih, hS = s.S / max * ih;
    x.fillStyle = '#2a78d6'; if (hR > 0) x.fillRect(bx, T + ih - hR, w, Math.max(hR - 1, 0));
    x.fillStyle = '#eb6834'; if (hS > 0) x.fillRect(bx, T + ih - hR - hS, w, Math.max(hS - 1, 0));
    x.fillStyle = '#475662'; x.textAlign = 'center'; x.fillText(s.label, bx + w / 2, H - 12);
  });
  x.textAlign = 'left'; x.fillStyle = '#2a78d6'; x.fillRect(W - 190, 18, 12, 12); x.fillStyle = '#0f1b24'; x.fillText('Ricardo', W - 172, 29);
  x.fillStyle = '#eb6834'; x.fillRect(W - 104, 18, 12, 12); x.fillStyle = '#0f1b24'; x.fillText('Silvana', W - 86, 29);
  return c.toDataURL('image/png');
}
function canvasBars(voci, titolo) {
  const rowH = 28, W = 760, H = 64 + voci.length * rowH + 12, c = document.createElement('canvas'); c.width = W * 2; c.height = H * 2;
  const x = c.getContext('2d'); x.scale(2, 2); x.fillStyle = '#fff'; x.fillRect(0, 0, W, H);
  x.fillStyle = '#0f1b24'; x.font = '600 16px sans-serif'; x.fillText(titolo, 20, 28);
  const mx = Math.max(...voci.map(v => v[1]), 1), L = 190, R = 110;
  voci.forEach(([n, v], i) => {
    const yy = 54 + i * rowH; x.font = '13px sans-serif'; x.fillStyle = '#0f1b24'; x.textAlign = 'left'; x.fillText(n.length > 24 ? n.slice(0, 23) + '…' : n, 20, yy + 14);
    x.fillStyle = '#0e7a4b'; x.fillRect(L, yy + 2, Math.max(3, (W - L - R) * v / mx), 16);
    x.fillStyle = '#475662'; x.fillText(eur(v), L + Math.max(3, (W - L - R) * v / mx) + 8, yy + 15);
  });
  return c.toDataURL('image/png');
}

async function creaWorkbook() {
  const [sm, sy] = periodoCorrente().split('-').map(Number);
  const yearOf = m => (m >= sm ? sy : sy + 1);
  const wb = new ExcelJS.Workbook(); wb.creator = 'Bilancio Casa'; wb.created = new Date();
  wb.calcProperties = { fullCalcOnLoad: true };
  const fogli = [], perVoce = new Map();
  for (let m = 1; m <= 12; m++) {
    const y = yearOf(m), rowsP = PERSONE.map(p => righeMese(p, y, m));
    rowsP.flat().forEach(r => perVoce.set(norm(r.voce), (perVoce.get(norm(r.voce)) || 0) + r.extra + r.importo));
    fogli.push({ m, y, ...buildMonthSheet(wb, m, y, rowsP) });
  }
  // --- RIEPILOGO (ultimo foglio): tabella con formule + grafici
  const ws = wb.addWorksheet('RIEPILOGO');
  [14, 14, 14, 14, 14, 14, 3, 26, 14].forEach((w, i) => ws.getColumn(i + 1).width = w);
  const fine = sm === 1 ? sy : sy + 1, titolo = sm === 1 ? `${sy}` : `${sy}/${fine}`;
  ws.mergeCells('A1:F1'); const t = ws.getCell('A1'); t.value = `RIEPILOGO FAMILIARE ${titolo}`; styleCell(t, { fill: COL.blu, bold: true });
  ['MESE', 'RICARDO', 'SILVANA', 'TOTALE', 'DI CUI FISSE', 'DI CUI EXTRA'].forEach((h, i) => { const c = ws.getCell(3, i + 1); c.value = h; styleCell(c, { fill: COL.verde, bold: true }); });
  const ordine = []; for (let i = 0; i < 12; i++) { const m = ((sm - 1 + i) % 12) + 1; ordine.push(fogli.find(f => f.m === m)); }
  const serie = []; let tR = 0, tS = 0, tF = 0, tE = 0;
  ordine.forEach((f, i) => {
    const r = 4 + i, sh = `'${MESI[f.m - 1]}'`, a = f.sums[0], b = f.sums[1];
    const c0 = ws.getCell(r, 1); c0.value = `${MESI_IT[f.m - 1]} ${f.y}`; styleCell(c0, { bold: true, center: false });
    const put = (col, formula, res) => { const c = ws.getCell(r, col); c.value = { formula, result: res }; styleCell(c, { fmt: FMT_EUR, center: false }); };
    put(2, `${sh}!H${f.TT}`, a.totale); put(3, `${sh}!R${f.TT}`, b.totale); put(4, `B${r}+C${r}`, round2(a.totale + b.totale));
    put(5, `${sh}!I${f.T}+${sh}!S${f.T}`, round2(a.fisso + b.fisso)); put(6, `${sh}!F${f.T}+${sh}!P${f.T}`, round2(a.extra + b.extra));
    tR += a.totale; tS += b.totale; tF += a.fisso + b.fisso; tE += a.extra + b.extra;
    serie.push({ label: MESI_BREVI[f.m - 1], R: a.totale, S: b.totale });
  });
  const lt = ws.getCell(16, 1); lt.value = 'TOTALE'; styleCell(lt, { bold: true, center: false });
  [[2, tR], [3, tS], [4, tR + tS], [5, tF], [6, tE]].forEach(([col, v]) => { const c = ws.getCell(16, col); c.value = { formula: `SUM(${colL(col)}4:${colL(col)}15)`, result: round2(v) }; styleCell(c, { bold: true, fmt: FMT_EUR, fill: col === 4 ? COL.tot : undefined, center: false }); });
  // voci: valori al momento dell'esportazione
  ws.getCell('H3').value = 'VOCE (totale periodo)'; styleCell(ws.getCell('H3'), { fill: COL.verde, bold: true });
  ws.getCell('I3').value = 'TOTALE'; styleCell(ws.getCell('I3'), { fill: COL.verde, bold: true });
  let voci = [...perVoce.entries()].filter(v => v[1] > 0).sort((a, b) => b[1] - a[1]);
  if (voci.length > 12) voci = [...voci.slice(0, 11), ['ALTRE', voci.slice(11).reduce((a, v) => a + v[1], 0)]];
  voci.forEach(([n, v], i) => { const a = ws.getCell(4 + i, 8); a.value = n; styleCell(a, { center: false }); const b = ws.getCell(4 + i, 9); b.value = round2(v); styleCell(b, { fmt: FMT_EUR, center: false }); });
  const n1 = ws.getCell(4 + voci.length, 8); n1.value = '(valori alla data di esportazione)'; n1.font = { italic: true, color: { argb: 'FF7A8894' }, size: 9 };
  const id1 = wb.addImage({ base64: canvasStack(serie), extension: 'png' });
  ws.addImage(id1, { tl: { col: 0, row: 18 }, ext: { width: 640, height: 286 } });
  if (voci.length) {
    const id2 = wb.addImage({ base64: canvasBars(voci.slice(0, 12), 'Dove vanno i soldi (periodo)'), extension: 'png' });
    ws.addImage(id2, { tl: { col: 0, row: 33 }, ext: { width: 640, height: Math.round((64 + Math.min(12, voci.length) * 28 + 12) * 640 / 760) } });
  }
  return { wb, nome: `FAMILIARE - ${fine}.xlsx` };
}
let ultimoFile = null;
async function esportaExcel() {
  const btn = $('#btnExp'); btn.disabled = true; btn.textContent = 'Creo il file…';
  try {
    const { wb, nome } = await creaWorkbook();
    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    ultimoFile = new File([blob], nome, { type: blob.type });
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = nome; document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 5000);
    toast('Excel creato: ' + nome);
    if (navigator.canShare && navigator.canShare({ files: [ultimoFile] })) $('#btnShare').hidden = false;
  } catch (e) { console.error(e); toast('Non sono riuscito a creare il file'); }
  btn.disabled = false; btn.textContent = "Crea l'Excel";
}

// ------------------------------------------------------------ FASE 2: importa un file FAMILIARE esistente
const cellVal = c => { let v = c.value; if (v && typeof v === 'object' && !(v instanceof Date)) { if ('result' in v) v = v.result; else if (v.richText) v = v.richText.map(t => t.text).join(''); else if ('text' in v) v = v.text; } return v; };
const numVal = v => typeof v === 'number' ? v : (parseNum(v) || 0);
async function importaExcel(file) {
  const wb = new ExcelJS.Workbook(); await wb.xlsx.load(await file.arrayBuffer());
  const righe = [], note = []; let prevYear = null, trovati = 0;
  for (const nome of MESI) {
    const ws = wb.worksheets.find(w => norm(w.name) === nome); if (!ws) continue;
    trovati++; const m = MESI.indexOf(nome) + 1, sheetRows = [];
    for (let r = 4; r <= Math.min(ws.rowCount, 150); r++) PERSONE.forEach((p, b) => {
      const o = b * 10, voce = cellVal(ws.getCell(r, 4 + o)); if (!voce) return;
      const lab = String(voce).trim(); if (!lab || norm(lab) === 'TOTALE') return;
      const dv = cellVal(ws.getCell(r, 2 + o)), ex = numVal(cellVal(ws.getCell(r, 6 + o))), im = numVal(cellVal(ws.getCell(r, 8 + o))), fi = numVal(cellVal(ws.getCell(r, 9 + o)));
      if (!(ex > 0) && !(im > 0)) return;
      let d = null; if (dv instanceof Date) d = { y: dv.getUTCFullYear(), m: dv.getUTCMonth() + 1, d: dv.getUTCDate() };
      sheetRows.push({ p, lab, d, ex, im, fi });
    });
    // anno di ogni blocco (Ricardo / Silvana): quello delle sue date; le righe senza data lo ereditano
    const annoDi = rows => { const c = {}; rows.forEach(r => { if (r.d) c[r.d.y] = (c[r.d.y] || 0) + 1; }); const k = Object.keys(c).map(Number); return k.length ? k.sort((a, b) => c[b] - c[a])[0] : null; };
    const annoFoglio = annoDi(sheetRows) ?? prevYear ?? new Date().getFullYear();
    const anniPresenti = [...new Set(sheetRows.filter(r => r.d).map(r => r.d.y))].sort();
    if (anniPresenti.length > 1) note.push(`${nome}: ci sono date di anni diversi (${anniPresenti.join(' e ')}). Le ho lasciate come nel file.`);
    prevYear = annoFoglio;
    PERSONE.forEach(p => {
      const blocco = sheetRows.filter(r => r.p === p), anno = annoDi(blocco) ?? annoFoglio;
      blocco.forEach(r => righe.push({ ...r, nd: !r.d, y: r.d ? r.d.y : anno, m: r.d ? r.d.m : m, dd: r.d ? r.d.d : 1 }));
    });
  }
  if (!trovati) return { err: 'Non trovo i fogli GENNAIO … DICEMBRE in questo file.' };
  const prima = new Map(); S.expenses.forEach(e => { const k = [e.persona, e.d, e.voce, e.tipo, e.importo].join('|'); prima.set(k, (prima.get(k) || 0) + 1); });
  let nuove = 0, doppie = 0, senzaData = 0; const voto = {}, occ = {};
  const hashStr = t => { let h = 5381; for (let i = 0; i < t.length; i++) h = ((h << 5) + h + t.charCodeAt(i)) >>> 0; return h.toString(36); };
  const add = (r, tipo, importo) => {
    const d = `${r.y}-${pad(r.m)}-${pad(r.dd)}`, voce = norm(r.lab), k = [r.p, d, voce, tipo, importo].join('|');
    (voto[voce] ||= { IMPORTO: 0, EXTRA: 0 })[tipo === 'EXTRA' ? 'EXTRA' : 'IMPORTO']++;
    const n = occ[k] = (occ[k] || 0) + 1;
    if ((prima.get(k) || 0) > 0) { prima.set(k, prima.get(k) - 1); doppie++; return; }
    S.expenses.push({ id: 'x' + hashStr(k) + '-' + n, d, persona: r.p, voce, tipo, importo: round2(importo), nota: '', nd: r.nd, u: Date.now() }); nuove++; if (r.nd) senzaData++;
  };
  righe.forEach(r => { if (r.ex > 0) add(r, 'EXTRA', r.ex); if (r.im > 0) add(r, r.fi > 0 ? 'FISSO' : 'IMPORTO', r.im); });
  Object.entries(voto).forEach(([v, c]) => { if (!(v in S.tipoPerVoce) || !DEFAULT_TIPO[v]) S.tipoPerVoce[v] = c.EXTRA > c.IMPORTO ? 'EXTRA' : 'IMPORTO'; });
  save();
  return { nuove, doppie, senzaData, note, fogli: trovati };
}

// ------------------------------------------------------------ copia di sicurezza
function backup() {
  const blob = new Blob([JSON.stringify(S, null, 1)], { type: 'application/json' });
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `bilancio-casa-${todayISO()}.json`; document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000); toast('Copia salvata');
}
async function restore(file) {
  try {
    const j = JSON.parse(await file.text());
    if (!Array.isArray(j.expenses)) throw new Error('formato');
    if (!confirm(`Ripristinare ${j.expenses.length} spese? I dati attuali verranno sostituiti.`)) return;
    S = Object.assign(defaults(), j); S.del = S.del || {}; save(); render(); toast('Dati ripristinati');
  } catch (e) { toast('File di copia non valido'); }
}


// ------------------------------------------------------------ sincronizzazione telefono <-> PC (archivio privato GitHub)
const SYNC_KEY = 'bilancio-casa.sync', SYNC_FILE = 'bilancio-casa-dati.json';
let sync = (() => { try { return JSON.parse(localStorage.getItem(SYNC_KEY)) || {}; } catch (e) { return {}; } })();
const saveSync = () => { try { localStorage.setItem(SYNC_KEY, JSON.stringify(sync)); } catch (e) { /* ignora */ } };
let syncBusy = false, syncTimer = null, lastSyncTry = 0;
const syncOn = () => !!(sync.repo && sync.token);
const syncHeaders = () => ({ Authorization: 'Bearer ' + sync.token, Accept: 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28' });
const syncURL = () => `https://api.github.com/repos/${sync.repo}/contents/${SYNC_FILE}`;
function syncStatusText() {
  if (syncBusy) return 'Sincronizzo…';
  if (sync.err) return '⚠ ' + sync.err;
  return sync.last ? 'Ultima sincronizzazione: ' + new Date(sync.last).toLocaleString('it-IT', { dateStyle: 'short', timeStyle: 'short' }) : 'Non ancora sincronizzato';
}
function updateSyncUI() { const e = $('#syncStat'); if (e) e.textContent = syncStatusText(); }
function schedSync() { if (!syncOn()) return; clearTimeout(syncTimer); syncTimer = setTimeout(() => sincronizza(), 2500); }
class SyncErr extends Error { }
async function ghGet() {
  const res = await fetch(syncURL(), { headers: syncHeaders(), cache: 'no-store' });
  if (res.status === 404) {
    const rr = await fetch(`https://api.github.com/repos/${sync.repo}`, { headers: syncHeaders(), cache: 'no-store' });
    if (!rr.ok) throw new SyncErr(rr.status === 401 ? 'Token non valido o scaduto.' : 'Archivio non trovato: controlla il nome e i permessi del token.');
    return { data: null, sha: null };
  }
  if (res.status === 401) throw new SyncErr('Token non valido o scaduto.');
  if (res.status === 403) throw new SyncErr('Il token non ha i permessi su questo archivio.');
  if (!res.ok) throw new SyncErr('GitHub ha risposto con un errore (' + res.status + ').');
  const j = await res.json();
  if (!j.content) throw new SyncErr('Il file dei dati è troppo grande per la sincronizzazione.');
  const bin = atob(j.content.replace(/\s/g, '')), txt = new TextDecoder().decode(Uint8Array.from(bin, c => c.charCodeAt(0)));
  return { data: JSON.parse(txt), sha: j.sha };
}
async function ghPut(obj, sha) {
  const bytes = new TextEncoder().encode(JSON.stringify(obj)); let bin = '';
  for (let i = 0; i < bytes.length; i += 8192) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 8192));
  const body = { message: 'Sincronizzazione Bilancio Casa', content: btoa(bin) }; if (sha) body.sha = sha;
  const res = await fetch(syncURL(), { method: 'PUT', headers: { ...syncHeaders(), 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
  if (res.status === 409 || res.status === 422) return { conflict: true };
  if (res.status === 401) throw new SyncErr('Token non valido o scaduto.');
  if (res.status === 403 || res.status === 404) throw new SyncErr('Il token non può scrivere su questo archivio (serve il permesso Contents: Read and write).');
  if (!res.ok) throw new SyncErr('GitHub ha risposto con un errore (' + res.status + ').');
  return { ok: true, sha: (await res.json()).content.sha };
}
// unisce due copie: per ogni spesa vince la modifica più recente; le cancellazioni valgono su entrambi
function mergeAll(L, R, adoptRemoteRecurring) {
  const del = { ...(R.del || {}) }; Object.entries(L.del || {}).forEach(([k, t]) => { del[k] = Math.max(del[k] || 0, t); });
  const mergeList = (a, b) => { const m = new Map(); [...(b || []), ...(a || [])].forEach(x => { const c = m.get(x.id); if (!c || (x.u || 0) >= (c.u || 0)) m.set(x.id, x); }); return [...m.values()].filter(x => !(del[x.id] !== undefined && del[x.id] >= (x.u || 0))); };
  const expenses = mergeList(L.expenses, R.expenses);
  let recurring = adoptRemoteRecurring ? (R.recurring || []).filter(x => !(del[x.id] !== undefined && del[x.id] >= (x.u || 0))) : mergeList(L.recurring, R.recurring);
  const seen = new Set(); recurring = recurring.filter(r => { const k = [r.persona, norm(r.voce), r.tipo, r.importo, r.giorno, r.da, r.a].join('|'); if (seen.has(k)) return false; seen.add(k); return true; });
  return Object.assign({}, L, { expenses, recurring, del, voci: [...new Set([...(L.voci || []), ...(R.voci || [])])], tipoPerVoce: { ...(R.tipoPerVoce || {}), ...(L.tipoPerVoce || {}) } });
}
const byId = (a, b) => (a.id < b.id ? -1 : a.id > b.id ? 1 : 0);
const payload = st => ({ v: 1, expenses: [...(st.expenses || [])].sort(byId), recurring: [...(st.recurring || [])].sort(byId), del: st.del || {}, voci: [...(st.voci || [])].sort(), tipoPerVoce: st.tipoPerVoce || {} });
const stable = o => JSON.stringify(o, (k, v) => (v && typeof v === 'object' && !Array.isArray(v)) ? Object.fromEntries(Object.entries(v).sort()) : v);
async function sincronizza(manuale) {
  if (!syncOn()) { if (manuale) toast('Prima collega l\'archivio'); return; }
  if (syncBusy) return; syncBusy = true; sync.err = ''; updateSyncUI();
  const prima = stable(payload(S));
  try {
    for (let tent = 0; tent < 4; tent++) {
      const r = await ghGet();
      if (r.data) { S = mergeAll(S, r.data, !sync.last && S.expenses.length === 0); save(true); }
      const out = payload(S);
      if (r.data && stable(payload(r.data)) === stable(out)) break;      // già allineato: niente da scrivere
      const put = await ghPut(out, r.sha);
      if (put.ok) break;
      if (tent === 3) throw new SyncErr('Conflitto con l\'altro dispositivo: riprova tra poco.');
    }
    sync.last = Date.now(); saveSync();
    if (stable(payload(S)) !== prima) { render(); if (manuale) toast('Dati aggiornati'); } else if (manuale) toast('Già tutto allineato');
  } catch (e) {
    sync.err = e instanceof SyncErr ? e.message : (navigator.onLine ? 'Sincronizzazione non riuscita.' : 'Sei offline: riprovo appena torna la rete.');
    if (!(e instanceof SyncErr)) console.warn(e);
    saveSync();
  }
  syncBusy = false; lastSyncTry = Date.now(); updateSyncUI();
}
function openSync() {
  sheet.hidden = false; form = null;
  sheet.innerHTML = `<div class="scrim" data-act="close"></div><section class="panel" role="dialog" aria-modal="true" aria-label="Collega l'archivio"><div class="grab"></div><div class="card" style="box-shadow:none">
    <h2>Collega l'archivio privato</h2>
    <p>Ti servono un repository <b>privato</b> su GitHub (per i dati) e un token con accesso solo a quello. La guida passo passo è nel file LEGGIMI_SINCRONIZZAZIONE.md.</p>
    <div class="field"><label for="sRepo">Archivio dei dati (utente/nome)</label><input class="inp" id="sRepo" placeholder="es. mario/bilancio-dati" autocapitalize="off" autocomplete="off" value="${esc(sync.repo || '')}"></div>
    <div class="field"><label for="sTok">Token di accesso</label><input class="inp" id="sTok" type="password" placeholder="github_pat_…" autocapitalize="off" autocomplete="off"></div>
    <p class="hint">Il token resta solo su questo dispositivo e non entra nelle copie di sicurezza.</p>
    <div class="actions" style="margin-top:12px"><button class="btn" data-act="close">Annulla</button><button class="btn primary" data-act="syncSave">Collega e sincronizza</button></div></div></section>`;
}

// ------------------------------------------------------------ eventi
document.addEventListener('click', async ev => {
  const nav = ev.target.closest('#nav button'); if (nav) { ui.tab = nav.dataset.tab; render(); $('#view').scrollTop = 0; window.scrollTo(0, 0); return; }
  const el = ev.target.closest('[data-act]'); if (!el) return;
  const act = el.dataset.act;
  switch (act) {
    case 'camera': $('#inCam').click(); break;
    case 'files': $('#inFiles').click(); break;
    case 'manuale': queue = []; openForm(newForm()); break;
    case 'close': queue = []; closeForm(); break;
    case 'salta': if (queue.length) processNext(); else closeForm(); break;
    case 'persona': form.persona = el.dataset.p; form.dirty.persona = true; renderSheet(); break;
    case 'voce': form.voce = el.dataset.v; form.dirty.voce = true; if (!form.dirty.tipo) form.tipo = tipoDi(form.voce); renderSheet(); break;
    case 'tipo': form.tipo = el.dataset.t; form.dirty.tipo = true; renderSheet(); break;
    case 'salva': salvaForm(); break;
    case 'elimina': if (confirm('Eliminare questa spesa?')) { S.del = S.del || {}; S.del[form.id] = Date.now(); S.expenses = S.expenses.filter(e => e.id !== form.id); save(); closeForm(); render(); toast('Spesa eliminata'); } break;
    case 'edit': { const e = S.expenses.find(x => x.id === el.dataset.id); if (e) openForm(newForm({ id: e.id, importo: String(e.importo).replace('.', ','), d: e.d, persona: e.persona, voce: e.voce, tipo: e.tipo, nota: e.nota || '', dirty: { importo: 1, d: 1, voce: 1, tipo: 1 } })); break; }
    case 'fPersona': ui.persona = el.dataset.p; render(); break;
    case 'barra': { const s = seriesMensili(meseCorrente()).find(x => x.ym === el.dataset.ym); ui.sel = el.dataset.ym; const ro = $('#readout'); if (ro) ro.innerHTML = readoutHTML(s); break; }
    case 'export': esportaExcel(); break;
    case 'share': if (ultimoFile) { try { await navigator.share({ files: [ultimoFile], title: ultimoFile.name }); } catch (e) { /* annullato */ } } break;
    case 'importXlsx': $('#inXlsx').click(); break;
    case 'backup': backup(); break;
    case 'restore': $('#inJson').click(); break;
    case 'syncSetup': openSync(); break;
    case 'syncSave': { const repo = $('#sRepo').value.trim().replace(/^https?:\/\/github\.com\//, '').replace(/\/$/, ''), tok = $('#sTok').value.trim() || sync.token; if (!/^[\w.-]+\/[\w.-]+$/.test(repo) || !tok) { toast('Servono archivio (utente/nome) e token'); break; } sync = { repo, token: tok, last: 0, err: '' }; saveSync(); closeForm(); render(); await sincronizza(true); render(); break; }
    case 'syncNow': await sincronizza(true); break;
    case 'syncOff': if (confirm('Scollegare questo dispositivo? I dati restano qui e nell\'archivio.')) { sync = {}; saveSync(); render(); } break;
    case 'newRec': openRec(null); break;
    case 'editRec': openRec(S.recurring.find(r => r.id === el.dataset.id)); break;
    case 'saveRec': saveRec(el.dataset.id); break;
    case 'delRec': { const id = $('[data-act=saveRec]').dataset.id; if (confirm('Eliminare questa voce fissa?')) { S.del = S.del || {}; S.del[id] = Date.now(); S.recurring = S.recurring.filter(r => r.id !== id); save(); closeForm(); render(); } break; }
  }
});
document.addEventListener('input', ev => {
  if (!form) return; const t = ev.target;
  if (t.id === 'fImp') { form.importo = t.value; form.dirty.importo = true; }
  else if (t.id === 'fData') { form.d = t.value; form.dirty.d = true; }
  else if (t.id === 'fNota') { form.nota = t.value; }
  else if (t.id === 'fVoce') { form.voce = t.value.toUpperCase(); form.dirty.voce = true; }
});
document.addEventListener('change', ev => {
  const t = ev.target;
  if (t.id === 'fVoce' && form) { form.voce = norm(t.value); if (!form.dirty.tipo && form.voce) form.tipo = tipoDi(form.voce); renderSheet(); }
  else if (t.id === 'selMese') { ui.mesePick = t.value; render(); }
  else if (t.id === 'selMeseG') { ui.mese = t.value; ui.sel = null; render(); }
  else if (t.id === 'selPer') { S.settings.periodo = t.value; save(); }
});
$('#inCam').addEventListener('change', e => { addFiles(e.target.files); e.target.value = ''; });
$('#inFiles').addEventListener('change', e => { addFiles(e.target.files); e.target.value = ''; });
$('#inJson').addEventListener('change', e => { if (e.target.files[0]) restore(e.target.files[0]); e.target.value = ''; });
$('#inXlsx').addEventListener('change', async e => {
  const f = e.target.files[0]; e.target.value = ''; if (!f) return;
  if (/\.xls$/i.test(f.name)) { toast('Salva il file come .xlsx e riprova'); return; }
  if (!/\.xls[xm]$/i.test(f.name)) { toast('Scegli un file Excel (.xlsx)'); return; }
  const box = $('#impReport'); box.innerHTML = '<div class="report">Leggo il file…</div>';
  try {
    const r = await importaExcel(f);
    if (r.err) box.innerHTML = `<div class="report">${esc(r.err)}</div>`;
    else { render(); $('#impReport').innerHTML = `<div class="report">${esc(`Importate ${r.nuove} spese da ${r.fogli} fogli.${r.doppie ? `\nGià presenti, saltate: ${r.doppie}.` : ''}${r.senzaData ? `\nSenza data nel file (le metto al giorno 1 del mese): ${r.senzaData}.` : ''}${r.note.length ? '\n' + r.note.join('\n') : ''}`)}</div>`; }
  } catch (err) { console.error(err); box.innerHTML = '<div class="report">Non riesco a leggere questo file. È un .xlsx valido?</div>'; }
});
document.addEventListener('keydown', e => { if (e.key === 'Escape' && !sheet.hidden) { queue = []; closeForm(); } });

// ------------------------------------------------------------ avvio
$('#inXlsx').removeAttribute('accept');
render();
if ('serviceWorker' in navigator && location.protocol !== 'file:') navigator.serviceWorker.register('sw.js').catch(() => { });
if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => { });
if (syncOn()) sincronizza();
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible' && syncOn() && Date.now() - lastSyncTry > 20000) sincronizza(); });
window.addEventListener('online', () => { if (syncOn()) sincronizza(); });
window.__app = { get S() { return S; }, set S(v) { S = v; }, creaWorkbook, importaExcel, sincronizza, mergeAll, get sync() { return sync; }, findAmount, findDate, suggestVoce, render, ui };
