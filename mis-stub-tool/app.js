// MiS Stub Tool - runs entirely in the browser. No network calls are made.


pdfjsLib.GlobalWorkerOptions.workerSrc = 'lib/pdf.worker.min.js';

const $ = id => document.getElementById(id);
const state = { db: null, pdfs: [], matches: [], selected: new Set() };

// ---------- helpers ----------
const normId = v => String(v == null ? '' : v).replace(/\D/g, '').replace(/^0+/, '');
const normHdr = v => String(v == null ? '' : v).replace(/\s+/g, ' ').trim().toUpperCase();
const pad2 = n => String(n).padStart(2, '0');
const fmtDate = d => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
const money = n => n == null ? '' : n.toLocaleString('en-ZA', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const cleanName = s => s.replace(/[\\/:*?"<>|\x00-\x1f]/g, '').replace(/\s+/g, ' ').trim().replace(/[. ]+$/, '');
const tick = () => new Promise(r => setTimeout(r, 0));

function parseStubDate(s) { // dd/mm/yyyy
  const m = /(\d{2})\/(\d{2})\/(\d{4})/.exec(s || '');
  return m ? new Date(+m[3], +m[2] - 1, +m[1]) : null;
}
function toDate(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === 'string') {
    const d = new Date(v);
    if (!isNaN(d)) return d;
  }
  return null;
}

// ---------- database ----------
async function loadDb(file) {
  const wb = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
  const byDisb = new Map(), rows = [];
  for (const name of wb.SheetNames) {
    const grid = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
    let h = -1;
    for (let i = 0; i < Math.min(grid.length, 60); i++) {
      if (grid[i].some(c => normHdr(c) === 'DISBURSEMENT NO')) { h = i; break; }
    }
    if (h < 0) continue;
    const col = {};
    grid[h].forEach((c, i) => { const k = normHdr(c); if (k && !(k in col)) col[k] = i; });
    const need = ['SUPPLIER', 'DISBURSEMENT NO'];
    if (need.some(k => !(k in col))) continue;
    for (let i = h + 1; i < grid.length; i++) {
      const r = grid[i];
      const supplier = r[col['SUPPLIER']];
      const disb = normId(r[col['DISBURSEMENT NO']]);
      if (!supplier || !disb) continue;
      const rec = {
        sheet: name, supplier: String(supplier).replace(/\s+/g, ' ').trim(), disb,
        payRef: normId(r[col['PAYMENT REF.']]),
        amount: typeof r[col['AMOUNT']] === 'number' ? r[col['AMOUNT']] : null,
        actionDate: toDate(r[col['ACTION DATE']]),
        used: false
      };
      rows.push(rec);
      if (!byDisb.has(disb)) byDisb.set(disb, []);
      byDisb.get(disb).push(rec);
    }
  }
  if (!rows.length) throw new Error('Could not find a "SUPPLIER" / "DISBURSEMENT NO" header row in this workbook.');
  return { rows, byDisb, sheets: wb.SheetNames.length };
}

// ---------- stub parsing ----------
function parseStubText(t) {
  const disb = /DISBURSEMENT NO\s*:\s*(\d+)/.exec(t);
  if (!disb) return null;
  const act = /ACTION DATE\s*:\s*(\d{2}\/\d{2}\/\d{4})/.exec(t);
  const payee = /PAYEE NAME\s*:\s*(.*?)\s+BANK NAME/.exec(t);
  const pays = new Set((t.match(/(?<!\d)0\d{8}(?!\d)/g) || []).map(normId));
  const tot = /TOTAL\s+([\d,]+\.\d{2})/.exec(t);
  return {
    disb: normId(disb[1]),
    date: act ? parseStubDate(act[1]) : null,
    payee: payee ? payee[1].trim() : '',
    pays,
    total: tot ? parseFloat(tot[1].replace(/,/g, '')) : null
  };
}

// pdf.js can split one word into fragments (e.g. "ST" + "UB"). Rejoin fragments that touch
// and put a single space only where there is a real gap, so numbers and labels stay intact.
function rebuildText(items) {
  let out = '', lastY = null, lastEnd = null;
  for (const it of items) {
    if (!it.str) { if (it.hasEOL) out += ' '; continue; }
    const x = it.transform[4], y = it.transform[5], h = Math.abs(it.transform[3]) || 8;
    if (out) {
      const sameLine = lastY !== null && Math.abs(y - lastY) < h * 0.5;
      const gap = sameLine ? x - lastEnd : Infinity;
      if (!sameLine || gap > h * 0.25) out += ' ';
    }
    out += it.str;
    lastY = y; lastEnd = x + (it.width || 0);
  }
  return out.replace(/\s+/g, ' ');
}

async function scan() {
  const db = state.db;
  state.matches = []; state.selected.clear();
  db.rows.forEach(r => r.used = false);
  const warns = [], unread = [];
  let totalPages = 0, readable = 0;
  const prog = $('prog'); prog.classList.remove('hidden');
  const grand = state.pdfs.reduce((s, p) => s + (p.pages || 0), 0) || 1;
  let done = 0;

  for (const f of state.pdfs) {
    let current = null; // payment the previous page belonged to (if ours), for continuation pages
    const doc = await pdfjsLib.getDocument({ data: new Uint8Array(await f.file.arrayBuffer()) }).promise;
    for (let p = 1; p <= doc.numPages; p++) {
      totalPages++;
      const page = await doc.getPage(p);
      const tc = await page.getTextContent();
      const text = rebuildText(tc.items);
      page.cleanup();
      const s = parseStubText(text);
      if (!s) {
        if (/PAYMENT STUB/i.test(text)) {
          // continuation page: a payment with many lines runs over onto extra pages without a header
          readable++;
          if (current) {
            current.pageIdxs.push(p - 1);
            const t2 = /TOTAL\s+([\d,]+\.\d{2})/.exec(text);
            if (t2) current.total = parseFloat(t2[1].replace(/,/g, ''));
          }
        } else unread.push(`${f.file.name} page ${p}`);
      } else {
        current = null;
        readable++;
        const cands = (db.byDisb.get(s.disb) || []);
        if (cands.length) {
          // confirm with the payment number when the database has one
          const rec = cands.find(c => !c.payRef || s.pays.has(c.payRef));
          if (rec) {
            rec.used = true;
            current = {
              id: state.matches.length, file: f.file, pageIdxs: [p - 1], srcName: f.file.name, srcPage: p,
              supplier: rec.supplier, disb: s.disb, date: s.date || rec.actionDate, total: s.total,
              verifiedByPayRef: !!rec.payRef
            };
            state.matches.push(current);
          } else {
            warns.push(`${f.file.name} page ${p}: disbursement ${s.disb} is in the database (${cands[0].supplier}) but its payment ref ${cands[0].payRef} is not on the stub, so the page was ignored.`);
          }
        }
      }
      done++;
      if (done % 10 === 0) { prog.value = done / grand; $('scanStatus').textContent = `Reading ${f.file.name} (page ${p}/${doc.numPages})…`; await tick(); }
    }
    await doc.destroy();
  }
  prog.value = 1;
  $('scanStatus').textContent = 'Done.';

  // build stub-name -> final file name
  state.matches.sort((a, b) => (a.supplier.localeCompare(b.supplier)) || ((a.date || 0) - (b.date || 0)) || a.disb.localeCompare(b.disb));
  const seen = new Map();
  for (const m of state.matches) {
    const date = m.date ? fmtDate(m.date) : 'no-date';
    let base = `${cleanName(m.supplier)} - ${date} - ${m.disb.padStart(9, '0')}`;
    const n = (seen.get(base) || 0) + 1; seen.set(base, n);
    if (n > 1) base += ` (${n})`;
    m.fileName = base + '.pdf';
    m.folder = cleanName(m.supplier);
    state.selected.add(m.id);
  }
  const unmatched = db.rows.filter(r => !r.used);
  showResults({ totalPages, readable, unmatched, warns, unread });
}

// ---------- UI ----------
function showResults({ totalPages, readable, unmatched, warns, unread }) {
  $('results').classList.remove('hidden');
  const payees = new Set(state.matches.map(m => m.supplier)).size;
  $('stats').innerHTML = [
    [totalPages.toLocaleString(), 'stub pages scanned'],
    [state.matches.length.toLocaleString(), 'our payments found'],
    [payees.toLocaleString(), 'different payees'],
    [(totalPages - state.matches.reduce((s, m) => s + m.pageIdxs.length, 0)).toLocaleString(), 'pages ignored (not ours)'],
    [unmatched.length.toLocaleString(), 'database rows with no stub'],
  ].map(([n, l]) => `<div class="stat"><b>${n}</b>${l}</div>`).join('');

  $('sumUnmatched').textContent = `${unmatched.length} database rows had no matching stub in these PDFs (older payments, or stubs not supplied)`;
  $('unmatchedRows').innerHTML = unmatched.map(r => `<tr><td>${esc(r.sheet)}</td><td>${esc(r.supplier)}</td><td>${r.disb}</td><td>${r.payRef}</td><td class="num">${money(r.amount)}</td></tr>`).join('');
  $('sumWarn').textContent = `${warns.length} payment-reference mismatches`;
  $('sumWarn').className = warns.length ? 'warn' : 'mute';
  $('warnList').innerHTML = warns.map(w => `<li>${esc(w)}</li>`).join('');
  $('sumUnread').textContent = `${unread.length} pages could not be read as text (cannot be matched)`;
  $('sumUnread').className = unread.length ? 'warn' : 'mute';
  $('unreadList').innerHTML = unread.map(w => `<li>${esc(w)}</li>`).join('') + (unread.length ? '<li class="mute">These pages have no usable text layer (e.g. image-only or garbled print-to-PDF). Open them manually.</li>' : '');
  renderRows();
}

function visibleMatches() {
  const q = $('filter').value.trim().toUpperCase();
  return state.matches.filter(m => !q || m.supplier.toUpperCase().includes(q) || m.disb.includes(q) || (m.date && fmtDate(m.date).includes(q)));
}
function renderRows() {
  const vis = visibleMatches();
  $('rows').innerHTML = vis.map(m => `<tr>
    <td><input type="checkbox" data-id="${m.id}" ${state.selected.has(m.id) ? 'checked' : ''}></td>
    <td>${esc(m.supplier)}</td><td>${m.date ? fmtDate(m.date) : '<span class="warn">unknown</span>'}</td>
    <td>${m.disb.padStart(9, '0')}</td><td class="num">${money(m.total)}</td>
    <td class="mute">${esc(m.srcName)} p${m.srcPage}${m.pageIdxs.length > 1 ? '–' + (m.pageIdxs[m.pageIdxs.length - 1] + 1) : ''}</td>
    <td>${m.verifiedByPayRef ? '<span class="ok" title="Payment ref also matches">✓</span>' : '<span class="warn" title="Matched on disbursement no. only">~</span>'}</td></tr>`).join('');
  updateCount();
}
function updateCount() {
  $('selCount').textContent = `${state.selected.size} of ${state.matches.length} selected`;
  $('btnZip').disabled = $('btnDir').disabled = state.selected.size === 0;
}
$('rows').addEventListener('change', e => {
  const id = +e.target.dataset.id;
  e.target.checked ? state.selected.add(id) : state.selected.delete(id);
  updateCount();
});
$('filter').addEventListener('input', renderRows);
$('btnAll').onclick = () => { visibleMatches().forEach(m => state.selected.add(m.id)); renderRows(); };
$('btnNone').onclick = () => { visibleMatches().forEach(m => state.selected.delete(m.id)); renderRows(); };

// ---------- output ----------
async function buildPdfs(onItem) {
  const sel = state.matches.filter(m => state.selected.has(m.id));
  const cache = new Map(); // file -> PDFDocument (kept only for the current file)
  const order = [...sel].sort((a, b) => state.pdfs.findIndex(p => p.file === a.file) - state.pdfs.findIndex(p => p.file === b.file));
  let n = 0;
  for (const m of order) {
    if (!cache.has(m.file)) { cache.clear(); cache.set(m.file, await PDFLib.PDFDocument.load(await m.file.arrayBuffer(), { ignoreEncryption: true })); }
    const src = cache.get(m.file);
    // Pure page cut: page content is copied untouched. Do not stamp our own producer/dates;
    // carry the original document info over instead.
    const out = await PDFLib.PDFDocument.create({ updateMetadata: false });
    try {
      const ref = src.context.trailerInfo.Info, info = ref && src.context.lookup(ref);
      if (info) out.context.trailerInfo.Info = out.context.register(info.clone(out.context));
    } catch (e) { /* metadata is optional */ }
    const pg = await out.copyPages(src, m.pageIdxs);
    pg.forEach(x => out.addPage(x));
    const bytes = await out.save({ updateFieldAppearances: false });
    await onItem(m, bytes);
    n++;
    if (n % 5 === 0) { $('outStatus').textContent = `Writing ${n}/${order.length}…`; await tick(); }
  }
  return n;
}
const pathFor = m => $('optFolders').checked ? `${m.folder}/${m.fileName}` : m.fileName;

$('btnZip').onclick = async () => {
  $('btnZip').disabled = true;
  try {
    const zip = new JSZip();
    const n = await buildPdfs((m, bytes) => { zip.file(pathFor(m), bytes); });
    $('outStatus').textContent = 'Compressing…'; await tick();
    const blob = await zip.generateAsync({ type: 'blob', compression: 'STORE' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = `Our payment stubs ${fmtDate(new Date())}.zip`;
    document.body.appendChild(a); a.click(); a.remove();
    $('outStatus').innerHTML = `<span class="ok">Saved ${n} PDFs to a ZIP.</span>`;
  } catch (e) { $('outStatus').innerHTML = `<span class="bad">${esc(e.message)}</span>`; }
  updateCount();
};

if (!window.showDirectoryPicker) $('btnDir').classList.add('hidden');
$('btnDir').onclick = async () => {
  try {
    const root = await window.showDirectoryPicker({ mode: 'readwrite' });
    const n = await buildPdfs(async (m, bytes) => {
      let dir = root;
      if ($('optFolders').checked) dir = await root.getDirectoryHandle(m.folder, { create: true });
      const fh = await dir.getFileHandle(m.fileName, { create: true });
      const w = await fh.createWritable(); await w.write(bytes); await w.close();
    });
    $('outStatus').innerHTML = `<span class="ok">Saved ${n} PDFs into the chosen folder.</span>`;
  } catch (e) { if (e.name !== 'AbortError') $('outStatus').innerHTML = `<span class="bad">${esc(e.message)}</span>`; }
};

// ---------- inputs ----------
function wireDrop(dropEl, inputEl, handler) {
  dropEl.onclick = () => inputEl.click();
  inputEl.onchange = () => handler([...inputEl.files]);
  ['dragenter', 'dragover'].forEach(ev => dropEl.addEventListener(ev, e => { e.preventDefault(); dropEl.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev => dropEl.addEventListener(ev, e => { e.preventDefault(); dropEl.classList.remove('over'); }));
  dropEl.addEventListener('drop', async e => handler(await filesFromDrop(e.dataTransfer)));
}
async function filesFromDrop(dt) {
  const out = [];
  const walk = async entry => {
    if (entry.isFile) out.push(await new Promise(r => entry.file(r)));
    else if (entry.isDirectory) {
      const rd = entry.createReader();
      let batch;
      do { batch = await new Promise(r => rd.readEntries(r)); for (const e of batch) await walk(e); } while (batch.length);
    }
  };
  const items = [...(dt.items || [])].map(i => i.webkitGetAsEntry && i.webkitGetAsEntry()).filter(Boolean);
  if (items.length) for (const it of items) await walk(it); else out.push(...dt.files);
  return out;
}
function refreshScanBtn() { $('btnScan').disabled = !(state.db && state.pdfs.length); }

wireDrop($('dropDb'), $('fileDb'), async files => {
  const f = files.find(x => /\.xls[xm]?$/i.test(x.name)); if (!f) return;
  $('dbInfo').textContent = 'Reading…';
  try {
    state.db = await loadDb(f);
    $('dbInfo').innerHTML = `<span class="ok">✓ ${esc(f.name)}</span>: ${state.db.rows.length} payment rows across ${state.db.sheets} sheet(s).`;
  } catch (e) { state.db = null; $('dbInfo').innerHTML = `<span class="bad">${esc(e.message)}</span>`; }
  refreshScanBtn();
});
wireDrop($('dropPdf'), $('filePdf'), async files => {
  const pdfs = files.filter(x => /\.pdf$/i.test(x.name));
  if (!pdfs.length) return;
  $('pdfInfo').textContent = 'Counting pages…';
  state.pdfs = [];
  for (const file of pdfs.sort((a, b) => a.name.localeCompare(b.name))) {
    try {
      const d = await pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
      state.pdfs.push({ file, pages: d.numPages }); await d.destroy();
    } catch (e) { state.pdfs.push({ file, pages: 0 }); }
  }
  const tp = state.pdfs.reduce((s, p) => s + p.pages, 0);
  $('pdfInfo').innerHTML = `<span class="ok">✓ ${state.pdfs.length} PDF(s)</span>, ${tp.toLocaleString()} pages in total.`;
  refreshScanBtn();
});

$('btnScan').onclick = async () => {
  $('btnScan').disabled = true; $('results').classList.add('hidden');
  try { await scan(); } catch (e) { $('scanStatus').innerHTML = `<span class="bad">${esc(e.message)}</span>`; }
  refreshScanBtn();
};
