(() => {
'use strict';
const $ = s => document.querySelector(s);
const state = { companies: [], selected: null };
const LS = { guide: 'sep_guide', p1: 'sep_p1', p2: 'sep_p2' };

/* ---------- service worker + splash ---------- */
if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
setTimeout(() => {
  $('#splash').classList.add('hide');
  $('#app').hidden = false;
  setTimeout(() => $('#splash').remove(), 700);
}, 5000);

/* ---------- helpers ---------- */
const nbsp = /[\u00a0\u200c\u200b\u200e\u200f]/g;
const cleanCred = v => {
  if (v == null) return '';
  v = String(v).replace(nbsp, ' ').trim();
  return /^[*\-–—_.]+$/.test(v) ? '' : v;
};
const cleanName = v => String(v ?? '').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
const normalize = s => String(s ?? '').toLowerCase()
  .replace(/ي/g, 'ی').replace(/ك/g, 'ک').replace(/[\u200c\u00a0]/g, ' ')
  .replace(/[۰-۹]/g, d => d.charCodeAt(0) - 0x6f0).replace(/[٠-٩]/g, d => d.charCodeAt(0) - 0x660)
  .replace(/\s+/g, ' ').trim();
const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

/* ---------- xlsx parsing (offline, JSZip only) ---------- */
const colIdx = ref => { let n = 0; for (const ch of ref.replace(/[0-9]/g, '')) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };

async function parseXlsx(file) {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const dp = new DOMParser();
  const read = async p => { const f = zip.file(p); return f ? dp.parseFromString(await f.async('string'), 'application/xml') : null; };
  const shared = [];
  const sd = await read('xl/sharedStrings.xml');
  if (sd) for (const si of sd.getElementsByTagName('si'))
    shared.push([...si.getElementsByTagName('t')].filter(t => t.parentNode.nodeName !== 'rPh').map(t => t.textContent).join(''));

  let best = null, bestCount = -1;
  for (const f of zip.file(/^xl\/worksheets\/sheet\d+\.xml$/)) {
    const doc = dp.parseFromString(await f.async('string'), 'application/xml');
    const rows = []; let count = 0;
    for (const c of doc.getElementsByTagName('c')) {
      const ref = c.getAttribute('r'); if (!ref) continue;
      const r = parseInt(ref.replace(/[A-Z]/gi, ''), 10) - 1, ci = colIdx(ref.toUpperCase());
      const t = c.getAttribute('t');
      const vEl = c.getElementsByTagName('v')[0];
      let val = null;
      if (t === 'inlineStr') val = [...c.getElementsByTagName('t')].map(x => x.textContent).join('');
      else if (vEl) val = t === 's' ? shared[+vEl.textContent] : vEl.textContent;
      if (val == null || val === '') continue;
      (rows[r] = rows[r] || [])[ci] = val; count++;
    }
    if (count > bestCount) { best = rows; bestCount = count; }
  }
  if (!best || bestCount < 1) throw new Error('فایل اکسل خالی است.');
  return Array.from(best, r => r || []);
}

function buildCompanies(rows) {
  let h = -1;
  for (let i = 0; i < Math.min(rows.length - 1, 6); i++)
    if ((rows[i + 1] || []).some(v => /کاربری/.test(String(v ?? '')))) { h = i; break; }
  if (h < 0) throw new Error('ساختار فایل شناخته نشد. ردیف دوم باید شامل «نام کاربری» و «رمز عبور» باشد.');

  const sysRow = rows[h], fieldRow = rows[h + 1], systems = {}, order = [];
  let cur = '';
  const maxCol = Math.max(sysRow.length, fieldRow.length);
  for (let c = 1; c < maxCol; c++) {
    const n = cleanName(sysRow[c]); if (n) cur = n;
    const f = String(fieldRow[c] ?? '');
    const kind = /کاربری/.test(f) ? 'user' : /رمز/.test(f) ? 'pass' : null;
    if (!cur || !kind) continue;
    if (!systems[cur]) { systems[cur] = {}; order.push(cur); }
    systems[cur][kind] = c;
  }
  const out = [];
  for (let r = h + 2; r < rows.length; r++) {
    const row = rows[r] || [], name = cleanName(row[0]);
    if (!name) continue;
    const list = [];
    for (const s of order) {
      const user = cleanCred(row[systems[s].user]), pass = cleanCred(row[systems[s].pass]);
      if (user || pass) list.push({ name: s, user, pass });
    }
    out.push({ name, norm: normalize(name), systems: list });
  }
  if (!out.length) throw new Error('هیچ شرکتی در فایل پیدا نشد.');
  return out;
}

/* ---------- UI ---------- */
$('#uploadBtn').onclick = () => $('#file').click();
$('#file').onchange = async e => {
  const f = e.target.files[0]; if (!f) return;
  const msg = $('#fileMsg'); msg.className = 'msg'; msg.textContent = 'در حال خواندن فایل...';
  try {
    state.companies = buildCompanies(await parseXlsx(f));
    msg.className = 'msg ok';
    msg.textContent = `${f.name} — ${state.companies.length.toLocaleString('fa-IR')} شرکت بارگذاری شد.`;
    $('#q').disabled = false; $('#q').value = ''; $('#q').focus();
    select(null); $('#results').innerHTML = '';
  } catch (err) {
    msg.className = 'msg err'; msg.textContent = err.message || 'خطا در خواندن فایل.';
  }
  e.target.value = '';
};

$('#q').oninput = () => {
  const toks = normalize($('#q').value).split(' ').filter(Boolean);
  const box = $('#results'); box.innerHTML = '';
  if (!toks.length) { select(null); return; }
  const hits = state.companies.filter(c => toks.every(t => c.norm.includes(t)));
  hits.forEach(c => {
    const d = document.createElement('div'); d.textContent = c.name;
    d.onclick = () => select(c); d.dataset.n = c.name; box.appendChild(d);
  });
  if (!hits.length) box.innerHTML = '<div style="cursor:default;color:#5b6b7c">موردی پیدا نشد</div>';
  if (hits.length === 1) select(hits[0]); else if (!hits.includes(state.selected)) select(null);
};

function select(c) {
  state.selected = c;
  document.querySelectorAll('#results div').forEach(d => d.classList.toggle('sel', c && d.dataset.n === c.name));
  const pv = $('#preview');
  $('#dlBtn').disabled = !c; $('#dlMsg').textContent = '';
  if (!c) { pv.hidden = true; pv.innerHTML = ''; return; }
  pv.hidden = false;
  pv.innerHTML = `<div class="company" style="margin-top:14px">${esc(c.name)}</div>` + (c.systems.length
    ? `<table><tr><th>ردیف</th><th>سامانه</th><th>نام کاربری</th><th>رمز عبور</th></tr>${c.systems.map((s, i) =>
        `<tr><td>${(i + 1).toLocaleString('fa-IR')}</td><td>${esc(s.name)}</td><td class="lat">${esc(s.user || '—')}</td><td class="lat">${esc(s.pass || '—')}</td></tr>`).join('')}</table>`
    : '<div class="msg err">برای این شرکت اطلاعات سامانه‌ای ثبت نشده است.</div>');
}

/* settings */
$('#guide').value = localStorage.getItem(LS.guide) || '';
$('#p1').value = localStorage.getItem(LS.p1) || '';
$('#p2').value = localStorage.getItem(LS.p2) || '';
$('#saveBtn').onclick = () => {
  localStorage.setItem(LS.guide, $('#guide').value);
  localStorage.setItem(LS.p1, $('#p1').value.trim());
  localStorage.setItem(LS.p2, $('#p2').value.trim());
  const m = $('#saveMsg'); m.className = 'msg ok'; m.textContent = 'ذخیره شد و برای همه فایل‌های بعدی استفاده می‌شود.';
  setTimeout(() => m.textContent = '', 3000);
};

/* ---------- PDF (canvas -> JPEG -> PDF, fully offline) ---------- */
const W = 1654, H = 2339, M = 110, CW = W - 2 * M, FOOT = 210;
const BLUE = '#14509f', LIGHT = '#dbe8f7', ZEBRA = '#f5f8fc', TXT = '#1c2733';
const UI = 'Yekan, Tahoma, sans-serif', LAT = "Menlo, Consolas, 'Courier New', monospace";

const setF = (ctx, size, fam = UI) => { ctx.font = `${size}px ${fam}`; };
function fit(ctx, text, maxW, size, fam, min = 18) {
  let s = size; setF(ctx, s, fam);
  while (s > min && ctx.measureText(text).width > maxW) { s -= 2; setF(ctx, s, fam); }
  return s;
}
function text(ctx, str, x, y, { align = 'center', dir = 'rtl', color = TXT, size = 34, fam = UI, maxW } = {}) {
  ctx.direction = dir; ctx.textAlign = align; ctx.textBaseline = 'middle'; ctx.fillStyle = color;
  if (maxW) fit(ctx, str, maxW, size, fam); else setF(ctx, size, fam);
  ctx.fillText(str, x, y);
}
function wrap(ctx, str, maxW) {
  const lines = [];
  for (const para of String(str).split(/\r?\n/)) {
    if (!para.trim()) { lines.push(''); continue; }
    let line = '';
    for (const w of para.split(/\s+/)) {
      const t = line ? line + ' ' + w : w;
      if (line && ctx.measureText(t).width > maxW) { lines.push(line); line = w; } else line = t;
    }
    lines.push(line);
  }
  return lines;
}
const loadImg = src => new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = src; });

async function renderPdf(company, logo) {
  const guide = (localStorage.getItem(LS.guide) || '').trim();
  const p1 = (localStorage.getItem(LS.p1) || '').trim() || '..........';
  const p2 = (localStorage.getItem(LS.p2) || '').trim() || '..........';
  const iso = v => '\u2066' + v + '\u2069';
  const footerText = `جهت کسب اطلاعات بیشتر با ${iso(p1)} و ${iso(p2)} در زمان اداری تماس حاصل فرمایید.`;
  const date = new Date().toLocaleDateString('fa-IR-u-ca-persian', { year: 'numeric', month: '2-digit', day: '2-digit' });

  const pages = []; let ctx, y;
  const limit = () => H - FOOT - 30;
  function addPage(first) {
    const c = document.createElement('canvas'); c.width = W; c.height = H;
    ctx = c.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, W, H);
    pages.push(c);
    if (first) {
      ctx.drawImage(logo, (W - 200) / 2, 70, 200, 200);
      text(ctx, 'سامانه‌های کارگزاران هدایت تراکنش سپ', W / 2, 335, { size: 60, color: BLUE });
      ctx.fillStyle = BLUE; ctx.fillRect(M, 385, CW, 5); ctx.fillStyle = '#7fb0e6'; ctx.fillRect(M, 396, CW, 2);
      y = 440;
    } else {
      ctx.drawImage(logo, W - M - 110, 60, 110, 110);
      text(ctx, 'سامانه‌های کارگزاران هدایت تراکنش سپ', W - M - 140, 115, { align: 'right', size: 42, color: BLUE });
      ctx.fillStyle = BLUE; ctx.fillRect(M, 190, CW, 4);
      y = 240;
    }
  }
  const ensure = h => { if (y + h > limit()) addPage(false); };

  function row(h, cells) {
    ensure(h);
    let x = W - M;
    for (const c of cells) {
      const x0 = x - c.w;
      ctx.fillStyle = c.fill || '#fff'; ctx.fillRect(x0, y, c.w, h);
      ctx.strokeStyle = BLUE; ctx.lineWidth = 2; ctx.strokeRect(x0, y, c.w, h);
      text(ctx, c.t, x0 + c.w / 2, y + h / 2 + 2, { size: c.size || 34, fam: c.fam || UI, dir: c.dir || 'rtl', color: c.color || TXT, maxW: c.w - 36 });
      x = x0;
    }
    y += h;
  }
  function bar(title) {
    ensure(90 + 90);
    ctx.fillStyle = BLUE; ctx.fillRect(M, y, CW, 72);
    text(ctx, title, W / 2, y + 38, { size: 38, color: '#fff' });
    y += 72;
  }

  addPage(true);

  // company info
  row(86, [{ w: 300, t: 'نام شرکت', fill: LIGHT, color: BLUE }, { w: CW - 300, t: company.name, size: 38 }]);
  row(76, [{ w: 300, t: 'تاریخ صدور', fill: LIGHT, color: BLUE }, { w: CW - 300, t: date }]);
  y += 50;

  // credentials
  bar('اطلاعات ورود به سامانه‌ها');
  const cw = [120, 360, (CW - 480) / 2, (CW - 480) / 2];
  row(80, [
    { w: cw[0], t: 'ردیف', fill: LIGHT, color: BLUE }, { w: cw[1], t: 'نام سامانه', fill: LIGHT, color: BLUE },
    { w: cw[2], t: 'نام کاربری', fill: LIGHT, color: BLUE }, { w: cw[3], t: 'رمز عبور', fill: LIGHT, color: BLUE }]);
  if (!company.systems.length) row(90, [{ w: CW, t: 'اطلاعات سامانه‌ای برای این شرکت ثبت نشده است.' }]);
  company.systems.forEach((s, i) => {
    const fill = i % 2 ? ZEBRA : '#fff';
    row(96, [
      { w: cw[0], t: (i + 1).toLocaleString('fa-IR'), fill },
      { w: cw[1], t: s.name, fill, size: 36 },
      { w: cw[2], t: s.user || '—', fill, fam: LAT, dir: 'ltr', size: 34 },
      { w: cw[3], t: s.pass || '—', fill, fam: LAT, dir: 'ltr', size: 34 }]);
  });
  y += 60;

  // guide
  if (guide) {
    bar('راهنمای سامانه‌ها');
    const PAD = 34, LH = 60;
    ctx.direction = 'rtl'; setF(ctx, 32);
    let lines = wrap(ctx, guide, CW - 2 * PAD);
    while (lines.length) {
      let room = limit() - y - 2 * PAD;
      if (room < LH * 2) { addPage(false); room = limit() - y - 2 * PAD; }
      const n = Math.max(1, Math.min(lines.length, Math.floor(room / LH)));
      const chunk = lines.splice(0, n), bh = chunk.length * LH + 2 * PAD;
      ctx.fillStyle = '#fff'; ctx.fillRect(M, y, CW, bh);
      ctx.strokeStyle = BLUE; ctx.lineWidth = 2; ctx.strokeRect(M, y, CW, bh);
      chunk.forEach((l, i) => text(ctx, l, W - M - PAD, y + PAD + LH * i + LH / 2, { align: 'right', size: 32 }));
      y += bh;
      if (lines.length) addPage(false);
    }
  }

  // footer + page numbers
  pages.forEach((c, i) => {
    const g = c.getContext('2d');
    g.fillStyle = BLUE; g.fillRect(M, H - FOOT + 10, CW, 3);
    g.direction = 'rtl'; setF(g, 31);
    const fl = wrap(g, footerText, CW - 80);
    fl.forEach((l, k) => text(g, l, W / 2, H - FOOT + 62 + k * 50, { size: 31, color: '#33475b' }));
    text(g, `صفحه ${(i + 1).toLocaleString('fa-IR')} از ${pages.length.toLocaleString('fa-IR')}`, W / 2, H - 42, { size: 24, color: '#7a8896' });
  });
  return pages;
}

function pdfFromJpegs(imgs) {
  const enc = new TextEncoder(), chunks = [], offs = []; let len = 0;
  const push = d => { const u = typeof d === 'string' ? enc.encode(d) : d; chunks.push(u); len += u.length; };
  const n = imgs.length;
  push('%PDF-1.4\n');
  offs[1] = len; push('1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n');
  offs[2] = len; push(`2 0 obj\n<< /Type /Pages /Kids [${imgs.map((_, i) => `${3 + 3 * i} 0 R`).join(' ')}] /Count ${n} >>\nendobj\n`);
  imgs.forEach((im, i) => {
    const p = 3 + 3 * i, c = p + 1, x = p + 2, cs = `q 595 0 0 842 0 0 cm /Im${i} Do Q`;
    offs[p] = len; push(`${p} 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /XObject << /Im${i} ${x} 0 R >> >> /Contents ${c} 0 R >>\nendobj\n`);
    offs[c] = len; push(`${c} 0 obj\n<< /Length ${cs.length} >>\nstream\n${cs}\nendstream\nendobj\n`);
    offs[x] = len; push(`${x} 0 obj\n<< /Type /XObject /Subtype /Image /Width ${im.w} /Height ${im.h} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${im.bytes.length} >>\nstream\n`);
    push(im.bytes); push('\nendstream\nendobj\n');
  });
  const total = 2 + 3 * n, xref = len;
  let t = `xref\n0 ${total + 1}\n0000000000 65535 f \n`;
  for (let i = 1; i <= total; i++) t += String(offs[i]).padStart(10, '0') + ' 00000 n \n';
  push(t); push(`trailer\n<< /Size ${total + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF`);
  return new Blob(chunks, { type: 'application/pdf' });
}

$('#dlBtn').onclick = async () => {
  const c = state.selected; if (!c) return;
  const btn = $('#dlBtn'), msg = $('#dlMsg');
  btn.disabled = true; msg.className = 'msg'; msg.textContent = 'در حال ساخت PDF...';
  try {
    await document.fonts.load('30px Yekan', 'سامانه');
    const logo = await loadImg('logo.png');
    const canvases = await renderPdf(c, logo);
    const imgs = [];
    for (const cv of canvases) {
      const b = await new Promise(r => cv.toBlob(r, 'image/jpeg', 0.93));
      imgs.push({ bytes: new Uint8Array(await b.arrayBuffer()), w: cv.width, h: cv.height });
    }
    const url = URL.createObjectURL(pdfFromJpegs(imgs));
    const a = document.createElement('a'); a.href = url; a.download = ('سامانه های کارگزاران - ' + c.name).replace(/[\\/:*?"<>|\u200c]/g, ' ').replace(/\s+/g, ' ').trim() + '.pdf';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
    msg.className = 'msg ok'; msg.textContent = 'فایل PDF آماده شد.';
  } catch (err) {
    msg.className = 'msg err'; msg.textContent = 'خطا در ساخت PDF.'; console.error(err);
  }
  btn.disabled = false;
};
})();
