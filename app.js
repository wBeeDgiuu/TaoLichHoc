'use strict';
const $ = s => document.querySelector(s);

// Giờ học theo bảng giờ của trường (lý thuyết/Zoom và thực hành/phòng máy/thí nghiệm)
const LT = ['06:30-07:20','07:20-08:10','08:10-09:00','09:30-10:20','10:20-11:10','11:10-12:00','12:30-13:20','13:20-14:10','14:10-15:00','15:30-16:20','16:20-17:10','17:10-18:00','18:10-19:00','19:00-19:50','19:50-20:40','20:40-21:30'];
const TH = ['07:00-07:50','07:50-08:40','08:40-09:30','09:30-10:20','10:20-11:10','11:10-12:00','12:30-13:20','13:20-14:10','14:10-15:00','15:00-15:50','15:50-16:40','16:40-17:30','18:00-18:50','18:50-19:40','19:40-20:30','20:30-21:20'];
const DS = {2:'Thứ 2',3:'Thứ 3',4:'Thứ 4',5:'Thứ 5',6:'Thứ 6',7:'Thứ 7',8:'Chủ nhật'};
// Phòng thực hành: phòng máy, xưởng, phòng thí nghiệm, sân bãi, GDQP...
const RTH = /máy tính|xưởng|thực hành|cơ điện tử|^P\.|ĐTCS|VĐK|PLC|ĐLCB|^G\d+\(|cắt, vẽ|Mannequi|Sân|Hồ bơi|VOTHUAT|^ST\d/i;

let subjects = [], classes = new Map(), sel = new Map(), open = new Set(), week1 = new Date(2027, 0, 11), nWeeks = 30;

const norm = v => String(v ?? '').replace(/\s+/g, ' ').trim().toLowerCase();
const strip = s => s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').toLowerCase();
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const hue = s => [...s].reduce((a, c) => (a * 31 + c.charCodeAt(0)) % 360, 7);
const fmtWeeks = w => { const r = []; for (let i = 0; i < w.length;) { let j = i; while (w[j + 1] === w[j] + 1) j++; r.push(j > i ? `${w[i]}-${w[j]}` : `${w[i]}`); i = j + 1; } return r.join(', '); };
const timeOf = b => { const T = b.kind === 'TH' ? TH : LT; return T[b.from - 1].slice(0, 5) + ' - ' + T[b.to - 1].slice(6); };
const kindOf = c => [...new Set(c.blocks.map(b => b.kind))].join('+');
const pad = n => String(n).padStart(2, '0');

function parseDate(v) {
  if (typeof v === 'number') return new Date(Math.round((v - 25569) * 864e5) + new Date().getTimezoneOffset() * 6e4);
  const m = String(v).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  return m ? new Date(+m[3], m[2] - 1, +m[1]) : null;
}

// Đọc bảng dữ liệu: mỗi dòng Excel là một buổi học; nhiều dòng cùng Mã LHP = một lớp học phần
function load(rows) {
  const h = rows.findIndex(r => r.some(c => norm(c) === 'mã lhp'));
  if (h < 0) throw new Error('Không tìm thấy cột "Mã LHP". Hãy chọn đúng file thời khóa biểu của trường.');
  const A = rows[h], B = rows[h + 1] || [], col = {}, wk = [];
  for (let i = 0; i < Math.max(A.length, B.length); i++) {
    const t = norm(B[i]) || norm(A[i]);
    if (/^t\d+$/.test(t)) wk.push([i, +t.slice(1)]); else if (t) col[t] = i;
  }
  for (const k of ['mã lhp', 'tên môn học', 'phòng', 'thứ', 'từ tiết', 'đến tiết'])
    if (col[k] === undefined) throw new Error('Thiếu cột "' + k + '" trong file.');
  const cm = new Map(); let w1 = null;
  for (const r of rows.slice(h + 2)) {
    const code = String(r[col['mã lhp']] ?? '').trim();
    if (!/^\d{6,}$/.test(code)) continue;
    const name = String(r[col['tên môn học']]).trim(), room = String(r[col['phòng']]).trim();
    const b = { day: +r[col['thứ']], from: +r[col['từ tiết']], to: +r[col['đến tiết']], room,
      gv: String(r[col['giảng viên']] ?? '').trim(), weeks: wk.filter(([i]) => r[i]).map(x => x[1]) };
    if (!(b.day >= 2 && b.day <= 8) || !(b.from >= 1 && b.to <= 16 && b.from <= b.to) || !b.weeks.length) continue;
    b.kind = /lý thuyết/i.test(room) ? 'LT' : (RTH.test(room) || /^(thực hành|thí nghiệm)/i.test(name)) ? 'TH' : 'LT';
    if (!w1) { const d = parseDate(r[col['bắt đầu']]); if (d) w1 = new Date(d.getFullYear(), d.getMonth(), d.getDate() - (b.day - 2) - (b.weeks[0] - 1) * 7); }
    let c = cm.get(code);
    if (!c) cm.set(code, c = { code, name, cls: String(r[col['lớp học']] ?? '').trim(), tc: +r[col['số tc']] || 0, el: !!norm(r[col['tính chất học phần']]), blocks: [] });
    c.blocks.push(b);
  }
  if (!cm.size) throw new Error('File không có dòng dữ liệu hợp lệ.');
  if (w1) week1 = w1;
  setData(cm, wk.length || 30);
}

function setData(cm, n) {
  nWeeks = n; classes = cm;
  const sm = new Map();
  for (const c of cm.values()) (sm.get(c.name) || sm.set(c.name, []).get(c.name)).push(c);
  subjects = [...sm].map(([name, cl]) => ({ name, cl, hay: strip(name + ' ' + cl.map(c => c.code + ' ' + c.cls + ' ' + c.blocks.map(b => b.gv).join(' ')).join(' ')) }))
    .sort((a, b) => a.name.localeCompare(b.name, 'vi'));
}

// ---- Lưu trữ: dữ liệu file Excel đã nhập nằm trong IndexedDB, nên tắt trang mở lại vẫn còn ----
const DB = {
  db: null,
  open() { return new Promise(res => { try { const r = indexedDB.open('huit_lich', 1); r.onupgradeneeded = () => r.result.createObjectStore('kv'); r.onsuccess = () => { DB.db = r.result; res(); }; r.onerror = () => res(); } catch (e) { res(); } }); },
  get(k) { return new Promise(res => { if (!DB.db) return res(); try { const q = DB.db.transaction('kv').objectStore('kv').get(k); q.onsuccess = () => res(q.result); q.onerror = () => res(); } catch (e) { res(); } }); },
  set(k, v) { return new Promise(res => { if (!DB.db) return res(); try { const t = DB.db.transaction('kv', 'readwrite'); t.objectStore('kv').put(v, k); t.oncomplete = t.onerror = t.onabort = () => res(); } catch (e) { res(); } }); }
};
let meta = { fileName: '', savedAt: 0 };
const saveData = () => DB.set('data', { classes: [...classes.values()], week1: week1.getTime(), nWeeks, meta });
const saveState = () => DB.set('state', { sel: [...sel.values()].map(c => c.code), skip: [...skip], ovr: [...ovr], pLT: $('#pLT').value, pTH: $('#pTH').value });
const fmtDT = t => { const d = new Date(t); return `${pad(d.getHours())}:${pad(d.getMinutes())} ${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()}`; };
const infoText = () => `${subjects.length} môn · ${classes.size} lớp học phần` + (meta.savedAt ? `\nFile: ${meta.fileName || '(không tên)'} · nhập lúc ${fmtDT(meta.savedAt)}` : '');

// Hai buổi học trùng khi cùng thứ, giao nhau về tiết và có ít nhất một tuần chung
function conflict(c) {
  for (const s of sel.values()) {
    if (s.name === c.name) continue; // chọn lớp khác của cùng môn = thay thế lớp cũ
    for (const a of c.blocks) for (const b of s.blocks)
      if (a.day === b.day && a.from <= b.to && b.from <= a.to) {
        const common = a.weeks.filter(w => b.weeks.includes(w));
        if (common.length) return { with: s, a, b, common };
      }
  }
  return null;
}

function toast(html, ok) {
  const t = $('#toast'); t.className = ok ? 'ok' : ''; t.innerHTML = html; t.hidden = false;
  clearTimeout(toast.t); toast.t = setTimeout(() => t.hidden = true, 9000);
}

function toggle(code) {
  const c = classes.get(code);
  if (sel.get(c.name) === c) sel.delete(c.name);
  else {
    const x = conflict(c);
    if (x) {
      toast(`<b>Trùng lịch, không thêm được.</b><br>${esc(c.name)} (${c.code}) trùng với ${esc(x.with.name)} (${x.with.code}): ${DS[x.a.day]}, tiết ${x.a.from}-${x.a.to} và tiết ${x.b.from}-${x.b.to}, tuần ${fmtWeeks(x.common)}.`);
      return;
    }
    sel.set(c.name, c);
  }
  refresh(); saveState();
}

function classHtml(c) {
  const on = sel.get(c.name) === c, bad = !on && conflict(c), k = kindOf(c);
  return `<div class="c${on ? ' on' : ''}${bad ? ' bad' : ''}" data-code="${c.code}"><div class="ch"><code>${c.code}</code><span>${esc(c.cls)}</span><i class="tag ${k}">${k}</i>${c.el ? '<i class="tag el">E-learning</i>' : ''}${on ? '<em>Đã chọn</em>' : bad ? `<em class="r">Trùng ${esc(bad.with.name)}</em>` : ''}</div>`
    + c.blocks.map(b => `<p>${DS[b.day]} · tiết ${b.from}-${b.to} · ${timeOf(b)} · Tuần ${fmtWeeks(b.weeks)} · ${esc(b.room)}<br><span>GV: ${esc(b.gv || 'chưa có')}</span></p>`).join('') + '</div>';
}

function fillSub(sec) {
  sec.querySelector('.cl').innerHTML = subjects[sec.dataset.i].cl.map(classHtml).join('');
}

function renderList() {
  const list = $('#list'), top = list.scrollTop, terms = strip($('#q').value.trim()).split(/\s+/).filter(Boolean);
  const vis = subjects.map((s, i) => [s, i]).filter(([s]) => terms.every(t => s.hay.includes(t)));
  list.innerHTML = vis.map(([s, i]) => `<section class="sub${sel.has(s.name) ? ' has' : ''}" data-i="${i}"><button class="sh"><b>${esc(s.name)}</b><small>${s.cl.length} lớp${sel.has(s.name) ? ' · đã chọn' : ''}</small></button><div class="cl"${open.has(i) || vis.length <= 3 ? '' : ' hidden'}></div></section>`).join('')
    || '<p class="empty">Không có môn nào khớp.</p>';
  list.querySelectorAll('.cl:not([hidden])').forEach(el => fillSub(el.parentNode));
  list.scrollTop = top;
  $('#info').innerHTML = esc(vis.length === subjects.length ? infoText() : `${vis.length}/${subjects.length} môn khớp · ${infoText()}`).replace(/\n/g, '<br>');
}

function renderTT() {
  const w = +$('#week').value, items = [];
  for (const c of sel.values()) { // gộp các buổi cùng thứ + cùng tiết của một lớp (khác tuần/khác phòng) thành một ô
    const m = new Map();
    for (const b of c.blocks) if (!w || b.weeks.includes(w)) {
      const k = b.day + '-' + b.from + '-' + b.to;
      let x = m.get(k); if (!x) { x = { c, b, parts: [] }; m.set(k, x); items.push(x); }
      x.parts.push(b);
    }
  }
  const td = todayCol(), f2 = d => pad(d.getDate()) + '/' + pad(d.getMonth() + 1);
  let h = '<div class="hd"></div>' + [2,3,4,5,6,7,8].map(d => `<div class="hd${d === td && (!w || w === curWeek()) ? ' today' : ''}">${DS[d]}${w ? `<small>${f2(new Date(week1.getTime() + ((w - 1) * 7 + d - 2) * 864e5))}</small>` : ''}</div>`).join('')
    + '<div class="tl">' + Array.from({length: 16}, (_, i) => `<div class="${i < 6 ? 's' : i < 12 ? 'c2' : 'e'}">${i + 1}</div>`).join('') + '</div>';
  for (let d = 2; d <= 8; d++) {
    const it = items.filter(x => x.b.day === d).sort((x, y) => x.b.from - y.b.from || y.b.to - x.b.to), lanes = [];
    it.forEach(x => { let k = lanes.findIndex(e => e < x.b.from); if (k < 0) { k = lanes.length; lanes.push(0); } lanes[k] = x.b.to; x.k = k; });
    h += `<div class="dc${d === td && (!w || w === curWeek()) ? ' today' : ''}" data-d="${d}">` + it.map(x => {
      const n = 1 + Math.max(...it.filter(y => y.b.from <= x.b.to && x.b.from <= y.b.to).map(y => y.k)), b = x.b;
      return `<div class="b" data-name="${esc(x.c.name)}" style="--hue:${hue(x.c.name)};top:calc(${b.from - 1}*var(--h));min-height:calc(${b.to - b.from + 1}*var(--h) - 2px);left:${x.k / n * 100}%;width:calc(${100 / n}% - 2px)" title="${esc(x.c.name)}">`
        + `<button class="x" title="Bỏ môn này">×</button><b>${esc(x.c.name)}</b><span>${timeOf(b)}</span>${x.parts.map(p => `<span>${w ? '' : 'Tuần ' + fmtWeeks(p.weeks) + ' · '}${esc(p.room.split(' - ')[0])}</span>`).join('')}<span>${x.c.code}</span><span>${esc([...new Set(x.parts.map(p => p.gv.replace(/^[^-]*-\s*/, '')))].filter(Boolean).join(', '))}</span></div>`;
    }).join('') + '</div>';
  }
  $('#tt').className = 'tt'; $('#tt').innerHTML = h; updateNow();
  const tc = [...sel.values()].reduce((a, c) => a + c.tc, 0);
  $('#sum').textContent = sel.size ? `Đã chọn ${sel.size} môn · ${tc} tín chỉ` : 'Chưa chọn môn nào. Bấm vào một lớp ở danh sách bên trái.';
}

const day0 = d => new Date(d.getFullYear(), d.getMonth(), d.getDate());
const todayCol = () => { const g = new Date().getDay(); return g === 0 ? 8 : g + 1; }; // 2..8 (8 = Chủ nhật)
function curWeek() { const n = Math.floor(Math.round((day0(new Date()) - week1) / 864e5) / 7) + 1; return n >= 1 && n <= nWeeks ? n : 0; }
function fillWeeks(pick) {
  const f = d => pad(d.getDate()) + '/' + pad(d.getMonth() + 1), cw = curWeek();
  $('#week').innerHTML = '<option value="0">Tất cả các tuần</option>' + Array.from({length: nWeeks}, (_, i) => {
    const s = new Date(week1.getTime() + i * 7 * 864e5), e = new Date(s.getTime() + 6 * 864e5);
    return `<option value="${i + 1}">Tuần ${i + 1} (${f(s)} - ${f(e)})${i + 1 === cw ? ' · hiện tại' : ''}</option>`;
  }).join('');
  $('#week').value = String(pick ?? cw); // mặc định: tuần hiện tại theo ngày thật
}

// Đồng hồ thật + vạch giờ hiện tại trên thời khóa biểu
const THU = ['Chủ nhật', 'Thứ 2', 'Thứ 3', 'Thứ 4', 'Thứ 5', 'Thứ 6', 'Thứ 7'];
function updateNow() {
  const n = new Date(), mins = n.getHours() * 60 + n.getMinutes();
  document.querySelectorAll('.now').forEach(e => e.remove());
  const col = document.querySelector('.dc.today'); if (!col) return;
  const m = t => { const [a, b] = t.split('-').map(x => +x.slice(0, 2) * 60 + +x.slice(3)); return [a, b]; };
  let pos = -1;
  for (let i = 0; i < 16 && pos < 0; i++) { const [a, b] = m(LT[i]); if (mins < a) pos = i; else if (mins < b) pos = i + (mins - a) / (b - a); }
  if (pos < 0) return;
  const l = document.createElement('div'); l.className = 'now'; l.style.top = `calc(${pos}*var(--h))`; l.title = 'Bây giờ'; col.appendChild(l);
}
let lastDay = '';
function tick() {
  const n = new Date();
  $('#clock').textContent = `${THU[n.getDay()]}, ${pad(n.getDate())}/${pad(n.getMonth() + 1)}/${n.getFullYear()} · ${pad(n.getHours())}:${pad(n.getMinutes())}:${pad(n.getSeconds())}`;
  const key = n.toDateString();
  if (lastDay && key !== lastDay) { const w = +$('#week').value; fillWeeks(w === 0 ? 0 : undefined); renderTT(); } // sang ngày mới: cập nhật tuần/cột hôm nay
  else if (n.getSeconds() === 0 || !lastDay) updateNow();
  lastDay = key;
}

// Bấm vào ô môn trên lịch: danh sách bên trái nhảy tới môn đó, mở sẵn các lớp
function focusSubject(name) {
  const i = subjects.findIndex(s => s.name === name); if (i < 0) return;
  const find = () => $('#list').querySelector(`.sub[data-i="${i}"]`);
  if (!find()) { $('#q').value = ''; }
  open.add(i); renderList();
  const sec = find(); if (!sec) return;
  const list = $('#list');
  list.scrollTo({ top: list.scrollTop + sec.getBoundingClientRect().top - list.getBoundingClientRect().top - 4, behavior: 'smooth' });
  if (innerWidth <= 860) $('#side').scrollIntoView({ behavior: 'smooth', block: 'start' }); // màn hình hẹp: danh sách nằm phía trên
  const t = sec.querySelector('.c.on') || sec; t.classList.add('flash'); setTimeout(() => t.classList.remove('flash'), 1800);
}

function refresh() { renderList(); renderTT(); }

$('#file').onchange = e => {
  const f = e.target.files[0]; if (!f) return;
  $('#info').textContent = 'Đang đọc file, vui lòng chờ vài giây...';
  setTimeout(async () => {
    try {
      const wb = XLSX.read(await f.arrayBuffer(), { type: 'array' });
      load(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' }));
      sel.clear(); open.clear(); skip.clear(); ovr.clear(); $('#q').disabled = false; $('#q').value = '';
      meta = { fileName: f.name, savedAt: Date.now() };
      fillWeeks(); refresh(); saveData(); saveState(); toast(`Đã nhập ${subjects.length} môn, ${classes.size} lớp học phần.`, true);
    } catch (err) { $('#info').textContent = 'Không đọc được file.'; toast(esc(err.message)); }
  }, 30);
  e.target.value = '';
};
let qt; $('#q').oninput = () => { clearTimeout(qt); qt = setTimeout(renderList, 200); };
$('#week').onchange = () => { renderTT(); };
$('#clear').onclick = () => { sel.clear(); refresh(); saveState(); };
$('#toast').onclick = () => $('#toast').hidden = true;
$('#list').onclick = e => {
  const sec = e.target.closest('.sub'); if (!sec) return;
  const c = e.target.closest('.c');
  if (c) return toggle(c.dataset.code);
  if (e.target.closest('.sh')) {
    const el = sec.querySelector('.cl');
    if (el.hidden) { fillSub(sec); el.hidden = false; open.add(+sec.dataset.i); } else { el.hidden = true; open.delete(+sec.dataset.i); }
  }
};
$('#tt').onclick = e => {
  const b = e.target.closest('.b'); if (!b) return;
  if (e.target.closest('.x')) { sel.delete(b.dataset.name); refresh(); saveState(); } else focusSubject(b.dataset.name);
};

// Dự toán học phí: buổi học 3 tiết = lý thuyết (863.000 đ/TC), còn lại = thực hành (1.100.000 đ/TC)
const money = n => n.toLocaleString('vi-VN') + ' đ';
const skip = new Set(), ovr = new Map();
function feeType(c) { // theo độ dài buổi học phổ biến nhất của lớp, có thể đổi tay
  if (ovr.has(c.name)) return ovr.get(c.name);
  const cnt = {}; c.blocks.forEach(b => { const n = b.to - b.from + 1; cnt[n] = (cnt[n] || 0) + 1; });
  const n = +Object.keys(cnt).sort((a, b) => cnt[b] - cnt[a] || a - b)[0];
  return n === 3 ? 'LT' : 'TH';
}
function renderFee() {
  const p = { LT: +$('#pLT').value || 0, TH: +$('#pTH').value || 0 }, tcs = { LT: 0, TH: 0 }, sub = { LT: 0, TH: 0 }; let sum = 0;
  const rows = [...sel.values()].map(c => {
    const on = !skip.has(c.name), t = feeType(c), m = c.tc * p[t];
    if (on) { tcs[t] += c.tc; sub[t] += m; sum += m; }
    return `<tr><td><input type="checkbox" data-n="${esc(c.name)}"${on ? ' checked' : ''}></td><td>${esc(c.name)}<br><small>${c.code} · ${c.tc} TC × ${money(p[t])}</small></td>`
      + `<td><select data-t="${esc(c.name)}"><option value="LT"${t === 'LT' ? ' selected' : ''}>Lý thuyết</option><option value="TH"${t === 'TH' ? ' selected' : ''}>Thực hành</option></select></td><td>${c.tc}</td><td>${money(on ? m : 0)}</td></tr>`;
  }).join('');
  $('#feebody').innerHTML = rows || '<tr><td colspan="5">Chưa chọn môn nào.</td></tr>';
  $('#feetotal').innerHTML = `<span>Lý thuyết ${tcs.LT} TC = ${money(sub.LT)} · Thực hành ${tcs.TH} TC = ${money(sub.TH)}</span>Tổng cộng (${tcs.LT + tcs.TH} TC): ${money(sum)}`;
}
$('#fee').onclick = () => { renderFee(); $('#dlg').showModal(); };
for (const k of ['LT', 'TH']) $('#p' + k).oninput = () => { renderFee(); saveState(); };
$('#feebody').onchange = e => {
  const t = e.target;
  if (t.dataset.n !== undefined) t.checked ? skip.delete(t.dataset.n) : skip.add(t.dataset.n);
  else if (t.dataset.t !== undefined) ovr.set(t.dataset.t, t.value);
  renderFee(); saveState();
};

// Giao diện sáng / tối
function setTheme(t) {
  document.documentElement.dataset.theme = t; $('#theme').textContent = t === 'dark' ? '☀️' : '🌙';
  try { localStorage.setItem('huit_theme', t); } catch (e) {}
}
$('#theme').onclick = () => setTheme(document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark');
setTheme(document.documentElement.dataset.theme === 'dark' ? 'dark' : 'light');

// Khởi động: khôi phục dữ liệu + lịch đã chọn từ lần trước
(async function init() {
  fillWeeks(); renderTT(); tick(); setInterval(tick, 1000);
  await DB.open();
  const d = await DB.get('data');
  if (d && d.classes && d.classes.length) {
    week1 = new Date(d.week1); meta = d.meta || meta;
    setData(new Map(d.classes.map(c => [c.code, c])), d.nWeeks);
    const st = (await DB.get('state')) || {};
    (st.skip || []).forEach(x => skip.add(x)); (st.ovr || []).forEach(([k, v]) => ovr.set(k, v));
    if (st.pLT) $('#pLT').value = st.pLT; if (st.pTH) $('#pTH').value = st.pTH;
    (st.sel || []).forEach(code => { const c = classes.get(code); if (c) sel.set(c.name, c); });
    $('#q').disabled = false; fillWeeks(); refresh();
  }
})();