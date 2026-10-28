'use strict';
/* =====================================================================
   Habitude — daily habits, tasks, coins & streaks.
   Your data lives in this browser (localStorage) and, once connected,
   in your own Google Sheet through your own Apps Script web app.
   Nothing is sent anywhere else (except the optional Claude API call
   when you generate a skin with your own key).
   ===================================================================== */

const STORE = 'habitude.state.v1';
const LOCAL = 'habitude.local.v1';
const AI_MODEL = 'claude-sonnet-5-5';          // used only for optional AI skin generation
const XLSX_URL = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';

/* ---------------- small utilities ---------------- */
const $ = (s, r = document) => r.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const uid = p => `${p}_${Date.now().toString(36)}${Math.random().toString(36).slice(2, 7)}`;
const round2 = n => Math.round((+n || 0) * 100) / 100;
const fmtN = n => { const r = round2(n); return Number.isInteger(r) ? r.toLocaleString() : r.toLocaleString(undefined, { maximumFractionDigits: 2 }); };
const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
const pad = n => String(n).padStart(2, '0');
const toStr = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const parseD = s => { const [y, m, d] = String(s).split('-').map(Number); return new Date(y, (m || 1) - 1, d || 1); };
const isDateStr = s => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ''));
const todayStr = () => toStr(new Date());
const addDays = (s, n) => { const d = parseD(s); d.setDate(d.getDate() + n); return toStr(d); };
const diffDays = (a, b) => Math.round((parseD(b) - parseD(a)) / 864e5);
const weekStart = s => { const d = parseD(s); d.setDate(d.getDate() - ((d.getDay() + 6) % 7)); return toStr(d); };
const fmtDate = (s, o = { weekday: 'long', month: 'long', day: 'numeric' }) => parseD(s).toLocaleDateString(undefined, o);
const shortD = s => parseD(s).toLocaleDateString(undefined, { month: 'short', day: 'numeric' });
const slug = s => String(s || 'file').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'file';
const PRI = { 3: 'Must', 2: 'Should', 1: 'Optional' };
const PRI_CLS = { 3: 'must', 2: 'should', 1: '' };
const TPRI = { 3: 'High', 2: 'Medium', 1: 'Low' };

const IC = {
  left: '<svg viewBox="0 0 24 24"><path d="M15 18l-6-6 6-6"/></svg>',
  right: '<svg viewBox="0 0 24 24"><path d="M9 18l6-6-6-6"/></svg>',
  check: '<svg viewBox="0 0 24 24"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg>',
  plus: '<svg viewBox="0 0 24 24"><path d="M12 5v14M5 12h14"/></svg>',
  up: '<svg viewBox="0 0 24 24"><path d="M6 15l6-6 6 6"/></svg>',
  down: '<svg viewBox="0 0 24 24"><path d="M6 9l6 6 6-6"/></svg>',
  edit: '<svg viewBox="0 0 24 24"><path d="M4 20h4L19 9l-4-4L4 16z"/></svg>',
  x: '<svg viewBox="0 0 24 24"><path d="M6 6l12 12M18 6L6 18"/></svg>',
  dl: '<svg viewBox="0 0 24 24"><path d="M12 4v11M7 10l5 5 5-5M5 20h14"/></svg>',
  ul: '<svg viewBox="0 0 24 24"><path d="M12 15V4M7 9l5-5 5 5M5 20h14"/></svg>',
  flame: '<svg viewBox="0 0 24 24"><path d="M12 22c4 0 7-2.7 7-6.8 0-3.6-2.5-6.2-4.4-8.4-.5 2-1.6 3.3-2.6 3.8C12 7 10.6 4.4 9 2c-.3 3-3.9 6.1-3.9 10.6C5 19.3 8 22 12 22z"/></svg>',
  brush: '<svg viewBox="0 0 24 24"><path d="M14 4l6 6-8.5 8.5a3 3 0 0 1-4.2 0l-1.8-1.8a3 3 0 0 1 0-4.2z"/><path d="M4 20l3-3"/></svg>',
  sparkle: '<svg viewBox="0 0 24 24"><path d="M12 3l1.8 5.2L19 10l-5.2 1.8L12 17l-1.8-5.2L5 10l5.2-1.8z"/></svg>',
};
const coinEl = '<span class="coin" aria-hidden="true"></span>';

/* ---------------- color helpers ---------------- */
const HEX = /^#[0-9a-f]{6}$/i;
const hexToRgb = h => { h = h.replace('#', ''); return [0, 2, 4].map(i => parseInt(h.slice(i, i + 2), 16)); };
const rgbToHex = (r, g, b) => '#' + [r, g, b].map(v => clamp(Math.round(v), 0, 255).toString(16).padStart(2, '0')).join('');
function rgbToHsl(r, g, b) {
  r /= 255; g /= 255; b /= 255;
  const mx = Math.max(r, g, b), mn = Math.min(r, g, b); let h = 0, s = 0; const l = (mx + mn) / 2;
  if (mx !== mn) { const d = mx - mn; s = l > .5 ? d / (2 - mx - mn) : d / (mx + mn);
    h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4; h *= 60; }
  return [h, s * 100, l * 100];
}
function hsl(h, s, l) {
  h = ((h % 360) + 360) % 360; s = clamp(s, 0, 100) / 100; l = clamp(l, 0, 100) / 100;
  const k = n => (n + h / 30) % 12, a = s * Math.min(l, 1 - l);
  const f = n => l - a * Math.max(-1, Math.min(k(n) - 3, Math.min(9 - k(n), 1)));
  return rgbToHex(f(0) * 255, f(8) * 255, f(4) * 255);
}
const hexHsl = h => rgbToHsl(...hexToRgb(h));
const lum = hex => { const c = hexToRgb(hex).map(v => { v /= 255; return v <= .03928 ? v / 12.92 : ((v + .055) / 1.055) ** 2.4; }); return .2126 * c[0] + .7152 * c[1] + .0722 * c[2]; };
const contrast = (a, b) => { const x = lum(a), y = lum(b); return (Math.max(x, y) + .05) / (Math.min(x, y) + .05); };
const inkOn = bg => contrast(bg, '#ffffff') >= 3.2 ? '#ffffff' : '#18140f';
function ensureContrast(fg, bg, min) {
  if (contrast(fg, bg) >= min) return fg;
  const [h, s, l] = hexHsl(fg); const dir = lum(bg) > .35 ? -1 : 1;
  for (let i = 1; i <= 50; i++) { const c = hsl(h, s, l + dir * i * 2); if (contrast(c, bg) >= min) return c; }
  return dir < 0 ? '#111111' : '#f5f5f5';
}
const hueDist = (a, b) => { const d = Math.abs(a - b) % 360; return d > 180 ? 360 - d : d; };

/* ---------------- fonts & skins ---------------- */
const FONTS = { 'Fraunces': 'serif', 'DM Sans': 'sans-serif', 'Cormorant Garamond': 'serif', 'Lora': 'serif', 'Playfair Display': 'serif',
  'Inter': 'sans-serif', 'Space Grotesk': 'sans-serif', 'Work Sans': 'sans-serif', 'Nunito': 'sans-serif', 'Fredoka': 'sans-serif',
  'Caveat': 'cursive', 'IBM Plex Mono': 'monospace' };
const loadedFonts = new Set(['Fraunces', 'DM Sans']);
function loadFont(f) {
  if (!FONTS[f] || loadedFonts.has(f)) return; loadedFonts.add(f);
  const l = document.createElement('link'); l.rel = 'stylesheet';
  l.href = `https://fonts.googleapis.com/css2?family=${f.replace(/ /g, '+')}:wght@400;500;600;700&display=swap`;
  document.head.appendChild(l);
}
const fontStack = f => `'${f}', ${FONTS[f] || 'sans-serif'}`;
const SKIN_KEYS = ['bg', 'surface', 'surface2', 'text', 'muted', 'accent', 'coin', 'border'];
const SKIN_LABELS = { bg: 'Background', surface: 'Cards', surface2: 'Subtle fill', text: 'Text', muted: 'Secondary text', accent: 'Accent', coin: 'Coins', border: 'Borders' };

const BUILTIN_SKINS = [
  { id: 'sk_parchment', name: 'Parchment', builtIn: true, mode: 'light', display: 'Fraunces', body: 'DM Sans', radius: 14, bgImage: '', bgOpacity: 0,
    vars: { bg: '#f3ede2', surface: '#fffaf1', surface2: '#ebe2d2', text: '#2a241d', muted: '#76695a', accent: '#a8472a', coin: '#c08a26', border: '#dccfb9' } },
  { id: 'sk_nightgarden', name: 'Night Garden', builtIn: true, mode: 'dark', display: 'Fraunces', body: 'DM Sans', radius: 16, bgImage: '', bgOpacity: 0,
    vars: { bg: '#11150e', surface: '#1a2016', surface2: '#252d1f', text: '#e9eadf', muted: '#9ba48e', accent: '#a3d16f', coin: '#e3b549', border: '#2f3828' } },
  { id: 'sk_moontide', name: 'Moon Tide', builtIn: true, mode: 'dark', display: 'Cormorant Garamond', body: 'Nunito', radius: 20, bgImage: '', bgOpacity: 0,
    vars: { bg: '#18141f', surface: '#221d2c', surface2: '#2d2639', text: '#efe8f6', muted: '#a99fbb', accent: '#e08fb8', coin: '#e9c46a', border: '#3a3249' } },
  { id: 'sk_cleanslate', name: 'Clean Slate', builtIn: true, mode: 'light', display: 'Space Grotesk', body: 'Inter', radius: 10, bgImage: '', bgOpacity: 0,
    vars: { bg: '#f4f5f7', surface: '#ffffff', surface2: '#eceef2', text: '#1b1f27', muted: '#646b7a', accent: '#3b5bdb', coin: '#d99a16', border: '#dde1e8' } },
];

/* ---------------- state ---------------- */
function defaultState() {
  return { meta: { version: 1, updatedAt: 0 },
    settings: { weeklyStep: 0.25, maxMult: 3, defaultSkinId: 'sk_parchment' },
    categories: [], habits: [], templates: [], skins: [], taskGroups: [], tasks: [], days: {} };
}
function normalize(s) {
  const d = defaultState(); s = s && typeof s === 'object' ? s : {};
  const out = { ...d, ...s, meta: { ...d.meta, ...(s.meta || {}) }, settings: { ...d.settings, ...(s.settings || {}) } };
  for (const k of ['categories', 'habits', 'templates', 'skins', 'taskGroups', 'tasks']) out[k] = Array.isArray(s[k]) ? s[k] : [];
  out.days = s.days && typeof s.days === 'object' ? s.days : {};
  out.templates.forEach(t => { t.habitIds = Array.isArray(t.habitIds) ? t.habitIds : []; if (t.showTasks === undefined) t.showTasks = true; });
  out.habits.forEach(h => { h.coins = clamp(+h.coins || 1, 1, 5); h.priority = clamp(+h.priority || 2, 1, 3); });
  out.tasks.forEach(t => { t.priority = clamp(+t.priority || 2, 1, 3); t.status = t.status === 'closed' ? 'closed' : 'open'; });
  out.skins = out.skins.map(sanitizeSkin).filter(Boolean);
  for (const [k, v] of Object.entries(out.days)) { if (!isDateStr(k) || !v) { delete out.days[k]; continue; } v.habits = v.habits || {}; v.tasks = Array.isArray(v.tasks) ? v.tasks : []; }
  return out;
}
let S = (() => { try { const raw = localStorage.getItem(STORE); return raw ? normalize(JSON.parse(raw)) : null; } catch { return null; } })();
const FIRST_RUN = !S;
if (!S) S = defaultState();
let L = (() => { try { return JSON.parse(localStorage.getItem(LOCAL)) || {}; } catch { return {}; } })();
L = { scriptUrl: '', appKey: '', claudeKey: '', lastSync: 0, ...L };
const V = { page: 'day', date: todayStr(), taskStatus: 'open', taskSort: 'priority', taskGroup: 'all',
  q: { metric: 'coins', group: 'day', range: '30', from: '', to: '', category: 'all', template: 'all' } };

function saveLocal() { try { localStorage.setItem(LOCAL, JSON.stringify(L)); } catch {} }
function saveState() {
  try { localStorage.setItem(STORE, JSON.stringify(S)); }
  catch { toast('Browser storage is full — try removing a skin background image.'); }
}
function commit(fn, { sync = true } = {}) {
  fn(S); S.meta.updatedAt = Date.now(); saveState(); if (sync) scheduleSync(); queueRender();
}

/* lookups */
const tpl = id => S.templates.find(t => t.id === id);
const habitById = id => S.habits.find(h => h.id === id);
const catById = id => S.categories.find(c => c.id === id);
const taskById = id => S.tasks.find(t => t.id === id);
const groupById = id => S.taskGroups.find(g => g.id === id);
const allSkins = () => [...BUILTIN_SKINS, ...S.skins];
const getSkin = id => allSkins().find(s => s.id === id);

function dayHabitIds(date) {
  const d = S.days[date]; if (!d || !d.templateId) return [];
  const ids = Array.isArray(d.habitIds) ? d.habitIds : (tpl(d.templateId)?.habitIds || []);
  return ids.filter(id => habitById(id));
}
/* when a template changes, refresh the habit snapshot for today and future days that use it (history stays as it was) */
function refreshSnapshots(s, tid) {
  const t = s.templates.find(x => x.id === tid); const today = todayStr();
  for (const [date, d] of Object.entries(s.days)) if (d.templateId === tid && date >= today) d.habitIds = t ? t.habitIds.slice() : [];
}

/* ---------------- scoring engine ----------------
   - Each completed habit earns coins (1–5) × its streak multiplier.
   - Multiplier = 1 + weeklyStep × (full weeks the streak has run), capped at maxMult.
   - A streak only looks at days whose template includes that habit; other days are invisible to it.
   - Missed on a past day:  Optional (P1) → nothing happens.
                            Should (P2)   → streak resets.
                            Must (P3)     → streak resets and its coin value is deducted.
   - Today is never counted as missed until the day is over.
   - Tasks completed on the Day page earn the coins assigned when they were added. */
let _sc = null, _scKey = '';
function scores() {
  const k = S.meta.updatedAt + '|' + todayStr();
  if (k !== _scKey) { _scKey = k; _sc = computeScores(); }
  return _sc;
}
function computeScores() {
  const today = todayStr(); const step = +S.settings.weeklyStep || 0.25; const max = +S.settings.maxMult || 3;
  const dates = Object.keys(S.days).filter(d => d <= today).sort();
  const perDay = {}, habit = {};
  S.habits.forEach(h => habit[h.id] = { streak: 0, start: null, best: 0, done: 0, eligible: 0, earned: 0, penalty: 0, mult: 1 });
  const multFor = (ph, date) => ph.streak > 0 ? Math.min(max, 1 + step * Math.floor((diffDays(ph.start, date) + 1) / 7)) : 1;
  for (const date of dates) {
    const d = S.days[date]; const ids = dayHabitIds(date);
    const pd = { habitCoins: 0, penalty: 0, taskCoins: 0, done: 0, scheduled: ids.length, eligible: 0, tasksDone: 0, habits: {} };
    for (const id of ids) {
      const h = habitById(id), ph = habit[id]; const done = !!d.habits[id];
      if (done) {
        if (ph.streak === 0) ph.start = date;
        ph.streak++; ph.best = Math.max(ph.best, ph.streak);
        const mult = Math.min(max, 1 + step * Math.floor((diffDays(ph.start, date) + 1) / 7));
        const earned = round2(h.coins * mult);
        ph.mult = mult; ph.earned += earned; ph.done++; ph.eligible++;
        pd.habitCoins += earned; pd.done++; pd.eligible++;
        pd.habits[id] = { done: true, earned, mult, streak: ph.streak };
      } else if (date < today) {
        ph.eligible++; pd.eligible++; let pen = 0;
        if (h.priority >= 2) { ph.streak = 0; ph.start = null; ph.mult = 1; }
        if (h.priority === 3) { pen = h.coins; ph.penalty += pen; pd.penalty += pen; }
        pd.habits[id] = { done: false, missed: true, earned: 0, penalty: pen, streak: ph.streak };
      } else {
        pd.habits[id] = { done: false, pending: true, earned: 0, streak: ph.streak, nextMult: ph.streak > 0 ? Math.min(max, 1 + step * Math.floor((diffDays(ph.start, date) + 1) / 7)) : 1 };
      }
    }
    for (const t of d.tasks) if (t.done) { pd.taskCoins += +t.coins || 0; pd.tasksDone++; }
    pd.habitCoins = round2(pd.habitCoins);
    pd.total = round2(pd.habitCoins + pd.taskCoins - pd.penalty);
    perDay[date] = pd;
  }
  let total = 0; for (const p of Object.values(perDay)) total += p.total;
  // current multiplier for display (as of today)
  for (const ph of Object.values(habit)) ph.multNow = multFor(ph, today);
  return { perDay, habit, total: round2(total) };
}

/* ---------------- rendering core ---------------- */
let renderQueued = false, modalFn = null, lastSkinKey = '';
function queueRender() { if (renderQueued) return; renderQueued = true; setTimeout(() => { renderQueued = false; render(); }, 0); }
function render() {
  const ae = document.activeElement; const fk = ae?.dataset?.fk;
  let sel = null; try { sel = ae && 'selectionStart' in ae ? ae.selectionStart : null; } catch {}
  applySkin(activeSkin());
  $('#main').innerHTML = (PAGES[V.page] || PAGES.day)();
  document.querySelectorAll('.tabbar button').forEach(b => b.classList.toggle('on', b.dataset.page === V.page));
  if (modalFn) $('#modalBody').innerHTML = modalFn();
  if (fk) { const el = document.querySelector(`[data-fk="${CSS.escape(fk)}"]`); if (el) { el.focus({ preventScroll: true }); try { if (sel != null) el.setSelectionRange(sel, sel); } catch {} } }
  updateSyncUI();
}
function openModal(fn) {
  modalFn = fn; const m = $('#modal'); $('#modalBody').innerHTML = fn(); m.hidden = false;
  document.body.classList.add('modal-open'); requestAnimationFrame(() => m.classList.add('show'));
  $('.sheet', m).scrollTop = 0;
}
function closeModal() {
  modalFn = null; skinDraft = null; const m = $('#modal'); m.classList.remove('show');
  document.body.classList.remove('modal-open'); setTimeout(() => { if (!modalFn) { m.hidden = true; $('#modalBody').innerHTML = ''; } }, 200);
}
const rerenderModal = () => { if (modalFn) $('#modalBody').innerHTML = modalFn(); };
let toastT;
function toast(msg, ms = 2600) { const t = $('#toast'); t.textContent = msg; t.hidden = false; clearTimeout(toastT); toastT = setTimeout(() => t.hidden = true, ms); }
const modalHead = (title, back) => `<div class="mhead">${back ? `<button class="iconbtn sm" data-act="${back.act}" ${back.attrs || ''} aria-label="Back">${IC.left}</button>` : ''}<h2 class="grow">${esc(title)}</h2><button class="btn sm ghost" data-act="closeModal">Done</button></div>`;
const seg = (act, attrs, opts, cur) => `<div class="seg">${opts.map(([v, l]) => `<button type="button" class="${String(v) === String(cur) ? 'on' : ''}" data-act="${act}" ${attrs} data-v="${esc(v)}">${l}</button>`).join('')}</div>`;
const opt = (v, l, cur) => `<option value="${esc(v)}"${String(v) === String(cur) ? ' selected' : ''}>${esc(l)}</option>`;
const catOptions = (cur, none = 'No category') => opt('', none, cur || '') + S.categories.map(c => opt(c.id, c.name, cur)).join('');
const groupOptions = (cur, none = 'No group') => opt('', none, cur || '') + S.taskGroups.map(g => opt(g.id, g.name, cur)).join('');
const swatch = sk => `<span class="swatches">${['bg', 'surface2', 'accent', 'coin'].map(k => `<i style="background:${sk.vars[k]}"></i>`).join('')}</span>`;

/* ---------------- skins: apply ---------------- */
function activeSkin() {
  const d = S.days[V.date]; const t = d && tpl(d.templateId);
  return (t && t.skinId && getSkin(t.skinId)) || getSkin(S.settings.defaultSkinId) || BUILTIN_SKINS[0];
}
function skinStyle(sk) {
  const v = sk.vars;
  return `--bg:${v.bg};--surface:${v.surface};--surface-2:${v.surface2};--text:${v.text};--muted:${v.muted};--accent:${v.accent};--accent-ink:${inkOn(v.accent)};--coin:${v.coin};--border:${v.border};--radius:${sk.radius}px;--font-display:${fontStack(sk.display)};--font-body:${fontStack(sk.body)};color-scheme:${sk.mode === 'dark' ? 'dark' : 'light'}`;
}
function applySkin(sk) {
  const key = sk.id + JSON.stringify(sk.vars) + sk.display + sk.body + sk.radius + sk.bgOpacity + (sk.bgImage || '').length;
  if (key === lastSkinKey) return; lastSkinKey = key;
  loadFont(sk.display); loadFont(sk.body);
  const root = document.documentElement; skinStyle(sk).split(';').forEach(p => { const i = p.indexOf(':'); root.style.setProperty(p.slice(0, i), p.slice(i + 1)); });
  const bg = $('#bgimg'); bg.style.backgroundImage = sk.bgImage ? `url("${sk.bgImage}")` : 'none'; bg.style.opacity = sk.bgImage ? sk.bgOpacity : 0;
  document.querySelector('meta[name=theme-color]')?.setAttribute('content', sk.vars.bg);
}

/* =====================================================================
   DAY PAGE
   ===================================================================== */
function pageDay() {
  const date = V.date, today = todayStr(), d = S.days[date], t = d && tpl(d.templateId);
  const future = date > today;
  let h = `<section class="daybar">
    <button class="iconbtn" data-act="shiftDay" data-n="-1" aria-label="Previous day">${IC.left}</button>
    <label class="datepick"><span class="dp-main">${date === today ? 'Today' : date === addDays(today, -1) ? 'Yesterday' : date === addDays(today, 1) ? 'Tomorrow' : esc(fmtDate(date, { weekday: 'long' }))}</span>
      <span class="dp-sub">${esc(fmtDate(date, { weekday: date === today ? 'long' : undefined, month: 'long', day: 'numeric', year: 'numeric' }))}</span>
      <input type="date" value="${date}" data-ch="pickDate" aria-label="Choose a date"></label>
    <button class="iconbtn" data-act="shiftDay" data-n="1" aria-label="Next day">${IC.right}</button></section>`;
  if (date !== today) h += `<div class="backtoday"><button class="chip" data-act="goToday">Back to today</button></div>`;
  if (!S.templates.length) return h + `<div class="card empty"><div class="big">Let's set up your first kind of day</div><p class="muted">Day templates hold the habits for a type of day — a work day, a day off, a low-energy day.</p><button class="btn" data-act="newTemplate">${IC.plus} Create a day template</button></div>`;
  if (!t) return h + templateChooser(date, d);

  const sc = scores(); const pd = sc.perDay[date];
  const ids = dayHabitIds(date); const doneN = ids.filter(id => d.habits[id]).length;
  const pct = ids.length ? doneN / ids.length : 0; const C = 2 * Math.PI * 36;
  const coins = pd ? pd.total : 0;
  h += `<section class="card hero">
    <div class="ring"><svg viewBox="0 0 84 84"><circle class="track" cx="42" cy="42" r="36" fill="none"/><circle class="val" cx="42" cy="42" r="36" fill="none" stroke-linecap="round" stroke-dasharray="${C}" stroke-dashoffset="${C * (1 - pct)}"/></svg><b>${doneN}/${ids.length}</b></div>
    <div class="grow"><div class="kicker">${esc(t.name)} <button class="link" data-act="changeTemplate">change</button></div>
      <div class="coins-big">${coinEl}${fmtN(coins)}</div>
      <div class="muted small">coins ${future ? 'planned for this day' : date === today ? 'earned today' : 'earned this day'}${pd && pd.penalty ? ` · <span class="overdue">−${fmtN(pd.penalty)} for missed must-dos</span>` : ''}</div></div></section>`;

  const habits = ids.map(habitById);
  h += `<section><div class="sec-head"><h2>Habits</h2><button class="iconbtn sm" data-act="editTemplate" data-id="${t.id}" aria-label="Edit this template">${IC.edit}</button></div>`;
  if (!habits.length) h += `<div class="card empty small muted">This template has no habits yet. <button class="link" data-act="editTemplate" data-id="${t.id}">Add some</button></div>`;
  else if (t.groupByCategory) {
    const groups = [...S.categories.map(c => ({ c, list: habits.filter(x => x.categoryId === c.id) })), { c: null, list: habits.filter(x => !catById(x.categoryId)) }];
    for (const g of groups) if (g.list.length) h += `<div class="cat-head"><span class="dot-c" style="--c:${g.c ? g.c.color : 'var(--muted)'}"></span>${esc(g.c ? g.c.name : 'Other')}</div><div class="list">${g.list.map(x => habitRow(x, date, pd, future)).join('')}</div>`;
  } else h += `<div class="list">${habits.map(x => habitRow(x, date, pd, future)).join('')}</div>`;
  h += `</section>`;

  if (t.showTasks !== false) {
    const list = d.tasks.map(x => ({ ...x, task: taskById(x.taskId) })).filter(x => x.task);
    h += `<section><div class="sec-head"><h2>Tasks</h2><button class="btn sm soft" data-act="addTasksToDay">${IC.plus} From task list</button></div>
      <div class="list">${list.length ? list.map(x => dayTaskRow(x, future)).join('') : `<p class="muted small">No tasks on this day yet.</p>`}</div>
      <form class="quick" data-submit="quickTask"><input type="text" name="title" placeholder="Quick-add a task for this day…" aria-label="New task" required>
      <select name="coins" aria-label="Coins">${[1, 2, 3, 4, 5].map(n => opt(n, `${n} coin${n > 1 ? 's' : ''}`, 3)).join('')}</select><button class="btn sm" type="submit">Add</button></form></section>`;
  }
  return h;
}
function templateChooser(date, d) {
  return `<section><h2 style="margin-bottom:4px">What kind of day is this?</h2><p class="muted small">Pick a template — it decides which habits show up${d && d.tasks.length ? '' : ' and how the app looks'}.</p>
  <div class="chooser">${S.templates.map(t => { const sk = getSkin(t.skinId) || getSkin(S.settings.defaultSkinId) || BUILTIN_SKINS[0];
    return `<button class="tplcard" data-act="chooseTemplate" data-id="${t.id}"><span class="sw">${['bg', 'surface2', 'accent', 'coin'].map(k => `<i style="background:${sk.vars[k]}"></i>`).join('')}</span><span class="tc-body"><b>${esc(t.name)}</b><span class="muted small">${t.habitIds.filter(habitById).length} habits</span></span></button>`; }).join('')}</div></section>`;
}
function habitRow(h, date, pd, future) {
  const d = S.days[date]; const done = !!d.habits[h.id]; const info = pd?.habits[h.id] || {}; const cat = catById(h.categoryId);
  const missed = !done && info.missed;
  let right;
  if (done) right = `+${fmtN(info.earned ?? h.coins)}${coinEl}${info.mult > 1 ? `<em>×${info.mult}</em>` : ''}`;
  else if (missed) right = h.priority === 3 ? `<span class="hb-coins neg">−${h.coins}</span>` : `<span class="muted small">missed</span>`;
  else { const m = info.nextMult || 1; right = `${h.coins}${coinEl}${m > 1 ? `<em>×${m}</em>` : ''}`; }
  const streak = info.streak || 0;
  return `<button class="habit ${done ? 'done' : ''} ${missed ? 'missed' : ''}" data-act="toggleHabit" data-id="${h.id}" style="--c:${cat ? cat.color : 'var(--border)'}" ${future ? 'disabled' : ''} aria-pressed="${done}">
    <span class="check">${IC.check}</span>
    <span class="hb-main"><span class="hb-name">${esc(h.name)}</span><span class="hb-meta">${cat ? `<span>${esc(cat.name)}</span>` : ''}${h.priority !== 1 ? `<span class="tag ${PRI_CLS[h.priority]}">${PRI[h.priority]}</span>` : `<span class="tag">Optional</span>`}${streak > 1 ? `<span class="streak">${IC.flame}${streak}</span>` : ''}</span></span>
    <span class="hb-coins">${right}</span></button>`;
}
function dayTaskRow(x, future) {
  const t = x.task; const g = groupById(t.groupId);
  return `<div class="trow ${x.done ? 'done' : ''}"><button class="tcheck" data-act="toggleDayTask" data-id="${t.id}" ${future ? 'disabled' : ''} aria-label="Complete">${IC.check}</button>
    <div class="t-main clickable" data-act="editTask" data-id="${t.id}"><span class="t-title">${esc(t.title)}</span><span class="t-meta">${g ? `<span class="dot-c" style="--c:${g.color}"></span>${esc(g.name)}` : ''}${t.due ? `<span class="${t.due < todayStr() && !x.done ? 'overdue' : ''}">due ${shortD(t.due)}</span>` : ''}</span></div>
    <span class="coinpill">${coinEl}${x.coins}</span><button class="iconbtn sm" data-act="removeDayTask" data-id="${t.id}" aria-label="Remove from this day">${IC.x}</button></div>`;
}
function ensureDay(s, date) { if (!s.days[date]) s.days[date] = { templateId: '', habits: {}, tasks: [] }; return s.days[date]; }

/* =====================================================================
   TASKS PAGE
   ===================================================================== */
function sortTasks(list, how) {
  const due = t => t.due || '9999-99-99';
  const gi = t => { const i = S.taskGroups.findIndex(g => g.id === t.groupId); return i < 0 ? 999 : i; };
  const byP = (a, b) => b.priority - a.priority || due(a).localeCompare(due(b));
  if (how === 'due') return list.sort((a, b) => due(a).localeCompare(due(b)) || b.priority - a.priority);
  if (how === 'group') return list.sort((a, b) => gi(a) - gi(b) || byP(a, b));
  if (how === 'closed') return list.sort((a, b) => String(b.closedAt || '').localeCompare(String(a.closedAt || '')));
  return list.sort(byP);
}
function pageTasks() {
  const open = S.tasks.filter(t => t.status === 'open'), closed = S.tasks.filter(t => t.status === 'closed');
  let list = V.taskStatus === 'open' ? open : closed;
  if (V.taskGroup !== 'all') list = list.filter(t => V.taskGroup === 'none' ? !groupById(t.groupId) : t.groupId === V.taskGroup);
  list = sortTasks(list.slice(), V.taskStatus === 'closed' && V.taskSort === 'priority' ? 'closed' : V.taskSort);
  let h = `<section><div class="sec-head" style="margin-top:8px"><h2>Tasks</h2><button class="btn sm ghost" data-act="manageGroups">Groups</button></div>
    <details class="addtask card" ${S.tasks.length ? '' : 'open'}><summary class="spread"><b>${IC.plus.replace('<svg', '<svg style="width:18px;height:18px;vertical-align:-4px"')} New task</b><span class="muted small">tap to ${S.tasks.length ? 'open' : 'close'}</span></summary>
      <form data-submit="newTask" class="stack" style="margin-top:12px">
        <input type="text" name="title" placeholder="What needs doing?" required aria-label="Task title">
        <div class="grid3"><label class="f"><span>Group</span><select name="groupId">${groupOptions('')}</select></label>
        <label class="f"><span>Due</span><input type="date" name="due"></label>
        <label class="f"><span>Priority</span><select name="priority">${[3, 2, 1].map(p => opt(p, TPRI[p], 2)).join('')}</select></label></div>
        <textarea name="notes" placeholder="Notes (optional)" style="min-height:48px"></textarea>
        <button class="btn" type="submit">Add task</button></form></details></section>
    <section><div class="toolbar"><div class="seg">${[['open', `Open · ${open.length}`], ['closed', `Done · ${closed.length}`]].map(([v, l]) => `<button class="${V.taskStatus === v ? 'on' : ''}" data-act="taskStatus" data-v="${v}">${l}</button>`).join('')}</div>
      <select data-ch="taskSort" aria-label="Sort">${[['priority', 'Sort: priority'], ['due', 'Sort: due date'], ['group', 'Sort: group']].map(([v, l]) => opt(v, l, V.taskSort)).join('')}</select>
      <select data-ch="taskGroup" aria-label="Filter by group">${opt('all', 'All groups', V.taskGroup)}${opt('none', 'No group', V.taskGroup)}${S.taskGroups.map(g => opt(g.id, g.name, V.taskGroup)).join('')}</select></div>`;
  if (!list.length) h += `<div class="card empty muted small">${V.taskStatus === 'open' ? 'Nothing open. Enjoy it.' : 'No completed tasks yet.'}</div>`;
  else if (V.taskSort === 'group') {
    let cur = null; h += list.map(t => { let head = ''; if (t.groupId !== cur) { cur = t.groupId; const g = groupById(cur); head = `<div class="group-head"><span class="dot-c" style="--c:${g ? g.color : 'var(--muted)'}"></span>${esc(g ? g.name : 'No group')}</div>`; } return head + `<div style="margin-bottom:8px">${taskCard(t)}</div>`; }).join('');
  } else h += `<div class="list">${list.map(taskCard).join('')}</div>`;
  return h + `</section>`;
}
function taskCard(t) {
  const g = groupById(t.groupId); const today = todayStr(); const onToday = S.days[today]?.tasks.some(x => x.taskId === t.id);
  return `<article class="task pr${t.priority}"><button class="tcheck ${t.status === 'closed' ? 'on' : ''}" data-act="toggleTaskClosed" data-id="${t.id}" aria-label="${t.status === 'closed' ? 'Reopen' : 'Mark done'}">${IC.check}</button>
    <div class="t-main clickable" data-act="editTask" data-id="${t.id}"><span class="t-title">${esc(t.title)}</span><span class="t-meta">
      <span class="tag ${t.priority === 3 ? 'must' : t.priority === 2 ? 'should' : ''}">${TPRI[t.priority]}</span>
      ${g ? `<span class="dot-c" style="--c:${g.color}"></span>${esc(g.name)}` : ''}
      ${t.due && t.status === 'open' ? `<span class="${t.due < today ? 'overdue' : ''}">due ${shortD(t.due)}</span>` : ''}
      ${t.status === 'closed' && t.closedAt ? `<span>done ${isDateStr(t.closedAt) ? shortD(t.closedAt) : ''}</span>` : ''}</span></div>
    ${t.status === 'open' ? (onToday ? `<span class="tag">on today</span>` : `<button class="btn sm soft" data-act="taskToToday" data-id="${t.id}">Add to Today's List</button>`) : ''}</article>`;
}

/* =====================================================================
   PROGRESS PAGE
   ===================================================================== */
function pageProgress() {
  const sc = scores(); const today = todayStr(); const ws = weekStart(today);
  let week = 0, tasks30 = 0; for (const [dt, p] of Object.entries(sc.perDay)) { if (dt >= ws) week += p.total; if (dt > addDays(today, -30)) tasks30 += p.tasksDone; }
  const active = S.habits.filter(h => S.templates.some(t => t.habitIds.includes(h.id)));
  const best = active.reduce((a, h) => sc.habit[h.id].streak > (a?.s || 0) ? { h, s: sc.habit[h.id].streak } : a, null);
  const daily = runQuery({ metric: 'coins', group: 'day', range: '30', category: 'all', template: 'all' });
  let h = `<section style="margin-top:8px"><div class="tiles">
    <div class="tile feature"><div class="tl">Coins collected</div><div class="tv">${coinEl}${fmtN(sc.total)}</div></div>
    <div class="tile"><div class="tl">This week</div><div class="tv">${fmtN(week)}</div></div>
    <div class="tile"><div class="tl">Today</div><div class="tv">${fmtN(sc.perDay[today]?.total || 0)}</div></div>
    <div class="tile"><div class="tl">Best streak now</div><div class="tv">${best ? best.s : 0}</div><div class="tiny muted">${best ? esc(best.h.name) : '—'}</div></div>
    <div class="tile"><div class="tl">Tasks · 30 days</div><div class="tv">${tasks30}</div></div>
    <div class="tile"><div class="tl">Completion · 30d</div><div class="tv">${runQuery({ metric: 'completion', group: 'template', range: '30', category: 'all', template: 'all' }).overall}%</div></div>
  </div></section>
  <section class="card"><div class="sec-head"><h3>Coins, last 30 days</h3></div>${chartBar(daily.rows)}</section>
  <section class="card"><div class="sec-head"><h3>Habit streaks</h3><span class="muted tiny">multiplier grows each full week</span></div>
    ${active.length ? `<div class="tbl-wrap"><table class="tbl"><thead><tr><th>Habit</th><th class="num">Streak</th><th class="num">Best</th><th class="num">Bonus</th><th class="num">Done</th></tr></thead><tbody>
    ${active.map(h => ({ h, x: sc.habit[h.id] })).sort((a, b) => b.x.streak - a.x.streak || b.x.best - a.x.best).map(({ h, x }) => { const c = catById(h.categoryId);
      return `<tr><td><span class="row" style="gap:7px;flex-wrap:nowrap"><span class="dot-c" style="--c:${c ? c.color : 'var(--muted)'}"></span>${esc(h.name)}</span></td><td class="num">${x.streak}</td><td class="num">${x.best}</td><td class="num">×${x.multNow}</td><td class="num">${x.eligible ? Math.round(x.done / x.eligible * 100) + '%' : '—'}</td></tr>`; }).join('')}</tbody></table></div>` : '<p class="muted small">Add habits to a template to start tracking streaks.</p>'}</section>
  <section class="card"><div class="sec-head"><h3>Categories · 30 days</h3><span class="muted tiny">completion rate</span></div>${chartHBar(runQuery({ metric: 'completion', group: 'category', range: '30', category: 'all', template: 'all' }).rows, { unit: '%', max: 100 })}</section>
  ${queryCard()}
  <section class="card"><div class="sec-head"><h3>Export</h3></div>
    <div class="export-grid">
      <button class="btn" data-act="exportFull" data-kind="visual">${IC.dl} Full visual report<small>Every chart & table, printable as PDF</small></button>
      <button class="btn ghost" data-act="exportPartial" data-kind="visual">${IC.dl} Custom visual report<small>Pick what to include</small></button>
      <button class="btn" data-act="exportFull" data-kind="sheet">${IC.dl} Full data spreadsheet<small>Every data type, one tab each</small></button>
      <button class="btn ghost" data-act="exportPartial" data-kind="sheet">${IC.dl} Custom spreadsheet<small>Pick what to include</small></button>
    </div></section>`;
  return h;
}

/* ---------- query engine ---------- */
const METRICS = {
  coins: { name: 'Coins earned', kind: 'both', val: b => round2(b.coins) },
  completion: { name: 'Habit completion %', kind: 'habit', val: b => b.eligible ? Math.round(b.done / b.eligible * 100) : 0, unit: '%' },
  habitsDone: { name: 'Habits completed', kind: 'habit', val: b => b.done },
  tasksDone: { name: 'Tasks completed', kind: 'task', val: b => b.tasksDone },
  penalties: { name: 'Coins lost to missed must-dos', kind: 'habit', val: b => round2(b.penalty) },
};
const GROUPS = {
  day: { name: 'Day', kind: 'time' }, week: { name: 'Week', kind: 'time' }, month: { name: 'Month', kind: 'time' },
  habit: { name: 'Habit', kind: 'habit' }, category: { name: 'Category', kind: 'habit' },
  template: { name: 'Day template', kind: 'both' }, taskGroup: { name: 'Task group', kind: 'task' },
};
function firstDate() { const ks = Object.keys(S.days).sort(); return ks[0] || todayStr(); }
function rangeDates(q) {
  const today = todayStr();
  if (q.range === 'custom') { const a = isDateStr(q.from) ? q.from : firstDate(), b = isDateStr(q.to) ? q.to : today; return a <= b ? [a, b] : [b, a]; }
  if (q.range === 'all') return [firstDate() < today ? firstDate() : today, today];
  return [addDays(today, -(+q.range - 1)), today];
}
let _facts = null, _factsKey = '';
function facts() {
  if (_factsKey === _scKey && _facts) return _facts;
  const sc = scores(); const hf = [], tf = [];
  for (const [date, pd] of Object.entries(sc.perDay)) {
    const d = S.days[date];
    for (const [id, x] of Object.entries(pd.habits)) { const h = habitById(id); hf.push({ date, habitId: id, categoryId: h?.categoryId || '', templateId: d.templateId, done: x.done, earned: x.earned || 0, penalty: x.penalty || 0, pending: !!x.pending }); }
    for (const t of d.tasks) { const task = taskById(t.taskId); tf.push({ date, taskId: t.taskId, groupId: task?.groupId || '', templateId: d.templateId, done: !!t.done, coins: +t.coins || 0 }); }
  }
  _factsKey = _scKey; return _facts = { hf, tf };
}
function bucketKey(group, f) {
  switch (group) {
    case 'day': return [f.date, shortD(f.date)];
    case 'week': { const w = weekStart(f.date); return [w, 'Wk ' + shortD(w)]; }
    case 'month': return [f.date.slice(0, 7), parseD(f.date.slice(0, 7) + '-01').toLocaleDateString(undefined, { month: 'short', year: '2-digit' })];
    case 'habit': return [f.habitId, habitById(f.habitId)?.name || 'Deleted habit'];
    case 'category': return [f.categoryId || '_', catById(f.categoryId)?.name || 'No category'];
    case 'template': return [f.templateId || '_', tpl(f.templateId)?.name || 'No template'];
    case 'taskGroup': return [f.groupId || '_', groupById(f.groupId)?.name || 'No group'];
  }
}
function runQuery(q) {
  const M = METRICS[q.metric] || METRICS.coins, G = GROUPS[q.group] || GROUPS.day;
  if ((M.kind === 'habit' && G.kind === 'task') || (M.kind === 'task' && G.kind === 'habit')) return { rows: [], error: `“${M.name}” can't be grouped by ${G.name.toLowerCase()}.`, overall: 0 };
  const [from, to] = rangeDates(q); const { hf, tf } = facts(); const B = new Map();
  const get = (k, label) => { if (!B.has(k)) B.set(k, { key: k, label, coins: 0, done: 0, eligible: 0, penalty: 0, tasksDone: 0 }); return B.get(k); };
  const inR = f => f.date >= from && f.date <= to && (q.template === 'all' || f.templateId === q.template);
  if (G.kind === 'time') { let d = from; let guard = 0; while (d <= to && guard++ < 4000) { const [k, l] = bucketKey(q.group, { date: d }); get(k, l); d = addDays(d, 1); } }
  const tot = { coins: 0, done: 0, eligible: 0, penalty: 0, tasksDone: 0 };
  const useH = M.kind !== 'task' && G.kind !== 'task', useT = M.kind !== 'habit' && G.kind !== 'habit' && q.category === 'all';
  if (useH) for (const f of hf) { if (!inR(f) || (q.category !== 'all' && f.categoryId !== q.category)) continue; const b = get(...bucketKey(q.group, f));
    for (const o of [b, tot]) { o.coins += f.earned - f.penalty; o.penalty += f.penalty; if (!f.pending) { o.eligible++; if (f.done) o.done++; } } }
  if (useT) for (const f of tf) { if (!inR(f) || !f.done) continue; const b = get(...bucketKey(q.group, f)); for (const o of [b, tot]) { o.tasksDone++; o.coins += f.coins; } }
  let rows = [...B.values()].map(b => ({ label: b.label, key: b.key, value: M.val(b) }));
  if (G.kind === 'time') rows.sort((a, b) => a.key.localeCompare(b.key)); else rows.sort((a, b) => b.value - a.value);
  return { rows, overall: M.val(tot), unit: M.unit || '', metric: M.name, group: G.name, from, to };
}
function queryCard() {
  const q = V.q; const r = runQuery(q); const G = GROUPS[q.group];
  return `<section class="card"><div class="sec-head"><h3>Ask your data</h3></div>
    <div class="qgrid">
      <label class="f"><span>Show</span><select data-ch="q" data-k="metric">${Object.entries(METRICS).map(([k, m]) => opt(k, m.name, q.metric)).join('')}</select></label>
      <label class="f"><span>Grouped by</span><select data-ch="q" data-k="group">${Object.entries(GROUPS).map(([k, g]) => opt(k, g.name, q.group)).join('')}</select></label>
      <label class="f"><span>Time range</span><select data-ch="q" data-k="range">${[['7', 'Last 7 days'], ['30', 'Last 30 days'], ['90', 'Last 90 days'], ['365', 'Last year'], ['all', 'All time'], ['custom', 'Custom…']].map(([v, l]) => opt(v, l, q.range)).join('')}</select></label>
      ${q.range === 'custom' ? `<label class="f"><span>From</span><input type="date" data-ch="q" data-k="from" value="${esc(q.from)}"></label><label class="f"><span>To</span><input type="date" data-ch="q" data-k="to" value="${esc(q.to)}"></label>` : ''}
      <label class="f"><span>Category</span><select data-ch="q" data-k="category">${opt('all', 'All categories', q.category)}${S.categories.map(c => opt(c.id, c.name, q.category)).join('')}</select></label>
      <label class="f"><span>Day template</span><select data-ch="q" data-k="template">${opt('all', 'All templates', q.template)}${S.templates.map(t => opt(t.id, t.name, q.template)).join('')}</select></label>
    </div>
    <div style="margin-top:14px">${r.error ? `<p class="muted">${esc(r.error)}</p>` : `<p class="small"><b>${esc(r.metric)}</b> by ${esc(r.group.toLowerCase())}, ${shortD(r.from)} – ${shortD(r.to)}: <b>${fmtN(r.overall)}${r.unit}</b> overall</p>
      ${G.kind === 'time' ? chartBar(r.rows, { unit: r.unit }) : chartHBar(r.rows, { unit: r.unit, max: r.unit === '%' ? 100 : undefined })}
      <details style="margin-top:8px"><summary class="small muted" style="cursor:pointer">Show as table</summary><div class="tbl-wrap"><table class="tbl"><thead><tr><th>${esc(r.group)}</th><th class="num">${esc(r.metric)}</th></tr></thead><tbody>${r.rows.map(x => `<tr><td>${esc(x.label)}</td><td class="num">${fmtN(x.value)}${r.unit}</td></tr>`).join('')}</tbody></table></div></details>`}</div></section>`;
}

/* ---------- SVG charts (no library) ---------- */
const chartW = () => Math.round(clamp((window.innerWidth || 640) - 66, 320, 680));
function chartBar(data, { w = chartW(), h = Math.round(w * .5), color = 'var(--accent)', neg = 'var(--danger)', muted = 'var(--muted)', grid = 'var(--border)', unit = '' } = {}) {
  if (!data.length || data.every(d => !d.value)) return '<p class="muted small" style="margin:0">No data in this range yet.</p>';
  const P = { l: 40, r: 6, t: 10, b: 26 }; const iw = w - P.l - P.r, ih = h - P.t - P.b;
  const mx = Math.max(0, ...data.map(d => d.value)), mn = Math.min(0, ...data.map(d => d.value));
  const nice = v => { if (v <= 0) return 0; const p = 10 ** Math.floor(Math.log10(v)); const n = v / p; return (n <= 1 ? 1 : n <= 2 ? 2 : n <= 5 ? 5 : 10) * p; };
  const top = nice(mx) || 1, bot = mn < 0 ? -nice(-mn) : 0; const y = v => P.t + ih * (1 - (v - bot) / (top - bot));
  const bw = iw / data.length; const every = Math.ceil(data.length / Math.max(4, Math.floor(w / 80)));
  let s = `<svg class="chart" viewBox="0 0 ${w} ${h}" role="img" aria-label="Bar chart">`;
  for (let i = 0; i <= 4; i++) { const v = bot + (top - bot) * i / 4; s += `<line x1="${P.l}" x2="${w - P.r}" y1="${y(v)}" y2="${y(v)}" style="stroke:${grid};stroke-width:1${v === 0 ? '' : ';stroke-dasharray:2 4'}"/><text x="${P.l - 7}" y="${y(v) + 4}" text-anchor="end" style="fill:${muted}">${fmtN(v)}${unit}</text>`; }
  data.forEach((d, i) => { const x = P.l + i * bw + bw * .14, bwid = Math.max(1.5, bw * .72); const y0 = y(Math.max(0, d.value)), y1 = y(Math.min(0, d.value));
    s += `<rect x="${x}" y="${y0}" width="${bwid}" height="${Math.max(d.value ? 1.5 : 0, y1 - y0)}" rx="${Math.min(4, bwid / 3)}" style="fill:${d.value < 0 ? neg : color}"><title>${esc(d.label)}: ${fmtN(d.value)}${unit}</title></rect>`;
    if (i % every === 0) s += `<text x="${x + bwid / 2}" y="${h - 7}" text-anchor="middle" style="fill:${muted}">${esc(d.label)}</text>`; });
  return s + '</svg>';
}
function chartHBar(data, { w = chartW(), row = 30, color = 'var(--accent)', text = 'var(--text)', muted = 'var(--muted)', track = 'var(--surface-2)', unit = '', max } = {}) {
  if (!data.length) return '<p class="muted small" style="margin:0">No data in this range yet.</p>';
  data = data.slice(0, 25); const L = Math.min(150, w * .36), R = 50; const mx = max || Math.max(1, ...data.map(d => Math.abs(d.value))); const h = data.length * row + 4;
  let s = `<svg class="chart" viewBox="0 0 ${w} ${h}" role="img" aria-label="Bar chart">`;
  data.forEach((d, i) => { const yy = i * row + 4; const bw = (w - L - R) * Math.max(0, d.value) / mx; const max = Math.floor(L / 7.2); const label = d.label.length > max ? d.label.slice(0, max - 1) + '…' : d.label;
    s += `<text x="0" y="${yy + row / 2 + 1}" style="fill:${text};font-size:12.5px">${esc(label)}</text><rect x="${L}" y="${yy + 6}" width="${w - L - R}" height="${row - 12}" rx="${(row - 12) / 2}" style="fill:${track}"/>
      <rect x="${L}" y="${yy + 6}" width="${Math.max(d.value > 0 ? 4 : 0, bw)}" height="${row - 12}" rx="${(row - 12) / 2}" style="fill:${color}"><title>${esc(d.label)}: ${fmtN(d.value)}${unit}</title></rect>
      <text x="${w}" y="${yy + row / 2 + 1}" text-anchor="end" style="fill:${muted};font-weight:600">${fmtN(d.value)}${unit}</text>`; });
  return s + '</svg>';
}

/* ---------- export datasets ---------- */
const DATASETS = [
  ['overview', 'Overview totals'], ['dailyCoins', 'Daily coins & completion'], ['habitStats', 'Habit streaks & completion'],
  ['categoryStats', 'Category breakdown'], ['templateUsage', 'Day-template usage'], ['habitLog', 'Habit log (every check-in)'],
  ['dayTasks', 'Daily task log'], ['openTasks', 'Open tasks'], ['closedTasks', 'Completed tasks'],
];
function dataset(id, from, to) {
  const sc = scores(); const days = Object.keys(sc.perDay).filter(d => d >= from && d <= to).sort(); const { hf, tf } = facts();
  const H = hf.filter(f => f.date >= from && f.date <= to), T = tf.filter(f => f.date >= from && f.date <= to);
  const q = (metric, group) => runQuery({ metric, group, range: 'custom', from, to, category: 'all', template: 'all' });
  switch (id) {
    case 'overview': { const tot = days.reduce((a, d) => { const p = sc.perDay[d]; a.c += p.total; a.h += p.habitCoins; a.t += p.taskCoins; a.p += p.penalty; a.d += p.done; a.e += p.eligible; a.td += p.tasksDone; return a; }, { c: 0, h: 0, t: 0, p: 0, d: 0, e: 0, td: 0 });
      return { title: 'Overview totals', headers: ['Measure', 'Value'], rows: [['Date range', `${from} to ${to}`], ['Days logged', days.length], ['Total coins', round2(tot.c)], ['Coins from habits', round2(tot.h)], ['Coins from tasks', tot.t], ['Coins lost to missed must-dos', tot.p], ['Habits completed', tot.d], ['Habit completion rate', tot.e ? Math.round(tot.d / tot.e * 100) + '%' : '—'], ['Tasks completed', tot.td], ['Lifetime coins (all time)', sc.total]] }; }
    case 'dailyCoins': return { title: 'Daily coins & completion', headers: ['Date', 'Day template', 'Habits done', 'Habits scheduled', 'Habit coins', 'Task coins', 'Penalties', 'Total coins'],
      rows: days.map(d => { const p = sc.perDay[d]; return [d, tpl(S.days[d].templateId)?.name || '', p.done, p.scheduled, p.habitCoins, p.taskCoins, p.penalty, p.total]; }),
      chart: { type: 'bar', data: q('coins', 'day').rows } };
    case 'habitStats': { const rows = S.habits.map(h => { const f = H.filter(x => x.habitId === h.id); const el = f.filter(x => !x.pending); const dn = el.filter(x => x.done).length; const x = sc.habit[h.id];
        return [h.name, catById(h.categoryId)?.name || '', h.coins, PRI[h.priority], el.length, dn, el.length ? Math.round(dn / el.length * 100) : '', round2(f.reduce((a, b) => a + b.earned, 0)), f.reduce((a, b) => a + b.penalty, 0), x.streak, x.best, x.multNow]; }).filter(r => r[4] || r[9]);
      return { title: 'Habit streaks & completion', headers: ['Habit', 'Category', 'Coins', 'Priority', 'Times scheduled', 'Times done', 'Completion %', 'Coins earned', 'Coins lost', 'Current streak', 'Best streak', 'Current bonus ×'], rows,
        chart: { type: 'hbar', unit: '%', max: 100, data: rows.filter(r => r[6] !== '').map(r => ({ label: r[0], value: r[6] })).sort((a, b) => b.value - a.value) } }; }
    case 'categoryStats': { const c = q('completion', 'category').rows, k = q('coins', 'category').rows;
      return { title: 'Category breakdown', headers: ['Category', 'Completion %', 'Coins earned'], rows: c.map(r => [r.label, r.value, k.find(x => x.key === r.key)?.value ?? 0]), chart: { type: 'hbar', unit: '%', max: 100, data: c } }; }
    case 'templateUsage': { const m = {}; days.forEach(d => { const n = tpl(S.days[d].templateId)?.name || 'No template'; m[n] = m[n] || { n: 0, c: 0 }; m[n].n++; m[n].c += sc.perDay[d].total; });
      const rows = Object.entries(m).sort((a, b) => b[1].n - a[1].n).map(([n, v]) => [n, v.n, round2(v.c), round2(v.c / v.n)]);
      return { title: 'Day-template usage', headers: ['Day template', 'Days used', 'Total coins', 'Average coins per day'], rows, chart: { type: 'hbar', data: rows.map(r => ({ label: r[0], value: r[1] })) } }; }
    case 'habitLog': return { title: 'Habit log', headers: ['Date', 'Habit', 'Category', 'Day template', 'Status', 'Coins earned', 'Coins lost'],
      rows: H.sort((a, b) => a.date.localeCompare(b.date)).map(f => [f.date, habitById(f.habitId)?.name || '', catById(f.categoryId)?.name || '', tpl(f.templateId)?.name || '', f.done ? 'Done' : f.pending ? 'Not yet' : 'Missed', f.earned, f.penalty]) };
    case 'dayTasks': return { title: 'Daily task log', headers: ['Date', 'Task', 'Group', 'Coins', 'Done'],
      rows: T.sort((a, b) => a.date.localeCompare(b.date)).map(f => [f.date, taskById(f.taskId)?.title || '', groupById(f.groupId)?.name || '', f.coins, f.done ? 'Yes' : 'No']),
      chart: { type: 'bar', data: q('tasksDone', 'day').rows } };
    case 'openTasks': return { title: 'Open tasks', headers: ['Task', 'Group', 'Priority', 'Due', 'Notes'], rows: sortTasks(S.tasks.filter(t => t.status === 'open'), 'priority').map(t => [t.title, groupById(t.groupId)?.name || '', TPRI[t.priority], t.due || '', t.notes || '']) };
    case 'closedTasks': return { title: 'Completed tasks', headers: ['Task', 'Group', 'Priority', 'Completed', 'Notes'], rows: sortTasks(S.tasks.filter(t => t.status === 'closed' && (!isDateStr(t.closedAt) || (t.closedAt >= from && t.closedAt <= to))), 'closed').map(t => [t.title, groupById(t.groupId)?.name || '', TPRI[t.priority], t.closedAt || '', t.notes || '']) };
  }
}
function exportModal(kind, partial) {
  const st = exportModal.st = exportModal.st || { ids: Object.fromEntries(DATASETS.map(([id]) => [id, true])), range: 'all', from: '', to: '' };
  const visual = kind === 'visual';
  return `${modalHead(visual ? (partial ? 'Custom visual report' : 'Full visual report') : (partial ? 'Custom data spreadsheet' : 'Full data spreadsheet'))}
    ${partial ? `<p class="small muted">Tick the data you want included.</p>${DATASETS.map(([id, n]) => `<label class="check-row"><input type="checkbox" data-ch="expPick" data-id="${id}" ${st.ids[id] ? 'checked' : ''}> ${esc(n)}</label>`).join('')}
      <div class="row" style="margin:6px 0 4px"><button class="link small" data-act="expAll" data-v="1">Select all</button><button class="link small" data-act="expAll" data-v="0">Clear</button></div>
      <hr class="hr"><div class="grid3"><label class="f"><span>Date range</span><select data-ch="expRange" data-k="range">${[['all', 'All time'], ['7', 'Last 7 days'], ['30', 'Last 30 days'], ['90', 'Last 90 days'], ['custom', 'Custom…']].map(([v, l]) => opt(v, l, st.range)).join('')}</select></label>
      ${st.range === 'custom' ? `<label class="f"><span>From</span><input type="date" data-ch="expRange" data-k="from" value="${esc(st.from)}"></label><label class="f"><span>To</span><input type="date" data-ch="expRange" data-k="to" value="${esc(st.to)}"></label>` : ''}</div>`
    : `<p class="small muted">Includes all ${DATASETS.length} data types for all time.</p>`}
    <div class="stack" style="margin-top:16px">${visual
      ? `<button class="btn block" data-act="runExport" data-kind="print" data-partial="${partial ? 1 : 0}">Open printable report (save as PDF)</button><button class="btn ghost block" data-act="runExport" data-kind="html" data-partial="${partial ? 1 : 0}">${IC.dl} Download report file (.html)</button>`
      : `<button class="btn block" data-act="runExport" data-kind="xlsx" data-partial="${partial ? 1 : 0}">${IC.dl} Download spreadsheet (.xlsx)</button><button class="btn ghost block" data-act="runExport" data-kind="csv" data-partial="${partial ? 1 : 0}">${IC.dl} Download as CSV instead</button>`}</div>`;
}
function exportSelection(partial) {
  const st = exportModal.st; const ids = partial ? DATASETS.map(d => d[0]).filter(id => st.ids[id]) : DATASETS.map(d => d[0]);
  const [from, to] = partial ? rangeDates({ range: st.range, from: st.from, to: st.to }) : rangeDates({ range: 'all' });
  return { ids, from, to };
}
function reportHTML(ids, from, to) {
  const cs = getComputedStyle(document.documentElement); const v = k => cs.getPropertyValue(k).trim();
  const col = { bg: v('--bg'), surface: v('--surface'), text: v('--text'), muted: v('--muted'), accent: v('--accent'), border: v('--border'), surface2: v('--surface-2'), coin: v('--coin') };
  const sk = activeSkin();
  const secs = ids.map(id => { const ds = dataset(id, from, to); let chart = '';
    if (ds.chart && ds.chart.data.length) chart = ds.chart.type === 'bar' ? chartBar(ds.chart.data, { w: 760, h: 260, color: col.accent, muted: col.muted, grid: col.border, neg: '#c0392b', unit: ds.chart.unit || '' }) : chartHBar(ds.chart.data, { w: 760, color: col.accent, text: col.text, muted: col.muted, track: col.surface2, unit: ds.chart.unit || '', max: ds.chart.max });
    const rows = ds.rows.slice(0, 400);
    return `<section><h2>${esc(ds.title)}</h2>${chart ? `<div class="chart-box">${chart}</div>` : ''}${rows.length ? `<table><thead><tr>${ds.headers.map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${rows.map(r => `<tr>${r.map(c => `<td>${esc(c)}</td>`).join('')}</tr>`).join('')}</tbody></table>${ds.rows.length > 400 ? `<p class="m">Showing the first 400 of ${ds.rows.length} rows — the spreadsheet export has them all.</p>` : ''}` : '<p class="m">No data in this range.</p>'}</section>`; }).join('');
  return `<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Habitude report ${from} to ${to}</title>
  <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=${sk.display.replace(/ /g, '+')}:wght@600&family=${sk.body.replace(/ /g, '+')}:wght@400;600&display=swap">
  <style>body{margin:0;background:${col.bg};color:${col.text};font:15px/1.45 ${fontStack(sk.body)};-webkit-print-color-adjust:exact;print-color-adjust:exact}.wrap{max-width:860px;margin:0 auto;padding:36px 24px}
  h1,h2{font-family:${fontStack(sk.display)};font-weight:600}h1{font-size:34px;margin:0}h2{font-size:21px;margin:0 0 12px}.m{color:${col.muted};font-size:13px}
  section{background:${col.surface};border:1px solid ${col.border};border-radius:14px;padding:18px;margin-top:18px;break-inside:avoid-page}
  table{width:100%;border-collapse:collapse;font-size:13px;margin-top:10px}th{text-align:left;color:${col.muted};font-size:11px;text-transform:uppercase;letter-spacing:.05em;border-bottom:1px solid ${col.border};padding:6px}td{padding:6px;border-bottom:1px solid ${col.border}}
  svg{width:100%;height:auto}svg text{font:11px ${fontStack(sk.body)}}.chart-box{margin:6px 0 4px}.bar{height:6px;background:${col.accent};border-radius:3px;width:64px;margin:10px 0}
  @media print{body{background:#fff}.wrap{padding:0}section{box-shadow:none}.noprint{display:none}}</style></head>
  <body><div class="wrap"><div class="bar"></div><h1>Habit & task report</h1><p class="m">${esc(fmtDate(from, { month: 'long', day: 'numeric', year: 'numeric' }))} – ${esc(fmtDate(to, { month: 'long', day: 'numeric', year: 'numeric' }))} · generated ${esc(new Date().toLocaleString())}</p>
  <p class="noprint"><button onclick="print()" style="font:inherit;padding:8px 14px;border-radius:8px;border:1px solid ${col.border};background:${col.accent};color:${inkOn(col.accent.startsWith('#') ? col.accent : '#a8472a')};cursor:pointer">Print / Save as PDF</button></p>${secs}</div></body></html>`;
}
function download(name, content, type) {
  const blob = content instanceof Blob ? content : new Blob([content], { type }); const url = URL.createObjectURL(blob);
  const a = Object.assign(document.createElement('a'), { href: url, download: name }); document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 4000);
}
function loadScript(src) { return new Promise((res, rej) => { const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = () => rej(new Error('Could not load ' + src)); document.head.appendChild(s); }); }
const csvCell = c => { c = String(c ?? ''); return /[",\n]/.test(c) ? `"${c.replace(/"/g, '""')}"` : c; };
async function runExport(kind, partial) {
  const { ids, from, to } = exportSelection(partial); if (!ids.length) return toast('Tick at least one data type.');
  const stamp = todayStr();
  if (kind === 'print' || kind === 'html') {
    const html = reportHTML(ids, from, to);
    if (kind === 'html') return download(`habitude-report-${stamp}.html`, html, 'text/html');
    const w = window.open('', '_blank'); if (!w) { download(`habitude-report-${stamp}.html`, html, 'text/html'); return toast('Pop-up blocked — downloaded the report file instead.'); }
    w.document.open(); w.document.write(html); w.document.close(); return;
  }
  const sets = ids.map(id => dataset(id, from, to));
  if (kind === 'csv') return download(`habitude-data-${stamp}.csv`, sets.map(ds => [[ds.title], ds.headers, ...ds.rows].map(r => r.map(csvCell).join(',')).join('\n')).join('\n\n'), 'text/csv');
  try {
    if (!window.XLSX) { toast('Preparing spreadsheet…'); await loadScript(XLSX_URL); }
    const wb = XLSX.utils.book_new(); const used = new Set();
    for (const ds of sets) { let n = ds.title.replace(/[\[\]:*?/\\]/g, '').slice(0, 31); while (used.has(n)) n = n.slice(0, 29) + Math.floor(Math.random() * 90 + 10); used.add(n);
      const ws = XLSX.utils.aoa_to_sheet([ds.headers, ...ds.rows]); ws['!cols'] = ds.headers.map(h => ({ wch: Math.max(12, h.length + 2) })); XLSX.utils.book_append_sheet(wb, ws, n); }
    XLSX.writeFile(wb, `habitude-data-${stamp}.xlsx`);
  } catch { toast('Spreadsheet library unavailable offline — downloading CSV instead.'); runExport('csv', partial); }
}

/* =====================================================================
   SETTINGS PAGE
   ===================================================================== */
function pageSettings() {
  const def = getSkin(S.settings.defaultSkinId) || BUILTIN_SKINS[0];
  return `<section style="margin-top:8px"><div class="sec-head"><h2>Day templates</h2><button class="btn sm" data-act="newTemplate">${IC.plus} Create template</button></div>
    <div class="card">${S.templates.length ? S.templates.map(t => { const sk = getSkin(t.skinId) || def; return `<div class="set-row">${swatch(sk)}<div class="grow"><b>${esc(t.name)}</b><div class="muted tiny">${t.habitIds.filter(habitById).length} habits · ${esc(sk.name)}${t.skinId ? '' : ' (default)'}${t.showTasks !== false ? ' · tasks list' : ''}</div></div><button class="btn sm ghost" data-act="editTemplate" data-id="${t.id}">Edit</button></div>`; }).join('') : '<p class="muted small" style="margin:0">No templates yet. Create one for each kind of day you have.</p>'}</div></section>

  <section><div class="sec-head"><h2>Habits & categories</h2></div>
    <div class="card"><div class="set-row"><div class="grow"><b>All habits</b><div class="muted tiny">${S.habits.length} habits · edit coins, priority, category, and which templates use them</div></div><button class="btn sm ghost" data-act="allHabits">Open</button></div>
    <div class="set-row"><div class="grow"><b>Categories</b><div class="muted tiny row" style="gap:6px;margin-top:4px">${S.categories.map(c => `<span class="row" style="gap:4px"><span class="dot-c" style="--c:${c.color}"></span>${esc(c.name)}</span>`).join('') || 'None yet'}</div></div><button class="btn sm ghost" data-act="manageCategories">Edit</button></div></div></section>

  <section><div class="sec-head"><h2>Skins</h2><div class="row"><button class="btn sm ghost" data-act="uploadSkin">${IC.ul} Upload</button><button class="btn sm" data-act="newSkin">${IC.brush} Create skin</button></div></div>
    <div class="skingrid">${allSkins().map(sk => skinCard(sk)).join('')}</div></section>

  <section><div class="sec-head"><h2>Scoring</h2></div><div class="card stack">
    <div class="grid2"><label class="f"><span>Bonus added per full week of streak</span><select data-ch="setting" data-k="weeklyStep">${[0.25, 0.5, 0.75, 1].map(v => opt(v, `+${v}×`, S.settings.weeklyStep)).join('')}</select></label>
    <label class="f"><span>Maximum multiplier</span><select data-ch="setting" data-k="maxMult">${[1.25, 1.5, 1.75, 2, 2.25, 2.5, 2.75, 3].map(v => opt(v, `${v}×`, S.settings.maxMult)).join('')}</select></label></div>
    <div class="help">With +${S.settings.weeklyStep}× per week up to ${S.settings.maxMult}×, a habit kept up for 1 week earns ${Math.min(S.settings.maxMult, 1 + +S.settings.weeklyStep)}× its coins, after 2 weeks ${Math.min(S.settings.maxMult, 1 + 2 * S.settings.weeklyStep)}×, and so on. Days where a habit isn't in that day's template don't count for or against its streak.
      <ul style="padding-left:18px;margin:8px 0 0"><li><b>Must</b> (priority 3): missing it resets its streak <i>and</i> deducts its coins.</li><li><b>Should</b> (priority 2): missing it resets its streak.</li><li><b>Optional</b> (priority 1): skip it freely — streak and coins unaffected.</li></ul></div></div></section>

  <section id="sheets"><div class="sec-head"><h2>Google Sheets sync</h2></div><div class="card stack">
    <label class="f"><span>Apps Script web app URL</span><input type="url" data-ch="localSet" data-k="scriptUrl" data-fk="scriptUrl" value="${esc(L.scriptUrl)}" placeholder="https://script.google.com/macros/s/…/exec"></label>
    <label class="f"><span>App key (the APP_KEY you set in the script)</span><input type="password" data-ch="localSet" data-k="appKey" data-fk="appKey" value="${esc(L.appKey)}" autocomplete="off"></label>
    <div class="row"><button class="btn sm" data-act="testConnection">Test & sync now</button>${L.scriptUrl && L.appKey ? `<button class="btn sm ghost" data-act="copySetupCode">Copy setup code</button>` : ''}${L.lastSync ? `<span class="muted tiny">Last synced ${esc(new Date(L.lastSync).toLocaleString())}</span>` : ''}</div>
    <div class="help"><b>Adding another device or a home-screen icon?</b> Tap <i>Copy setup code</i> here, open the new copy of the app, choose <i>Connect to my Google Sheet</i>, and paste. On iPhone the home-screen app keeps its own storage, separate from Safari, so it needs this once.</div>
    <div class="help">Your data always saves in this browser first. Once connected, every change is copied to your own Google Sheet, so you can use the app on your phone and computer and keep a permanent record. The README explains the 5-minute setup.</div></div></section>

  <section><div class="sec-head"><h2>AI skin designer <span class="muted small">(optional)</span></h2></div><div class="card stack">
    <label class="f"><span>Claude API key</span><input type="password" data-ch="localSet" data-k="claudeKey" data-fk="claudeKey" value="${esc(L.claudeKey)}" placeholder="sk-ant-…" autocomplete="off"></label>
    <div class="help">Without a key, skins are made on your device by reading the colors in your image and keywords in your prompt (“dark”, “pastel”, “playful”, “elegant”…). With a key, Claude looks at the image and your prompt and designs the palette and fonts. The key stays on this device only — it's never written to your Google Sheet.</div></div></section>

  <section><div class="sec-head"><h2>Your data</h2></div><div class="card">
    <div class="set-row"><div class="grow"><b>Back up everything</b><div class="muted tiny">A single file you can restore later</div></div><button class="btn sm ghost" data-act="backup">${IC.dl} Download</button></div>
    <div class="set-row"><div class="grow"><b>Restore from backup</b><div class="muted tiny">Replaces what's here now</div></div><button class="btn sm ghost" data-act="restore">${IC.ul} Choose file</button></div>
    <div class="set-row"><div class="grow"><b>Start over</b><div class="muted tiny">Erase everything on this device</div></div><button class="btn sm danger" data-act="resetAll">Reset</button></div></div></section>

  <section><div class="sec-head"><h2>Put it on your home screen</h2></div><div class="card help">
    <p><b>iPhone:</b> open this page in Safari → Share → <i>Add to Home Screen</i>.</p><p style="margin:0"><b>Android:</b> open in Chrome → ⋮ menu → <i>Add to Home screen</i> (or <i>Install app</i>).</p></div></section>`;
}
function skinCard(sk) {
  const uses = S.templates.filter(t => t.skinId === sk.id).map(t => t.name); const isDef = S.settings.defaultSkinId === sk.id;
  return `<div class="skincard"><div class="skinprev" style="background-color:${sk.vars.bg};${sk.bgImage ? `background-image:linear-gradient(${sk.vars.bg}${Math.round((1 - sk.bgOpacity) * 255).toString(16).padStart(2, '0')},${sk.vars.bg}${Math.round((1 - sk.bgOpacity) * 255).toString(16).padStart(2, '0')}),url('${sk.bgImage}')` : ''}">
    <span class="sp-name" style="font-family:${esc(fontStack(sk.display))};color:${sk.vars.text}">${esc(sk.name)}</span><span class="sp-pill"><i style="background:${sk.vars.accent}"></i><i style="background:${sk.vars.coin}"></i><i style="background:${sk.vars.surface};outline:1px solid ${sk.vars.border}"></i></span></div>
    <div class="tiny muted" style="padding:6px 10px 0;min-height:22px">${isDef ? '<b style="color:var(--accent)">Default</b>' : ''}${isDef && uses.length ? ' · ' : ''}${uses.length ? esc(uses.join(', ')) : ''}${!isDef && !uses.length ? (sk.builtIn ? 'Built-in' : 'Not in use') : ''}</div>
    <div class="sk-actions"><button class="btn sm soft" data-act="applySkin" data-id="${sk.id}">Apply</button><span class="row" style="gap:0">
    ${sk.builtIn ? '' : `<button class="iconbtn sm" data-act="editSkin" data-id="${sk.id}" aria-label="Edit skin">${IC.edit}</button>`}
    <button class="iconbtn sm" data-act="downloadSkin" data-id="${sk.id}" aria-label="Download skin">${IC.dl}</button>
    ${sk.builtIn ? '' : `<button class="iconbtn sm" data-act="deleteSkin" data-id="${sk.id}" aria-label="Delete skin">${IC.x}</button>`}</span></div></div>`;
}

/* ---------- template editor ---------- */
function tplEditor(id) {
  const t = tpl(id); if (!t) { setTimeout(closeModal); return ''; }
  const inT = new Set(t.habitIds); const others = S.habits.filter(h => !inT.has(h.id) && !h.archived);
  const hs = t.habitIds.map(habitById).filter(Boolean);
  return `${modalHead('Day template')}<div class="stack">
    <label class="f"><span>Name</span><input type="text" data-ch="tplField" data-id="${t.id}" data-k="name" data-fk="tplname" value="${esc(t.name)}" maxlength="40"></label>
    <label class="f"><span>Skin for this kind of day</span><select data-ch="tplField" data-id="${t.id}" data-k="skinId">${opt('', `Default skin (${(getSkin(S.settings.defaultSkinId) || BUILTIN_SKINS[0]).name})`, t.skinId || '')}${allSkins().map(s => opt(s.id, s.name, t.skinId)).join('')}</select></label>
    <label class="check-row"><input type="checkbox" data-ch="tplToggle" data-id="${t.id}" data-k="showTasks" ${t.showTasks !== false ? 'checked' : ''}> Include a tasks list on this kind of day</label>
    <label class="check-row" style="padding-top:0"><input type="checkbox" data-ch="tplToggle" data-id="${t.id}" data-k="groupByCategory" ${t.groupByCategory ? 'checked' : ''}> Sort habits under category headings</label>
    <div><div class="spread" style="margin:6px 0 8px"><h3>Habits <span class="muted small">(${hs.length})</span></h3><span class="muted tiny">use arrows to reorder</span></div>
    ${hs.map((h, i) => { const c = catById(h.categoryId); return `<div class="hrow" style="--c:${c ? c.color : 'var(--border)'}"><div class="grow"><b>${esc(h.name)}</b><span class="muted tiny">${h.coins} coin${h.coins > 1 ? 's' : ''} · ${PRI[h.priority]}${c ? ' · ' + esc(c.name) : ''}</span></div>
      <button class="iconbtn sm" data-act="tplMove" data-id="${t.id}" data-i="${i}" data-n="-1" ${i === 0 ? 'disabled' : ''} aria-label="Move up">${IC.up}</button><button class="iconbtn sm" data-act="tplMove" data-id="${t.id}" data-i="${i}" data-n="1" ${i === hs.length - 1 ? 'disabled' : ''} aria-label="Move down">${IC.down}</button>
      <button class="iconbtn sm" data-act="editHabit" data-id="${h.id}" data-back="${t.id}" aria-label="Edit habit">${IC.edit}</button><button class="iconbtn sm" data-act="tplRemoveHabit" data-id="${t.id}" data-h="${h.id}" aria-label="Remove from template">${IC.x}</button></div>`; }).join('') || '<p class="muted small">No habits yet — add your first below.</p>'}</div>
    <form class="card stack" data-submit="tplAddHabit" data-id="${t.id}" style="box-shadow:none;background:var(--bg)"><b>Add a new habit</b>
      <input type="text" name="name" placeholder="e.g. Walk the dogs" required maxlength="60" aria-label="Habit name">
      <div class="grid3"><label class="f"><span>Coins</span><select name="coins">${[1, 2, 3, 4, 5].map(n => opt(n, n, 2)).join('')}</select></label>
      <label class="f"><span>Priority</span><select name="priority">${[3, 2, 1].map(p => opt(p, `${p} · ${PRI[p]}`, 2)).join('')}</select></label>
      <label class="f"><span>Category</span><select name="categoryId">${catOptions('')}</select></label></div>
      <button class="btn" type="submit">${IC.plus} Add habit</button></form>
    ${others.length ? `<div class="row"><select data-fk="tplExisting" id="tplExisting" class="grow" aria-label="Existing habit">${others.map(h => opt(h.id, h.name, '')).join('')}</select><button class="btn sm ghost" data-act="tplAddExisting" data-id="${t.id}">Add existing</button></div>` : ''}
    <hr class="hr"><div class="row"><button class="btn sm ghost" data-act="dupTemplate" data-id="${t.id}">Duplicate template</button><span class="grow"></span><button class="btn sm danger" data-act="deleteTemplate" data-id="${t.id}">Delete template</button></div></div>`;
}
function habitEditor(id, back) {
  const h = habitById(id); if (!h) { setTimeout(closeModal); return ''; }
  return `${modalHead('Habit', back ? { act: 'editTemplate', attrs: `data-id="${back}"` } : { act: 'allHabits' })}<div class="stack">
    <label class="f"><span>Name</span><input type="text" data-ch="habitField" data-id="${h.id}" data-k="name" data-fk="hname" value="${esc(h.name)}" maxlength="60"></label>
    <div><div class="f" style="margin-bottom:5px;font-size:13px;font-weight:600;color:var(--muted)">Coins (1–5)</div>${seg('habitSet', `data-id="${h.id}" data-k="coins"`, [1, 2, 3, 4, 5].map(n => [n, n]), h.coins)}</div>
    <div><div style="margin-bottom:5px;font-size:13px;font-weight:600;color:var(--muted)">Priority</div>${seg('habitSet', `data-id="${h.id}" data-k="priority"`, [[3, '3 · Must'], [2, '2 · Should'], [1, '1 · Optional']], h.priority)}</div>
    <label class="f"><span>Category</span><select data-ch="habitField" data-id="${h.id}" data-k="categoryId">${catOptions(h.categoryId)}</select></label>
    <div><div style="margin-bottom:2px;font-size:13px;font-weight:600;color:var(--muted)">Show on these kinds of day</div>${S.templates.map(t => `<label class="check-row"><input type="checkbox" data-ch="habitInTpl" data-id="${h.id}" data-t="${t.id}" ${t.habitIds.includes(h.id) ? 'checked' : ''}> ${esc(t.name)}</label>`).join('') || '<p class="muted small">No templates yet.</p>'}</div>
    <hr class="hr"><div class="row"><button class="btn sm ghost" data-act="archiveHabit" data-id="${h.id}">Remove from all templates</button><span class="grow"></span><button class="btn sm danger" data-act="deleteHabit" data-id="${h.id}">Delete forever</button></div>
    <p class="help" style="margin:0">Removing keeps its history and coins. Deleting also erases its past check-ins.</p></div>`;
}
function allHabitsModal() {
  return `${modalHead('All habits')}<div class="stack">${S.habits.length ? `<div>${S.habits.map(h => { const c = catById(h.categoryId); const ts = S.templates.filter(t => t.habitIds.includes(h.id)).map(t => t.name);
    return `<div class="hrow" style="--c:${c ? c.color : 'var(--border)'}"><div class="grow"><b>${esc(h.name)}</b><span class="muted tiny">${h.coins} coins · ${PRI[h.priority]} · ${ts.length ? esc(ts.join(', ')) : 'not in any template'}</span></div><button class="iconbtn sm" data-act="editHabit" data-id="${h.id}" aria-label="Edit">${IC.edit}</button></div>`; }).join('')}</div>` : '<p class="muted">No habits yet.</p>'}
    <form class="row" data-submit="newHabitGlobal"><input type="text" name="name" class="grow" placeholder="New habit name" required style="flex:1" aria-label="New habit"><button class="btn sm" type="submit">Add</button></form></div>`;
}
function listEditor(title, coll, addAct) {
  const items = S[coll];
  return `${modalHead(title)}<div class="stack">${items.map(x => `<div class="row" style="flex-wrap:nowrap"><input type="color" value="${esc(x.color)}" data-ch="listField" data-coll="${coll}" data-id="${x.id}" data-k="color" aria-label="Color">
    <input type="text" value="${esc(x.name)}" data-ch="listField" data-coll="${coll}" data-id="${x.id}" data-k="name" data-fk="${coll}-${x.id}" maxlength="40" style="flex:1" aria-label="Name">
    <button class="iconbtn sm" data-act="listDelete" data-coll="${coll}" data-id="${x.id}" aria-label="Delete">${IC.x}</button></div>`).join('') || '<p class="muted small">Nothing here yet.</p>'}
    <form class="row" data-submit="${addAct}" style="flex-wrap:nowrap"><input type="color" name="color" value="${PALETTE[items.length % PALETTE.length]}" aria-label="Color"><input type="text" name="name" placeholder="New name" required maxlength="40" style="flex:1" aria-label="Name"><button class="btn sm" type="submit">Add</button></form></div>`;
}
const PALETTE = ['#5f9e6e', '#4f8fbf', '#c98a3b', '#8a6bbf', '#d0607a', '#3f9c9c', '#b5a133', '#c0533a', '#6d7a8c'];

/* ---------- skin creator / editor ---------- */
let skinDraft = null;
function skinEditor() {
  const d = skinDraft; if (!d) return '';
  const ai = !!L.claudeKey;
  let h = `${modalHead(d.id ? 'Edit skin' : 'Create a skin')}<div class="stack">
    <label class="f"><span>Skin name</span><input type="text" data-in="skinField" data-k="name" data-fk="skname" value="${esc(d.name)}" maxlength="40" placeholder="e.g. Moss & Candlelight"></label>
    <div><div style="margin-bottom:6px;font-size:13px;font-weight:600;color:var(--muted)">Style reference image (JPG or PNG)</div>
      ${d.refImage ? `<img class="thumb" src="${d.refImage}" alt="Style reference">` : ''}
      <label class="btn ghost sm" style="margin-top:8px">${IC.ul} ${d.refImage ? 'Choose a different image' : 'Choose image'}<input type="file" accept="image/png,image/jpeg" data-ch="skinImage" hidden></label></div>
    <label class="f"><span>Describe the feel you want</span><textarea data-in="skinField" data-k="prompt" data-fk="skprompt" maxlength="500" placeholder="e.g. dark and moody, candlelit forest, elegant serif headings">${esc(d.prompt)}</textarea></label>
    <label class="check-row"><input type="checkbox" data-ch="skinUseBg" ${d.useBg ? 'checked' : ''}> Use the image as a soft background</label>
    <button class="btn" data-act="skinGenerate" ${d.busy ? 'disabled' : ''}>${d.busy ? '<span class="spin"></span> Designing…' : `${IC.sparkle} ${d.vars ? 'Regenerate' : 'Generate'} skin${ai ? ' with Claude' : ''}`}</button>
    ${!ai ? '<p class="help" style="margin:0">Made on your device from the image colors and your words. Add a Claude API key in Settings for AI-designed skins.</p>' : ''}`;
  if (d.vars) {
    const sk = previewSkin(d);
    h += `<hr class="hr"><h3>Preview</h3><div class="preview-box" style="${skinStyle(sk)};background:var(--bg)">${sk.bgImage ? `<div class="pv-bg" style="background-image:url('${sk.bgImage}');opacity:${sk.bgOpacity}"></div>` : ''}
      <div class="kicker" style="color:var(--muted)">Tuesday</div><h3>Today's habits</h3>
      <div class="list" style="margin-top:10px"><div class="habit done" style="--c:var(--accent)"><span class="check">${IC.check}</span><span class="hb-main"><span class="hb-name" style="color:var(--text)">Morning pages</span><span class="hb-meta"><span class="tag should">Should</span><span class="streak">${IC.flame}9</span></span></span><span class="hb-coins" style="color:var(--text)">+2.5${coinEl}<em>×1.25</em></span></div>
      <div class="habit" style="--c:var(--coin)"><span class="check"></span><span class="hb-main"><span class="hb-name" style="color:var(--text)">Walk the dogs</span><span class="hb-meta"><span class="tag must">Must</span></span></span><span class="hb-coins" style="color:var(--text)">3${coinEl}</span></div></div>
      <div class="row" style="margin-top:12px"><span class="btn sm">Button</span><span class="chip">Chip</span></div></div>
    <h3>Fine-tune</h3><div class="colorgrid">${SKIN_KEYS.map(k => `<label><input type="color" value="${d.vars[k]}" data-ch="skinVar" data-k="${k}"> ${SKIN_LABELS[k]}</label>`).join('')}</div>
    <div class="grid2"><label class="f"><span>Heading font</span><select data-ch="skinSet" data-k="display">${Object.keys(FONTS).map(f => opt(f, f, d.display)).join('')}</select></label>
      <label class="f"><span>Body font</span><select data-ch="skinSet" data-k="body">${Object.keys(FONTS).map(f => opt(f, f, d.body)).join('')}</select></label></div>
    <div class="grid2"><label class="f"><span>Light or dark</span><select data-ch="skinSet" data-k="mode">${opt('light', 'Light', d.mode)}${opt('dark', 'Dark', d.mode)}</select></label>
      <label class="f"><span>Corner roundness · ${d.radius}px</span><input type="range" min="0" max="28" step="2" value="${d.radius}" data-ch="skinSet" data-k="radius"></label></div>
    ${d.bgImage ? `<label class="f"><span>Background image strength · ${Math.round(d.bgOpacity * 100)}%</span><input type="range" min="0.04" max="0.6" step="0.02" value="${d.bgOpacity}" data-ch="skinSet" data-k="bgOpacity"></label>` : ''}
    <button class="btn block" data-act="skinSave">${d.id ? 'Save changes' : 'Save skin'}</button>`;
  }
  return h + '</div>';
}
function previewSkin(d) { return { ...d, vars: { ...d.vars }, radius: +d.radius, bgOpacity: +d.bgOpacity }; }
function sanitizeSkin(o) {
  if (!o || typeof o !== 'object' || !o.vars) return null;
  const vars = {}; for (const k of SKIN_KEYS) { if (!HEX.test(o.vars[k] || '')) return null; vars[k] = o.vars[k].toLowerCase(); }
  const img = typeof o.bgImage === 'string' && /^data:image\/(jpeg|png|webp);base64,[A-Za-z0-9+/=]+$/.test(o.bgImage) && o.bgImage.length < 60000 ? o.bgImage : '';
  return { id: typeof o.id === 'string' && /^sk_[a-z0-9]+$/i.test(o.id) ? o.id : uid('sk'), name: String(o.name || 'Untitled skin').slice(0, 40), prompt: String(o.prompt || '').slice(0, 500),
    mode: o.mode === 'dark' ? 'dark' : 'light', vars, display: FONTS[o.display] ? o.display : 'Fraunces', body: FONTS[o.body] ? o.body : 'DM Sans',
    radius: clamp(Math.round(+o.radius || 14), 0, 28), bgImage: img, bgOpacity: img ? clamp(+o.bgOpacity || .12, 0, .6) : 0 };
}
function readFile(file) { return new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsDataURL(file); }); }
function loadImg(src) { return new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = src; }); }
async function resizeImage(src, maxW, q) { const img = await loadImg(src); const s = Math.min(1, maxW / Math.max(img.width, img.height)); const c = document.createElement('canvas');
  c.width = Math.max(1, Math.round(img.width * s)); c.height = Math.max(1, Math.round(img.height * s)); const x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); x.drawImage(img, 0, 0, c.width, c.height); return c.toDataURL('image/jpeg', q); }
async function compressForBg(src) { for (const [w, q] of [[760, .62], [620, .55], [480, .5], [360, .45], [260, .4]]) { const d = await resizeImage(src, w, q); if (d.length < 45000) return d; } return ''; }
async function imagePalette(src) {
  const img = await loadImg(src); const c = document.createElement('canvas'); c.width = c.height = 72; const x = c.getContext('2d', { willReadFrequently: true }); x.drawImage(img, 0, 0, 72, 72);
  const px = x.getImageData(0, 0, 72, 72).data; const B = new Map();
  for (let i = 0; i < px.length; i += 4) { if (px[i + 3] < 128) continue; const k = (px[i] >> 4) << 8 | (px[i + 1] >> 4) << 4 | (px[i + 2] >> 4); const b = B.get(k) || [0, 0, 0, 0]; b[0] += px[i]; b[1] += px[i + 1]; b[2] += px[i + 2]; b[3]++; B.set(k, b); }
  const cand = [...B.values()].sort((a, b) => b[3] - a[3]).map(b => ({ rgb: [b[0] / b[3], b[1] / b[3], b[2] / b[3]], count: b[3] }));
  const out = []; for (const cc of cand) { const near = out.find(o => Math.hypot(...o.rgb.map((v, i) => v - cc.rgb[i])) < 46); if (near) near.count += cc.count; else out.push({ ...cc }); if (out.length >= 10) break; }
  return out.map(o => { const [h, s, l] = rgbToHsl(...o.rgb); return { hex: rgbToHex(...o.rgb), count: o.count, h, s, l }; }).sort((a, b) => b.count - a.count);
}
function skinFromPalette(pal, prompt) {
  const p = (prompt || '').toLowerCase(); const has = re => re.test(p);
  const dom = pal[0] || { h: 30, s: 20, l: 80 };
  const mode = has(/\b(dark|night|moody|noir|midnight|gothic|witch\w*|shadow\w*|deep|candle\w*)\b/) ? 'dark' : has(/\b(light|bright|airy|pastel|soft|sunny|daylight|clean|white|fresh)\b/) ? 'light' : (pal.slice(0, 3).reduce((a, c) => a + c.l * c.count, 0) / Math.max(1, pal.slice(0, 3).reduce((a, c) => a + c.count, 0)) < 42 ? 'dark' : 'light');
  const dark = mode === 'dark';
  const vivid = pal.filter(c => c.s > 22 && c.l > 15 && c.l < 88).sort((a, b) => b.s * Math.sqrt(b.count) - a.s * Math.sqrt(a.count));
  const acc = vivid[0] || pal.find(c => c !== dom) || dom; const coinSrc = vivid.find(c => hueDist(c.h, acc.h) > 35) || { h: 42, s: 72, l: 52 };
  let sat = has(/\b(pastel|muted|soft|calm|gentle|dusty)\b/) ? .65 : has(/\b(bold|vibrant|neon|bright|electric|saturated)\b/) ? 1.25 : 1;
  const bh = dom.h, bs = Math.min(dom.s, 38);
  const v = {
    bg: dark ? hsl(bh, bs * .55, 8.5) : hsl(bh, bs * .7, 94), surface: dark ? hsl(bh, bs * .55, 12.5) : hsl(bh, bs * .55, 98.5),
    surface2: dark ? hsl(bh, bs * .55, 18) : hsl(bh, bs * .6, 89.5), border: dark ? hsl(bh, bs * .45, 23) : hsl(bh, bs * .45, 84),
    text: dark ? hsl(bh, 14, 92) : hsl(bh, 20, 13), muted: dark ? hsl(bh, 12, 66) : hsl(bh, 14, 40),
  };
  v.accent = ensureContrast(hsl(acc.h, clamp(acc.s * sat, 38, 88), clamp(acc.l, dark ? 55 : 34, dark ? 72 : 52)), v.surface, 3.2);
  v.coin = ensureContrast(hsl(coinSrc.h, clamp(coinSrc.s * sat, 45, 85), dark ? 62 : 46), v.surface, 2.4);
  v.text = ensureContrast(v.text, v.bg, 9); v.muted = ensureContrast(v.muted, v.surface, 4.5);
  const font = has(/\b(elegant|classic|literary|vintage|victorian|gothic|witch\w*|romantic|old)\b/) ? ['Cormorant Garamond', 'Lora'] : has(/\b(playful|fun|cute|bubbly|kid\w*|whimsical)\b/) ? ['Fredoka', 'Nunito']
    : has(/\b(modern|minimal\w*|clean|tech\w*|sleek|futur\w*)\b/) ? ['Space Grotesk', 'Inter'] : has(/\b(handwrit\w*|journal|cozy|scrapbook|sketch\w*)\b/) ? ['Caveat', 'Nunito'] : has(/\b(editorial|magazine|luxur\w*|fashion)\b/) ? ['Playfair Display', 'Work Sans']
    : has(/\b(retro|typewriter|terminal|code|hacker)\b/) ? ['IBM Plex Mono', 'Work Sans'] : ['Fraunces', 'DM Sans'];
  const radius = has(/\b(sharp|brutal\w*|geometric|angular|square)\b/) ? 4 : has(/\b(round\w*|soft|bubbly|cute|pill)\b/) ? 22 : 14;
  return { mode, vars: v, display: font[0], body: font[1], radius };
}
async function aiSkin(src, prompt) {
  const small = await resizeImage(src, 512, .8);
  const fonts = Object.keys(FONTS).join(', ');
  const text = `Design a color theme for a habit-tracking app, inspired by this reference image${prompt ? ` and this description: "${prompt}"` : ''}.
Reply with ONLY a JSON object, no prose:
{"name": short evocative name (max 4 words), "mode": "light" or "dark", "bg": hex, "surface": hex (cards), "surface2": hex (subtle fills), "text": hex, "muted": hex (secondary text), "accent": hex (buttons, checkmarks), "coin": hex (gold-ish reward color that suits the palette), "border": hex, "display": heading font, "body": body font, "radius": corner radius 0-28}
Fonts must be chosen from: ${fonts}.
Text must be very readable on bg and surface (contrast 7:1 or better); accent must stand out on surface.`;
  const r = await fetch('https://api.anthropic.com/v1/messages', { method: 'POST',
    headers: { 'content-type': 'application/json', 'x-api-key': L.claudeKey, 'anthropic-version': '2023-06-01', 'anthropic-dangerous-direct-browser-access': 'true' },
    body: JSON.stringify({ model: AI_MODEL, max_tokens: 600, messages: [{ role: 'user', content: [{ type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: small.split(',')[1] } }, { type: 'text', text }] }] }) });
  if (!r.ok) throw new Error('Claude API error ' + r.status);
  const j = await r.json(); const txt = (j.content || []).map(c => c.text || '').join(''); const m = txt.match(/\{[\s\S]*\}/); if (!m) throw new Error('No JSON');
  const o = JSON.parse(m[0]); const vars = {}; for (const k of SKIN_KEYS) { if (!HEX.test(o[k] || '')) throw new Error('Bad color ' + k); vars[k] = o[k].toLowerCase(); }
  vars.text = ensureContrast(vars.text, vars.bg, 7); vars.muted = ensureContrast(vars.muted, vars.surface, 4.5); vars.accent = ensureContrast(vars.accent, vars.surface, 3);
  return { mode: o.mode === 'dark' ? 'dark' : 'light', vars, display: FONTS[o.display] ? o.display : 'Fraunces', body: FONTS[o.body] ? o.body : 'DM Sans', radius: clamp(+o.radius || 14, 0, 28), aiName: typeof o.name === 'string' ? o.name.slice(0, 40) : '' };
}

/* =====================================================================
   GOOGLE SHEETS SYNC
   ===================================================================== */
let syncTimer = null, syncState = 'local', lastPull = 0, syncBusy = false, syncHold = false;
/* How sync stays safe:
   - L.remoteAt = the sheet's version (Meta.updatedAt) this device last saw. L.dirty = this device has unsaved edits.
   - Saves send baseAt; the sheet refuses them if another device saved in between, and the two copies are merged.
   - An empty or clearly smaller copy never silently replaces a fuller one: the app asks, with "combine both" as the safe choice. */
function updateSyncUI() {
  const b = $('#syncBtn'); if (!b) return;
  const st = !L.scriptUrl ? 'local' : syncState; b.dataset.s = st;
  $('#syncLabel').textContent = { local: 'On this device', pending: 'Saving…', saving: 'Saving…', synced: 'Synced', error: 'Sync problem', pulling: 'Checking…', waiting: 'Needs your choice' }[st] || 'Synced';
}
function setSync(s) { syncState = s; updateSyncUI(); }
function scheduleSync() { if (!L.scriptUrl) return; L.dirty = true; saveLocal(); if (syncHold) return; setSync('pending'); clearTimeout(syncTimer); syncTimer = setTimeout(() => pushRemote(), 2200); }
async function api(action, extra = {}) {
  const res = await fetch(L.scriptUrl, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ action, key: L.appKey, ...extra }), redirect: 'follow' });
  if (!res.ok) throw new Error('HTTP ' + res.status);
  const j = await res.json(); if (j.v && L.serverV !== j.v) { L.serverV = j.v; saveLocal(); } if (!j.ok) { const e = new Error(j.error || 'Request failed'); e.conflict = !!j.conflict; e.refusedEmpty = !!j.refusedEmpty; throw e; } return j;
}
const dayCount = s => Object.values(s.days).filter(d => d.templateId || Object.keys(d.habits).length || d.tasks.length).length;
const isEmptyState = s => !s.habits.length && !s.tasks.length && !s.templates.length && !dayCount(s);
const plural = (n, w) => `${n} ${w}${n === 1 ? '' : 's'}`;
const sizeOf = s => `${plural(s.habits.length, 'habit')}, ${plural(s.tasks.length, 'task')}, ${plural(dayCount(s), 'day')} logged`;
// Would taking `remote` lose things this device has? (days are never deleted, so fewer days = something went wrong)
function remoteLooksSmaller(remote) {
  if (isEmptyState(S)) return false;
  if (dayCount(remote) < dayCount(S)) return true;
  const n = x => x.habits.length + x.tasks.length + x.templates.length;
  return n(remote) < n(S) * 0.6;
}
function mergeStates(local, remote) {
  const out = normalize(JSON.parse(JSON.stringify(remote)));
  for (const k of ['categories', 'habits', 'templates', 'skins', 'taskGroups', 'tasks']) {
    const byId = new Map(out[k].map(x => [x.id, x]));
    for (const x of local[k]) byId.set(x.id, JSON.parse(JSON.stringify(x)));
    const order = [...local[k].map(x => x.id), ...out[k].map(x => x.id).filter(id => !local[k].some(y => y.id === id))];
    out[k] = order.map(id => byId.get(id)).filter(Boolean);
  }
  for (const [date, ld] of Object.entries(local.days)) {
    const rd = out.days[date];
    if (!rd) { out.days[date] = JSON.parse(JSON.stringify(ld)); continue; }
    rd.templateId = ld.templateId || rd.templateId;
    if (Array.isArray(ld.habitIds)) rd.habitIds = ld.habitIds.slice();
    rd.habits = { ...rd.habits, ...ld.habits };
    const tasks = new Map(rd.tasks.map(t => [t.taskId, t])); for (const t of ld.tasks) { const r = tasks.get(t.taskId); tasks.set(t.taskId, r ? { ...t, done: t.done || r.done } : { ...t }); }
    rd.tasks = [...tasks.values()];
  }
  out.settings = { ...remote.settings, ...local.settings };
  return out;
}
function adoptState(next, msg) {
  S = normalize(next); saveState(); _scKey = ''; L.remoteAt = S.meta.updatedAt; L.dirty = false; L.lastSync = Date.now(); saveLocal();
  setSync('synced'); render(); if (msg) toast(msg);
}
async function pushRemote({ force = false, retried = false } = {}) {
  if (!L.scriptUrl || syncHold) return; clearTimeout(syncTimer);
  if (isEmptyState(S) && !force) { L.dirty = false; saveLocal(); setSync('synced'); return; } // never save "nothing" over a sheet
  setSync('saving');
  try {
    // An older backend script can't detect conflicts, so check the sheet ourselves first.
    if (L.serverV !== 2 && !force && typeof L.remoteAt === 'number') {
      const j = await api('load'); const r = j.data && j.data.Meta ? fromSheets(j.data) : null;
      if (r && !isEmptyState(r) && r.meta.updatedAt !== L.remoteAt) { S = mergeStates(S, r); S.meta.updatedAt = Date.now(); saveState(); _scKey = ''; L.remoteAt = r.meta.updatedAt; saveLocal(); render(); }
    }
    const sentAt = S.meta.updatedAt;
    await api('save', { data: toSheets(), baseAt: typeof L.remoteAt === 'number' ? L.remoteAt : null, force });
    L.remoteAt = sentAt; L.lastSync = Date.now(); if (S.meta.updatedAt === sentAt) L.dirty = false; saveLocal(); setSync(L.dirty ? 'pending' : 'synced');
    if (L.dirty) scheduleSync();
  } catch (e) {
    console.warn(e);
    if (e.conflict && !retried) { lastPull = 0; return pullRemote({ force: true }); }   // someone else saved first → merge, then save
    if (e.refusedEmpty) { setSync('synced'); return; }
    setSync('error');
  }
}
async function pullRemote({ force = false, announce = false } = {}) {
  if (!L.scriptUrl || syncHold || syncBusy || (!force && Date.now() - lastPull < 30000)) return;
  lastPull = Date.now(); syncBusy = true; setSync('pulling');
  try {
    const j = await api('load'); const remote = j.data && j.data.Meta ? fromSheets(j.data) : null;
    syncBusy = false;
    if (!remote || isEmptyState(remote)) {                              // sheet is empty → fill it from here (if we have anything)
      if (!isEmptyState(S)) { await pushRemote({ force: true }); if (announce) toast('Connected — your data is now in your sheet.'); }
      else { L.lastSync = Date.now(); saveLocal(); setSync('synced'); if (announce) toast('Connected. Your sheet is empty so far.'); }
      return;
    }
    if (isEmptyState(S)) return adoptState(remote, 'Connected — your habits and tasks are loaded.'), closeModal();
    if (remote.meta.updatedAt === S.meta.updatedAt) { L.remoteAt = remote.meta.updatedAt; L.lastSync = Date.now(); L.dirty = false; saveLocal(); setSync('synced'); if (announce) toast('Everything is up to date.'); return; }
    const firstTime = !L.lastSync || typeof L.remoteAt !== 'number';
    const sheetChanged = remote.meta.updatedAt !== L.remoteAt;
    if (firstTime) return askSync(remote, 'first');
    if (!sheetChanged) { if (L.dirty || S.meta.updatedAt > remote.meta.updatedAt) await pushRemote(); else setSync('synced'); return; }
    // the sheet was changed by another device
    if (!L.dirty) {
      if (remoteLooksSmaller(remote)) return askSync(remote, 'smaller');
      return adoptState(remote, 'Updated with changes from your other device.');
    }
    // both sides changed → combine, then save the combined copy
    S = mergeStates(S, remote); S.meta.updatedAt = Date.now(); saveState(); _scKey = ''; L.remoteAt = remote.meta.updatedAt; saveLocal(); render();
    await pushRemote({ retried: true });
  } catch (e) { syncBusy = false; console.warn(e); setSync('error'); if (announce) toast('Could not reach your sheet: ' + e.message, 5000); }
}
let pendingRemote = null;
function askSync(remote, why) { pendingRemote = remote; syncHold = true; clearTimeout(syncTimer); setSync('waiting'); openModal(() => syncChoiceModal(why)); }
function syncChoiceModal(why) {
  const r = pendingRemote; if (!r) return '';
  return `<div class="mhead"><h2 class="grow">Which data should we keep?</h2></div><div class="stack">
    <p class="small">${why === 'first' ? 'This device and your Google Sheet both have data, and they don’t match yet.' : 'Your Google Sheet has <b>less</b> data than this device — it may have been overwritten by another tab or device.'} Nothing has been changed yet.</p>
    <div class="card" style="box-shadow:none"><div class="set-row"><div class="grow"><b>This device</b><div class="muted tiny">${sizeOf(S)}</div></div></div><div class="set-row"><div class="grow"><b>Google Sheet</b><div class="muted tiny">${sizeOf(r)}</div></div></div></div>
    <button class="btn block" data-act="syncChoose" data-v="merge">Combine both (nothing is lost)</button>
    <button class="btn ghost block" data-act="syncChoose" data-v="remote">Use the Google Sheet’s data on this device</button>
    <button class="btn danger block" data-act="syncChoose" data-v="local">Replace the Google Sheet with this device’s data</button>
    <p class="help" style="margin:0">Not sure? Choose <b>Combine both</b>. If you combine the example setup with your real data you can delete the example habits afterwards.</p></div>`;
}
async function resolveSync(choice) {
  const r = pendingRemote; if (!r) return closeModal();
  if (choice === 'local' && !confirm(`Replace your Google Sheet (${sizeOf(r)}) with this device’s data (${sizeOf(S)})?`)) return;
  pendingRemote = null; syncHold = false; closeModal();
  if (choice === 'remote') return adoptState(r, 'Loaded your data from Google Sheets.');
  if (choice === 'merge') { S = mergeStates(S, r); S.meta.updatedAt = Date.now(); saveState(); _scKey = ''; L.remoteAt = r.meta.updatedAt; L.dirty = true; saveLocal(); render(); await pushRemote({ retried: true }); toast('Combined — everything from both is kept.'); return; }
  L.remoteAt = r.meta.updatedAt; L.dirty = true; saveLocal(); await pushRemote({ force: true }); toast('Your sheet now matches this device.');
}
function toSheets() {
  const out = {}; const put = (name, headers, rows) => out[name] = { headers, rows: rows.map(r => headers.map(h => r[h] === undefined || r[h] === null ? '' : r[h])) };
  const sc = scores();
  put('Meta', ['key', 'value'], [{ key: 'updatedAt', value: S.meta.updatedAt }, { key: 'version', value: 1 }, { key: 'settings', value: JSON.stringify(S.settings) }]);
  put('Categories', ['id', 'name', 'color'], S.categories);
  put('Habits', ['id', 'name', 'coins', 'priority', 'categoryId', 'archived'], S.habits.map(h => ({ ...h, archived: h.archived ? 1 : 0 })));
  put('Templates', ['id', 'name', 'skinId', 'showTasks', 'groupByCategory', 'habitIds'], S.templates.map(t => ({ ...t, showTasks: t.showTasks !== false ? 1 : 0, groupByCategory: t.groupByCategory ? 1 : 0, habitIds: JSON.stringify(t.habitIds) })));
  put('Skins', ['id', 'name', 'prompt', 'mode', 'vars', 'display', 'body', 'radius', 'bgImage', 'bgOpacity'], S.skins.map(s => ({ ...s, vars: JSON.stringify(s.vars) })));
  put('TaskGroups', ['id', 'name', 'color'], S.taskGroups);
  put('Tasks', ['id', 'title', 'groupId', 'groupName', 'due', 'priority', 'notes', 'status', 'createdAt', 'closedAt'], S.tasks.map(t => ({ ...t, groupName: groupById(t.groupId)?.name || '' })));
  const dates = Object.keys(S.days).sort();
  put('DayLog', ['date', 'templateId', 'templateName', 'habitIds'], dates.map(d => ({ date: d, templateId: S.days[d].templateId || '', templateName: tpl(S.days[d].templateId)?.name || '', habitIds: S.days[d].habitIds ? JSON.stringify(S.days[d].habitIds) : '' })));
  put('HabitLog', ['date', 'habitId', 'habitName', 'completedAt'], dates.flatMap(d => Object.entries(S.days[d].habits).map(([id, at]) => ({ date: d, habitId: id, habitName: habitById(id)?.name || '', completedAt: at }))));
  put('DayTasks', ['date', 'taskId', 'taskTitle', 'coins', 'done'], dates.flatMap(d => S.days[d].tasks.map(t => ({ date: d, taskId: t.taskId, taskTitle: taskById(t.taskId)?.title || '', coins: t.coins, done: t.done ? 1 : 0 }))));
  // read-only summaries for browsing the sheet directly
  put('DailySummary', ['date', 'template', 'habitsDone', 'habitsScheduled', 'habitCoins', 'taskCoins', 'penalties', 'totalCoins'], Object.entries(sc.perDay).sort().map(([d, p]) => ({ date: d, template: tpl(S.days[d].templateId)?.name || '', habitsDone: p.done, habitsScheduled: p.scheduled, habitCoins: p.habitCoins, taskCoins: p.taskCoins, penalties: p.penalty, totalCoins: p.total })));
  put('HabitStats', ['habit', 'category', 'coins', 'priority', 'currentStreak', 'bestStreak', 'multiplier', 'timesDone', 'timesScheduled', 'completionPct', 'coinsEarned', 'coinsLost'], S.habits.map(h => { const x = sc.habit[h.id]; return { habit: h.name, category: catById(h.categoryId)?.name || '', coins: h.coins, priority: PRI[h.priority], currentStreak: x.streak, bestStreak: x.best, multiplier: x.multNow, timesDone: x.done, timesScheduled: x.eligible, completionPct: x.eligible ? Math.round(x.done / x.eligible * 100) : '', coinsEarned: round2(x.earned), coinsLost: x.penalty }; }));
  return out;
}
function fromSheets(D) {
  const rows = n => Array.isArray(D[n]) ? D[n] : []; const str = v => v === null || v === undefined ? '' : String(v); const num = v => +v || 0; const J = (v, d) => { try { return v ? JSON.parse(v) : d; } catch { return d; } };
  const meta = Object.fromEntries(rows('Meta').map(r => [r.key, r.value]));
  const s = defaultState(); s.meta.updatedAt = num(meta.updatedAt); s.settings = { ...s.settings, ...J(str(meta.settings), {}) };
  s.categories = rows('Categories').map(r => ({ id: str(r.id), name: str(r.name), color: str(r.color) || '#888888' })).filter(r => r.id);
  s.habits = rows('Habits').map(r => ({ id: str(r.id), name: str(r.name), coins: num(r.coins), priority: num(r.priority), categoryId: str(r.categoryId), archived: !!num(r.archived) })).filter(r => r.id);
  s.templates = rows('Templates').map(r => ({ id: str(r.id), name: str(r.name), skinId: str(r.skinId), showTasks: !!num(r.showTasks), groupByCategory: !!num(r.groupByCategory), habitIds: J(str(r.habitIds), []) })).filter(r => r.id);
  s.skins = rows('Skins').map(r => ({ id: str(r.id), name: str(r.name), prompt: str(r.prompt), mode: str(r.mode), vars: J(str(r.vars), null), display: str(r.display), body: str(r.body), radius: num(r.radius), bgImage: str(r.bgImage), bgOpacity: num(r.bgOpacity) }));
  s.taskGroups = rows('TaskGroups').map(r => ({ id: str(r.id), name: str(r.name), color: str(r.color) || '#888888' })).filter(r => r.id);
  s.tasks = rows('Tasks').map(r => ({ id: str(r.id), title: str(r.title), groupId: str(r.groupId), due: str(r.due), priority: num(r.priority), notes: str(r.notes), status: str(r.status), createdAt: num(r.createdAt), closedAt: str(r.closedAt) })).filter(r => r.id);
  for (const r of rows('DayLog')) { const d = str(r.date); if (!isDateStr(d)) continue; s.days[d] = { templateId: str(r.templateId), habits: {}, tasks: [] }; const hi = J(str(r.habitIds), null); if (Array.isArray(hi)) s.days[d].habitIds = hi; }
  for (const r of rows('HabitLog')) { const d = str(r.date); if (!isDateStr(d)) continue; (s.days[d] = s.days[d] || { templateId: '', habits: {}, tasks: [] }).habits[str(r.habitId)] = num(r.completedAt) || 1; }
  for (const r of rows('DayTasks')) { const d = str(r.date); if (!isDateStr(d)) continue; (s.days[d] = s.days[d] || { templateId: '', habits: {}, tasks: [] }).tasks.push({ taskId: str(r.taskId), coins: num(r.coins), done: !!num(r.done) }); }
  return s;
}

/* =====================================================================
   ACTIONS
   ===================================================================== */
const ACT = {
  go: el => { V.page = el.dataset.page; render(); window.scrollTo({ top: 0 }); },
  shiftDay: el => { V.date = addDays(V.date, +el.dataset.n); render(); },
  goToday: () => { V.date = todayStr(); render(); },
  closeModal: () => closeModal(),
  syncClick: () => { if (!L.scriptUrl) { V.page = 'settings'; render(); setTimeout(() => $('#sheets')?.scrollIntoView({ behavior: 'smooth' }), 50); } else pullRemote({ force: true, announce: true }); },
  chooseTemplate: el => { const t = tpl(el.dataset.id); commit(s => { const d = ensureDay(s, V.date); d.templateId = t.id; d.habitIds = t.habitIds.slice(); }); },
  changeTemplate: () => commit(s => { const d = s.days[V.date]; if (d) { d.templateId = ''; delete d.habitIds; } }),
  toggleHabit: el => { if (V.date > todayStr()) return; const id = el.dataset.id; commit(s => { const d = s.days[V.date]; if (d.habits[id]) delete d.habits[id]; else d.habits[id] = Date.now(); }); },
  toggleDayTask: el => { const id = el.dataset.id; commit(s => { const e = s.days[V.date].tasks.find(x => x.taskId === id); const t = s.tasks.find(x => x.id === id); e.done = !e.done; if (t) { t.status = e.done ? 'closed' : 'open'; t.closedAt = e.done ? V.date : ''; } }); },
  removeDayTask: el => commit(s => { const d = s.days[V.date]; d.tasks = d.tasks.filter(x => x.taskId !== el.dataset.id); }),
  addTasksToDay: () => openModal(dayTaskPicker),
  pickToDay: el => { const id = el.dataset.id; const coins = +(document.querySelector(`[data-coinsel="${id}"]`)?.value || 3);
    commit(s => { const d = ensureDay(s, V.date); if (!d.tasks.some(x => x.taskId === id)) d.tasks.push({ taskId: id, coins, done: false }); }); toast('Added to ' + (V.date === todayStr() ? 'today' : shortD(V.date))); },
  taskToToday: el => openModal(() => todayCoinPicker(el.dataset.id)),
  confirmToToday: el => { const id = el.dataset.id; const coins = +el.dataset.v; const today = todayStr();
    commit(s => { const d = ensureDay(s, today); if (!d.tasks.some(x => x.taskId === id)) d.tasks.push({ taskId: id, coins, done: false }); }); closeModal(); toast(`Added to today's list · ${coins} coin${coins > 1 ? 's' : ''}`); },
  taskStatus: el => { V.taskStatus = el.dataset.v; render(); },
  toggleTaskClosed: el => commit(s => { const t = s.tasks.find(x => x.id === el.dataset.id); const closing = t.status === 'open'; t.status = closing ? 'closed' : 'open'; t.closedAt = closing ? todayStr() : '';
    for (const d of Object.values(s.days)) for (const e of d.tasks) if (e.taskId === t.id && !closing) e.done = false; }),
  editTask: el => openModal(() => taskEditor(el.dataset.id)),
  deleteTask: el => { if (!confirm('Delete this task everywhere, including from past days?')) return; const id = el.dataset.id; closeModal(); commit(s => { s.tasks = s.tasks.filter(t => t.id !== id); for (const d of Object.values(s.days)) d.tasks = d.tasks.filter(x => x.taskId !== id); }); },
  manageGroups: () => openModal(() => listEditor('Task groups', 'taskGroups', 'addGroup')),
  manageCategories: () => openModal(() => listEditor('Habit categories', 'categories', 'addCategory')),
  listDelete: el => { const { coll, id } = el.dataset; if (!confirm('Delete this? Items using it will become unsorted.')) return;
    commit(s => { s[coll] = s[coll].filter(x => x.id !== id); if (coll === 'categories') s.habits.forEach(h => { if (h.categoryId === id) h.categoryId = ''; }); else s.tasks.forEach(t => { if (t.groupId === id) t.groupId = ''; }); }); },
  newTemplate: () => { const id = uid('tp'); commit(s => s.templates.push({ id, name: `Day type ${s.templates.length + 1}`, skinId: '', showTasks: true, groupByCategory: false, habitIds: [] })); openModal(() => tplEditor(id)); setTimeout(() => { const i = $('[data-fk="tplname"]'); i?.focus(); i?.select(); }, 60); },
  editTemplate: el => openModal(() => tplEditor(el.dataset.id)),
  tplMove: el => { const i = +el.dataset.i, n = +el.dataset.n; commit(s => { const t = s.templates.find(x => x.id === el.dataset.id); const ids = t.habitIds.filter(id => s.habits.some(h => h.id === id)); const j = i + n; if (j < 0 || j >= ids.length) return; [ids[i], ids[j]] = [ids[j], ids[i]]; t.habitIds = ids; refreshSnapshots(s, t.id); }); },
  tplRemoveHabit: el => commit(s => { const t = s.templates.find(x => x.id === el.dataset.id); t.habitIds = t.habitIds.filter(h => h !== el.dataset.h); refreshSnapshots(s, t.id); }),
  tplAddExisting: el => { const hid = $('#tplExisting')?.value; if (!hid) return; commit(s => { const t = s.templates.find(x => x.id === el.dataset.id); if (!t.habitIds.includes(hid)) t.habitIds.push(hid); refreshSnapshots(s, t.id); }); },
  dupTemplate: el => { const t = tpl(el.dataset.id); const id = uid('tp'); commit(s => s.templates.push({ ...t, id, name: t.name + ' copy', habitIds: t.habitIds.slice() })); openModal(() => tplEditor(id)); },
  deleteTemplate: el => { const t = tpl(el.dataset.id); if (!confirm(`Delete the “${t.name}” template? Past days keep their check-ins.`)) return; closeModal();
    commit(s => { s.templates = s.templates.filter(x => x.id !== t.id); const today = todayStr(); for (const [d, day] of Object.entries(s.days)) if (day.templateId === t.id && d >= today && !Object.keys(day.habits).length) { day.templateId = ''; delete day.habitIds; } }); },
  editHabit: el => { const back = el.dataset.back; openModal(() => habitEditor(el.dataset.id, back)); },
  habitSet: el => commit(s => { const h = s.habits.find(x => x.id === el.dataset.id); h[el.dataset.k] = +el.dataset.v; }),
  archiveHabit: el => commit(s => { s.templates.forEach(t => { if (t.habitIds.includes(el.dataset.id)) { t.habitIds = t.habitIds.filter(x => x !== el.dataset.id); refreshSnapshots(s, t.id); } }); }),
  deleteHabit: el => { const h = habitById(el.dataset.id); if (!confirm(`Delete “${h.name}” and all of its past check-ins? This can't be undone.`)) return; closeModal();
    commit(s => { s.habits = s.habits.filter(x => x.id !== h.id); s.templates.forEach(t => t.habitIds = t.habitIds.filter(x => x !== h.id)); for (const d of Object.values(s.days)) { delete d.habits[h.id]; if (d.habitIds) d.habitIds = d.habitIds.filter(x => x !== h.id); } }); },
  allHabits: () => openModal(allHabitsModal),
  setDefaultSkin: el => commit(s => { s.settings.defaultSkinId = el.dataset.id; }),
  applySkin: el => openModal(() => applySkinModal(el.dataset.id)),
  newSkin: () => { skinDraft = { name: '', prompt: '', useBg: true, refImage: '', vars: null, radius: 14, bgOpacity: .12 }; openModal(skinEditor); },
  editSkin: el => { const sk = getSkin(el.dataset.id); skinDraft = { ...sk, vars: { ...sk.vars }, useBg: !!sk.bgImage, refImage: sk.bgImage || '' }; openModal(skinEditor); },
  skinGenerate: async () => {
    const d = skinDraft; if (!d.refImage) return toast('Choose a style reference image first.'); d.busy = true; rerenderModal();
    let res = null; if (L.claudeKey) { try { res = await aiSkin(d.refImage, d.prompt); } catch (e) { console.warn(e); toast('Claude skin design failed — made one on-device instead.', 3500); } }
    if (!res) res = skinFromPalette(await imagePalette(d.refImage), d.prompt);
    if (!skinDraft) return; Object.assign(d, res); if (!d.name) d.name = res.aiName || (d.prompt ? d.prompt.split(/[,.]/)[0].slice(0, 28) : 'My skin');
    d.bgImage = d.useBg ? await compressForBg(d.refImage) : ''; if (d.bgImage && !d.bgOpacity) d.bgOpacity = .12; if (d.bgImage) d.bgOpacity = d.bgOpacity || (d.mode === 'dark' ? .16 : .12);
    d.busy = false; loadFont(d.display); loadFont(d.body); rerenderModal();
  },
  skinSave: () => { const sk = sanitizeSkin({ ...skinDraft, id: skinDraft.id || uid('sk') }); if (!sk) return toast('Something is off with this skin’s colors.');
    const isNew = !S.skins.some(x => x.id === sk.id); commit(s => { const i = s.skins.findIndex(x => x.id === sk.id); if (i >= 0) s.skins[i] = sk; else s.skins.push(sk); });
    if (isNew) openModal(() => applySkinModal(sk.id)); else { closeModal(); toast('Skin saved.'); } skinDraft = null; },
  downloadSkin: el => { const sk = getSkin(el.dataset.id); const { id, builtIn, ...rest } = sk; download(`${slug(sk.name)}.habitude-skin.json`, JSON.stringify({ format: 'habitude-skin', version: 1, skin: rest }, null, 1), 'application/json'); },
  uploadSkin: () => $('#fileSkin').click(),
  deleteSkin: el => { const sk = getSkin(el.dataset.id); if (!confirm(`Delete the “${sk.name}” skin?`)) return; commit(s => { s.skins = s.skins.filter(x => x.id !== sk.id); s.templates.forEach(t => { if (t.skinId === sk.id) t.skinId = ''; }); if (s.settings.defaultSkinId === sk.id) s.settings.defaultSkinId = 'sk_parchment'; }); },
  skinTplToggle: el => commit(s => { const t = s.templates.find(x => x.id === el.dataset.t); t.skinId = t.skinId === el.dataset.id ? '' : el.dataset.id; }),
  testConnection: () => { if (!L.scriptUrl) return toast('Paste your Apps Script web app URL first.'); pullRemote({ force: true, announce: true }); },
  backup: () => download(`habitude-backup-${todayStr()}.json`, JSON.stringify({ format: 'habitude-backup', ...S }), 'application/json'),
  restore: () => $('#fileBackup').click(),
  resetAll: () => { if (!confirm('Erase all habits, tasks, templates and history on this device?')) return; if (!confirm('Really erase everything? (Your Google Sheet keeps its copy until the next sync.)')) return;
    S = defaultState(); saveState(); _scKey = ''; L.dirty = false; delete L.remoteAt; saveLocal(); V.date = todayStr(); render(); openModal(welcomeModal); },
  syncChoose: el => resolveSync(el.dataset.v),
  openConnect: () => openModal(connectModal),
  backToWelcome: () => openModal(welcomeModal),
  copySetupCode: async () => { const code = 'HBT1:' + btoa(unescape(encodeURIComponent(JSON.stringify({ u: L.scriptUrl, k: L.appKey }))));
    try { await navigator.clipboard.writeText(code); toast('Setup code copied — paste it into the other copy of the app.', 3500); }
    catch { prompt('Copy this setup code:', code); } },
  seedExample: () => { S = exampleState(); S.meta.updatedAt = Date.now(); saveState(); closeModal(); render(); if (L.scriptUrl) pullRemote({ force: true }); },
  startBlank: () => { closeModal(); render(); },
  expAll: el => { const v = el.dataset.v === '1'; DATASETS.forEach(([id]) => exportModal.st.ids[id] = v); rerenderModal(); },
  exportFull: el => openModal(() => exportModal(el.dataset.kind, false)),
  exportPartial: el => openModal(() => exportModal(el.dataset.kind, true)),
  runExport: el => runExport(el.dataset.kind, el.dataset.partial === '1'),
};
const CH = {
  pickDate: el => { if (isDateStr(el.value)) { V.date = el.value; render(); } },
  taskSort: el => { V.taskSort = el.value; render(); },
  taskGroup: el => { V.taskGroup = el.value; render(); },
  q: el => { V.q[el.dataset.k] = el.value; render(); },
  setting: el => commit(s => { s.settings[el.dataset.k] = +el.value; }),
  localSet: el => { L[el.dataset.k] = el.value.trim(); if (el.dataset.k !== 'claudeKey') { L.lastSync = 0; delete L.remoteAt; } saveLocal(); queueRender(); if (el.dataset.k === 'scriptUrl' && L.scriptUrl && L.appKey) pullRemote({ force: true, announce: true }); },
  tplField: el => commit(s => { const t = s.templates.find(x => x.id === el.dataset.id); t[el.dataset.k] = el.dataset.k === 'name' ? (el.value.trim() || 'Untitled day') : el.value; }),
  tplToggle: el => commit(s => { const t = s.templates.find(x => x.id === el.dataset.id); t[el.dataset.k] = el.checked; }),
  habitField: el => commit(s => { const h = s.habits.find(x => x.id === el.dataset.id); h[el.dataset.k] = el.dataset.k === 'name' ? (el.value.trim() || 'Untitled habit') : el.value; }),
  habitInTpl: el => commit(s => { const t = s.templates.find(x => x.id === el.dataset.t); const id = el.dataset.id; t.habitIds = el.checked ? [...new Set([...t.habitIds, id])] : t.habitIds.filter(x => x !== id); refreshSnapshots(s, t.id); }),
  listField: el => commit(s => { const x = s[el.dataset.coll].find(i => i.id === el.dataset.id); x[el.dataset.k] = el.dataset.k === 'name' ? (el.value.trim() || 'Untitled') : el.value; }),
  skinImage: async el => { const f = el.files[0]; if (!f) return; if (!/^image\/(png|jpeg)$/.test(f.type)) return toast('Please choose a JPG or PNG.'); if (f.size > 20e6) return toast('That image is over 20 MB.');
    skinDraft.refImage = await resizeImage(await readFile(f), 1400, .85); rerenderModal(); },
  skinUseBg: el => { skinDraft.useBg = el.checked; if (!el.checked) skinDraft.bgImage = ''; else if (skinDraft.vars && skinDraft.refImage) compressForBg(skinDraft.refImage).then(d => { skinDraft.bgImage = d; skinDraft.bgOpacity = skinDraft.bgOpacity || .12; rerenderModal(); }); rerenderModal(); },
  skinVar: el => { skinDraft.vars[el.dataset.k] = el.value; rerenderModal(); },
  skinSet: el => { const k = el.dataset.k; skinDraft[k] = ['radius', 'bgOpacity'].includes(k) ? +el.value : el.value; if (k === 'display' || k === 'body') loadFont(el.value); rerenderModal(); },
  expPick: el => { exportModal.st.ids[el.dataset.id] = el.checked; },
  expRange: el => { exportModal.st[el.dataset.k] = el.value; rerenderModal(); },
};
const IN = { skinField: el => { if (skinDraft) skinDraft[el.dataset.k] = el.value; } };
const SUB = {
  connectSheet: f => { const fromCode = parseSetupCode(f.code.value); const url = fromCode ? fromCode.u : f.url.value.trim(); const key = fromCode ? fromCode.k : f.key.value.trim();
    if (f.code.value.trim() && !fromCode) return toast("That setup code doesn't look right — try copying it again.");
    if (!/^https:\/\/script\.google(usercontent)?\.com\//.test(url) || !key) return toast('Paste a setup code, or both the web app URL and key.');
    L.scriptUrl = url; L.appKey = key; L.lastSync = 0; delete L.remoteAt; saveLocal(); toast('Connecting…'); pullRemote({ force: true, announce: true }); },
  quickTask: f => { const title = f.title.value.trim(); if (!title) return; const coins = +f.coins.value; const id = uid('tk');
    commit(s => { s.tasks.push({ id, title, groupId: '', due: '', priority: 2, notes: '', status: 'open', createdAt: Date.now(), closedAt: '' }); ensureDay(s, V.date).tasks.push({ taskId: id, coins, done: false }); }); },
  newTask: f => { const fd = new FormData(f); const title = String(fd.get('title')).trim(); if (!title) return;
    commit(s => s.tasks.push({ id: uid('tk'), title, groupId: String(fd.get('groupId') || ''), due: String(fd.get('due') || ''), priority: +fd.get('priority') || 2, notes: String(fd.get('notes') || '').trim(), status: 'open', createdAt: Date.now(), closedAt: '' }));
    toast('Task added.'); },
  editTask: f => { const fd = new FormData(f); const id = f.dataset.id;
    commit(s => { const t = s.tasks.find(x => x.id === id); Object.assign(t, { title: String(fd.get('title')).trim() || t.title, groupId: String(fd.get('groupId') || ''), due: String(fd.get('due') || ''), priority: +fd.get('priority') || 2, notes: String(fd.get('notes') || '').trim() }); }); closeModal(); },
  addGroup: f => { const name = f.name.value.trim(); if (!name) return; commit(s => s.taskGroups.push({ id: uid('gr'), name, color: f.color.value })); },
  addCategory: f => { const name = f.name.value.trim(); if (!name) return; commit(s => s.categories.push({ id: uid('ct'), name, color: f.color.value })); },
  tplAddHabit: f => { const name = f.name.value.trim(); if (!name) return; const id = uid('hb'); const tid = f.dataset.id;
    commit(s => { s.habits.push({ id, name, coins: +f.coins.value, priority: +f.priority.value, categoryId: f.categoryId.value, archived: false }); const t = s.templates.find(x => x.id === tid); t.habitIds.push(id); refreshSnapshots(s, tid); });
    setTimeout(() => $('form[data-submit="tplAddHabit"] input[name=name]')?.focus(), 30); },
  newHabitGlobal: f => { const name = f.name.value.trim(); if (!name) return; const id = uid('hb'); commit(s => s.habits.push({ id, name, coins: 2, priority: 2, categoryId: '', archived: false })); openModal(() => habitEditor(id)); },
};

function dayTaskPicker() {
  const d = S.days[V.date]; const on = new Set((d?.tasks || []).map(x => x.taskId)); const list = sortTasks(S.tasks.filter(t => t.status === 'open' && !on.has(t.id)), 'priority');
  const def = { 3: 5, 2: 3, 1: 1 };
  return `${modalHead(`Add tasks to ${V.date === todayStr() ? 'today' : shortD(V.date)}`)}<p class="small muted">Set how many coins each one is worth, then add it.</p>
    <div class="list">${list.map(t => { const g = groupById(t.groupId); return `<div class="trow"><div class="t-main"><span class="t-title">${esc(t.title)}</span><span class="t-meta"><span class="tag ${t.priority === 3 ? 'must' : t.priority === 2 ? 'should' : ''}">${TPRI[t.priority]}</span>${g ? esc(g.name) : ''}${t.due ? ` · due ${shortD(t.due)}` : ''}</span></div>
      <select data-coinsel="${t.id}" style="width:auto;padding:6px 8px" aria-label="Coins">${[1, 2, 3, 4, 5].map(n => opt(n, `${n}`, def[t.priority])).join('')}</select><button class="btn sm" data-act="pickToDay" data-id="${t.id}">Add</button></div>`; }).join('') || '<p class="muted">Every open task is already on this day. Add new ones on the Tasks page.</p>'}</div>`;
}
function todayCoinPicker(id) {
  const t = taskById(id); const def = { 3: 5, 2: 3, 1: 1 }[t.priority];
  return `${modalHead("Add to Today's List")}<p><b>${esc(t.title)}</b></p><p class="small muted">How many coins is finishing this worth today?</p>
    <div class="row">${[1, 2, 3, 4, 5].map(n => `<button class="btn ${n === def ? '' : 'ghost'}" data-act="confirmToToday" data-id="${id}" data-v="${n}" style="min-width:52px">${n}</button>`).join('')}</div>`;
}
function taskEditor(id) {
  const t = taskById(id); if (!t) { setTimeout(closeModal); return ''; }
  return `${modalHead('Task')}<form class="stack" data-submit="editTask" data-id="${t.id}">
    <input type="text" name="title" value="${esc(t.title)}" required aria-label="Title">
    <div class="grid3"><label class="f"><span>Group</span><select name="groupId">${groupOptions(t.groupId)}</select></label>
    <label class="f"><span>Due</span><input type="date" name="due" value="${esc(t.due)}"></label>
    <label class="f"><span>Priority</span><select name="priority">${[3, 2, 1].map(p => opt(p, TPRI[p], t.priority)).join('')}</select></label></div>
    <textarea name="notes" placeholder="Notes">${esc(t.notes)}</textarea>
    <button class="btn" type="submit">Save</button></form><div class="row" style="margin-top:12px"><span class="grow"></span><button class="btn sm danger" data-act="deleteTask" data-id="${t.id}">Delete task</button></div>`;
}
function applySkinModal(id) {
  const sk = getSkin(id); if (!sk) { setTimeout(closeModal); return ''; }
  return `${modalHead(`Apply “${sk.name}”`)}<div class="stack">
    <div class="set-row">${swatch(sk)}<div class="grow"><b>Default skin</b><div class="muted tiny">Used by any template without its own skin</div></div>${S.settings.defaultSkinId === id ? '<span class="badge-on">In use</span>' : `<button class="btn sm" data-act="setDefaultSkin" data-id="${id}">Use as default</button>`}</div>
    <div><div style="margin:4px 0;font-size:13px;font-weight:600;color:var(--muted)">Use for these kinds of day</div>${S.templates.map(t => `<label class="check-row"><input type="checkbox" data-act="skinTplToggle" data-id="${id}" data-t="${t.id}" ${t.skinId === id ? 'checked' : ''}> ${esc(t.name)}</label>`).join('') || '<p class="muted small">Create a day template first.</p>'}</div></div>`;
}
function connectModal() {
  return `${modalHead('Connect to my Google Sheet')}<form class="stack" data-submit="connectSheet">
    <p class="small">Already using Habitude on another device or in your browser? In that copy go to <b>Settings → Google Sheets sync → Copy setup code</b>, then paste it here.</p>
    <label class="f"><span>Setup code</span><textarea name="code" placeholder="HBT1:…" style="min-height:64px" autocomplete="off"></textarea></label>
    <details><summary class="small muted" style="cursor:pointer">Or enter the web app URL and key yourself</summary><div class="stack" style="margin-top:10px">
      <label class="f"><span>Apps Script web app URL</span><input type="url" name="url" placeholder="https://script.google.com/macros/s/…/exec"></label>
      <label class="f"><span>App key</span><input type="password" name="key" autocomplete="off"></label></div></details>
    <button class="btn block" type="submit">Connect & load my data</button>
    <button class="btn ghost block" type="button" data-act="backToWelcome">Back</button></form>`;
}
function parseSetupCode(code) {
  const m = String(code || '').trim().match(/HBT1:([A-Za-z0-9+/=_-]+)/); if (!m) return null;
  try { const o = JSON.parse(decodeURIComponent(escape(atob(m[1].replace(/-/g, '+').replace(/_/g, '/'))))); return o && o.u && o.k ? o : null; } catch { return null; }
}
function welcomeModal() {
  return `<div class="stack" style="padding:6px 2px"><div class="brand" style="font-size:28px"><span class="brand-mark"></span>Habitude</div>
    <p>Track your habits and tasks, earn coins, and build streaks that multiply your rewards.</p>
    <p class="muted small">Everything you enter stays in this browser. When you're ready, connect your own Google Sheet in Settings to keep a permanent copy and sync across devices.</p>
    <button class="btn block" data-act="seedExample">Start with an example setup</button><button class="btn ghost block" data-act="startBlank">Start with a blank slate</button><button class="btn soft block" data-act="openConnect">Connect to my Google Sheet</button>
    <p class="help center" style="margin:0">The example includes three kinds of day, a few habits and tasks — edit or delete any of it.</p></div>`;
}
function exampleState() {
  const s = defaultState(); const c = (name, color) => { const id = uid('ct'); s.categories.push({ id, name, color }); return id; };
  const health = c('Health', '#5f9e6e'), hyg = c('Hygiene', '#4f8fbf'), dog = c('Dog care', '#c98a3b'), home = c('Home care', '#8a6bbf'), social = c('Social', '#d0607a'), mind = c('Mind', '#3f9c9c');
  const hb = (name, coins, priority, categoryId) => { const id = uid('hb'); s.habits.push({ id, name, coins, priority, categoryId, archived: false }); return id; };
  const water = hb('Drink 8 glasses of water', 2, 2, health), vit = hb('Take vitamins', 1, 3, health), walk = hb('Walk the dogs', 3, 3, dog), feed = hb('Feed the dogs', 1, 3, dog),
    brush = hb('Brush & floss', 1, 3, hyg), skin = hb('Skincare routine', 1, 1, hyg), dishes = hb('Clear the kitchen', 2, 2, home), stretch = hb('Stretch 10 minutes', 2, 2, health),
    pages = hb('Morning pages', 2, 2, mind), med = hb('Meditate', 3, 1, mind), friend = hb('Reach out to a friend', 1, 1, social), clean = hb('Deep-clean one area', 3, 1, home), bed = hb('In bed by 10:30', 2, 2, health);
  s.templates.push({ id: uid('tp'), name: 'Work day', skinId: '', showTasks: true, groupByCategory: false, habitIds: [vit, feed, walk, brush, water, pages, dishes, bed] });
  s.templates.push({ id: uid('tp'), name: 'Day off', skinId: 'sk_nightgarden', showTasks: true, groupByCategory: true, habitIds: [vit, feed, walk, brush, skin, water, pages, med, friend, clean, stretch] });
  s.templates.push({ id: uid('tp'), name: 'Luteal phase', skinId: 'sk_moontide', showTasks: false, groupByCategory: false, habitIds: [vit, feed, walk, brush, skin, water, stretch, med, bed] });
  const g = (name, color) => { const id = uid('gr'); s.taskGroups.push({ id, name, color }); return id; };
  const hm = g('Home', '#8a6bbf'), wk = g('Work', '#4f8fbf'), er = g('Errands', '#c98a3b'); const t = todayStr();
  [['Schedule vet check-up', er, addDays(t, 5), 3], ['Order more dog food', er, addDays(t, 2), 2], ['Draft next newsletter', wk, addDays(t, 3), 2], ['Clean out the fridge', hm, '', 1], ['Fix the porch light', hm, addDays(t, 10), 1]]
    .forEach(([title, groupId, due, priority]) => s.tasks.push({ id: uid('tk'), title, groupId, due, priority, notes: '', status: 'open', createdAt: Date.now(), closedAt: '' }));
  return s;
}

/* ---------------- global listeners ---------------- */
document.addEventListener('click', e => {
  const el = e.target.closest('[data-act]');
  if (!el) { if (e.target.id === 'modal' && !syncHold) closeModal(); return; }
  if (el.disabled) return; const fn = ACT[el.dataset.act]; if (!fn) return;
  if (el.type !== 'checkbox') e.preventDefault(); fn(el, e);
});
document.addEventListener('change', e => { const el = e.target.closest('[data-ch]'); if (el && CH[el.dataset.ch]) CH[el.dataset.ch](el, e); });
document.addEventListener('input', e => { const el = e.target.closest('[data-in]'); if (el && IN[el.dataset.in]) IN[el.dataset.in](el, e); });
document.addEventListener('submit', e => { const f = e.target.closest('form[data-submit]'); if (f && SUB[f.dataset.submit]) { e.preventDefault(); SUB[f.dataset.submit](f); } });
document.addEventListener('keydown', e => { if (e.key === 'Escape' && modalFn && !syncHold) closeModal(); });
$('#fileSkin').addEventListener('change', async e => {
  const f = e.target.files[0]; e.target.value = ''; if (!f) return;
  try { const o = JSON.parse(await f.text()); const sk = sanitizeSkin({ ...(o.skin || o), id: uid('sk') }); if (!sk) throw 0;
    commit(s => s.skins.push(sk)); toast(`Added “${sk.name}” to your skins.`); openModal(() => applySkinModal(sk.id)); }
  catch { toast("That file isn't a Habitude skin."); }
});
$('#fileBackup').addEventListener('change', async e => {
  const f = e.target.files[0]; e.target.value = ''; if (!f) return;
  try { const o = JSON.parse(await f.text()); if (o.format !== 'habitude-backup' || !Array.isArray(o.habits)) throw 0; if (!confirm('Replace everything here with this backup?')) return;
    S = normalize(o); delete S.format; S.meta.updatedAt = Date.now(); saveState(); _scKey = ''; scheduleSync(); render(); toast('Backup restored.'); }
  catch { toast("That file isn't a Habitude backup."); }
});
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') { if (V.date < todayStr() && V.page === 'day' && !modalFn) { /* keep the date they chose */ } pullRemote(); }
  else if (syncState === 'pending') pushRemote(); });
window.addEventListener('pagehide', () => { if (syncState === 'pending') pushRemote(); });

window.addEventListener('storage', e => {
  try {
    if (e.key === STORE && e.newValue) { S = normalize(JSON.parse(e.newValue)); _scKey = ''; queueRender(); }
    if (e.key === LOCAL && e.newValue) { const o = JSON.parse(e.newValue); L = { ...L, ...o }; updateSyncUI(); }
  } catch {}
});
const PAGES = { day: pageDay, tasks: pageTasks, progress: pageProgress, settings: pageSettings };
render();
if (FIRST_RUN) openModal(welcomeModal);
pullRemote({ force: true });
try { navigator.storage?.persist?.(); } catch {}
if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js').catch(() => {});
let _rw = window.innerWidth; window.addEventListener('resize', () => { if (Math.abs(window.innerWidth - _rw) > 60 && V.page === 'progress') { _rw = window.innerWidth; render(); } });
