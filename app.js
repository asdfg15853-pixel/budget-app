/* 나의 자산 — 가계부·자산 웹앱
 * 데이터와 계산은 구글 시트(Apps Script 웹 앱)에 있고, 이 앱은 화면과 입력만 맡아요.
 * 비밀키는 처음 한 번 입력하면 이 기기에만 저장돼요 (코드·저장소에는 없음).
 */
'use strict';

const API = 'https://script.google.com/macros/s/AKfycbwK4qetpriYFGnhtJaVvmnH-7PoPrEM-XS-oGAq5Ri3r93KnwrSnognlYIMh7LAwhD8/exec';
const SHEET_URL = 'https://docs.google.com/spreadsheets/d/1opK41aXB8vl_7LGYoYTdYdV5QQHiGS3nxYsvmYzDAcQ/edit';
const VERSION = '1.1.0';

const CAT_ICON = {
  '주거': '🏠', '정기결제': '🔁', '식비': '🍚', '생활·쇼핑': '🛍️', '교통': '🚇', '건강·운동': '💪', '취미·배움': '🎨',
  '여행': '✈️', '경조사·선물': '🎁', '급여': '💼', '상여': '🎉', '금융수입': '🏦', '기타수입': '➕',
  '투자·저축': '📈', '카드대금': '💳', '계좌이동': '🔄', '대출상환': '📉'
};
const PAY_SHORT = { '네이버 현대카드': '현대카드', '신한 체크': '신한체크', '현금': '현금·계좌' };
// 대시보드 표 중 두 개가 옆에 붙어 있는 섹션: 이 열 앞에서 나눠서 위아래로 보여 줘요
const DASH_SPLIT = { '⑤': [3], '⑨': [3] };

/* ───────────── 작은 도구 ───────────── */

const $ = (s, el = document) => el.querySelector(s);
function h(tag, attrs, ...kids) {
  const el = document.createElement(tag);
  if (attrs) for (const k in attrs) {
    const v = attrs[k];
    if (v === null || v === undefined || v === false) continue;
    if (k === 'class') el.className = v;
    else if (k === 'html') el.innerHTML = v;
    else if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
    else if (k === 'style' && typeof v === 'object') Object.assign(el.style, v);
    else el.setAttribute(k, v === true ? '' : v);
  }
  for (const c of kids.flat(Infinity)) if (c !== null && c !== undefined && c !== false) el.append(c.nodeType ? c : String(c));
  return el;
}
const won = n => (n < 0 ? '-' : '') + Math.abs(Math.round(Number(n) || 0)).toLocaleString('ko-KR') + '원';
const signed = n => (n > 0 ? '+' : '') + won(n);
const pct = (x, d = 1) => (x === null || !isFinite(x)) ? '–' : (x * 100).toFixed(d) + '%';
function man(n) {           // 1,234,000 → 123.4만
  const a = Math.abs(n);
  if (a >= 1e8) return (n / 1e8).toFixed(a >= 1e9 ? 1 : 2).replace(/\.?0+$/, '') + '억';
  if (a >= 1e4) return (n / 1e4).toFixed(a >= 1e6 ? 0 : 1).replace(/\.0$/, '') + '만';
  return Math.round(n).toLocaleString('ko-KR');
}
const pad2 = n => String(n).padStart(2, '0');
const ymdOf = d => d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate());
function parseYmd(s) { const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d); }
const DOW = ['일', '월', '화', '수', '목', '금', '토'];
function dayLabel(s, withYear) {
  const d = parseYmd(s);
  return (withYear ? d.getFullYear() + '년 ' : '') + (d.getMonth() + 1) + '월 ' + d.getDate() + '일 (' + DOW[d.getDay()] + ')';
}
function monthAdd(ym, k) { const [y, m] = ym.split('-').map(Number); const d = new Date(y, m - 1 + k, 1); return d.getFullYear() + '-' + pad2(d.getMonth() + 1); }
const store = {
  get(k, def) { try { const v = localStorage.getItem('ba_' + k); return v === null ? def : JSON.parse(v); } catch (e) { return def; } },
  set(k, v) { try { localStorage.setItem('ba_' + k, JSON.stringify(v)); } catch (e) { /* 저장 공간 없음 */ } },
  del(k) { try { localStorage.removeItem('ba_' + k); } catch (e) { /* 무시 */ } }
};
const today = () => ymdOf(new Date());

/* ───────────── 상태 ───────────── */

const S = {
  key: store.get('key', ''),
  tab: 'add',
  boot: store.get('c_boot', null),
  month: null,
  ledger: null,
  dash: store.get('c_dash', null),
  sheets: { hold: store.get('c_hold', null), cash: store.get('c_cash', null), set: store.get('c_set', null) },
  assetSeg: store.get('assetSeg', 'hold'),
  selDay: null,
  add: null,
  busy: {}
};
function resetAdd(keepDate) {
  const keep = keepDate && S.add;
  S.add = { kind: keep && S.add.kind === '고정' ? '고정' : '지출', cat: '', amt: '', name: '', pay: store.get('lastPay', '네이버 현대카드'), rc: false,
            date: keep ? S.add.date : today(), day: keep ? S.add.day : parseYmd(today()).getDate() };
}

/* ───────────── API ───────────── */

async function api(action, body = {}, retry = 1) {
  const res = await fetch(API, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(Object.assign({ key: S.key, action }, body)),
    redirect: 'follow'
  });
  const txt = await res.text();
  let j;
  try { j = JSON.parse(txt); } catch (e) { throw new Error('시트 응답을 읽지 못했어요. 잠시 뒤 다시 해 주세요'); }
  // 스크립트 새 버전이 퍼지는 몇 초 동안 옛 버전이 답할 수 있어요 (문자 처리 응답이 옴) → 한 번 더 시도
  if (j.ok && j.result !== undefined && j.parsed !== undefined && retry > 0) {
    await new Promise(r => setTimeout(r, 1500));
    return api(action, body, retry - 1);
  }
  if (!j.ok) throw new Error(j.error || '처리하지 못했어요');
  return j;
}

async function loadBoot(month) {
  const j = await api('app_boot', { month: month || undefined });
  S.boot = j;
  S.ledger = j.ledger;
  S.month = j.ledger.month;
  if (!month || month === ymdOf(new Date()).slice(0, 7)) store.set('c_boot', j);
  return j;
}
async function loadLedger(month) {
  const j = await api('app_ledger', { month });
  S.ledger = j.ledger; S.month = j.ledger.month;
}
async function loadDash() { const j = await api('app_dash'); S.dash = j; store.set('c_dash', j); }
async function loadSheet(k) { const j = await api('app_sheet', { sheet: k }); S.sheets[k] = j; store.set('c_' + k, j); }

/* 연결이 안 될 때 입력을 모아 뒀다가 다시 보내요 */
function outbox() { return store.get('outbox', []); }
async function flushOutbox() {
  const box = outbox();
  if (!box.length || !navigator.onLine) return;
  const left = [];
  for (const it of box) {
    try { await api('app_add', it); } catch (e) { left.push(it); }
  }
  store.set('outbox', left);
  if (box.length !== left.length) {
    toast((box.length - left.length) + '건을 시트에 보냈어요');
    refreshQuiet();
  }
  if (S.tab === 'add') render();
}

/* ───────────── 토스트·시트(모달) ───────────── */

let toastTimer = null;
function toast(msg, opt = {}) {
  const t = $('#toast');
  t.replaceChildren(h('span', null, msg));
  if (opt.undo) t.append(h('button', { onclick: () => { hideToast(); opt.undo(); } }, '실행 취소'));
  t.classList.add('in');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(hideToast, opt.undo ? 5000 : 2600);
}
function hideToast() { $('#toast').classList.remove('in'); }

function openSheet({ title, left = '취소', right = null, onRight = null, body }) {
  const root = $('#sheet-root');
  const scrim = h('div', { class: 'scrim' });
  const rightBtn = right ? h('button', { class: 'r' }, right) : h('span', { style: { minWidth: '60px' } });
  const sheet = h('div', { class: 'sheet', role: 'dialog', 'aria-label': title },
    h('div', { class: 'grab' }),
    h('div', { class: 'bar' }, h('button', { onclick: () => close() }, left), h('h3', null, title), rightBtn),
    body);
  root.append(scrim, sheet);
  requestAnimationFrame(() => { scrim.classList.add('in'); sheet.classList.add('in'); });
  scrim.addEventListener('click', () => close());
  function close() {
    scrim.classList.remove('in'); sheet.classList.remove('in');
    setTimeout(() => { scrim.remove(); sheet.remove(); }, 280);
  }
  if (right && onRight) rightBtn.addEventListener('click', async () => {
    rightBtn.disabled = true;
    try { const keep = await onRight(); if (keep !== true) close(); }
    catch (e) { toast(e.message); }
    finally { rightBtn.disabled = false; }
  });
  return { close, sheet, rightBtn };
}

function sw(checked, onchange) {
  const inp = h('input', { type: 'checkbox' });
  inp.checked = !!checked;
  if (onchange) inp.addEventListener('change', () => onchange(inp.checked, inp));
  return { el: h('label', { class: 'switch' }, inp, h('span')), inp };
}

/* ───────────── 화면 전환 ───────────── */

function setTab(t) {
  S.tab = t;
  document.querySelectorAll('#tabs button').forEach(b => b.classList.toggle('on', b.dataset.tab === t));
  try { history.replaceState(null, '', '?tab=' + t); } catch (e) { /* 무시 */ }
  window.scrollTo(0, 0);
  render();
  autoLoad(t);
}

function render() {
  const v = $('#view');
  $('#tabs').classList.toggle('hidden', !S.key);
  if (!S.key) { v.replaceChildren(viewWelcome()); return; }
  const fn = { add: viewAdd, ledger: viewLedger, dash: viewDash, assets: viewAssets, settings: viewSettings }[S.tab] || viewAdd;
  v.replaceChildren(fn());
}

const lastLoad = {};
async function autoLoad(t, force) {
  const fresh = k => !force && lastLoad[k] && Date.now() - lastLoad[k] < 60000;
  try {
    if ((t === 'add' || t === 'ledger') && !fresh('boot')) {
      S.busy.boot = true; if (t === 'ledger') render();
      await loadBoot(S.tab === 'ledger' && S.month ? S.month : undefined); lastLoad.boot = Date.now();
    }
    if (t === 'dash' && !fresh('dash')) { S.busy.dash = true; render(); await loadDash(); lastLoad.dash = Date.now(); }
    if (t === 'assets') {
      const k = S.assetSeg;
      if (!fresh(k)) { S.busy[k] = true; render(); await loadSheet(k); lastLoad[k] = Date.now(); }
    }
    if (t === 'settings' && !fresh('set')) { S.busy.set = true; render(); await loadSheet('set'); lastLoad.set = Date.now(); }
    S.err = null;
  } catch (e) {
    S.err = e.message;
  } finally {
    S.busy = {};
    if (S.tab === t) render();
  }
}
function refreshQuiet() { lastLoad.boot = 0; lastLoad.dash = 0; lastLoad.hold = 0; lastLoad.cash = 0; lastLoad.set = 0; autoLoad(S.tab, true); }
function errBox() { return S.err ? h('div', { class: 'err' }, '불러오지 못했어요: ', S.err) : null; }

/* ───────────── 처음 설정 ───────────── */

function viewWelcome() {
  const inp = h('input', { type: 'text', inputmode: 'text', autocomplete: 'off', autocapitalize: 'off', spellcheck: 'false', placeholder: '비밀키' });
  const msg = h('div');
  const btn = h('button', { class: 'save', onclick: go }, '시작하기');
  async function go() {
    const k = inp.value.trim();
    if (!k) return;
    btn.disabled = true; msg.replaceChildren(h('div', { class: 'loading' }, h('span', { class: 'spin' }), '확인하는 중'));
    S.key = k;
    try {
      await loadBoot();
      store.set('key', k);
      resetAdd();
      setTab('add');
    } catch (e) {
      S.key = '';
      msg.replaceChildren(h('div', { class: 'err' }, e.message));
      btn.disabled = false;
    }
  }
  inp.addEventListener('keydown', e => { if (e.key === 'Enter') go(); });
  return h('div', { class: 'welcome' },
    h('img', { class: 'logo', src: 'icons/icon-192.png', alt: '' }),
    h('h1', null, '나의 자산'),
    h('p', null, '구글 시트 설정 탭 67행의 비밀키를 넣어 주세요.', h('br'), '이 기기에만 저장돼요.'),
    inp, btn, msg);
}

/* ───────────── ① 빠른 입력 ───────────── */

function catsOf(kind) { return ((S.boot && S.boot.cats) || []).filter(c => c.kind === kind); }
function fixedCats() { return catsOf('지출').filter(c => c.fixed === '고정'); }

/* 한 달 예산: 고정비까지 포함한 전체 지출 기준. 남은 돈 = 예산 − 쓴 돈 − 이번 달 아직 안 나간 고정비 */
function budgetInfo() {
  const b = S.boot && S.boot.budget;
  if (!b || b.month !== today().slice(0, 7)) return null;
  const t = parseYmd(today());
  const daysLeft = new Date(t.getFullYear(), t.getMonth() + 1, 0).getDate() - t.getDate() + 1;
  const remain = b.amount - b.spent - b.upcoming;
  return { amount: b.amount, spent: b.spent, upcoming: b.upcoming, remain, daysLeft,
           perDay: remain > 0 ? Math.floor(remain / daysLeft / 100) * 100 : 0 };
}
function saveBootCache() { if (S.boot && S.boot.ledger && S.boot.ledger.month === today().slice(0, 7)) store.set('c_boot', S.boot); }
function activeFixed(kind) {
  const cur = today().slice(0, 7);
  return ((S.boot && S.boot.fixed) || []).filter(f => f.k === kind && !f.stop && !(f.e && f.e <= cur));
}

function viewAdd() {
  if (!S.add) resetAdd();
  const A = S.add;
  if (!S.boot) return h('div', { class: 'loading' }, h('span', { class: 'spin' }), '불러오는 중');
  const n = Number(A.amt || 0);
  if (A.kind === '고정' && !S.boot.fixed) A.kind = '지출';
  const isFx = A.kind === '고정';

  // 왼쪽 위: 보통은 날짜, 고정비는 '매월 며칠'
  let leftChip;
  if (isFx) {
    const sel = h('select', { 'aria-label': '매월 결제일' }, Array.from({ length: 31 }, (_, i) =>
      h('option', { value: i + 1, selected: A.day === i + 1 ? true : null }, '매월 ' + (i + 1) + '일' + (i === 30 ? ' (말일)' : ''))));
    sel.addEventListener('change', () => { A.day = Number(sel.value); render(); });
    leftChip = h('label', { class: 'date-chip' }, '🔁 매월 ', A.day + '일', sel);
  } else {
    const dateInp = h('input', { type: 'date', value: A.date, max: '2100-12-31', 'aria-label': '날짜' });
    dateInp.addEventListener('change', () => { A.date = dateInp.value || today(); render(); });
    leftChip = h('label', { class: 'date-chip' }, '📅 ', A.date === today() ? '오늘' : dayLabel(A.date), dateInp);
  }

  // 고정 탭은 시트 스크립트가 고정비 기능을 지원할 때만 (boot에 fixed가 있을 때)
  const kinds = S.boot.fixed ? ['지출', '수입', '이체', '고정'] : ['지출', '수입', '이체'];
  const kindSeg = h('div', { class: 'seg' }, kinds.map(k =>
    h('button', { class: A.kind === k ? 'on' : '', onclick: () => { A.kind = k; A.cat = ''; render(); } }, k)));

  const catList = isFx ? fixedCats() : catsOf(A.kind);
  const cats = h('div', { class: 'chips cats' }, catList.map(c =>
    h('button', { class: 'chip' + (A.cat === c.name ? ' on' : ''), onclick: () => { A.cat = A.cat === c.name ? '' : c.name; render(); } },
      (CAT_ICON[c.name] || '') + ' ' + c.name)));

  const pays = (S.boot.pays || []);
  const paySeg = h('div', { class: 'seg' }, pays.map(p =>
    h('button', { class: A.pay === p.name ? 'on' : '', onclick: () => { A.pay = p.name; store.set('lastPay', p.name); render(); } },
      PAY_SHORT[p.name] || p.name)));
  const rcBtn = (A.pay === '현금' && A.kind === '지출')
    ? h('button', { class: 'toggle-pill' + (A.rc ? ' on' : ''), onclick: () => { A.rc = !A.rc; render(); } }, A.rc ? '✓ 현금영수증' : '현금영수증')
    : null;

  const name = h('input', { class: 'name', type: 'text', value: A.name, enterkeyhint: 'done',
    placeholder: isFx ? '이름 (필수) 예: 넷플릭스, 관리비' : '내용 (선택) 예: 점심, 쿠팡' });
  name.addEventListener('input', () => { A.name = name.value; saveBtn.disabled = !canSave(); });
  name.addEventListener('keydown', e => { if (e.key === 'Enter') name.blur(); });

  const key = d => () => {
    if (d === 'del') A.amt = A.amt.slice(0, -1);
    else if ((A.amt + d).replace(/^0+/, '').length <= 10) A.amt = (A.amt + d).replace(/^0+/, '');
    updateAmount();
  };
  const amtV = h('span', { class: 'v' + (n ? '' : ' zero') }, n ? n.toLocaleString('ko-KR') : '0');
  const hint = h('div', { class: 'hint' }, n >= 10000 ? man(n) + '원' : '');
  const bline = h('button', { class: 'bline' });
  const saveBtn = h('button', { class: 'save', onclick: save }, isFx ? '고정비 등록' : '저장');
  const canSave = () => Number(A.amt || 0) > 0 && !!A.cat && (!isFx || !!A.name.trim());

  // 금액 아래 한 줄: 지출은 남은 예산, 고정비는 매달 자동 안내
  function drawLine() {
    const v = Number(A.amt || 0);
    if (isFx) {
      const cnt = activeFixed('지출').length;
      bline.className = 'bline';
      bline.onclick = () => openFixedList();
      bline.replaceChildren('매달 ', h('b', null, A.day + '일'), '에 자동으로 채워져요', cnt ? ' · 등록 ' + cnt + '개 ›' : '');
      return;
    }
    const info = A.kind === '지출' ? budgetInfo() : null;
    if (!info) { bline.className = 'bline hide'; bline.replaceChildren(); return; }
    bline.onclick = () => openBudget();
    if (!info.amount) { bline.className = 'bline set'; bline.replaceChildren('💰 한 달 예산 정하기 ›'); return; }
    const inMonth = A.date.slice(0, 7) === today().slice(0, 7);
    const after = info.remain - (inMonth ? v : 0);
    bline.className = 'bline' + (after < 0 ? ' neg' : '');
    if (v > 0 && inMonth) bline.replaceChildren(after < 0 ? '저장하면 예산 ' : '저장하면 남은 예산 ', h('b', null, won(Math.abs(after))), after < 0 ? ' 초과' : '');
    else if (info.remain < 0) bline.replaceChildren('이번 달 예산 ', h('b', null, won(-info.remain)), ' 초과');
    else bline.replaceChildren('남은 예산 ', h('b', null, won(info.remain)), ' · 하루 ', h('b', null, won(info.perDay)));
  }
  function updateAmount() {
    const v = Number(A.amt || 0);
    amtV.textContent = v ? v.toLocaleString('ko-KR') : '0';
    amtV.className = 'v' + (v ? '' : ' zero');
    hint.textContent = v >= 10000 ? man(v) + '원' : '';
    saveBtn.disabled = !canSave();
    drawLine();
  }
  saveBtn.disabled = !canSave();
  drawLine();

  async function save() {
    if (isFx) return saveFixed();
    const item = { d: A.date, c: A.cat, n: A.name.trim(), a: Number(A.amt), p: A.pay, rc: A.rc && A.pay === '현금' };
    saveBtn.disabled = true;
    try {
      const j = await api('app_add', item);
      const r = j.row;
      toast((A.cat) + ' ' + won(item.a) + ' 저장했어요', {
        undo: async () => {
          try { await api('app_delete', { r, check: { d: item.d, a: item.a } }); toast('지웠어요'); refreshQuiet(); }
          catch (e) { toast(e.message); }
        }
      });
      lastLoad.boot = 0;
    } catch (e) {
      if (!navigator.onLine || /fetch|network|Load failed/i.test(e.message)) {
        const box = outbox(); box.push(item); store.set('outbox', box);
        toast('연결이 안 돼서 모아 뒀어요. 연결되면 자동으로 보내요');
      } else { toast(e.message); saveBtn.disabled = false; return; }
    }
    // 예산 숫자는 다음 새로고침 전까지 화면에서 바로 반영
    const b = S.boot && S.boot.budget;
    if (b && A.kind === '지출' && item.d.slice(0, 7) === b.month) { b.spent += item.a; saveBootCache(); }
    resetAdd(true);
    render();
  }

  async function saveFixed() {
    const item = { n: A.name.trim(), c: A.cat, a: Number(A.amt), p: A.pay, d: A.day };
    saveBtn.disabled = true;
    try {
      const j = await api('app_fixed_add', item);
      toast(j.message, {
        undo: async () => {
          try { await api('app_fixed_delete', { i: j.i, check: { n: item.n } }); toast('등록을 취소했어요'); refreshQuiet(); }
          catch (e) { toast(e.message); }
        }
      });
      Object.assign(A, { amt: '', name: '', cat: '' });
      render();
      refreshQuiet();
    } catch (e) { toast(e.message); saveBtn.disabled = false; }
  }

  const keys = h('div', { class: 'pad-keys' },
    ['1', '2', '3', '4', '5', '6', '7', '8', '9', '00', '0', 'del'].map(d =>
      h('button', { onclick: key(d), 'aria-label': d === 'del' ? '지우기' : d }, d === 'del' ? '⌫' : d)));

  const box = outbox();
  return h('div', { class: 'add' + (isFx ? ' fx' : '') },
    h('div', { class: 'head' }, leftChip, kindSeg),
    h('div', { class: 'amount' }, amtV, h('span', { class: 'u' }, '원'), hint),
    bline,
    cats,
    h('div', { class: 'opts' }, paySeg, rcBtn),
    name,
    keys,
    saveBtn,
    box.length ? h('div', { class: 'outbox' }, '보내지 못한 입력 ' + box.length + '건 · 연결되면 자동으로 보내요') : null
  );
}

/* 한 달 예산 정하기 */
function openBudget() {
  const info = budgetInfo() || { amount: Number(S.boot && S.boot.budget && S.boot.budget.amount) || 0 };
  let v = info.amount ? String(info.amount) : '';
  const inp = h('input', { type: 'text', inputmode: 'numeric', placeholder: '예: 1,500,000', value: v ? Number(v).toLocaleString('ko-KR') : '' });
  const hint = h('div', { class: 'help' }, v ? man(Number(v)) + '원' : ' ');
  inp.addEventListener('input', () => {
    v = inp.value.replace(/[^\d]/g, '').replace(/^0+/, '').slice(0, 10);
    inp.value = v ? Number(v).toLocaleString('ko-KR') : '';
    hint.textContent = v ? man(Number(v)) + '원' : ' ';
  });
  const fx = activeFixed('지출').reduce((s, f) => s + f.a, 0);
  openSheet({
    title: '한 달 예산', right: '저장',
    onRight: async () => {
      const j = await api('app_budget', { amount: Number(v || 0) });
      if (S.boot && S.boot.budget) S.boot.budget.amount = j.amount;
      saveBootCache();
      toast(j.message);
      render();
    },
    body: h('div', null,
      h('div', { class: 'list' }, field('한 달 예산', inp)),
      hint,
      h('div', { class: 'help' }, '고정비까지 포함한 한 달 전체 지출 기준이에요. 가계부와 입력 화면에 남은 예산과 하루에 쓸 수 있는 돈이 나와요.'),
      fx ? h('div', { class: 'help' }, '참고로 매달 나가는 고정비는 ' + won(fx) + '이에요.') : null,
      h('div', { class: 'help' }, '설정 탭의 같은 칸이라 비상금 목표(생활비 × 개월 수)에도 함께 쓰여요.'))
  });
  setTimeout(() => inp.focus(), 320);
}

/* ───────────── ② 가계부 ───────────── */

function ledgerStats(rows) {
  const s = { exp: 0, inc: 0, inv: 0, byCat: {}, byDay: {}, card: 0 };
  rows.forEach(r => {
    if (r.g === '지출') {
      s.exp += r.a;
      s.byCat[r.c] = (s.byCat[r.c] || 0) + r.a;
      s.byDay[r.d] = (s.byDay[r.d] || 0) + r.a;
      if (r.p === '네이버 현대카드' && r.c !== '주거') s.card += r.a;
    } else if (r.g === '수입') s.inc += r.a;
    if (r.c === '투자·저축') s.inv += r.a;
  });
  s.left = s.inc - s.exp - s.inv;
  s.rate = s.inc > 0 ? (s.inc - s.exp) / s.inc : null;
  return s;
}

function nextPayday() {
  const pd = Number((S.boot && S.boot.card && S.boot.card.payday) || 1);
  const t = parseYmd(today());
  let d = new Date(t.getFullYear(), t.getMonth(), pd);
  if (d < t) d = new Date(t.getFullYear(), t.getMonth() + 1, pd);
  return d;
}

function viewLedger() {
  const L = S.ledger || (S.boot && S.boot.ledger);
  const ym = (L && L.month) || today().slice(0, 7);
  const [y, m] = ym.split('-').map(Number);
  const head = h('div', { class: 'month-nav' },
    h('h1', { class: 'large' }, m + '월', h('span', { class: 'caption', style: { fontSize: '17px', marginLeft: '6px', fontWeight: 500 } }, y + '년')),
    h('div', { class: 'arrows' },
      h('button', { 'aria-label': '이전 달', onclick: () => goMonth(-1) }, '‹'),
      h('button', { 'aria-label': '다음 달', onclick: () => goMonth(1) }, '›')));
  if (!L) return h('div', null, head, errBox(), h('div', { class: 'loading' }, h('span', { class: 'spin' }), '불러오는 중'));

  const rows = L.rows || [];
  const st = ledgerStats(rows);
  const out = [head, errBox()];

  // 요약 (이번 달이면 예산도 같이)
  const isCur = ym === today().slice(0, 7);
  const binfo = isCur ? budgetInfo() : null;
  out.push(h('div', { class: 'card pad kpi' },
    h('div', { class: 'lbl' }, '이번 달 지출'),
    h('div', { class: 'big' }, won(st.exp)),
    binfo ? budgetBlock(binfo) : null,
    h('div', { class: 'grid' },
      h('div', null, h('div', { class: 'lbl' }, '수입'), h('div', { class: 'v pos' }, won(st.inc))),
      h('div', null, h('div', { class: 'lbl' }, '투자·저축'), h('div', { class: 'v' }, won(st.inv))),
      h('div', null, h('div', { class: 'lbl' }, '저축률'), h('div', { class: 'v' }, pct(st.rate))))));
  if (S.busy.boot) out.push(h('div', { class: 'stale' }, h('span', { class: 'spin' }), '새로 불러오는 중'));

  // 확인할 결제
  const pend = (S.boot && S.boot.pending) || [];
  if (pend.length) {
    out.push(h('div', { class: 'section-title' }, '확인할 결제 ' + pend.length + '건'));
    out.push(h('div', { class: 'list pend' }, pend.map(p => h('div', { class: 'row' },
      h('div', { class: 'top-line' },
        h('div', { class: 'main' },
          h('div', { class: 't1' }, p.n || '(내용 없음)', p.pg ? h('span', { class: 'badge y' }, '결제대행') : null),
          h('div', { class: 't2' }, dayLabel(p.d), ' · ', PAY_SHORT[p.p] || p.p, p.guess ? ' · 추측: ' + p.guess : '')),
        h('div', { class: 'end amt' }, won(p.a))),
      h('div', { class: 'chips scroll' }, catsOf('지출').map(c =>
        h('button', { class: 'chip sm' + (p.guess === c.name ? ' on' : ''), onclick: () => classify(p, c.name) },
          (CAT_ICON[c.name] || '') + ' ' + c.name)))))));
  }

  // 현대카드
  const card = S.boot && S.boot.card;
  if (card && ym === today().slice(0, 7)) {
    const min = Number(card.min) || 300000;
    const np = nextPayday();
    out.push(h('div', { class: 'section-title' }, '네이버 현대카드'));
    out.push(h('div', { class: 'card pad' },
      h('div', { style: { display: 'flex', justifyContent: 'space-between' } },
        h('div', null, h('div', { class: 'caption' }, '이번 달 실적 (기준 ' + man(min) + ')'),
          h('div', { style: { fontSize: '20px', fontWeight: 700 }, class: 'num' }, won(st.card))),
        h('div', { style: { textAlign: 'right' } }, h('div', { class: 'caption' }, '다음 결제일'),
          h('div', { style: { fontSize: '17px', fontWeight: 600 } }, (np.getMonth() + 1) + '월 ' + np.getDate() + '일'))),
      h('div', { class: 'progress' }, h('div', { style: { width: Math.min(100, st.card / min * 100) + '%' } })),
      h('div', { class: 'caption', style: { marginTop: '6px' } },
        st.card >= min ? '✓ 실적 채웠어요' : won(min - st.card) + ' 더 쓰면 실적 달성')));
  }

  // 고정비
  if (isCur && S.boot && S.boot.fixed) out.push(fixedSection());

  // 카테고리
  const cats = Object.keys(st.byCat).sort((a, b) => st.byCat[b] - st.byCat[a]);
  if (cats.length) {
    const max = st.byCat[cats[0]];
    out.push(h('div', { class: 'section-title' }, '카테고리별 지출'));
    out.push(h('div', { class: 'list bars' }, cats.map(c => {
      const v = st.byCat[c], pv = (L.prevCat || {})[c] || 0, diff = v - pv;
      return h('div', { class: 'row' },
        h('div', { class: 'line' }, h('span', null, (CAT_ICON[c] || '') + ' ' + (c || '미분류')),
          h('span', { class: 'amt' }, won(v), h('span', { class: 'caption', style: { fontWeight: 400, marginLeft: '6px' } }, pct(v / st.exp, 0)))),
        h('div', { class: 'track' }, h('div', { class: 'fillbar', style: { width: (v / max * 100) + '%' } })),
        pv || diff ? h('div', { class: 'diff' }, '지난달 ' + won(pv) + ' · ' + (diff > 0 ? '▲ ' : diff < 0 ? '▼ ' : '') + won(Math.abs(diff))) : null);
    })));
  }

  // 달력
  out.push(h('div', { class: 'section-title' }, '달력 · 그날 쓴 돈'));
  out.push(h('div', { class: 'card cal' }, calendar(ym, st.byDay)));

  // 내역
  const list = S.selDay ? rows.filter(r => r.d === S.selDay) : rows;
  const days = [];
  list.forEach(r => { const last = days[days.length - 1]; if (!last || last.d !== r.d) days.push({ d: r.d, rows: [r] }); else last.rows.push(r); });
  if (S.selDay) out.push(h('div', { class: 'day-head' }, h('span', null, dayLabel(S.selDay) + '만 보는 중'),
    h('button', { style: { color: 'var(--blue)' }, onclick: () => { S.selDay = null; render(); } }, '전체 보기')));
  if (!days.length) out.push(h('div', { class: 'loading' }, '이 달에는 아직 내역이 없어요'));
  days.forEach(g => {
    const tot = g.rows.filter(r => r.g === '지출').reduce((s, r) => s + r.a, 0);
    out.push(h('div', { class: 'day-head' }, h('span', null, dayLabel(g.d)), h('span', { class: 'num' }, tot ? '지출 ' + won(tot) : '')));
    out.push(h('div', { class: 'list' }, g.rows.map(r => txRow(r))));
  });
  return h('div', null, out);
}

function budgetBlock(info) {
  if (!info.amount) {
    return h('button', { class: 'budget set', onclick: openBudget },
      h('span', { class: 'bt' }, '💰 한 달 예산 정하기'), h('span', { class: 'caption' }, '정하면 하루에 쓸 수 있는 돈을 알려 드려요 ›'));
  }
  const used = Math.min(1, info.spent / info.amount);
  const up = Math.min(1 - used, info.upcoming / info.amount);
  const over = info.remain < 0;
  const tone = over ? ' over' : (info.spent + info.upcoming) / info.amount >= 0.9 ? ' warn' : '';
  return h('button', { class: 'budget', onclick: openBudget },
    h('div', { class: 'bl' }, h('span', { class: 'lbl' }, over ? '예산 초과' : '남은 예산'), h('span', { class: 'lbl' }, '예산 ' + won(info.amount) + ' ›')),
    h('div', { class: 'bl' },
      h('span', { class: 'bv' + (over ? ' neg' : '') }, won(Math.abs(info.remain))),
      over ? null : h('span', { class: 'pd' }, '하루 ' + won(info.perDay))),
    h('div', { class: 'progress b' }, h('div', { class: 'u' + tone, style: { width: used * 100 + '%' } }), h('div', { class: 'f', style: { width: up * 100 + '%' } })),
    h('div', { class: 'caption' }, '쓴 돈 ' + won(info.spent) + (info.upcoming ? ' · 나갈 고정비 ' + won(info.upcoming) : '') + ' · ' + info.daysLeft + '일 남음'));
}

/* 고정비: 한 번 넣으면 매달 결제일에 가계부에 자동으로 채워져요 */
function fixedState(f) {
  const s = f.st || '';
  let m;
  if (f.stop || /중지/.test(s)) return '매월 ' + f.d + '일 · 잠시 중지';
  if ((m = s.match(/예정\s*(\d+\/\d+)/))) return m[1] + ' 예정';
  if ((m = s.match(/(\d+\/\d+)\s*반영/))) return m[1] + (f.k === '수입' ? ' 들어옴' : ' 나감');
  if (/가계부 값/.test(s)) return '이번 달은 직접 적음';
  return s;
}
function fixedRow(f) {
  const inc = f.k === '수입';
  const ended = !!f.e;
  return h('button', { class: 'row ind' + (ended || f.stop ? ' ro' : ''), onclick: () => editFixed(f) },
    h('div', { class: 'ico' }, CAT_ICON[f.c] || '🔁'),
    h('div', { class: 'main' },
      h('div', { class: 't1' }, f.n, ended ? h('span', { class: 'badge' }, '해지 · 이번 달까지') : null),
      h('div', { class: 't2' }, [PAY_SHORT[f.p] || f.p, fixedState(f) || '매월 ' + f.d + '일'].filter(Boolean).join(' · '))),
    h('div', { class: 'end amt' + (inc ? ' inc' : '') }, (inc ? '+' : '') + won(f.a)));
}
function fixedSection() {
  const all = S.boot.fixed || [];
  const exp = all.filter(f => f.k === '지출'), inc = all.filter(f => f.k === '수입');
  const tot = activeFixed('지출').reduce((s, f) => s + f.a, 0);
  return h('div', { id: 'fixed-sec' },
    h('div', { class: 'section-title row-title' }, h('span', null, '고정비 · 매달 자동'), tot ? h('span', { class: 'num' }, '월 ' + won(tot)) : null),
    exp.length ? h('div', { class: 'list' }, exp.map(fixedRow))
      : h('div', { class: 'card pad empty' }, '넷플릭스·관리비처럼 매달 나가는 돈을 한 번만 넣어 두면, 매달 결제일에 가계부에 자동으로 채워져요.'),
    inc.length ? h('div', { class: 'list', style: { marginTop: '8px' } }, inc.map(fixedRow)) : null,
    h('button', { class: 'add-btn', onclick: () => goAddFixed() }, '＋ 고정비 추가'));
}
function goAddFixed() {
  resetAdd();
  S.add.kind = '고정';
  setTab('add');
}
function openFixedList() {
  setTab('ledger');
  setTimeout(() => { const el = document.getElementById('fixed-sec'); if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' }); }, 60);
}

function editFixed(f) {
  const inc = f.k === '수입';
  const E = { n: f.n, c: f.c, a: String(f.a), p: f.p, d: f.d };
  const list = inc ? catsOf('수입') : fixedCats();
  if (!list.some(c => c.name === f.c)) list.push({ name: f.c });
  const catBox = h('div', { class: 'chips' });
  const drawCats = () => catBox.replaceChildren(...list.map(c =>
    h('button', { class: 'chip sm' + (E.c === c.name ? ' on' : ''), onclick: () => { E.c = c.name; drawCats(); } }, (CAT_ICON[c.name] || '') + ' ' + c.name)));
  drawCats();
  const name = h('input', { type: 'text', value: E.n, placeholder: '이름' }); name.addEventListener('input', () => E.n = name.value);
  const amt = h('input', { type: 'text', inputmode: 'numeric', value: Number(E.a).toLocaleString('ko-KR') });
  amt.addEventListener('input', () => { const v = amt.value.replace(/[^\d]/g, ''); E.a = v; amt.value = v ? Number(v).toLocaleString('ko-KR') : ''; });
  const day = h('select', null, Array.from({ length: 31 }, (_, i) => h('option', { value: i + 1, selected: E.d === i + 1 ? true : null }, (i + 1) + '일')));
  day.addEventListener('change', () => E.d = Number(day.value));
  const pay = h('select', null, (S.boot.pays || []).map(p => h('option', { value: p.name, selected: p.name === E.p ? true : null }, PAY_SHORT[p.name] || p.name)));
  pay.addEventListener('change', () => E.p = pay.value);
  const check = { n: f.n };

  const { close } = openSheet({
    title: inc ? '매달 들어오는 돈' : '고정비', right: '저장',
    onRight: async () => {
      const ch = {};
      if (E.n.trim() !== f.n) ch.n = E.n.trim();
      if (E.c !== f.c) ch.c = E.c;
      if (Number(E.a) !== f.a) ch.a = Number(E.a);
      if (E.d !== f.d) ch.d = E.d;
      if (E.p !== f.p) ch.p = E.p;
      if (!Object.keys(ch).length) return;
      const j = await api('app_fixed_save', { i: f.i, check, f: ch });
      toast(j.message);
      refreshQuiet();
    },
    body: h('div', null,
      h('div', { class: 'card pad' }, catBox),
      h('div', { class: 'list', style: { marginTop: '12px' } },
        field('이름', name), field('금액', amt), field('매월 결제일', day), field('결제수단', pay)),
      h('div', { class: 'help' }, f.s ? f.s.slice(0, 4) + '년 ' + Number(f.s.slice(5, 7)) + '월부터 매달 자동으로 채우는 중이에요. ' : '',
        '금액이나 결제일을 바꾸면 이번 달부터 새 값이 들어가고, 지난달까지 기록은 그대로 남아요.'),
      f.e
        ? h('button', { class: 'add-btn', onclick: async () => {
            try { const j = await api('app_fixed_end', { i: f.i, check, undo: true }); close(); toast(j.message); refreshQuiet(); }
            catch (e) { toast(e.message); }
          } }, '해지 취소하기')
        : h('button', { class: 'danger', onclick: async () => {
            if (!confirm(f.n + ' 해지할까요?\n지난 기록은 남고, 다음 결제부터 자동으로 안 채워져요.')) return;
            try {
              const j = await api('app_fixed_end', { i: f.i, check });
              close();
              toast(j.message, { undo: async () => {
                try {
                  if (j.deleted) await api('app_fixed_add', j.item); else await api('app_fixed_end', { i: f.i, check, undo: true });
                  toast('해지를 취소했어요'); refreshQuiet();
                } catch (e) { toast(e.message); }
              } });
              refreshQuiet();
            } catch (e) { toast(e.message); }
          } }, inc ? '그만 받기' : '해지하기'),
      h('button', { class: 'linkbtn', onclick: async () => {
        if (!confirm(f.n + '을(를) 완전히 지울까요?\n잘못 넣었을 때 쓰는 기능이에요. 지난달까지 자동으로 채워진 기록도 같이 사라져요.')) return;
        try { const j = await api('app_fixed_delete', { i: f.i, check }); close(); toast(j.message); refreshQuiet(); }
        catch (e) { toast(e.message); }
      } }, '잘못 넣었어요 · 기록까지 지우기'))
  });
}

function txRow(r) {
  const inc = r.g === '수입';
  const t2 = [PAY_SHORT[r.p] || r.p, r.c && r.n ? r.c : ''].filter(Boolean).join(' · ');
  return h('button', { class: 'row ind' + (r.fx ? ' ro' : ''), onclick: () => editTx(r) },
    h('div', { class: 'ico' }, CAT_ICON[r.c] || (r.c ? '•' : '❓')),
    h('div', { class: 'main' },
      h('div', { class: 't1' }, r.n || r.c || '(미분류)',
        r.fx ? h('span', { class: 'badge' }, '고정 자동') : null,
        r.ff ? h('span', { class: 'badge g' }, '고정') : null,
        !r.c ? h('span', { class: 'badge y' }, '확인') : null),
      h('div', { class: 't2' }, t2 || ' ')),
    h('div', { class: 'end amt' + (inc ? ' inc' : '') }, (inc ? '+' : '') + won(r.a)));
}

function calendar(ym, byDay) {
  const [y, m] = ym.split('-').map(Number);
  const first = new Date(y, m - 1, 1);
  const startPad = (first.getDay() + 6) % 7;                 // 월요일 시작
  const nDays = new Date(y, m, 0).getDate();
  const t = today();
  const cells = [];
  for (let i = 0; i < startPad; i++) cells.push(h('div', { class: 'd out' }));
  for (let d = 1; d <= nDays; d++) {
    const key = ym + '-' + pad2(d);
    const v = byDay[key] || 0;
    const tier = v >= 100000 ? 't3' : v >= 30000 ? 't2' : v > 0 ? 't1' : '';
    const past = key <= t;
    cells.push(h('button', {
      class: 'd ' + tier + (key === t ? ' today' : '') + (S.selDay === key ? ' sel' : '') + (!v && past ? ' nospend' : ''),
      onclick: () => { S.selDay = S.selDay === key ? null : key; render(); }
    }, h('div', { class: 'n' }, d), h('div', { class: 's' }, v ? man(v) : '')));
  }
  return [h('div', { class: 'wk' }, ['월', '화', '수', '목', '금', '토', '일'].map(x => h('div', null, x))), h('div', { class: 'days' }, cells)];
}

async function goMonth(k) {
  const cur = (S.ledger && S.ledger.month) || today().slice(0, 7);
  const nm = monthAdd(cur, k);
  S.selDay = null;
  S.ledger = { month: nm, rows: [], prevCat: {} };
  S.busy.boot = true; render();
  try { await loadLedger(nm); S.err = null; } catch (e) { S.err = e.message; }
  S.busy = {}; render();
}

async function classify(p, cat) {
  try {
    await api('app_update', { r: p.r, check: { d: p.d, a: p.a }, f: { c: cat } });
    S.boot.pending = S.boot.pending.filter(x => x !== p);
    toast((p.n || '') + ' → ' + cat);
    render();
    refreshQuiet();
  } catch (e) { toast(e.message); }
}

function editTx(r) {
  if (r.fx) {
    const fxItem = ((S.boot && S.boot.fixed) || []).find(f => f.n === r.n && f.c === r.c);
    openSheet({
      title: '고정비 자동', left: '닫기',
      body: h('div', null,
        h('div', { class: 'list' },
          field('내용', h('div', { class: 'ro' }, r.n)),
          field('금액', h('div', { class: 'ro' }, won(r.a))),
          field('날짜', h('div', { class: 'ro' }, dayLabel(r.d))),
          field('카테고리', h('div', { class: 'ro' }, r.c))),
        h('div', { class: 'help' }, '고정비로 매달 자동으로 채워진 줄이에요. 이번 달만 금액이 다르면 같은 이름으로 실제 금액을 입력하면 그 값으로 바뀌어요.'),
        fxItem ? h('button', { class: 'add-btn', onclick: () => { document.querySelector('.scrim').click(); setTimeout(() => editFixed(fxItem), 300); } },
          '고정비 설정 열기 (금액 변경·해지)') : null,
        h('button', { class: 'add-btn', onclick: () => {
          document.querySelector('.scrim').click();
          resetAdd(); Object.assign(S.add, { kind: '지출', cat: r.c, name: r.n, date: r.d, pay: r.p || S.add.pay });
          if (catsOf('수입').some(c => c.name === r.c)) S.add.kind = '수입';
          setTab('add');
        } }, '실제 금액으로 입력하기'))
    });
    return;
  }
  const E = { d: r.d, c: r.c, n: r.n, a: String(r.a), p: r.p, rc: r.rc, m: r.m };
  const kind = (S.boot.cats.find(c => c.name === r.c) || {}).kind || '지출';
  let curKind = kind;
  const catBox = h('div');
  function drawCats() {
    catBox.replaceChildren(
      h('div', { class: 'seg', style: { margin: '0 0 10px' } }, ['지출', '수입', '이체'].map(k =>
        h('button', { class: curKind === k ? 'on' : '', onclick: () => { curKind = k; drawCats(); } }, k))),
      h('div', { class: 'chips' }, catsOf(curKind).map(c =>
        h('button', { class: 'chip sm' + (E.c === c.name ? ' on' : ''), onclick: () => { E.c = c.name; drawCats(); } },
          (CAT_ICON[c.name] || '') + ' ' + c.name))));
  }
  drawCats();
  const date = h('input', { type: 'date', value: E.d }); date.addEventListener('change', () => E.d = date.value);
  const name = h('input', { type: 'text', value: E.n, placeholder: '내용' }); name.addEventListener('input', () => E.n = name.value);
  const amt = h('input', { type: 'text', inputmode: 'numeric', value: Number(E.a).toLocaleString('ko-KR') });
  amt.addEventListener('input', () => { const v = amt.value.replace(/[^\d-]/g, ''); E.a = v; amt.value = v && v !== '-' ? Number(v).toLocaleString('ko-KR') : v; });
  const pay = h('select', null, (S.boot.pays || []).map(p => h('option', { value: p.name, selected: p.name === E.p ? true : null }, PAY_SHORT[p.name] || p.name)));
  pay.addEventListener('change', () => E.p = pay.value);
  const rc = sw(E.rc, v => E.rc = v);
  const memo = h('input', { type: 'text', value: E.m, placeholder: '메모' }); memo.addEventListener('input', () => E.m = memo.value);
  const fixedSw = sw(r.ff, async (on, inp) => {
    inp.disabled = true;
    try {
      const j = await api('app_fixed', { r: r.r, check: { d: r.d, a: r.a }, on });
      inp.checked = j.on; r.ff = j.on; toast(j.message); refreshQuiet();
    } catch (e) { inp.checked = !on; toast(e.message); }
    inp.disabled = false;
  });

  const { close } = openSheet({
    title: '내역 고치기', right: '저장',
    onRight: async () => {
      const f = {};
      if (E.d !== r.d) f.d = E.d;
      if (E.c !== r.c) f.c = E.c;
      if (E.n !== r.n) f.n = E.n;
      if (Number(E.a) !== r.a) f.a = Number(E.a);
      if (E.p !== r.p) f.p = E.p;
      if (!!E.rc !== !!r.rc) f.rc = E.rc;
      if (E.m !== r.m) f.m = E.m;
      if (!Object.keys(f).length) return;
      await api('app_update', { r: r.r, check: { d: r.d, a: r.a }, f });
      toast('고쳤어요');
      refreshQuiet();
    },
    body: h('div', null,
      h('div', { class: 'card pad' }, catBox),
      h('div', { class: 'section-title' }, ' '),
      h('div', { class: 'list' },
        field('금액', amt), field('내용', name), field('날짜', date), field('결제수단', pay),
        E.p === '현금' || r.p === '현금' ? field('현금영수증', rc.el) : null,
        field('메모', memo)),
      (r.g === '지출' || kind === '지출') ? h('div', null,
        h('div', { class: 'list', style: { marginTop: '20px' } }, field('매달 나가는 돈', fixedSw.el)),
        h('div', { class: 'help' }, '켜면 고정지출 탭에 등록돼서 다음 달부터 자동으로 잡혀요.')) : null,
      h('button', { class: 'danger', onclick: async () => {
        if (!confirm('이 내역을 지울까요?')) return;
        try { const j = await api('app_delete', { r: r.r, check: { d: r.d, a: r.a } }); close(); toast(j.message); refreshQuiet(); }
        catch (e) { toast(e.message); }
      } }, '삭제'))
  });
}

function field(label, control, opt = {}) {
  return h('div', { class: 'field' + (opt.stack ? ' stack' : '') }, h('label', null, label), control);
}

/* ───────────── ③ 대시보드 (시트 그대로) ───────────── */

function viewDash() {
  const D = S.dash;
  const top = h('div', { class: 'top' }, h('span'), h('button', { class: 'act', onclick: () => autoLoad('dash', true), disabled: S.busy.dash ? true : null }, S.busy.dash ? '불러오는 중…' : '새로고침'));
  if (!D) return h('div', null, top, errBox(), h('div', { class: 'loading' }, h('span', { class: 'spin' }), '불러오는 중'));
  const rows = D.rows.map(r => r || []);
  const out = [top, errBox()];
  let sec = null, secTitle = '';
  const flushPairs = (ctx) => { if (ctx.pairs && ctx.pairs.children.length) ctx.card.append(ctx.pairs); ctx.pairs = null; };
  let ctx = null;
  const newSection = (title) => {
    sec = h('div', { class: 'dsec' }, h('h2', null, title));
    ctx = { card: h('div', { class: 'card' }), pairs: null, table: null };
    sec.append(ctx.card);
    out.push(sec);
    secTitle = title;
  };
  let i = 0;
  while (i < rows.length) {
    const r = rows[i];
    const cells = r.map((c, j) => c ? Object.assign({ j }, c) : null).filter(Boolean);
    if (!cells.length) { if (ctx) { flushPairs(ctx); ctx.table = null; } i++; continue; }
    const k0 = cells[0].k;
    if (k0 === 'title') { out.push(h('h1', { class: 'large' }, cells[0].v)); i++; continue; }
    if (k0 === 'sub' && !sec) { out.push(h('div', { class: 'sub' }, cells[0].v)); i++; continue; }
    if (k0 === 'sec') { newSection(cells[0].v); i++; continue; }
    if (!ctx) newSection('');
    if (k0 === 'para') { flushPairs(ctx); ctx.card.append(h('div', { class: 'para' }, cells[0].v)); i++; continue; }
    if (cells.some(c => c.k === 'th')) {
      flushPairs(ctx);
      const key = Object.keys(DASH_SPLIT).find(p => secTitle.startsWith(p));
      i = dashTables(rows, i, key ? DASH_SPLIT[key] : [], ctx.card);
      continue;
    }
    // 항목명-값 짝
    if (!ctx.pairs) ctx.pairs = h('div', { class: 'pairs' });
    const notes = [];
    for (let x = 0; x < cells.length; x++) {
      const c = cells[x];
      if ((c.k === 'lab' || c.k === 'rh') && cells[x + 1] && !['lab', 'rh', 'sec', 'th'].includes(cells[x + 1].k)) {
        ctx.pairs.append(pairEl(c.v, cells[x + 1])); x++;
      } else if (c.k === 'note' || c.k === 'sub') notes.push(c.v);
      else if (c.k === 'lab' || c.k === 'rh') ctx.pairs.append(pairEl(c.v, null));
      else if (cells.length === 1) { flushPairs(ctx); ctx.card.append(h('div', { class: 'para' }, c.v)); }
      else ctx.pairs.append(pairEl('', c));
    }
    if (notes.length) { flushPairs(ctx); ctx.card.append(h('div', { class: 'note', style: { paddingTop: '10px' } }, notes.join(' · '))); }
    i++;
  }
  if (ctx) flushPairs(ctx);
  // 추이
  const snap = D.snap || [];
  out.push(h('div', { class: 'dsec' }, h('h2', null, '추이 (스냅샷)'),
    h('div', { class: 'card' }, snap.length >= 2 ? trendChart(snap) :
      h('div', { class: 'note', style: { paddingTop: '12px' } }, '매월 1일 기록이 두 번 이상 쌓이면 금융자산과 목표 경로 그래프가 생겨요.'))));
  // 짝이 홀수면 마지막 칸을 넓게
  setTimeout(() => document.querySelectorAll('.pairs').forEach(p => {
    const n = p.children.length;
    if (n % 2 === 1) p.lastElementChild.classList.add('wide');
  }), 0);
  return h('div', { class: 'dash' }, out);
}

function pairEl(label, cell) {
  const v = cell ? cell.v : '';
  const tint = cell && cell.t ? ' t-' + cell.t : '';
  const isProg = /진행률|비중$/.test(label) && /^-?[\d.]+%$/.test(v);
  const long = String(v).length > 14 || String(label).length > 12;
  const el = h('div', { class: 'pair' + (/진행률/.test(label) ? ' big wide' : long ? ' wide' : '') },
    label ? h('div', { class: 'k' }, label) : null,
    h('div', { class: 'v' + tint }, v));
  if (isProg && /진행률/.test(label)) el.append(h('div', { class: 'progress' }, h('div', { style: { width: Math.max(0, Math.min(100, parseFloat(v))) + '%' } })));
  return el;
}

// 표 그리기: split = 나눌 열 번호들. 끝난 다음 줄 번호를 돌려줘요
function dashTables(rows, start, split, card) {
  const width = 7;
  const bounds = [0].concat(split, [width]);
  const groups = [];
  for (let g = 0; g < bounds.length - 1; g++) groups.push([bounds[g], bounds[g + 1]]);
  let i = start;
  const tables = groups.map(() => []);        // 그룹마다 [{head, body}]
  const isTableRow = r => {
    const cells = r.filter(Boolean);
    if (!cells.length) return false;
    return !cells.some(c => ['sec', 'para', 'title'].includes(c.k)) && !(cells[0].k === 'lab');
  };
  while (i < rows.length && isTableRow(rows[i] || [])) {
    const r = rows[i] || [];
    groups.forEach(([a, b], g) => {
      const seg = r.slice(a, b);
      if (!seg.some(Boolean)) return;
      const first = seg.find(Boolean);
      if (first.k === 'th') tables[g].push({ head: seg, body: [] });
      else if (tables[g].length) tables[g][tables[g].length - 1].body.push(seg);
    });
    i++;
  }
  tables.flat().forEach(t => {
    const used = t.head.map((c, j) => !!c || t.body.some(b => b[j]));
    const cols = used.map((u, j) => u ? j : -1).filter(j => j >= 0);
    card.append(h('div', { class: 'tbl-wrap' }, h('table', { class: 'tbl' },
      h('thead', null, h('tr', null, cols.map(j => h('th', null, t.head[j] ? t.head[j].v : '')))),
      h('tbody', null, t.body.map(b => h('tr', null, cols.map((j, x) => {
        const c = b[j];
        return h('td', { class: (c && (c.k === 'val' || c.k === 'rh') && x > 0 ? 'b' : '') + (c && c.t ? ' t-' + c.t : '') }, c ? c.v : '');
      })))))));
  });
  return i;
}

function trendChart(snap) {
  const W = 320, H = 140, P = 8;
  const xs = snap.map(s => parseYmd(s.d).getTime());
  const vals = snap.flatMap(s => [s.fin, s.path].filter(v => typeof v === 'number'));
  const x0 = Math.min(...xs), x1 = Math.max(...xs), y0 = Math.min(...vals) * 0.97, y1 = Math.max(...vals) * 1.03;
  const X = t => P + (W - 2 * P) * (x1 === x0 ? 0.5 : (t - x0) / (x1 - x0));
  const Y = v => H - P - (H - 2 * P) * ((v - y0) / (y1 - y0 || 1));
  const line = key => snap.filter(s => typeof s[key] === 'number').map((s, k) => (k ? 'L' : 'M') + X(parseYmd(s.d).getTime()).toFixed(1) + ' ' + Y(s[key]).toFixed(1)).join(' ');
  const svg = `<svg viewBox="0 0 ${W} ${H}" preserveAspectRatio="none"><path d="${line('path')}" fill="none" stroke="var(--label3)" stroke-width="2" stroke-dasharray="4 4"/><path d="${line('fin')}" fill="none" stroke="var(--blue)" stroke-width="2.5" stroke-linejoin="round"/></svg>`;
  const last = snap[snap.length - 1];
  return h('div', { class: 'chart' }, h('div', { html: svg }),
    h('div', { class: 'legend' }, h('span', null, h('i', { style: { background: 'var(--blue)' } }), '금융자산 ' + man(last.fin)),
      h('span', null, h('i', { style: { background: 'var(--label3)' } }), '목표 경로 ' + (last.path ? man(last.path) : '–'))));
}

/* ───────────── ④ 자산: 보유종목·현금·기타 ───────────── */

function viewAssets() {
  const seg = h('div', { class: 'seg', style: { marginBottom: '14px' } },
    [['hold', '보유종목'], ['cash', '현금·기타']].map(([k, l]) =>
      h('button', { class: S.assetSeg === k ? 'on' : '', onclick: () => { S.assetSeg = k; store.set('assetSeg', k); render(); autoLoad('assets'); } }, l)));
  const T = S.sheets[S.assetSeg];
  const head = [h('h1', { class: 'large' }, '자산'), seg, errBox()];
  if (!T) return h('div', null, head, h('div', { class: 'loading' }, h('span', { class: 'spin' }), '불러오는 중'));
  if (S.busy[S.assetSeg]) head.push(h('div', { class: 'stale' }, h('span', { class: 'spin' }), '새로 불러오는 중'));
  return h('div', null, head, S.assetSeg === 'hold' ? holdView(T) : cashView(T));
}

const col = (T, re) => T.head.findIndex(x => re.test(x));

function holdView(T) {
  const iName = col(T, /^종목명/), iTick = col(T, /^티커/), iAcc = col(T, /^계좌$/), iQty = col(T, /^수량/), iVal = col(T, /^평가금액/),
        iCost = col(T, /^매입금액/), iPnl = col(T, /^손익/), iRet = col(T, /^수익률/), iW = col(T, /^비중/), iPx = col(T, /^현재가/), iCur = col(T, /^통화/);
  const num = (r, i) => typeof r.v[i] === 'number' ? r.v[i] : 0;
  const tot = T.rows.reduce((s, r) => s + num(r, iVal), 0), cost = T.rows.reduce((s, r) => s + num(r, iCost), 0), pnl = T.rows.reduce((s, r) => s + num(r, iPnl), 0);
  const out = [h('div', { class: 'card pad sum' },
    h('div', { class: 'caption' }, '주식 평가금액'),
    h('div', { class: 'big' }, won(tot)),
    h('div', { class: (pnl >= 0 ? 'pos' : 'neg'), style: { fontWeight: 600, marginTop: '2px' } }, signed(pnl) + ' (' + pct(cost ? pnl / cost : null, 2) + ')'))];
  const groups = {};
  T.rows.forEach(r => { const g = r.v[iAcc] || '기타'; (groups[g] = groups[g] || []).push(r); });
  Object.keys(groups).forEach(g => {
    const gs = groups[g].reduce((s, r) => s + num(r, iVal), 0);
    out.push(h('div', { class: 'section-title', style: { display: 'flex', justifyContent: 'space-between' } }, h('span', null, g), h('span', { class: 'num' }, won(gs))));
    out.push(h('div', { class: 'list' }, groups[g].map(r => {
      const ret = r.v[iRet];
      const pxBad = typeof r.v[iPx] !== 'number' && r.v[iTick];
      return h('button', { class: 'row', onclick: () => editRow('hold', T, r) },
        h('div', { class: 'main' },
          h('div', { class: 't1' }, r.v[iName] || r.v[iTick]),
          h('div', { class: 't2' }, [r.v[iTick], r.d[iQty] ? r.d[iQty] + '주' : '', r.d[iW] ? '비중 ' + r.d[iW] : ''].filter(Boolean).join(' · '),
            pxBad ? h('span', { class: 'badge y' }, '시세 확인') : null)),
        h('div', { class: 'end' },
          h('div', { class: 'amt' }, r.d[iVal] ? r.d[iVal] + '원' : '–'),
          h('div', { class: 'caption ' + (typeof ret === 'number' ? (ret >= 0 ? 'pos' : 'neg') : '') }, r.d[iRet] || (r.v[iCur] === 'USD' ? 'USD' : ''))));
    })));
  });
  out.push(h('button', { class: 'add-btn', onclick: () => editRow('hold', T, null) }, '＋ 종목 추가'));
  out.push(h('div', { class: 'help' }, '시세는 구글 시세라 약 20분 늦어요. 티커 예: KRX:379810, NASDAQ:QQQ, NYSEARCA:SCHD'));
  return h('div', null, out);
}

function cashView(T) {
  const iName = col(T, /^항목/), iKind = col(T, /^구분/), iAmt = col(T, /^금액/), iInc = col(T, /포함/), iEff = col(T, /^반영/), iMemo = col(T, /^메모/);
  const out = [];
  out.push(h('div', { class: 'list' }, T.rows.map(r => {
    const eff = typeof r.v[iEff] === 'number' ? r.v[iEff] : (typeof r.v[iAmt] === 'number' ? r.v[iAmt] : 0);
    return h('button', { class: 'row', onclick: () => editRow('cash', T, r) },
      h('div', { class: 'main' },
        h('div', { class: 't1' }, r.v[iName], r.v[iInc] === true ? h('span', { class: 'badge g' }, '10억 집계') : null),
        h('div', { class: 't2' }, [r.v[iKind], r.v[iMemo]].filter(Boolean).join(' · '))),
      h('div', { class: 'end amt' + (eff < 0 ? ' neg' : '') }, won(eff)));
  })));
  out.push(h('button', { class: 'add-btn', onclick: () => editRow('cash', T, null) }, '＋ 항목 추가'));
  return h('div', null, out);
}

const NUM_HEAD = /수량|단가|환율|배당$|현재가|금액|원\)/;
function editRow(sheet, T, r) {
  const isNew = !r;
  const vals = {};
  const controls = [];
  T.head.forEach((name, j) => {
    if (!name) return;
    const cur = r ? r.v[j] : '';
    if (T.formula[j]) {
      if (r && r.d[j]) controls.push(field(name, h('div', { class: 'ro' }, r.d[j])));
      return;
    }
    const opts = T.opts[j];
    let ctl;
    if (opts === 'bool') {
      const s = sw(cur === true, v => vals[j] = { v, type: 'bool' });
      ctl = s.el;
    } else if (Array.isArray(opts)) {
      ctl = h('select', null, h('option', { value: '' }, '선택'), opts.map(o => h('option', { value: o, selected: o === cur ? true : null }, o)));
      ctl.addEventListener('change', () => vals[j] = { v: ctl.value, type: 'text' });
    } else {
      const isNum = typeof cur === 'number' || NUM_HEAD.test(name);
      ctl = h('input', { type: 'text', inputmode: isNum ? 'decimal' : 'text', value: cur === '' || cur === null ? '' : String(cur), placeholder: isNum ? '0' : '' });
      if (/티커/.test(name)) { ctl.setAttribute('autocapitalize', 'characters'); ctl.setAttribute('placeholder', 'KRX:379810'); }
      ctl.addEventListener('input', () => vals[j] = { v: isNum ? ctl.value.replace(/[^\d.-]/g, '') : ctl.value, type: isNum ? 'num' : 'text' });
    }
    controls.push(field(name, ctl));
  });
  const title = sheet === 'hold' ? (isNew ? '종목 추가' : (r.v[col(T, /^종목명/)] || '종목')) : (isNew ? '항목 추가' : r.v[0]);
  const { close } = openSheet({
    title, right: isNew ? '추가' : '저장',
    onRight: async () => {
      const cells = Object.keys(vals).map(j => ({ c: Number(j) + 1, v: vals[j].type === 'bool' ? !!vals[j].v : vals[j].v, type: vals[j].type === 'bool' ? 'text' : vals[j].type }));
      if (!cells.length) return;
      if (isNew) await api('app_row_add', { sheet, cells });
      else await api('app_set', { sheet, cells: cells.map(c => Object.assign({ r: r.r }, c)) });
      toast(isNew ? '추가했어요' : '저장했어요');
      lastLoad[sheet] = 0; lastLoad.dash = 0;
      await loadSheet(sheet).catch(() => {});
      render();
    },
    body: h('div', null,
      h('div', { class: 'list' }, controls),
      sheet === 'hold' ? h('div', { class: 'help' }, '회색 값은 시트가 자동으로 계산해요. 해외 종목은 통화를 USD로, 평균환율은 비워 두면 현재 환율을 써요.') : null,
      !isNew ? h('button', { class: 'danger', onclick: async () => {
        if (!confirm('이 줄을 비울까요? (자동 계산 칸은 남아요)')) return;
        try { await api('app_row_clear', { sheet, r: r.r, check: r.v[0] }); close(); toast('지웠어요'); lastLoad[sheet] = 0; await loadSheet(sheet); render(); }
        catch (e) { toast(e.message); }
      } }, '삭제') : null)
  });
}

/* ───────────── ⑤ 설정 ───────────── */

function viewSettings() {
  const T = S.sheets.set;
  const out = [h('h1', { class: 'large' }, '설정'), errBox()];
  if (!T) out.push(h('div', { class: 'loading' }, h('span', { class: 'spin' }), '불러오는 중'));
  else {
    if (S.busy.set) out.push(h('div', { class: 'stale' }, h('span', { class: 'spin' }), '새로 불러오는 중'));
    let list = null;
    const startList = (title) => { out.push(h('div', { class: 'section-title' }, title)); list = h('div', { class: 'list' }); out.push(list); };
    startList('목표와 가정');
    T.items.forEach(it => {
      if (it.sec) { startList(it.sec); return; }
      list.append(h(it.edit ? 'button' : 'div', { class: 'row' + (it.edit ? '' : ' ro'), onclick: it.edit ? () => editSetting(it) : null },
        h('div', { class: 'main' }, h('div', { class: 't1' }, it.label), it.desc ? h('div', { class: 't2' }, it.desc) : null),
        h('div', { class: 'end num', style: { color: it.edit ? 'var(--blue)' : 'var(--label2)' } }, it.d || '–')));
    });
  }
  out.push(h('div', { class: 'section-title' }, '앱'));
  out.push(h('div', { class: 'list' },
    h('a', { class: 'row', href: SHEET_URL, target: '_blank', rel: 'noopener', style: { color: 'inherit', textDecoration: 'none' } },
      h('div', { class: 'main' }, h('div', { class: 't1' }, '구글 시트 열기')), h('span', { class: 'chev' }, '›')),
    h('button', { class: 'row', onclick: () => { ['c_boot', 'c_dash', 'c_hold', 'c_cash', 'c_set'].forEach(store.del); location.reload(); } },
      h('div', { class: 'main' }, h('div', { class: 't1' }, '저장된 화면 지우고 새로 불러오기'))),
    h('button', { class: 'row', onclick: () => { if (confirm('이 기기에서 비밀키를 지울까요?')) { store.del('key'); S.key = ''; render(); } } },
      h('div', { class: 'main' }, h('div', { class: 't1', style: { color: 'var(--red)' } }, '비밀키 지우기'))),
    h('div', { class: 'row ro' }, h('div', { class: 'main' }, h('div', { class: 't1' }, '버전')), h('div', { class: 'end caption' }, VERSION))));
  return h('div', null, out);
}

function editSetting(it) {
  let ctl, get;
  if (it.type === 'opt') {
    ctl = h('select', null, it.opts.map((o, k) => h('option', { value: k, selected: o === it.v ? true : null }, it.optsD[k] || o)));
    get = () => ({ v: it.opts[Number(ctl.value)], type: 'text' });
  } else if (it.type === 'date') {
    ctl = h('input', { type: 'date', value: it.v || '' });
    get = () => ({ v: ctl.value, type: 'date' });
  } else if (it.type === 'pct') {
    ctl = h('input', { type: 'text', inputmode: 'decimal', value: typeof it.v === 'number' ? +(it.v * 100).toFixed(4) : '' });
    get = () => ({ v: ctl.value === '' ? '' : Number(ctl.value) / 100, type: 'pct' });
  } else if (it.type === 'num') {
    ctl = h('input', { type: 'text', inputmode: 'decimal', value: typeof it.v === 'number' ? it.v.toLocaleString('ko-KR') : String(it.v || '') });
    ctl.addEventListener('input', () => { const raw = ctl.value.replace(/[^\d.-]/g, ''); if (raw && !raw.endsWith('.')) ctl.value = Number(raw).toLocaleString('ko-KR'); });
    get = () => ({ v: ctl.value.replace(/[^\d.-]/g, ''), type: 'num' });
  } else {
    ctl = h('input', { type: 'text', value: it.v || '' });
    get = () => ({ v: ctl.value, type: 'text' });
  }
  openSheet({
    title: it.label, right: '저장',
    onRight: async () => {
      const g = get();
      await api('app_set', { sheet: 'set', cells: [{ r: it.r, c: 2, v: g.v, type: g.type }] });
      toast('저장했어요');
      lastLoad.set = 0; lastLoad.dash = 0;
      await loadSheet('set').catch(() => {});
      render();
    },
    body: h('div', null,
      h('div', { class: 'list' }, field(it.type === 'pct' ? '값 (%)' : '값', ctl)),
      it.desc ? h('div', { class: 'help' }, it.desc) : null)
  });
  setTimeout(() => ctl.focus && ctl.focus(), 300);
}

/* ───────────── 시작 ───────────── */

function init() {
  document.querySelectorAll('#tabs button').forEach(b => b.addEventListener('click', () => setTab(b.dataset.tab)));
  const q = new URLSearchParams(location.search).get('tab');
  S.tab = ['add', 'ledger', 'dash', 'assets', 'settings'].includes(q) ? q : 'add';
  resetAdd();
  if (S.boot) { S.ledger = S.boot.ledger; S.month = S.boot.ledger && S.boot.ledger.month; }
  document.querySelectorAll('#tabs button').forEach(b => b.classList.toggle('on', b.dataset.tab === S.tab));
  render();
  if (S.key) { autoLoad(S.tab); flushOutbox(); }
  window.addEventListener('online', flushOutbox);
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'visible' && S.key) {
      if (S.tab === 'add' && S.add && !S.add.amt) { S.add.date = ymdOf(new Date()); }
      autoLoad(S.tab); flushOutbox();
    }
  });
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
}
init();
