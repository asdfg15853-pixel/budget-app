/**
 * ─────────────────────────────────────────────────────────────
 *  나의 자산관리 · 자동 기록 스크립트 (v2: 문자 자동입력 추가)
 * ─────────────────────────────────────────────────────────────
 *  하는 일
 *   1) [자산관리] 메뉴
 *   2) 스냅샷 기록 (수동 / 매월 1일 자동)
 *   3) 카드 결제 문자 → 가계부 자동 입력 (아이폰 단축어가 이 스크립트로 문자를 보내요)
 *      - 신한 체크: 문자 형식을 직접 읽어요
 *      - 그 밖의 형식(현대카드 등): Claude API 키가 있으면 AI가 읽어요
 *      - 카테고리: 고정지출 > 가맹점 규칙 > AI 분류(키가 있을 때) 순서로 정해요
 *      - 노란 칸(확인 필요)에서 카테고리를 고르면 그 가맹점이 규칙으로 저장돼요 (다음부터 자동)
 *      - KG이니시스·네이버페이 같은 결제대행사는 이름만으론 알 수 없어서 규칙을 만들지 않고 매번 확인받아요
 *   4) 가계부 분류 단축어: 확인할 결제를 폰에서 하나씩 받아 카테고리를 고르면 시트에 적혀요
 *   5) 가계부 H열 '고정 등록'을 체크하면 그 줄이 고정지출 표로 들어가요 (체크를 풀면 다시 빠져요)
 *   6) 가계부 앱(깃허브 웹앱)이 쓰는 API: app_boot·app_ledger·app_dash·app_sheet·app_add 등 (맨 아래)
 *
 *  업데이트 방법 (기존 v1에서)
 *   1) [확장 프로그램] > [Apps Script] → 코드를 모두 지우고 이 파일 전체 붙여넣기 → 저장
 *   2) 시트로 돌아가 새로고침 → [자산관리] > [문자 자동입력 설정] → 권한 승인
 *   3) [배포] > [새 배포] > 유형: 웹 앱, 실행: 나, 액세스: 모든 사용자 → 배포 → 웹 앱 URL 복사
 *   4) [자산관리] > [문자 자동입력 설정]을 다시 열어 URL을 붙여넣기
 *   (매월 스냅샷 자동 기록은 그대로 유지돼요)
 */

const SNAP_SHEET = '스냅샷';
const HOLD_SHEET = '보유종목';
const PRICE_COL = 10;    // 보유종목 J열(현재가)
const TICKER_COL = 4;    // 보유종목 D열(티커)
const HANDLER = 'autoSnapshot';

// 스냅샷 B~G열에 들어갈 값 (대시보드·설정의 이름 있는 범위)
const SNAP_KEYS = ['KPI_FIN', 'KPI_PATH', 'KPI_NET', 'KPI_LEV', 'KPI_DIV', 'FX_USDKRW'];


/* ───────────────────────── 메뉴 ───────────────────────── */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('자산관리')
    .addItem('지금 스냅샷 기록', 'snapshotNow')
    .addSeparator()
    .addItem('매월 자동 기록 켜기', 'install')
    .addItem('매월 자동 기록 끄기', 'uninstall')
    .addItem('자동 기록 상태 확인', 'status')
    .addSeparator()
    .addItem('문자 자동입력 설정', 'smsSetup')
    .addItem('문자 자동입력 테스트', 'smsTest')
    .addToUi();
}


/* ───────────────────────── 기록 ───────────────────────── */

// 메뉴에서 누르는 수동 기록: 오늘 날짜
function snapshotNow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const r = record_(ss, todayParts_(ss), '수동');
  notify_(ss, r.ok
    ? `${r.date} 기준으로 기록했어요${r.replaced ? ' (같은 날짜 줄을 고쳐 썼어요)' : ''}.`
    : `${r.date} 기준으로 기록했지만 일부 시세를 못 받아왔어요. 기록 방식 칸에 표시해 두었어요.`);
}

// 매월 1일 오전 7시 트리거: 지난달 말일 날짜
function autoSnapshot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const t = todayParts_(ss);
  record_(ss, prevMonthEnd_(t), '자동');
}

function record_(ss, ymd, how) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const snap = ss.getSheetByName(SNAP_SHEET);
    if (!snap) throw new Error(`'${SNAP_SHEET}' 탭을 찾지 못했어요.`);

    refreshPrices_(ss);
    const ready = waitForPrices_(ss);

    const vals = SNAP_KEYS.map(k => {
      const rg = ss.getRangeByName(k);
      return rg ? rg.getValue() : '';
    });
    const ok = ready && vals.every(v => typeof v === 'number' && isFinite(v));
    const label = how + (ok ? '' : ' · 시세 확인 필요');

    const dateStr = fmtYmd_(ymd);
    const row = findRow_(snap, ss.getSpreadsheetTimeZone(), dateStr);
    const target = row || Math.max(snap.getLastRow() + 1, 2);
    snap.getRange(target, 1, 1, 8).setValues([[serial_(ymd)].concat(vals, [label])]);
    snap.getRange(target, 1).setNumberFormat('yyyy-mm-dd');

    // 시작점 아래 기록 줄을 날짜순으로 정렬 (추이 차트 선이 꼬이지 않게)
    const last = snap.getLastRow();
    if (last >= 4) snap.getRange(3, 1, last - 2, 8).sort({ column: 1, ascending: true });
    SpreadsheetApp.flush();

    return { ok: ok, date: dateStr, replaced: !!row };
  } finally {
    lock.releaseLock();
  }
}

// 같은 날짜의 기존 기록 줄 번호 (시작점 줄은 제외), 없으면 0
function findRow_(snap, tz, dateStr) {
  const last = snap.getLastRow();
  if (last < 2) return 0;
  const rows = snap.getRange(2, 1, last - 1, 8).getValues();
  for (let i = 0; i < rows.length; i++) {
    const a = rows[i][0], h = String(rows[i][7] || '');
    if (h.indexOf('시작점') === 0) continue;
    if (isDate_(a) && Utilities.formatDate(a, tz, 'yyyy-MM-dd') === dateStr) return i + 2;
  }
  return 0;
}


/* ───────────────────────── 시세 새로고침 ───────────────────────── */

// 현재가 수식을 지웠다가 다시 넣어 GOOGLEFINANCE를 새로 계산시켜요.
// 수식이 아닌 칸(직접 입력한 값)은 그대로 되돌려 놓아요.
function refreshPrices_(ss) {
  const hold = ss.getSheetByName(HOLD_SHEET);
  if (hold) {
    const last = hold.getLastRow();
    if (last >= 2) {
      const rng = hold.getRange(2, PRICE_COL, last - 1, 1);
      const formulas = rng.getFormulas();
      const values = rng.getValues();
      const restore = formulas.map((r, i) => [r[0] || values[i][0]]);
      if (formulas.some(r => r[0])) {
        try {
          rng.clearContent();
          SpreadsheetApp.flush();
        } finally {
          rng.setValues(restore);
        }
      }
    }
  }
  const fx = ss.getRangeByName('FX_USDKRW');
  const ff = fx && fx.getFormula();
  if (ff) {
    try {
      fx.clearContent();
      SpreadsheetApp.flush();
    } finally {
      fx.setFormula(ff);
    }
  }
  SpreadsheetApp.flush();
}

// 티커가 있는 줄의 현재가와 환율이 모두 숫자가 될 때까지 최대 40초 기다려요.
function waitForPrices_(ss) {
  const hold = ss.getSheetByName(HOLD_SHEET);
  for (let i = 0; i < 10; i++) {
    let pending = 0;
    if (hold && hold.getLastRow() >= 2) {
      const n = hold.getLastRow() - 1;
      const tick = hold.getRange(2, TICKER_COL, n, 1).getValues();
      const px = hold.getRange(2, PRICE_COL, n, 1).getValues();
      pending = tick.filter((t, j) => t[0] !== '' && typeof px[j][0] !== 'number').length;
    }
    const fxr = ss.getRangeByName('FX_USDKRW');
    const fx = fxr ? fxr.getValue() : 0;
    if (!pending && typeof fx === 'number' && fx > 0) return true;
    Utilities.sleep(4000);
    SpreadsheetApp.flush();
  }
  return false;
}


/* ───────────────────────── 자동 실행 켜기·끄기 ───────────────────────── */

function install() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  removeTriggers_();
  ScriptApp.newTrigger(HANDLER)
    .timeBased()
    .onMonthDay(1)
    .atHour(7)
    .inTimezone(ss.getSpreadsheetTimeZone())
    .create();
  notify_(ss, '자동 기록을 켰어요. 매월 1일 오전 7~8시에 지난달 말 기준으로 한 줄씩 쌓여요.');
}

function uninstall() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const n = removeTriggers_();
  notify_(ss, n ? '자동 기록을 껐어요. 메뉴의 [지금 스냅샷 기록]은 계속 쓸 수 있어요.' : '켜져 있는 자동 기록이 없어요.');
}

function status() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const on = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === HANDLER);
  const snap = ss.getSheetByName(SNAP_SHEET);
  let lastMsg = '아직 기록이 없어요.';
  if (snap && snap.getLastRow() >= 2) {
    const r = snap.getRange(snap.getLastRow(), 1, 1, 8).getValues()[0];
    const d = isDate_(r[0]) ? Utilities.formatDate(r[0], ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd') : String(r[0]);
    lastMsg = `마지막 기록: ${d} (${r[7] || '-'})`;
  }
  notify_(ss, (on ? '자동 기록: 켜짐 (매월 1일 오전 7~8시)' : '자동 기록: 꺼짐') + '\n' + lastMsg);
}

function removeTriggers_() {
  const ts = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === HANDLER);
  ts.forEach(t => ScriptApp.deleteTrigger(t));
  return ts.length;
}


/* ───────────────────────── 날짜·공통 ───────────────────────── */

// 시트 시간대 기준 오늘 {y, m, d}
function todayParts_(ss) {
  const s = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd').split('-');
  return { y: +s[0], m: +s[1], d: +s[2] };
}

// 지난달 말일 {y, m, d}
function prevMonthEnd_(t) {
  const u = new Date(Date.UTC(t.y, t.m - 1, 0));
  return { y: u.getUTCFullYear(), m: u.getUTCMonth() + 1, d: u.getUTCDate() };
}

// 시트 날짜 일련번호 (1899-12-30 = 0) — 시간대 영향 없이 날짜만 넣기 위해 숫자로 써요
function serial_(t) {
  return (Date.UTC(t.y, t.m - 1, t.d) - Date.UTC(1899, 11, 30)) / 86400000;
}

function isDate_(v) {
  return v !== null && typeof v === 'object' && typeof v.getTime === 'function' && !isNaN(v.getTime());
}

function fmtYmd_(t) {
  const p = n => (n < 10 ? '0' : '') + n;
  return `${t.y}-${p(t.m)}-${p(t.d)}`;
}

// 시트에서 메뉴로 실행하면 알림창, 편집기에서 실행하면 실행 로그에 남겨요
function notify_(ss, msg) {
  Logger.log(msg);
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    try { ss.toast(msg, '자산관리', 8); } catch (e2) { /* 무시 */ }
  }
}


/* ═════════════════════════ 문자 자동입력 ═════════════════════════ */

const SS_ID = '1opK41aXB8vl_7LGYoYTdYdV5QQHiGS3nxYsvmYzDAcQ';
const LEDGER_SHEET = '가계부';
const RULE_SHEET = '가맹점 규칙';
const FIXED_SHEET = '고정지출';
const LOG_SHEET = '문자 로그';
const SET_SHEET = '설정';
const CLAUDE_MODEL = 'claude-haiku-5-5';
const SMS_TZ = 'Asia/Seoul';
const CARD_SHINHAN = '신한 체크';
const CARD_HYUNDAI = '네이버 현대카드';
// 결제대행사(PG)·간편결제: 문자에 실제 가게 대신 이 이름이 찍혀서, 이름만으로는 뭘 샀는지 알 수 없어요.
// 이런 결제는 규칙을 만들지 않고 매번 '확인 필요'로 남겨요.
const PG_NAMES = ['이니시스', 'KCP', '한국사이버결제', '토스페이먼츠', '나이스페이', 'NICEPAY', '나이스정보통신',
  '다날', '네이버페이', 'NAVERPAY', '카카오페이', 'KAKAOPAY', '페이코', 'PAYCO', '스마일페이', 'SSG페이', 'KSNET',
  '갤럭시아', '모빌리언스', '페이레터', '헥토', '세틀뱅크', 'KICC', '한국정보통신', '이지페이', '페이팔', 'PAYPAL'];

function isPg_(merchant) {
  const m = norm_(merchant);
  return PG_NAMES.some(k => m.indexOf(norm_(k)) >= 0);
}

function book_() {
  try {
    const a = SpreadsheetApp.getActiveSpreadsheet();
    if (a) return a;
  } catch (e) { /* 웹 앱으로 실행되면 활성 시트가 없을 수 있어요 */ }
  return SpreadsheetApp.openById(SS_ID);
}

// 아이폰 단축어가 보내는 곳 (웹 앱 POST)
//  - 문자 자동입력: {key, text}
//  - 가계부 분류 단축어: {key, action:'classify', category} (row를 빼면 맨 앞의 확인할 건)
function doPost(e) {
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { body = {}; }
  if (e && e.parameter) {
    ['key', 'text', 'action', 'row', 'category'].forEach(k => { if (body[k] === undefined && e.parameter[k] !== undefined) body[k] = e.parameter[k]; });
  }
  if (!keyOk_(body.key)) return out_({ ok: false, error: '비밀키가 맞지 않아요' });
  try {
    if (String(body.action || '').indexOf('app_') === 0) return out_(appRoute_(body));   // 가계부 앱(웹앱)
    if (body.action === 'classify') return out_(classifyRow_(body.row, body.category));
    return out_(handleSms_(String(body.text || ''), new Date(), false));
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

// 웹 앱 GET
//  - 그냥 열면: 주소가 살아 있는지 확인
//  - 가계부 분류 단축어: ?key=…&action=next → 다음에 확인할 결제 한 건
function doGet(e) {
  const prm = (e && e.parameter) || {};
  if (prm.action === 'next') {
    if (!keyOk_(prm.key)) return out_({ ok: false, error: '비밀키가 맞지 않아요' });
    return out_(nextPending_());
  }
  return out_({ ok: true, message: '문자 자동입력 주소가 정상이에요' });
}

function keyOk_(k) {
  const s = PropertiesService.getScriptProperties().getProperty('SMS_SECRET');
  return !!s && String(k || '') === s;
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// 문자 한 통 처리. dry = true면 가계부·로그에 쓰지 않고 결과만 돌려줘요 (테스트용)
function handleSms_(text, now, dry) {
  const ss = book_();
  const lock = LockService.getDocumentLock();
  lock.waitLock(25000);
  try {
    const log = ss.getSheetByName(LOG_SHEET);
    let p = parseSms_(text, now);
    let via = '';
    if (!p) { p = aiParse_(text, now); via = ' (AI 해석)'; }
    if (!p) return finish_(log, now, text, null, '해석 실패', '', '', dry);
    if (p.ignore) return finish_(log, now, text, p, '무시: ' + p.ignore, '', '', dry);

    const key = [p.card, fmtYmd_(p.date), p.time, p.amount, norm_(p.merchant), p.cancel ? '취소' : '승인'].join('|');
    if (log && isDup_(log, key)) return finish_(log, now, text, p, '중복이라 건너뜀', '', key, dry);

    const c = classify_(ss, p);
    const memo = ['문자 자동 ' + p.time, c.note].filter(Boolean).join(' · ');
    const row = [serial_(p.date), c.category, c.name, p.cancel ? -p.amount : p.amount, p.card, '', memo];
    const at = dry ? '' : appendLedger_(ss, row);
    const res = finish_(log, now, text, p, '입력됨' + via, at, key, dry, row);
    // 확인이 필요한 결제면 단축어가 폰 알림으로 띄울 문구를 같이 돌려줘요
    res.notify = (!c.category || c.note.indexOf('확인 필요') >= 0)
      ? c.name + ' ' + won_(row[3]) + ' · 가계부에서 카테고리를 골라 주세요' : '';
    return res;
  } finally {
    lock.releaseLock();
  }
}

function finish_(log, now, text, p, result, at, key, dry, row) {
  if (!dry && log) {
    log.appendRow([now, text, p ? p.card || '' : '', p && !p.ignore ? (p.cancel ? '취소' : '승인') : '',
      p && p.date ? serial_(p.date) : '', p ? p.time || '' : '', p ? p.amount || '' : '',
      p ? p.merchant || '' : '', p ? p.cumulative || '' : '', result, at, key]);
  }
  return { ok: true, result: result, row: at, entry: row || null, parsed: p || null };
}

function isDup_(log, key) {
  const n = log.getLastRow();
  if (n < 2) return false;
  const v = log.getRange(2, 10, n - 1, 3).getValues();   // J 처리 결과, K 행, L 키
  return v.some(r => r[2] === key && String(r[0]).indexOf('입력됨') === 0);
}

/* ── 문자 해석 ── */

function parseSms_(text, now) {
  const t = String(text || '').replace(/\[Web발신\]/g, ' ').replace(/\s+/g, ' ').trim();
  if (!t) return null;
  // 해외 외화 결제(SOL트래블 등)는 지금은 넣지 않아요
  if (/해외(승인|취소)/.test(t) && /\b[A-Z]{3}\s?[\d,.]+/.test(t)) return { ignore: '해외 외화 결제' };
  // 신한 체크: [신한체크승인] 박*찬(1505) 09/24 22:48 11,100원 쿠팡 잔액1,719,024원
  const m = t.match(/\[신한체크(승인|취소)\]\s*\S*?\(\d{4}\)\s*(\d{1,2})\/(\d{1,2})\s+(\d{1,2}):(\d{2})\s+([\d,]+)원\s+(.+?)(?:\s*잔액\s*[\d,]+원)?\s*$/);
  if (m) return tx_(CARD_SHINHAN, m[1] === '취소', m[2], m[3], m[4], m[5], m[6], m[7], '', now);
  // 네이버 현대카드: 현대 네이버 승인 박*찬 1,800원 일시불 10/23 07:41 발트페이먼츠 누적3,311,407원
  const h = t.match(/현대\s*네이버\s*(승인취소|취소|승인)\s+\S+\s+([\d,]+)원\s*(?:일시불|\d+\s*개월)?\s*(\d{1,2})\/(\d{1,2})\s+(\d{1,2}):(\d{2})\s+(.+?)(?:\s*누적\s*([\d,]+)원)?\s*$/);
  if (h) return tx_(CARD_HYUNDAI, h[1].indexOf('취소') >= 0, h[3], h[4], h[5], h[6], h[2], h[7], h[8] || '', now);
  return null;
}

function tx_(card, cancel, mo, d, hh, mi, amt, merchant, cum, now) {
  const ym = Utilities.formatDate(now, SMS_TZ, 'yyyy-MM').split('-');
  let y = +ym[0];
  if (+mo > +ym[1]) y -= 1;                       // 1월에 12월 문자를 받은 경우
  return {
    card: card, cancel: !!cancel,
    date: { y: y, m: +mo, d: +d },
    time: ('0' + String(hh)).slice(-2) + ':' + ('0' + String(mi)).slice(-2),
    amount: toNum_(amt), merchant: String(merchant || '').trim(),
    cumulative: cum ? toNum_(cum) : ''
  };
}

function aiParse_(text, now) {
  if (!apiKey_()) return null;
  const today = Utilities.formatDate(now, SMS_TZ, 'yyyy-MM-dd');
  const prompt = [
    '카드 결제 알림 문자를 읽고 JSON 한 개만 답해. 설명은 쓰지 마.',
    '오늘 날짜: ' + today,
    '문자:', '"""' + text + '"""', '',
    '- 국내 원화 결제 승인 또는 취소 문자면:',
    '  {"payment":true,"card":"신한 체크" 또는 "네이버 현대카드" 또는 "기타","cancel":true/false,"month":숫자,"day":숫자,"time":"HH:MM","amount":원화 정수,"merchant":"가맹점 이름","cumulative":누적 이용금액 정수 또는 null}',
    '- amount에는 이번 결제 금액만. 누적 이용금액이나 잔액은 넣지 마.',
    '- 외화 결제, 결제 예정 안내, 승인 거절, 광고처럼 결제가 아니면: {"payment":false,"reason":"짧은 이유"}'
  ].join('\n');
  const r = claude_(prompt, 300);
  if (!r) return null;
  if (r.payment === false) return { ignore: String(r.reason || '결제 문자 아님') };
  if (!r.amount || !r.month || !r.day) return null;
  const hm = String(r.time || '00:00').split(':');
  let card = String(r.card || '');
  if (/현대/.test(text) || card === CARD_HYUNDAI) card = CARD_HYUNDAI;
  else if (/신한/.test(text) || card === CARD_SHINHAN) card = CARD_SHINHAN;
  return tx_(card, r.cancel === true, r.month, r.day, hm[0] || '0', hm[1] || '0', r.amount, r.merchant || '', r.cumulative || '', now);
}

/* ── 카테고리 정하기: 고정지출 → 가맹점 규칙 → AI ── */

function classify_(ss, p) {
  const fixed = matchFixed_(ss, p.merchant);
  if (fixed) return { category: fixed.category, name: fixed.name, note: '고정지출' };
  if (isPg_(p.merchant)) {
    const g = aiClassify_(ss, p);   // 키가 있으면 금액·시간으로 추측만 해 두고, 그래도 확인은 받아요
    return { category: g ? g.category : '', name: p.merchant,
             note: '결제대행 확인 필요' + (g ? '(' + g.category + '?)' : '') + ': 뭘 샀는지 알려 주세요' };
  }
  const rule = matchRule_(ss, p.merchant);
  if (rule) return { category: rule.category, name: rule.name || p.merchant, note: '' };
  const ai = aiClassify_(ss, p);
  if (ai) {
    return { category: ai.category, name: p.merchant,
             note: ai.sure ? 'AI 분류' : ('AI 확인 필요(' + ai.category + '?) ' + ai.reason).trim() };
  }
  return { category: '', name: p.merchant, note: '확인 필요: 처음 보는 가맹점' };
}

function matchFixed_(ss, merchant) {
  const sh = ss.getSheetByName(FIXED_SHEET);
  if (!sh || sh.getLastRow() < 2) return null;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 11).getValues();   // A~K
  const m = norm_(merchant);
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i], kw = norm_(r[10]);
    if (r[0] && kw && r[7] !== true && m.indexOf(kw) >= 0) return { name: String(r[0]), category: r[1] };
  }
  return null;
}

function matchRule_(ss, merchant) {
  const sh = ss.getSheetByName(RULE_SHEET);
  if (!sh || sh.getLastRow() < 2) return null;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  const m = norm_(merchant), up = String(merchant || '').toUpperCase();
  let best = null;
  rows.forEach(r => {
    const k = norm_(r[0]);
    if (!k || !r[1]) return;
    // CU·SRT처럼 짧은 영문 키워드는 단어 경계가 맞을 때만 (SECURE 같은 오탐 방지)
    const hit = /^[A-Z0-9\-]{1,3}$/.test(k)
      ? new RegExp('(^|[^A-Z0-9])' + k.replace(/-/g, '\\-') + '([^A-Z0-9]|$)').test(up)
      : m.indexOf(k) >= 0;
    if (hit && (!best || k.length > best.len)) best = { category: r[1], name: r[2], len: k.length };
  });
  return best;
}

function aiClassify_(ss, p) {
  if (!apiKey_()) return null;
  const set = ss.getSheetByName(SET_SHEET);
  const catRows = set.getRange('L2:O40').getValues().filter(r => r[0] && r[1] === '지출');
  const names = catRows.map(r => r[0]);
  const cats = catRows.map(r => '- ' + r[0] + ' (' + r[2] + '): ' + r[3]).join('\n');
  const prompt = [
    '개인 가계부의 지출 카테고리를 정해 줘. JSON 한 개만 답해: {"category":"카테고리 이름","sure":true 또는 false,"reason":"20자 이내 이유"}',
    '', '카테고리 (이 중 하나만):', cats, '',
    '헷갈리면 이 순서로: ①여행 중에 쓴 돈 → 여행 ②남을 위해 쓴 돈 → 경조사·선물 ③매달 같은 금액 자동결제 → 정기결제 ④나머지는 산 물건 종류대로',
    '', '이 사람이 정해 둔 가맹점 규칙 일부:', ruleExamples_(ss), '',
    '결제 정보:',
    '- 가맹점: ' + p.merchant,
    '- 금액: ' + p.amount + '원' + (p.cancel ? ' (취소)' : ''),
    '- 일시: ' + fmtYmd_(p.date) + ' ' + p.time,
    '- 카드: ' + p.card, '',
    '가맹점 이름만으로 뭘 샀는지 알기 어려우면(결제대행사, 사람 이름 등) 가장 그럴듯한 걸 고르고 sure는 false.'
  ].join('\n');
  const r = claude_(prompt, 150);
  if (!r || names.indexOf(r.category) < 0) return null;
  return { category: r.category, sure: r.sure === true, reason: String(r.reason || '').slice(0, 40) };
}

function ruleExamples_(ss) {
  const sh = ss.getSheetByName(RULE_SHEET);
  if (!sh || sh.getLastRow() < 2) return '(없음)';
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().filter(r => r[0] && r[1]);
  // 직접·검토로 추가된 규칙이 아래쪽에 쌓이니 최근 것 위주로 40개
  return rows.slice(-40).map(r => r[0] + ' → ' + r[1]).join('\n');
}

/* ── Claude API ── */

function apiKey_() {
  return PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY') || '';
}

function claude_(prompt, maxTokens) {
  const key = apiKey_();
  if (!key) return null;
  try {
    const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: CLAUDE_MODEL, max_tokens: maxTokens || 300,
                                messages: [{ role: 'user', content: prompt }] })
    });
    if (res.getResponseCode() !== 200) {
      console.warn('Claude API ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 300));
      return null;
    }
    const data = JSON.parse(res.getContentText());
    const txt = (data.content || []).filter(c => c.type === 'text').map(c => c.text).join('');
    const j = txt.match(/\{[\s\S]*\}/);
    return j ? JSON.parse(j[0]) : null;
  } catch (err) {
    console.warn('Claude API 오류: ' + err);
    return null;
  }
}

/* ── 가계부에 한 줄 쓰기 ── */

function appendLedger_(ss, row) {
  const sh = ss.getSheetByName(LEDGER_SHEET);
  const n = sh.getMaxRows() - 1;
  const col = sh.getRange(2, 1, n, 1).getValues();
  let last = 0;
  for (let i = col.length - 1; i >= 0; i--) {
    if (col[i][0] !== '' && col[i][0] !== null) { last = i + 1; break; }
  }
  const target = last + 2;
  if (target > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 200);
  sh.getRange(target, 1, 1, 7).setValues([row]);
  return target;
}

/* ── 메뉴: 설정 · 테스트 ── */

function smsSetup() {
  const ss = book_();
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('SMS_SECRET');
  if (!secret) {
    secret = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
    props.setProperty('SMS_SECRET', secret);
  }
  const hasKey = !!apiKey_();
  const a = ui.prompt('1/2 · Claude API 키',
    hasKey ? '키가 이미 저장돼 있어요. 바꾸려면 새 키를 붙여넣고, 그대로 두려면 비운 채 확인을 누르세요.'
           : 'Claude API 키(sk-ant-로 시작)를 붙여넣어 주세요.\n아직 없으면 비운 채 확인을 누르세요. 키가 없어도 신한 체크 해석과 가맹점 규칙 분류는 돼요.',
    ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return;
  if (a.getResponseText().trim()) props.setProperty('ANTHROPIC_API_KEY', a.getResponseText().trim());

  const set = ss.getSheetByName(SET_SHEET);
  const curUrl = String(set.getRange('B66').getValue() || '');
  const b = ui.prompt('2/2 · 웹 앱 URL',
    (curUrl ? '저장된 주소: ' + curUrl + '\n바꾸려면 새 주소를, 그대로 두려면 비운 채 확인.\n' : '')
    + '[배포] > [새 배포](또는 배포 관리)에서 복사한 웹 앱 URL(…/exec)을 붙여넣어 주세요.\n아직 배포 전이면 비운 채 확인을 누르고, 배포 후 이 메뉴를 다시 열면 돼요.',
    ui.ButtonSet.OK_CANCEL);
  if (b.getSelectedButton() !== ui.Button.OK) return;
  const url = b.getResponseText().trim() || curUrl;

  set.getRange('A65:C67').setValues([
    ['문자 자동입력 (단축어에 넣을 값)', '', ''],
    ['웹 앱 URL', url, '아이폰 단축어 \'URL 내용 가져오기\'의 주소'],
    ['비밀키', secret, '단축어 JSON의 key 값. 남에게 보여 주지 마세요']
  ]);
  set.getRange('A65:C65').setFontWeight('bold').setBackground('#EAF0F7');
  ui.alert('설정 완료',
    'AI 분류: ' + (apiKey_() ? '켜짐' : '꺼짐 (API 키 없음)') + '\n'
    + '웹 앱 URL: ' + (url || '아직 없음') + '\n'
    + '비밀키: ' + secret + '\n\n'
    + '두 값은 설정 탭 66~67행에도 적어 두었어요. 아이폰의 구글 시트 앱에서 복사해 단축어에 넣으면 돼요.',
    ui.ButtonSet.OK);
}

function smsTest() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('문자 자동입력 테스트',
    '결제 문자 내용을 그대로 붙여넣어 주세요.\n가계부에는 넣지 않고, 어떻게 읽히는지만 보여 드려요.',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const res = handleSms_(r.getResponseText(), new Date(), true);
  const e = res.entry, p = res.parsed;
  ui.alert('테스트 결과', e
    ? ['처리: ' + res.result,
       '날짜: ' + fmtYmd_(p.date) + ' ' + p.time,
       '카테고리: ' + (e[1] || '(빈칸)'),
       '내용: ' + e[2],
       '금액: ' + e[3],
       '결제수단: ' + e[4],
       '메모: ' + e[6]].join('\n')
    : '처리: ' + res.result, ui.ButtonSet.OK);
}

function toNum_(s) {
  return Number(String(s).replace(/[^\d.]/g, '')) || 0;
}

function norm_(s) {
  return String(s || '').replace(/\s+/g, '').toUpperCase();
}


/* ── 직접 고른 카테고리 학습 ── */

// 가계부에서 '확인 필요' 줄의 카테고리를 고르면, 그 가맹점을 규칙에 저장하고 노란 표시를 지워요.
// (이미 규칙으로 분류된 줄을 한 번 바꾼 건 그때만의 예외일 수 있어서 학습하지 않아요.
//  규칙 자체를 바꾸려면 가맹점 규칙 탭에서 고치면 돼요.)
const FIXED_FLAG_COL = 8;   // 가계부 H열 '고정 등록' 체크칸
const FIXED_ROWS = 50;      // 고정지출 표 2~51행

function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== LEDGER_SHEET) return;
    const c1 = e.range.getColumn(), c2 = e.range.getLastColumn();
    const r1 = Math.max(e.range.getRow(), 2), r2 = e.range.getLastRow();
    if (r2 < r1) return;
    if (c1 <= 2 && c2 >= 2) {                                             // B열(카테고리)을 고친 경우
      const vals = sh.getRange(r1, 2, r2 - r1 + 1, 6).getValues();       // B~G
      vals.forEach((v, i) => {
        if (v[0] && v[1]) confirmRow_(e.source, sh, r1 + i, String(v[1]), v[0], String(v[5] || ''));
      });
    }
    if (c1 === 3 && c2 === 3 && r1 === r2 && e.oldValue !== undefined) {    // C열(내용) 한 칸을 고친 경우
      if (sh.getRange(r1, FIXED_FLAG_COL).getValue() === true) syncFixedName_(e.source, e.oldValue, e.value, '');
    }
    if (c1 <= FIXED_FLAG_COL && c2 >= FIXED_FLAG_COL) {                   // H열(고정 등록)을 고친 경우
      const vals = sh.getRange(r1, 1, r2 - r1 + 1, FIXED_FLAG_COL).getValues();   // A~H
      vals.forEach((v, i) => toggleFixed_(e.source, sh, r1 + i, v));
    }
  } catch (err) {
    console.warn('onEdit 처리 실패: ' + err);
  }
}


/* ── 가계부에서 체크한 줄 → 고정지출 표 ── */

// 고정지출 표에서 이름이 oldName인 '살아 있는' 줄의 이름(과 카테고리)을 바꿔요
function syncFixedName_(ss, oldName, newName, newCat) {
  oldName = String(oldName || '').trim(); newName = String(newName || '').trim();
  if (!oldName || !newName) return '';
  const fx = ss.getSheetByName(FIXED_SHEET);
  if (!fx) return '';
  const rows = fx.getRange(2, 1, FIXED_ROWS, 11).getValues();
  const cur = appCurYm_(ss);
  const alive = x => !(appYm_(ss, x[6]) && appYm_(ss, x[6]) < cur);
  const i = rows.findIndex(x => String(x[0]).trim() === oldName && alive(x));
  if (i < 0) return '';
  if (newName !== oldName && rows.some((x, j) => j !== i && String(x[0]).trim() === newName && alive(x))) return '';
  const t = i + 2;
  if (newName !== oldName) {
    fx.getRange(t, 1).setValue(newName);
    if (String(rows[i][10]).trim() === oldName) fx.getRange(t, 11).setValue(newName);
  }
  if (newCat) fx.getRange(t, 2).setValue(newCat);
  return newName;
}

// 체크: 그 줄의 이름·카테고리·금액·결제수단·날짜(일)로 고정지출에 한 줄 추가. 시작월은 그 줄의 달이에요.
//       (그 달은 가계부에 실제 줄이 있으니 겹치지 않고, 다음 달부터 자동으로 잡혀요)
// 체크 해제: 이 체크로 들어간 고정지출 줄만 지워요. 손으로 넣은 고정지출 줄은 건드리지 않아요.
function toggleFixed_(ss, sh, row, v) {
  const fx = ss.getSheetByName(FIXED_SHEET);
  if (!fx) return '';
  const on = v[FIXED_FLAG_COL - 1] === true;
  const name = String(v[2] || '').trim();
  const mark = '가계부 ' + row + '행에서 등록';
  const rows = fx.getRange(2, 1, FIXED_ROWS, 11).getValues();             // A~K
  const same = rows.findIndex(x => String(x[0]).trim() === name && name !== '');

  if (!on) {
    if (same >= 0 && /^가계부 \d+행에서 등록/.test(String(rows[same][9]))) {
      const t = same + 2;
      fx.getRange(t, 1, 1, 8).setValues([['', '', '', '', '', '', '', false]]);
      fx.getRange(t, 10, 1, 2).setValues([['', '']]);
      return toast_(ss, '고정지출에서 뺐어요: ' + name);
    }
    return '';
  }

  const ymd = ymdOf_(ss, v[0]);
  const amt = Number(v[3]);
  if (!ymd || !name || !v[1] || !(amt > 0)) {
    sh.getRange(row, FIXED_FLAG_COL).setValue(false);
    return toast_(ss, '날짜·카테고리·내용·금액이 다 있는 줄만 고정지출로 넣을 수 있어요');
  }
  if (same >= 0) {
    return toast_(ss, '이미 고정지출에 있어요: ' + name + ' (금액이 바뀌었으면 고정지출 탭에서 고쳐 주세요)');
  }
  const empty = rows.findIndex(x => x[0] === '' || x[0] === null);
  if (empty < 0) {
    sh.getRange(row, FIXED_FLAG_COL).setValue(false);
    return toast_(ss, '고정지출 표가 꽉 찼어요 (' + FIXED_ROWS + '개)');
  }
  const pay = String(v[4] || '');
  // 문자로 들어온 줄이면 문자 속 가맹점명을, 손으로 적은 카드 결제면 이름을 '문자 가맹점명'으로 둬요
  const kw = smsMerchant_(ss, row, amt) || (pay && pay !== '현금' ? name : '');
  const today = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'M/d');
  const t = empty + 2;
  fx.getRange(t, 1, 1, 8).setValues([[name, v[1], amt, pay, ymd.d, serial_({ y: ymd.y, m: ymd.m, d: 1 }), '', false]]);
  fx.getRange(t, 10, 1, 2).setValues([[mark + ' (' + today + ')', kw]]);
  return toast_(ss, '고정지출에 넣었어요: ' + name + ' ' + won_(amt) + ' · 다음 달부터 매월 ' + ymd.d + '일에 자동으로 잡혀요');
}

// 문자 로그에서 그 가계부 줄로 들어간 문자의 가맹점명 (줄 번호와 금액이 둘 다 맞을 때만)
function smsMerchant_(ss, row, amt) {
  const log = ss.getSheetByName(LOG_SHEET);
  if (!log || log.getLastRow() < 2) return '';
  const v = log.getRange(2, 1, log.getLastRow() - 1, 11).getValues();
  for (let i = v.length - 1; i >= 0; i--) {
    if (Number(v[i][10]) === row && Number(v[i][6]) === amt && String(v[i][9]).indexOf('입력됨') === 0) return String(v[i][7] || '');
  }
  return '';
}

// 시트 날짜 칸 값 → {y, m, d}
function ymdOf_(ss, v) {
  if (isDate_(v)) {
    const s = Utilities.formatDate(v, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd').split('-');
    return { y: +s[0], m: +s[1], d: +s[2] };
  }
  if (typeof v === 'number' && v > 0) {
    const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(v) * 86400000);
    return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate() };
  }
  return null;
}

function toast_(ss, msg) {
  try { ss.toast(msg, '가계부', 6); } catch (e) { /* 폰 앱·웹 앱에서는 안 떠요 */ }
  return msg;
}

// 확인 필요였던 문자 자동 줄을 확정: 규칙 학습(결제대행 제외) + 노란 표시 지우기
function confirmRow_(ss, sh, row, name, category, memo) {
  if (memo.indexOf('문자 자동') !== 0 || memo.indexOf('확인 필요') < 0) return;
  if (!isPg_(name)) learnRule_(ss, name, category);
  sh.getRange(row, 7).setValue(memo.replace(/ · (AI |결제대행 )?확인 필요.*$/, ' · 직접 분류'));
}

function learnRule_(ss, merchant, category) {
  const sh = ss.getSheetByName(RULE_SHEET);
  const key = String(merchant || '').trim();
  if (!sh || !key) return;
  const n = sh.getLastRow();
  const rows = n >= 2 ? sh.getRange(2, 1, n - 1, 2).getValues() : [];
  const idx = rows.findIndex(r => norm_(r[0]) === norm_(key));
  const today = Utilities.formatDate(new Date(), SMS_TZ, 'yyyy-MM-dd');
  if (idx >= 0) {
    sh.getRange(idx + 2, 2).setValue(category);
    sh.getRange(idx + 2, 4, 1, 2).setValues([['직접 분류', today]]);
  } else {
    sh.appendRow([key, category, '', '직접 분류', today]);
  }
}


/* ── 가계부 분류 단축어: 확인할 결제를 하나씩 받아서 카테고리 고르기 ── */

// 날짜는 있는데 카테고리가 비었거나 '확인 필요'인 줄 (오래된 것부터)
function pendingRows_(sh) {
  const n = sh.getLastRow();
  if (n < 2) return [];
  const v = sh.getRange(2, 1, n - 1, 7).getValues();
  const out = [];
  v.forEach((r, i) => {
    if (r[0] !== '' && r[0] !== null && (r[1] === '' || String(r[6] || '').indexOf('확인 필요') >= 0)) out.push({ row: i + 2, r: r });
  });
  return out;
}

function nextPending_() {
  const ss = book_();
  const sh = ss.getSheetByName(LEDGER_SHEET);
  const list = pendingRows_(sh);
  if (!list.length) return { ok: true, done: true, count: 0, message: '확인할 결제가 없어요' };
  const it = list[0], r = it.r;
  const guess = (String(r[6] || '').match(/\(([^()?]+)\?\)/) || [])[1] || '';
  return {
    ok: true, done: false, count: list.length, row: it.row,
    label: dayLabel_(r[0]) + ' ' + r[2] + ' ' + won_(r[3]) + (guess ? ' (추측: ' + guess + ')' : '')
           + (list.length > 1 ? ' · 남은 ' + list.length + '건' : ''),
    pg: isPg_(r[2]),
    categories: categoryNames_(ss)
  };
}

function classifyRow_(row, category) {
  const ss = book_();
  const sh = ss.getSheetByName(LEDGER_SHEET);
  const names = categoryNames_(ss);
  row = Number(row);
  if (!row) {                                   // 줄 번호 없이 오면 지금 맨 앞의 확인할 건
    const first = pendingRows_(sh)[0];
    if (!first) return { ok: true, left: 0, message: '확인할 결제가 없어요' };
    row = first.row;
  }
  if (row < 2 || names.indexOf(category) < 0) return { ok: false, message: '분류하지 못했어요 (카테고리 확인)' };
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const r = sh.getRange(row, 1, 1, 7).getValues()[0];
    if (r[0] === '' || r[0] === null) return { ok: false, message: '그 줄이 비어 있어요' };
    sh.getRange(row, 2).setValue(category);
    confirmRow_(ss, sh, row, String(r[2]), category, String(r[6] || ''));
    const left = pendingRows_(sh).length;
    return { ok: true, left: left,
             message: r[2] + ' ' + won_(r[3]) + ' → ' + category + (left ? ' · 남은 확인 ' + left + '건' : ' · 다 정리됐어요') };
  } finally {
    lock.releaseLock();
  }
}

function categoryNames_(ss) {
  return ss.getSheetByName(SET_SHEET).getRange('L2:M40').getValues().filter(r => r[0] && r[1] === '지출').map(r => r[0]);
}

function won_(n) {
  const v = Math.round(Number(n) || 0);
  return (v < 0 ? '-' : '') + String(Math.abs(v)).replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '원';
}

function dayLabel_(v) {
  if (isDate_(v)) return Utilities.formatDate(v, SMS_TZ, 'M/d');
  if (typeof v === 'number') {
    const d = new Date(Date.UTC(1899, 11, 30) + v * 86400000);
    return (d.getUTCMonth() + 1) + '/' + d.getUTCDate();
  }
  return String(v);
}


/* ═════════════════════════ 가계부 앱(웹앱) API ═════════════════════════
 *  깃허브에 올린 앱이 이 주소로 {key, action:'app_…', …}를 POST해요.
 *  계산은 지금처럼 시트가 하고, 앱은 시트 값을 받아 보여 주고 입력만 보내요.
 */

const DASH_SHEET = '대시보드';
const SNAP_SHEET_APP = '스냅샷';
const ALL_SHEET = '가계부 전체';
const APP_SHEETS = { hold: '보유종목', cash: '현금·기타', set: '설정' };
// 설정 탭에서 앱에 안 보여 줄 줄 (내부용·비밀값)
const APP_SET_HIDE = ['웹 앱 URL', '비밀키', '문자 자동입력 (단축어에 넣을 값)', '가계부 월간: 보는 달', '가계부 월간: 달력 시작일',
  '현대카드 이용기간 마감일', '현대카드 마감 월'];
const INPUT_BG = '#fff8e0';   // 설정 탭의 직접 입력 칸 배경

function appRoute_(b) {
  const a = String(b.action);
  try {
    if (a === 'app_boot') return appBoot_(b);
    if (a === 'app_ledger') return { ok: true, ledger: appLedger_(book_(), b.month) };
    if (a === 'app_dash') return appDash_();
    if (a === 'app_sheet') return appSheet_(b.sheet);
    if (a === 'app_add') return appWrite_(() => appAdd_(b));
    if (a === 'app_update') return appWrite_(() => appUpdate_(b));
    if (a === 'app_delete') return appWrite_(() => appDelete_(b));
    if (a === 'app_fixed') return appWrite_(() => appFixed_(b));
    if (a === 'app_budget') return appWrite_(() => appBudgetSet_(b));
    if (a === 'app_fixed_add') return appWrite_(() => appFixedAdd_(b));
    if (a === 'app_fixed_save') return appWrite_(() => appFixedSave_(b));
    if (a === 'app_fixed_end') return appWrite_(() => appFixedEnd_(b, false));
    if (a === 'app_fixed_delete') return appWrite_(() => appFixedEnd_(b, true));
    if (a === 'app_set') return appWrite_(() => appSet_(b));
    if (a === 'app_row_add') return appWrite_(() => appRowAdd_(b));
    if (a === 'app_row_clear') return appWrite_(() => appRowClear_(b));
    return { ok: false, error: '알 수 없는 요청이에요: ' + a };
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  }
}

function appWrite_(fn) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ── 공통 ── */

function appTz_(ss) { return ss.getSpreadsheetTimeZone() || SMS_TZ; }

function appYmd_(ss, v) {
  if (isDate_(v)) return Utilities.formatDate(v, appTz_(ss), 'yyyy-MM-dd');
  const p = ymdOf_(ss, v);
  return p ? fmtYmd_(p) : '';
}

// 'YYYY-MM-DD' → 시트 날짜 일련번호
function appSerial_(s) {
  const m = String(s || '').match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!m) throw new Error('날짜 형식이 맞지 않아요: ' + s);
  return serial_({ y: +m[1], m: +m[2], d: +m[3] });
}

// 설정 탭 L:O 카테고리 표 (2행부터, 빈 줄 전까지 넉넉히 40행까지 읽어요)
function appCats_(ss) {
  return ss.getSheetByName(SET_SHEET).getRange('L2:O40').getValues()
    .filter(r => r[0]).map(r => ({ name: String(r[0]), kind: String(r[1]), fixed: String(r[2]), desc: String(r[3] || '') }));
}

function appPays_(ss) {
  return ss.getSheetByName(SET_SHEET).getRange('Q2:R4').getValues()
    .filter(r => r[0]).map(r => ({ name: String(r[0]), type: String(r[1]) }));
}

function appNamed_(ss, name) {
  const r = ss.getRangeByName(name);
  return r ? r.getValue() : '';
}

/* ── 처음 열 때 한 번에 ── */

function appBoot_(b) {
  const ss = book_();
  const tz = appTz_(ss);
  const sh = ss.getSheetByName(LEDGER_SHEET);
  const pend = pendingRows_(sh).map(it => {
    const r = it.r;
    const guess = (String(r[6] || '').match(/\(([^()?]+)\?\)/) || [])[1] || '';
    return { r: it.row, d: appYmd_(ss, r[0]), n: String(r[2] || ''), a: Number(r[3]) || 0, p: String(r[4] || ''),
             m: String(r[6] || ''), guess: guess, pg: isPg_(r[2]) };
  });
  const ledger = appLedger_(ss, b.month);
  const cur = appCurYm_(ss);
  const curRows = ledger.month === cur ? ledger.rows : appLedger_(ss, cur).rows;   // 예산은 늘 이번 달 기준
  const fixed = appFixedList_(ss);
  return {
    ok: true,
    today: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'),
    cats: appCats_(ss),
    pays: appPays_(ss),
    card: { payday: appNamed_(ss, 'CARD_PAYDAY'), min: appNamed_(ss, 'CARD_MIN'), period: appNamed_(ss, 'CARD_PERIOD') },
    pending: pend,
    ledger: ledger,
    fixed: fixed,
    budget: appBudget_(ss, cur, curRows, fixed)
  };
}

/* ── 가계부: 한 달 치 ──
 * '가계부 전체' 탭(가계부 + 고정지출 자동 줄)을 읽어요. 앞쪽 줄은 가계부 탭의 줄과 순서가 같아서
 * 가계부 탭 줄 번호(r)를 붙여 주고, 고정지출 자동 줄은 r = 0 (앱에서 고칠 수 없음)이에요.
 */
function appLedger_(ss, month) {
  const tz = appTz_(ss);
  const now = new Date();
  const ym = /^\d{4}-\d{2}$/.test(String(month || '')) ? String(month) : Utilities.formatDate(now, tz, 'yyyy-MM');
  const y = +ym.slice(0, 4), mo = +ym.slice(5, 7);
  const prev = mo === 1 ? (y - 1) + '-12' : y + '-' + ('0' + (mo - 1)).slice(-2);

  const led = ss.getSheetByName(LEDGER_SHEET);
  const n = Math.max(led.getLastRow() - 1, 0);
  const lv = n ? led.getRange(2, 1, n, 8).getValues() : [];
  const ledRows = [];
  lv.forEach((r, i) => { if (r[0] !== '' && r[0] !== null) ledRows.push({ row: i + 2, fixedFlag: r[7] === true }); });

  const all = ss.getSheetByName(ALL_SHEET);
  const an = Math.max(all.getLastRow() - 1, 0);
  const av = an ? all.getRange(2, 1, an, 13).getValues() : [];

  const rows = [], prevCat = {};
  av.forEach((r, i) => {
    if (r[0] === '' || r[0] === null) return;
    const d = appYmd_(ss, r[0]);
    if (!d) return;
    const mk = d.slice(0, 7);
    const isFx = String(r[6]) === '고정지출 자동';
    const src = (!isFx && i < ledRows.length) ? ledRows[i] : null;
    if (mk === prev && String(r[7]) === '지출') prevCat[r[1]] = (prevCat[r[1]] || 0) + (Number(r[3]) || 0);
    if (mk !== ym) return;
    rows.push({
      r: src ? src.row : 0, d: d, c: String(r[1] || ''), n: String(r[2] || ''), a: Number(r[3]) || 0,
      p: String(r[4] || ''), rc: r[5] === true, m: String(r[6] || ''), g: String(r[7] || ''), fv: String(r[8] || ''),
      pt: String(r[9] || ''), due: r[10] ? appYmd_(ss, r[10]) : '', fx: isFx, ff: src ? src.fixedFlag : false
    });
  });
  rows.sort((p, q) => (p.d < q.d ? 1 : p.d > q.d ? -1 : (q.r - p.r)));
  return { month: ym, rows: rows, prevCat: prevCat };
}

/* ── 가계부: 쓰기 ── */

function appCheckRow_(sh, row, check) {
  row = Number(row);
  if (!(row >= 2)) throw new Error('줄 번호가 없어요');
  const v = sh.getRange(row, 1, 1, 8).getValues()[0];
  if (v[0] === '' || v[0] === null) throw new Error('그 줄이 비어 있어요. 새로고침해 주세요');
  if (check) {
    const d = appYmd_(book_(), v[0]);
    if ((check.d && check.d !== d) || (check.a !== undefined && Number(check.a) !== Number(v[3]))) {
      throw new Error('시트 내용이 바뀌었어요. 새로고침한 뒤 다시 해 주세요');
    }
  }
  return v;
}

function appAdd_(b) {
  const ss = book_();
  const amt = Number(b.a);
  if (!isFinite(amt) || amt === 0) throw new Error('금액을 넣어 주세요');
  const cats = appCats_(ss).map(c => c.name), pays = appPays_(ss).map(p => p.name);
  if (b.c && cats.indexOf(b.c) < 0) throw new Error('없는 카테고리예요: ' + b.c);
  if (b.p && pays.indexOf(b.p) < 0) throw new Error('없는 결제수단이에요: ' + b.p);
  const row = [appSerial_(b.d), b.c || '', String(b.n || '').trim(), amt, b.p || '',
               b.rc ? true : '', String(b.m || '').trim() || '앱 입력'];
  const at = appendLedger_(ss, row);
  return { ok: true, row: at, message: (b.c || '미분류') + ' ' + won_(amt) + ' 저장했어요' };
}

function appUpdate_(b) {
  const ss = book_();
  const sh = ss.getSheetByName(LEDGER_SHEET);
  const v = appCheckRow_(sh, b.r, b.check);
  const f = b.f || {};
  const cats = appCats_(ss).map(c => c.name), pays = appPays_(ss).map(p => p.name);
  if (f.c !== undefined && f.c !== '' && cats.indexOf(f.c) < 0) throw new Error('없는 카테고리예요: ' + f.c);
  if (f.p !== undefined && f.p !== '' && pays.indexOf(f.p) < 0) throw new Error('없는 결제수단이에요: ' + f.p);
  const oldCat = v[1], memo = String(v[6] || '');
  const nv = [
    f.d !== undefined ? appSerial_(f.d) : v[0],
    f.c !== undefined ? f.c : v[1],
    f.n !== undefined ? String(f.n) : v[2],
    f.a !== undefined ? Number(f.a) : v[3],
    f.p !== undefined ? f.p : v[4],
    f.rc !== undefined ? (f.rc ? true : '') : v[5],
    f.m !== undefined ? String(f.m) : v[6]
  ];
  if (!isFinite(Number(nv[3])) || Number(nv[3]) === 0) throw new Error('금액이 맞지 않아요');
  sh.getRange(Number(b.r), 1, 1, 7).setValues([nv]);
  // 고정 등록(H열)된 줄의 이름·카테고리를 바꾸면 고정지출 표도 따라가요 (안 그러면 자동 줄이 한 번 더 생겨요)
  if (v[FIXED_FLAG_COL - 1] === true && (String(nv[2]).trim() !== String(v[2]).trim() || nv[1] !== v[1])) {
    syncFixedName_(ss, v[2], nv[2], nv[1] !== v[1] ? nv[1] : '');
  }
  // 문자로 들어온 '확인 필요' 줄의 카테고리를 고르면 시트에서 고를 때처럼 규칙을 배워요
  if (f.c && f.c !== oldCat && f.m === undefined) confirmRow_(ss, sh, Number(b.r), String(nv[2]), f.c, memo);
  return { ok: true, message: '고쳤어요' };
}

function appDelete_(b) {
  const ss = book_();
  const sh = ss.getSheetByName(LEDGER_SHEET);
  const v = appCheckRow_(sh, b.r, b.check);
  sh.deleteRow(Number(b.r));
  return { ok: true, message: (v[2] || '') + ' ' + won_(v[3]) + ' 지웠어요' };
}

function appFixed_(b) {
  const ss = book_();
  const sh = ss.getSheetByName(LEDGER_SHEET);
  const v = appCheckRow_(sh, b.r, b.check);
  const on = b.on === true || b.on === 'true';
  sh.getRange(Number(b.r), FIXED_FLAG_COL).setValue(on);
  v[FIXED_FLAG_COL - 1] = on;
  const msg = toggleFixed_(ss, sh, Number(b.r), v);
  const now = sh.getRange(Number(b.r), FIXED_FLAG_COL).getValue() === true;
  return { ok: true, on: now, message: msg || (on ? '고정지출에 넣었어요' : '고정 등록을 풀었어요') };
}


/* ── 한 달 예산 ──
 * 예산은 설정 탭의 LIVING_COST(한 달 생활비) 한 칸이에요. 고정비(주거·정기결제, 고정지출 자동 줄)는 빼고 생활비만 봐요.
 * 남은 생활비 = 예산 − 이번 달 생활비로 쓴 돈. 비상금 목표는 대시보드에서 (생활비 + 고정비) × 개월 수 */

function appBudget_(ss, ym, rows, fixed) {
  let living = 0, fixedSpent = 0;
  rows.forEach(r => {
    if (r.g !== '지출') return;
    if (r.fx || r.fv === '고정') fixedSpent += r.a; else living += r.a;
  });
  return { month: ym, amount: Number(appNamed_(ss, 'LIVING_COST')) || 0, living: living, fixedSpent: fixedSpent };
}

function appBudgetSet_(b) {
  const ss = book_();
  const a = Math.round(Number(b.amount));
  if (!isFinite(a) || a < 0 || a > 1e10) throw new Error('예산 금액을 확인해 주세요');
  const rg = ss.getRangeByName('LIVING_COST');
  if (!rg) throw new Error('설정 탭에서 예산 칸(LIVING_COST)을 찾지 못했어요');
  rg.setValue(a);
  return { ok: true, amount: a, message: a ? '한 달 생활비 예산을 ' + won_(a) + '으로 정했어요' : '예산을 비웠어요' };
}


/* ── 고정비 (고정지출 표) ──
 * 한 번 넣으면 가계부 전체 탭 수식이 시작월부터 매달 결제일에 자동으로 채워요.
 * 해지: 종료월만 적어서 지난 기록은 남겨요. 금액·결제일 변경: 지난달까지는 옛 값, 이번 달부터 새 줄로 이어 가요. */

function appCurYm_(ss) { return Utilities.formatDate(new Date(), appTz_(ss), 'yyyy-MM'); }
function appYmShift_(ym, k) {
  const d = new Date(Date.UTC(+ym.slice(0, 4), +ym.slice(5, 7) - 1 + k, 1));
  return d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2);
}
function appYmSerial_(ym) { return serial_({ y: +ym.slice(0, 4), m: +ym.slice(5, 7), d: 1 }); }
// 시작월·종료월 칸 값 → 'yyyy-MM' (비었으면 '')
function appYm_(ss, v) {
  if (v === '' || v === null || v === undefined) return '';
  if (typeof v === 'string') {
    const m = v.match(/^(\d{4})[-.\/]\s*(\d{1,2})/);
    return m ? m[1] + '-' + ('0' + m[2]).slice(-2) : '';
  }
  const t = ymdOf_(ss, v);
  return t ? t.y + '-' + ('0' + t.m).slice(-2) : '';
}

function appFixedList_(ss) {
  const sh = ss.getSheetByName(FIXED_SHEET);
  if (!sh) return [];
  const v = sh.getRange(2, 1, FIXED_ROWS, 8).getValues();
  const st = sh.getRange(2, 9, FIXED_ROWS, 1).getDisplayValues();     // I열 '이번 달' (예정 11/25, ✓ 11/1 반영 …)
  const kind = {};
  appCats_(ss).forEach(c => { kind[c.name] = c.kind; });
  const cur = appCurYm_(ss);
  const out = [];
  v.forEach((r, i) => {
    const n = String(r[0] || '').trim();
    if (!n) return;
    const e = appYm_(ss, r[6]);
    if (e && e < cur) return;                                             // 지난달까지 끝난 줄은 숨겨요
    out.push({ i: i + 2, n: n, c: String(r[1] || ''), a: Number(r[2]) || 0, p: String(r[3] || ''), d: Number(r[4]) || 0,
               s: appYm_(ss, r[5]), e: e, stop: r[7] === true, k: kind[r[1]] || '', st: String(st[i][0] || '') });
  });
  out.sort((p, q) => (p.k === q.k ? p.d - q.d : p.k === '지출' ? -1 : q.k === '지출' ? 1 : 0));
  return out;
}

function appFixedCheck_(ss, b) {
  const n = String(b.n || '').trim();
  if (!n) throw new Error('이름을 넣어 주세요 (예: 넷플릭스, 관리비)');
  if (n.length > 40) throw new Error('이름이 너무 길어요');
  const cat = appCats_(ss).find(c => c.name === b.c);
  if (!cat || (cat.kind !== '지출' && cat.kind !== '수입')) throw new Error('카테고리를 골라 주세요');
  const a = Math.round(Number(b.a));
  if (!(a > 0)) throw new Error('금액을 넣어 주세요');
  const d = Math.round(Number(b.d));
  if (!(d >= 1 && d <= 31)) throw new Error('결제일은 1~31일 중에서 골라 주세요');
  const p = String(b.p || '');
  if (p && appPays_(ss).map(x => x.name).indexOf(p) < 0) throw new Error('없는 결제수단이에요: ' + p);
  return { n: n, c: cat.name, a: a, p: p, d: d, k: cat.kind };
}

// 고정지출 표에서 이름이 같은 '살아 있는' 줄 (지난달까지 끝난 줄은 빼요)
function appFixedDup_(ss, v, n, skipRow) {
  const cur = appCurYm_(ss);
  return v.findIndex((r, i) => i + 2 !== skipRow && String(r[0] || '').trim() === n && !(appYm_(ss, r[6]) && appYm_(ss, r[6]) < cur));
}

function appFixedWrite_(ss, sh, v, f, ym, how) {
  const empty = v.findIndex(r => r[0] === '' || r[0] === null);
  if (empty < 0) throw new Error('고정지출 표가 꽉 찼어요 (' + FIXED_ROWS + '줄). 시트에서 끝난 줄을 정리해 주세요');
  const t = empty + 2;
  const today = Utilities.formatDate(new Date(), appTz_(ss), 'M/d');
  sh.getRange(t, 1, 1, 8).setValues([[f.n, f.c, f.a, f.p, f.d, appYmSerial_(ym), '', f.stop === true]]);
  sh.getRange(t, 6, 1, 2).setNumberFormat('yyyy-mm');
  sh.getRange(t, 10, 1, 2).setValues([[how + ' (' + today + ')', '']]);
  v[empty] = [f.n, f.c, f.a, f.p, f.d, appYmSerial_(ym), '', false];
  return t;
}

function appFixedRow_(sh, i, check) {
  const t = Number(i);
  if (!(t >= 2 && t <= FIXED_ROWS + 1)) throw new Error('고정비 줄 번호가 맞지 않아요');
  const r = sh.getRange(t, 1, 1, 8).getValues()[0];
  if (!String(r[0] || '').trim() || (check && String(check.n || '') !== String(r[0]).trim())) {
    throw new Error('고정비 표가 그새 바뀌었어요. 새로고침한 뒤 다시 해 주세요');
  }
  return r;
}

// 이번 달 결제일이 지났는지 (월말보다 큰 결제일은 말일로)
function appFixedPaidThisMonth_(ss, day) {
  const tz = appTz_(ss), now = new Date();
  const y = +Utilities.formatDate(now, tz, 'yyyy'), m = +Utilities.formatDate(now, tz, 'MM'), td = +Utilities.formatDate(now, tz, 'dd');
  const last = new Date(Date.UTC(y, m, 0)).getUTCDate();
  return td >= Math.min(Number(day) || 1, last);
}

function appFixedAdd_(b) {
  const ss = book_();
  const sh = ss.getSheetByName(FIXED_SHEET);
  const f = appFixedCheck_(ss, b);
  const v = sh.getRange(2, 1, FIXED_ROWS, 7).getValues();
  if (appFixedDup_(ss, v, f.n, 0) >= 0) throw new Error('같은 이름이 이미 고정비에 있어요: ' + f.n);
  const cur = appCurYm_(ss);
  const t = appFixedWrite_(ss, sh, v, f, cur, '앱에서 등록');
  const paid = appFixedPaidThisMonth_(ss, f.d);
  const what = f.k === '수입' ? '들어와요' : '나가요';
  return { ok: true, i: t, message: f.n + ' 등록 · 매달 ' + f.d + '일에 자동으로 ' + what + (paid ? ' (이번 달 포함)' : '') };
}

function appFixedSave_(b) {
  const ss = book_();
  const sh = ss.getSheetByName(FIXED_SHEET);
  const t = Number(b.i);
  const r = appFixedRow_(sh, t, b.check);
  const old = { n: String(r[0]).trim(), c: String(r[1] || ''), a: Number(r[2]) || 0, p: String(r[3] || ''), d: Number(r[4]) || 0 };
  const f = appFixedCheck_(ss, Object.assign({}, old, b.f || {}));
  f.stop = r[7] === true;
  const v = sh.getRange(2, 1, FIXED_ROWS, 7).getValues();
  if (f.n !== old.n && appFixedDup_(ss, v, f.n, t) >= 0) throw new Error('같은 이름이 이미 고정비에 있어요: ' + f.n);
  const cur = appCurYm_(ss), start = appYm_(ss, r[5]), end = appYm_(ss, r[6]);
  if ((f.a !== old.a || f.d !== old.d) && start && start < cur && !end) {
    // 지난달까지는 옛 금액·결제일 그대로 남기고, 이번 달부터 새 줄로 이어 가요
    const nt = appFixedWrite_(ss, sh, v, f, cur, '앱에서 금액·결제일 변경');
    sh.getRange(t, 7).setValue(appYmSerial_(appYmShift_(cur, -1)));
    sh.getRange(t, 7).setNumberFormat('yyyy-mm');
    return { ok: true, i: nt, message: '이번 달부터 바뀐 값으로 들어가요 (지난 기록은 그대로)' };
  }
  sh.getRange(t, 1, 1, 5).setValues([[f.n, f.c, f.a, f.p, f.d]]);
  return { ok: true, i: t, message: '고쳤어요' };
}

// hard=false: 해지 (종료월만 적어서 지난 기록은 남김) / hard=true: 기록까지 지우기
// b.undo: 해지 취소 (종료월 비우기)
function appFixedEnd_(b, hard) {
  const ss = book_();
  const sh = ss.getSheetByName(FIXED_SHEET);
  const t = Number(b.i);
  const r = appFixedRow_(sh, t, b.check);
  const item = { n: String(r[0]).trim(), c: String(r[1] || ''), a: Number(r[2]) || 0, p: String(r[3] || ''), d: Number(r[4]) || 0 };
  if (!hard && b.undo) {
    sh.getRange(t, 7).setValue('');
    return { ok: true, message: '해지를 취소했어요' };
  }
  const cur = appCurYm_(ss), start = appYm_(ss, r[5]);
  let endYm = '';
  if (!hard) {
    endYm = appFixedPaidThisMonth_(ss, item.d) ? cur : appYmShift_(cur, -1);
    if (!start || endYm < start) hard = true;                             // 한 번도 안 나간 줄은 그냥 지워요
  }
  if (hard) {
    sh.getRange(t, 1, 1, 8).setValues([['', '', '', '', '', '', '', false]]);
    sh.getRange(t, 10, 1, 2).setValues([['', '']]);
    return { ok: true, deleted: true, item: item, message: '고정비에서 지웠어요: ' + item.n };
  }
  sh.getRange(t, 7).setValue(appYmSerial_(endYm));
  sh.getRange(t, 7).setNumberFormat('yyyy-mm');
  const inc = (appCats_(ss).find(c => c.name === item.c) || {}).kind === '수입';
  const verb = inc ? '들어와요' : '나가요';
  return { ok: true, end: endYm, message: '해지했어요 · ' + (endYm === cur ? '이번 달까지만 ' + verb : '이번 달부터 안 ' + verb) };
}

/* ── 대시보드: 시트 모양 그대로 ──
 * 칸마다 글자와 '역할'(제목·섹션·표머리·항목명·값·설명)을 보내면 앱이 그대로 카드로 그려요.
 * 대시보드 탭을 고치면 앱 화면도 따라 바뀌어요.
 */
function appDash_() {
  const ss = book_();
  const sh = ss.getSheetByName(DASH_SHEET);
  const last = Math.min(sh.getLastRow(), 250);
  const rg = sh.getRange(1, 2, last, 7);                       // B~H
  const dv = rg.getDisplayValues(), bg = rg.getBackgrounds(), fw = rg.getFontWeights(),
        fs = rg.getFontSizes(), fc = rg.getFontColors();
  const rows = dv.map((r, i) => r.map((v, j) => {
    if (v === '' || v === null) return null;
    const k = appKind_(bg[i][j], fw[i][j], fs[i][j], fc[i][j]);
    const cell = { v: v, k: k };
    const tint = appTint_(fc[i][j]);
    if (tint) cell.t = tint;
    return cell;
  }));
  // 끝의 빈 줄 정리
  while (rows.length && rows[rows.length - 1].every(c => !c)) rows.pop();
  return { ok: true, rows: rows, snap: appSnap_(ss) };
}

function appKind_(bg, fw, fs, fc) {
  bg = String(bg || '').toLowerCase();
  const bold = fw === 'bold';
  if (Number(fs) >= 16) return 'title';
  if (bg === '#eaeff7') return 'sec';
  if (bg === '#e9edf5') return 'th';
  if (bg === '#f4f6f8') return bold ? 'rh' : 'lab';
  if (bg === '#f7f8fb') return 'para';
  if (Number(fs) <= 9) return 'sub';
  if (bold) return 'val';
  if (String(fc || '').toLowerCase() === '#6b7180') return 'note';
  return 'txt';
}

// 기본 글자색이 아니면 의미 있는 색으로 묶어서 보내요 (초록=좋음, 빨강=나쁨, 주황=주의)
function appTint_(fc) {
  const h = String(fc || '').toLowerCase();
  if (!/^#[0-9a-f]{6}$/.test(h)) return '';
  const r = parseInt(h.slice(1, 3), 16), g = parseInt(h.slice(3, 5), 16), b = parseInt(h.slice(5, 7), 16);
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  if (max - min < 40) return '';                                   // 회색 계열
  if (g > r && g > b) return 'green';
  if (r > g && r > b) return (g > 90 ? 'orange' : 'red');
  return '';
}

function appSnap_(ss) {
  const sh = ss.getSheetByName(SNAP_SHEET_APP);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()
    .filter(r => isDate_(r[0]) && typeof r[1] === 'number')
    .map(r => ({ d: appYmd_(ss, r[0]), fin: r[1], path: typeof r[2] === 'number' ? r[2] : null }));
}

/* ── 보유종목·현금·기타·설정: 표 그대로 ── */

function appSheet_(key) {
  const name = APP_SHEETS[key];
  if (!name) return { ok: false, error: '없는 탭이에요' };
  const ss = book_();
  const sh = ss.getSheetByName(name);
  if (key === 'set') return appSetSheet_(ss, sh);
  const lastCol = sh.getLastColumn();
  const head = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  let cols = head.length;
  while (cols > 0 && !head[cols - 1]) cols--;
  // 수식이 아래쪽까지 미리 깔려 있어서 getLastRow가 크게 나와요. 앞쪽 입력 칸(A~D)이 찬 마지막 줄까지만 읽어요.
  const all = Math.max(sh.getLastRow() - 1, 0);
  let n = 0;
  if (all) {
    const keyCols = sh.getRange(2, 1, all, Math.min(4, cols)).getValues();
    for (let i = keyCols.length - 1; i >= 0; i--) { if (keyCols[i].some(v => v !== '' && v !== null && v !== false)) { n = i + 1; break; } }
  }
  const rg = n ? sh.getRange(2, 1, n, cols) : null;
  const vals = rg ? rg.getValues() : [], disp = rg ? rg.getDisplayValues() : [], fx = rg ? rg.getFormulas() : [];
  const rows = [];
  vals.forEach((r, i) => {
    // 직접 입력 칸이 하나라도 차 있는 줄만
    const has = r.some((v, j) => !fx[i][j] && v !== '' && v !== null && v !== false);
    if (!has) return;
    rows.push({ r: i + 2, v: r.map(v => isDate_(v) ? appYmd_(ss, v) : v), d: disp[i] });
  });
  // 어떤 칸이 수식(자동 계산)인지는 2행 기준
  const f2 = sh.getRange(2, 1, 1, cols).getFormulas()[0].map(f => !!f);
  return { ok: true, sheet: key, head: head.slice(0, cols), formula: f2, opts: appOpts_(ss, sh, cols), rows: rows };
}

// 2행의 드롭다운 목록 → {열번호: [선택지]}
function appOpts_(ss, sh, cols) {
  const dvs = sh.getRange(2, 1, 1, cols).getDataValidations()[0];
  const out = {};
  dvs.forEach((dv, j) => {
    if (!dv) return;
    const t = dv.getCriteriaType(), a = dv.getCriteriaValues();
    if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) out[j] = a[0].map(String);
    else if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) out[j] = [].concat.apply([], a[0].getDisplayValues()).filter(String);
    else if (t === SpreadsheetApp.DataValidationCriteria.CHECKBOX) out[j] = 'bool';
  });
  return out;
}

function appSetSheet_(ss, sh) {
  const n = sh.getLastRow();
  const rg = sh.getRange(1, 1, n, 3);
  const vals = rg.getValues(), disp = rg.getDisplayValues();
  const bgB = sh.getRange(1, 2, n, 1).getBackgrounds();
  const nfB = sh.getRange(1, 2, n, 1).getNumberFormats();
  const dvB = sh.getRange(1, 2, n, 1).getDataValidations();
  const items = [];
  for (let i = 1; i < n; i++) {                       // 1행은 머리글
    const label = String(disp[i][0] || '').trim();
    if (!label || APP_SET_HIDE.indexOf(label) >= 0) continue;
    const isSec = !disp[i][1] && !disp[i][2];
    if (isSec) { items.push({ sec: label }); continue; }
    const edit = String(bgB[i][0]).toLowerCase() === INPUT_BG;
    const nf = String(nfB[i][0] || '');
    let type = 'text';
    if (/%/.test(nf)) type = 'pct';
    else if (/y/.test(nf) && /d/.test(nf)) type = 'date';
    else if (typeof vals[i][1] === 'number' || /[#0]/.test(nf)) type = 'num';
    const dv = dvB[i][0];
    let opts = null, optsD = null;
    if (dv) {
      const t = dv.getCriteriaType(), a = dv.getCriteriaValues();
      if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) { opts = a[0].slice(); optsD = a[0].map(String); }
      else if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
        const ov = [].concat.apply([], a[0].getValues()), od = [].concat.apply([], a[0].getDisplayValues());
        opts = []; optsD = [];
        ov.forEach((v, k) => { if (v !== '' && v !== null) { opts.push(v); optsD.push(od[k]); } });
      }
    }
    const raw = vals[i][1];
    items.push({ r: i + 1, label: label, v: isDate_(raw) ? appYmd_(ss, raw) : raw, d: disp[i][1], desc: disp[i][2],
                 edit: edit, type: opts ? 'opt' : type, opts: opts, optsD: optsD });
  }
  return { ok: true, sheet: 'set', items: items };
}

// 칸 고치기: {sheet, cells:[{r, c, v}]}  (c는 1부터)
function appSet_(b) {
  const name = APP_SHEETS[b.sheet];
  if (!name) throw new Error('고칠 수 없는 탭이에요');
  const ss = book_();
  const sh = ss.getSheetByName(name);
  const cells = Array.isArray(b.cells) ? b.cells : [];
  if (!cells.length) throw new Error('고칠 내용이 없어요');
  cells.forEach(x => {
    const r = Number(x.r), c = Number(x.c);
    if (!(r >= 2) || !(c >= 1)) throw new Error('칸 위치가 맞지 않아요');
    const cell = sh.getRange(r, c);
    if (cell.getFormula()) throw new Error('자동 계산 칸은 고칠 수 없어요');
    if (b.sheet === 'set' && (c !== 2 || String(cell.getBackground()).toLowerCase() !== INPUT_BG)) throw new Error('직접 입력 칸만 고칠 수 있어요');
    let v = x.v;
    if (x.type === 'date' && v) v = appSerial_(v);
    else if (x.type === 'num' || x.type === 'pct') v = (v === '' || v === null) ? '' : Number(v);
    cell.setValue(v === null || v === undefined ? '' : v);
  });
  SpreadsheetApp.flush();
  return { ok: true, message: '저장했어요' };
}

// 보유종목·현금·기타에 한 줄 추가: {sheet, cells:[{c, v, type}]}  → 비어 있는 첫 줄에 써요 (수식 칸은 그대로)
function appRowAdd_(b) {
  const name = APP_SHEETS[b.sheet];
  if (!name || b.sheet === 'set') throw new Error('줄을 더할 수 없는 탭이에요');
  const ss = book_();
  const sh = ss.getSheetByName(name);
  const cols = sh.getLastColumn();
  const n = Math.min(Math.max(sh.getLastRow(), 2), 300);
  const vals = sh.getRange(2, 1, n, cols).getValues(), fx = sh.getRange(2, 1, n, cols).getFormulas();
  let target = 0;
  for (let i = 0; i < vals.length; i++) {
    const empty = vals[i].every((v, j) => fx[i][j] || v === '' || v === null || v === false);
    if (empty) { target = i + 2; break; }
  }
  if (!target) target = n + 2;
  // 수식이 아래까지 안 내려와 있으면 바로 위 줄 수식을 복사
  const f2 = sh.getRange(2, 1, 1, cols).getFormulasR1C1()[0];
  const ft = sh.getRange(target, 1, 1, cols).getFormulas()[0];
  f2.forEach((f, j) => { if (f && !ft[j]) sh.getRange(target, j + 1).setFormulaR1C1(f); });
  appSet_({ sheet: b.sheet, cells: (b.cells || []).map(x => ({ r: target, c: x.c, v: x.v, type: x.type })) });
  return { ok: true, row: target, message: '추가했어요' };
}

// 보유종목·현금·기타 한 줄 비우기 (수식 칸은 남겨요)
function appRowClear_(b) {
  const name = APP_SHEETS[b.sheet];
  if (!name || b.sheet === 'set') throw new Error('비울 수 없는 탭이에요');
  const ss = book_();
  const sh = ss.getSheetByName(name);
  const r = Number(b.r);
  if (!(r >= 2)) throw new Error('줄 번호가 없어요');
  const cols = sh.getLastColumn();
  const fx = sh.getRange(r, 1, 1, cols).getFormulas()[0];
  const cur = sh.getRange(r, 1, 1, cols).getValues()[0];
  if (b.check && String(cur[0]) !== String(b.check)) throw new Error('시트 내용이 바뀌었어요. 새로고침해 주세요');
  fx.forEach((f, j) => {
    if (f) return;
    const cell = sh.getRange(r, j + 1);
    cell.setValue(typeof cur[j] === 'boolean' ? false : '');
  });
  return { ok: true, message: '지웠어요' };
}
