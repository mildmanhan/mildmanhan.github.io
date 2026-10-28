/**
 * 키다리용달 링크 페이지 상담예약 백엔드 (Google Apps Script + 구글 시트)
 * 홍보 영상은 여기서 다루지 않고 GitHub 저장소의 videos.json + videos/ 폴더로 관리합니다.
 *
 * 스크립트 속성(프로젝트 설정 > 스크립트 속성)에 넣을 값
 *   ADMIN_PASSWORD       관리자 페이지 비밀번호
 *   KAKAO_REST_KEY       카카오 디벨로퍼스 앱의 REST API 키
 *   KAKAO_CLIENT_SECRET  (선택) 카카오 앱에서 Client Secret을 '사용'으로 켰다면 그 코드
 *   SITE_URL             (선택) 기본값 https://mildmanhan.github.io
 */

const TZ = 'Asia/Seoul';
const SHEET_SLOTS = '일정';
const SHEET_BOOKINGS = '예약';
const PUBLIC_CACHE_KEY = 'public_v1';

const props = () => PropertiesService.getScriptProperties();
const siteUrl = () => props().getProperty('SITE_URL') || 'https://mildmanhan.github.io';

/* ───────── 최초 1회 실행: 시트 탭 + 매일 카카오 토큰 갱신 트리거 ───────── */
function setup() {
  const ss = SpreadsheetApp.getActive();
  const make = (name, header) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    sh.getRange('A:Z').setNumberFormat('@'); // 날짜·시간이 자동 변환되지 않도록 전부 텍스트
    if (sh.getLastRow() === 0) sh.appendRow(header);
    sh.setFrozenRows(1);
  };
  make(SHEET_SLOTS, ['날짜', '시간']);
  make(SHEET_BOOKINGS, ['id', '접수시각', '날짜', '시간', '이름', '연락처', '내용', '상태']);
  const blank = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 1) ss.deleteSheet(blank);

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'keepKakaoAlive')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('keepKakaoAlive').timeBased().everyDays(1).atHour(4).create();
}

/* ───────── 웹앱 입구 ───────── */
function doGet(e) {
  const p = e.parameter || {};
  if (p.code || p.error) return kakaoCallback_(p);
  if (p.action === 'public') return json_(publicData_());
  return json_({ ok: true, msg: 'kidari api' });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: '잘못된 요청' }); }
  try {
    if (body.action === 'book') return json_(book_(body));
    if (!checkPassword_(body.pw)) return json_({ ok: false, error: '비밀번호가 틀렸습니다', auth: false });
    switch (body.action) {
      case 'login': return json_({ ok: true });
      case 'adminData': return json_(adminData_());
      case 'setSlots': return json_(setSlots_(body));
      case 'setStatus': return json_(setStatus_(body));
      case 'kakaoAuthUrl': return json_(kakaoAuthUrl_());
      case 'kakaoTest': return json_(kakaoTest_());
    }
    return json_({ ok: false, error: '알 수 없는 요청' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

/* ───────── 공개 데이터 (방문자 페이지) ───────── */
function publicData_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(PUBLIC_CACHE_KEY);
  if (hit) return JSON.parse(hit);

  const today = todayStr_();
  const taken = takenSet_();
  const slots = {};
  readRows_(SHEET_SLOTS).forEach(r => {
    const d = cell_(r[0]), t = cell_(r[1]);
    if (!d || d < today) return;
    (slots[d] = slots[d] || []).push({ t, taken: taken.has(d + ' ' + t) });
  });
  Object.values(slots).forEach(list => list.sort((a, b) => a.t.localeCompare(b.t)));

  const data = { ok: true, slots, today };
  cache.put(PUBLIC_CACHE_KEY, JSON.stringify(data), 300);
  return data;
}

/* ───────── 예약 접수 ───────── */
function book_(b) {
  if (b.hp) return { ok: true }; // 스팸봇 함정 칸이 채워짐
  const name = clean_(b.name, 20);
  const phone = String(b.phone || '').replace(/[^0-9]/g, '');
  const memo = clean_(b.memo, 500);
  const date = String(b.date || '');
  const times = Array.isArray(b.times) ? b.times.map(String) : [];
  if (!name) return { ok: false, error: '성함을 적어 주세요' };
  if (!/^01[0-9]{8,9}$/.test(phone)) return { ok: false, error: '휴대폰 번호를 확인해 주세요' };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || date < todayStr_()) return { ok: false, error: '날짜를 다시 골라 주세요' };
  if (!times.length || times.length > 6) return { ok: false, error: '시간대를 1개 이상 골라 주세요' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const open = new Set(readRows_(SHEET_SLOTS).filter(r => cell_(r[0]) === date).map(r => cell_(r[1])));
    const taken = takenSet_();
    const bad = times.filter(t => !open.has(t) || taken.has(date + ' ' + t));
    if (bad.length) {
      clearPublicCache_();
      return { ok: false, error: bad.join(', ') + ' 시간은 방금 마감되었어요. 다른 시간을 골라 주세요', refresh: true };
    }
    times.sort();
    const id = Utilities.getUuid().slice(0, 8);
    const phoneFmt = phone.replace(/^(\d{3})(\d{3,4})(\d{4})$/, '$1-$2-$3');
    sheet_(SHEET_BOOKINGS).appendRow([id, nowStr_(), date, times.join(','), name, phoneFmt, memo, '접수']);
    clearPublicCache_();

    const msg = '📅 새 상담예약\n' +
      '날짜: ' + prettyDate_(date) + '\n' +
      '시간: ' + times.join(', ') + '\n' +
      '이름: ' + name + '\n' +
      '연락처: ' + phoneFmt + (memo ? '\n내용: ' + memo : '');
    try { kakaoSend_(msg); } catch (err) { console.error('카카오 알림 실패', err); }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// 취소되지 않은 예약이 잡은 "날짜 시간" 묶음
function takenSet_() {
  const s = new Set();
  readRows_(SHEET_BOOKINGS).forEach(r => {
    if (cell_(r[7]) === '취소') return;
    cell_(r[3]).split(',').filter(Boolean).forEach(t => s.add(cell_(r[2]) + ' ' + t));
  });
  return s;
}

/* ───────── 관리자 ───────── */
function checkPassword_(pw) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('pw_fails') || 0);
  if (fails >= 10) throw new Error('비밀번호를 너무 많이 틀렸어요. 10분 뒤에 다시 시도해 주세요');
  const real = props().getProperty('ADMIN_PASSWORD');
  if (!real) throw new Error('스크립트 속성에 ADMIN_PASSWORD를 먼저 넣어 주세요');
  if (String(pw || '') === real) return true;
  cache.put('pw_fails', String(fails + 1), 600);
  return false;
}

function adminData_() {
  const today = todayStr_();
  const slots = {};
  readRows_(SHEET_SLOTS).forEach(r => {
    const d = cell_(r[0]);
    if (d >= today) (slots[d] = slots[d] || []).push(cell_(r[1]));
  });
  const bookings = readRows_(SHEET_BOOKINGS).map(r => ({
    id: cell_(r[0]), at: cell_(r[1]), date: cell_(r[2]), times: cell_(r[3]),
    name: cell_(r[4]), phone: cell_(r[5]), memo: cell_(r[6]), status: cell_(r[7]),
  })).reverse().slice(0, 300);
  return {
    ok: true, today, slots, bookings,
    kakao: { connected: !!props().getProperty('KAKAO_REFRESH'), redirect: ScriptApp.getService().getUrl() },
  };
}

// dates 배열의 각 날짜 일정을 times로 통째로 바꿈 (times가 비면 그 날짜는 예약 닫힘)
function setSlots_(b) {
  const dates = (Array.isArray(b.dates) ? b.dates : []).filter(d => /^\d{4}-\d{2}-\d{2}$/.test(d));
  const times = (Array.isArray(b.times) ? b.times : []).filter(t => /^\d{2}:\d{2}$/.test(t)).sort();
  if (!dates.length) throw new Error('날짜를 골라 주세요');
  const sh = sheet_(SHEET_SLOTS);
  const today = todayStr_();
  const drop = new Set(dates);
  const keep = readRows_(SHEET_SLOTS).filter(r => !drop.has(cell_(r[0])) && cell_(r[0]) >= today);
  dates.forEach(d => times.forEach(t => keep.push([d, t])));
  keep.sort((a, b2) => (a[0] + a[1]).localeCompare(b2[0] + b2[1]));
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  if (keep.length) sh.getRange(2, 1, keep.length, 2).setNumberFormat('@').setValues(keep.map(r => [cell_(r[0]), cell_(r[1])]));
  clearPublicCache_();
  return adminData_();
}

function setStatus_(b) {
  const allowed = ['접수', '확정', '완료', '취소'];
  if (allowed.indexOf(b.status) < 0) throw new Error('상태 값 오류');
  const rows = readRows_(SHEET_BOOKINGS);
  const i = rows.findIndex(r => cell_(r[0]) === b.id);
  if (i >= 0) sheet_(SHEET_BOOKINGS).getRange(i + 2, 8).setValue(b.status);
  clearPublicCache_();
  return adminData_();
}

/* ───────── 카카오톡 나에게 보내기 ───────── */
function kakaoAuthUrl_() {
  const key = props().getProperty('KAKAO_REST_KEY');
  if (!key) throw new Error('스크립트 속성에 KAKAO_REST_KEY를 먼저 넣어 주세요');
  const state = Utilities.getUuid();
  CacheService.getScriptCache().put('kakao_state', state, 600);
  const url = 'https://kauth.kakao.com/oauth/authorize?response_type=code&scope=talk_message' +
    '&client_id=' + encodeURIComponent(key) +
    '&redirect_uri=' + encodeURIComponent(ScriptApp.getService().getUrl()) +
    '&state=' + state;
  return { ok: true, url };
}

function kakaoCallback_(p) {
  const page = (title, text) => HtmlService.createHtmlOutput(
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<div style="font-family:sans-serif;padding:40px 20px;text-align:center"><h2>' + title + '</h2><p>' + text + '</p></div>');
  const cache = CacheService.getScriptCache();
  if (p.error) return page('연결 취소', '카카오 연결이 취소되었어요. 관리자 페이지에서 다시 시도해 주세요.');
  if (!p.state || p.state !== cache.get('kakao_state')) {
    return page('연결 실패', '관리자 페이지의 [카카오 연결하기] 버튼을 눌러 10분 안에 진행해 주세요.');
  }
  cache.remove('kakao_state');
  const res = kakaoToken_({ grant_type: 'authorization_code', code: p.code, redirect_uri: ScriptApp.getService().getUrl() });
  if (!res.refresh_token) return page('연결 실패', '토큰을 받지 못했어요: ' + JSON.stringify(res));
  saveKakaoTokens_(res);
  try { kakaoSend_('✅ 키다리용달 예약 알림이 연결되었습니다.\n이제 상담예약이 들어오면 이 카톡으로 알려 드려요.'); } catch (err) {}
  return page('연결 완료 ✅', '카카오톡으로 확인 메시지를 보냈어요. 이 창은 닫으셔도 됩니다.');
}

function kakaoToken_(params) {
  params.client_id = props().getProperty('KAKAO_REST_KEY');
  const secret = props().getProperty('KAKAO_CLIENT_SECRET');
  if (secret) params.client_secret = secret;
  const r = UrlFetchApp.fetch('https://kauth.kakao.com/oauth/token', { method: 'post', payload: params, muteHttpExceptions: true });
  return JSON.parse(r.getContentText());
}

function saveKakaoTokens_(res) {
  const p = props();
  p.setProperty('KAKAO_ACCESS', res.access_token);
  p.setProperty('KAKAO_ACCESS_EXP', String(Date.now() + (res.expires_in - 300) * 1000));
  if (res.refresh_token) p.setProperty('KAKAO_REFRESH', res.refresh_token);
}

function kakaoAccess_(force) {
  const p = props();
  const refresh = p.getProperty('KAKAO_REFRESH');
  if (!refresh) throw new Error('카카오가 아직 연결되지 않았어요');
  if (!force && Number(p.getProperty('KAKAO_ACCESS_EXP') || 0) > Date.now()) return p.getProperty('KAKAO_ACCESS');
  const res = kakaoToken_({ grant_type: 'refresh_token', refresh_token: refresh });
  if (!res.access_token) throw new Error('카카오 토큰 갱신 실패: ' + JSON.stringify(res));
  saveKakaoTokens_(res);
  return res.access_token;
}

function kakaoSend_(text) {
  const template = {
    object_type: 'text',
    text: text.slice(0, 200),
    link: { web_url: siteUrl() + '/admin.html', mobile_web_url: siteUrl() + '/admin.html' },
    button_title: '예약 확인하기',
  };
  const r = UrlFetchApp.fetch('https://kapi.kakao.com/v2/api/talk/memo/default/send', {
    method: 'post',
    headers: { Authorization: 'Bearer ' + kakaoAccess_() },
    payload: { template_object: JSON.stringify(template) },
    muteHttpExceptions: true,
  });
  if (r.getResponseCode() !== 200) throw new Error('카카오 전송 실패: ' + r.getContentText());
}

function kakaoTest_() {
  kakaoSend_('🔔 테스트 알림입니다.\n상담예약이 들어오면 이렇게 카톡으로 알려 드려요.');
  return { ok: true };
}

// 매일 새벽 트리거: 토큰을 갱신해서 리프레시 토큰(2개월)이 만료되지 않게 유지
function keepKakaoAlive() {
  if (props().getProperty('KAKAO_REFRESH')) kakaoAccess_(true);
}

/* ───────── 도우미 ───────── */
function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('"' + name + '" 시트가 없어요. setup 함수를 먼저 실행해 주세요');
  return sh;
}

function readRows_(name) {
  const sh = sheet_(name);
  const n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, sh.getLastColumn()).getValues() : [];
}

function cell_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v == null ? '' : v).trim();
}

// 앞글자가 = + - @ 이면 시트가 수식으로 읽지 않도록 막고, 길이를 자름
function clean_(v, max) {
  const s = String(v || '').replace(/[\u0000-\u0009\u000B-\u001F]/g, '').trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function todayStr_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function nowStr_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }

function prettyDate_(d) {
  const [y, m, day] = d.split('-').map(Number);
  const w = '일월화수목금토'[new Date(y, m - 1, day).getDay()];
  return m + '월 ' + day + '일 (' + w + ')';
}

function clearPublicCache_() { CacheService.getScriptCache().remove(PUBLIC_CACHE_KEY); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
