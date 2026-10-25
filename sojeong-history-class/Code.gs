/**
 * 유쾌한 소정샘의 역사 Class - Code.gs
 * 구글 시트 "기록" 시트 1개로 모든 데이터를 관리합니다. (동시 접속 40명 기준 설계)
 *
 * [최초 설정]
 *  1) 프로젝트 설정 > 스크립트 속성에 아래 3개 추가
 *     - GEMINI_API_KEY : Gemini API 키
 *     - NOTION_TOKEN   : 노션 내부 통합(Integration) 시크릿  (과제 제출 DB에 '연결' 추가 필수)
 *     - TEACHER_PW     : 선생님 암호 (미설정 시 1234)
 *  2) 에디터에서 setup() 한 번 실행 → "기록" 시트와 머리글 자동 생성
 *  3) 배포 > 새 배포 > 웹 앱 (실행: 나 / 액세스: 모든 사용자)
 */

const CFG = {
  SHEET_ID: '11mGWsmflG2aHBZwKcZ7StCyfDskFr_EnLTx74y9M1FE',   // 비우면 이 스크립트가 연결된 스프레드시트 사용
  SHEET_NAME: '기록',
  TEACHER_PW_DEFAULT: '1234',
  GEMINI_MODEL: 'gemini-3.6-flash',   // ※ 실제 모델 ID를 AI Studio에서 확인해 맞추세요
  NOTION_DB_ID: '840e81aa88cf4c72b6199367471452d1',  // 노션 '학생 과제 수합표' 데이터베이스
  SESSION_SEC: 6 * 60 * 60,
  AI_COOLDOWN_SEC: 5,
  CACHE_SEC: 20
};

// 0순번 1일시 2구분 3수업일 4학반 5번호 6이름 7제목 8내용 9링크 10부가(JSON) 11상태
const HEADERS = ['순번', '일시', '구분', '수업일', '학반', '번호', '이름', '제목', '내용', '링크', '부가(JSON)', '상태'];
const COL = { SEQ: 0, TIME: 1, TYPE: 2, DATE: 3, CLS: 4, NO: 5, NAME: 6, TITLE: 7, BODY: 8, LINK: 9, EXTRA: 10, STATUS: 11 };
const ROOMS = ['사전학습', '심화탐구', '오늘의역사'];   // 방 게시판 구분값

/* ------------------------------ 진입점 ------------------------------ */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('유쾌한 소정샘의 역사 Class')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** 모든 클라이언트 호출은 이 함수 하나로 들어옵니다. */
function api(action, p) {
  p = p || {};
  try {
    switch (action) {
      case 'login':      return login_(p);
      case 'join':       return join_(p);
      case 'getMonth':   return getMonth_(p);
      case 'getLesson':  return getLesson_(p);
      case 'saveLesson': requireTeacher_(p); return saveLesson_(p);
      case 'getRoom':    return getRoom_(p);
      case 'saveItem':   requireTeacher_(p); return saveItem_(p);
      case 'deleteItem': requireTeacher_(p); return deleteItem_(p);
      case 'ask':        return ask_(p);
      case 'submit':     return submitReport_(p);
      case 'raise':      return raise_(p);
      case 'getHands':   requireTeacher_(p); return getHands_(p);
      case 'clearHands': requireTeacher_(p); return clearHands_(p);
      case 'getRoster':  requireTeacher_(p); return getRoster_();
      default: throw new Error('알 수 없는 요청');
    }
  } catch (e) {
    return { ok: false, error: String(e.message || e) };
  }
}

/* ------------------------------ 시트 도우미 ------------------------------ */

function setup() {
  sheet_();
  return '완료: "' + CFG.SHEET_NAME + '" 시트 준비됨';
}

function sheet_() {
  const ss = CFG.SHEET_ID ? SpreadsheetApp.openById(CFG.SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CFG.SHEET_NAME);
  if (!sh) {
    const all = ss.getSheets();
    if (all.length === 1) { sh = all[0]; sh.setName(CFG.SHEET_NAME); }   // 새 파일의 기본 탭을 '기록'으로
    else sh = ss.insertSheet(CFG.SHEET_NAME);
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#fde68a');
    sh.setFrozenRows(1);
  }
  return sh;
}

function now_() { return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'); }
function today_() { return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd'); }

/** 쓰기: 락 후 한 줄 추가. 순번은 마지막 행 기준으로 자동 증가. */
function append_(type, o) {
  o = o || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = sheet_();
    const last = sh.getLastRow();
    const seq = last <= 1 ? 1 : Number(sh.getRange(last, 1).getValue()) + 1;
    sh.appendRow([seq, now_(), type, o.date || '', o.cls || '', o.no || '', o.name || '', o.title || '',
      o.body || '', o.link || '', o.extra ? JSON.stringify(o.extra) : '', o.status || '']);
    CacheService.getScriptCache().removeAll(['lessons', 'room_사전학습', 'room_심화탐구', 'room_오늘의역사']);
    return seq;
  } finally {
    lock.releaseLock();
  }
}

/** 읽기: 해당 구분의 행만. (최근 limit행만 읽어 부하 줄임) */
function rows_(type, limit) {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const start = limit ? Math.max(2, last - limit + 1) : 2;
  const vals = sh.getRange(start, 1, last - start + 1, HEADERS.length).getValues();
  const out = [];
  vals.forEach(function (r, i) {
    if (r[COL.TYPE] === type) out.push({ row: start + i, v: r });
  });
  return out;
}

function cached_(key, fn) {
  const c = CacheService.getScriptCache();
  const hit = c.get(key);
  if (hit) return JSON.parse(hit);
  const val = fn();
  try { c.put(key, JSON.stringify(val), CFG.CACHE_SEC); } catch (e) {}
  return val;
}

function safeUrl_(u) {
  u = String(u || '').trim();
  return /^https?:\/\//i.test(u) ? u : '';
}

function dateStr_(v) {
  return v instanceof Date ? Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd') : String(v || '');
}

/* ------------------------------ 선생님 인증 ------------------------------ */

function login_(p) {
  const pw = PropertiesService.getScriptProperties().getProperty('TEACHER_PW') || CFG.TEACHER_PW_DEFAULT;
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('pwfail') || 0);
  if (fails >= 10) throw new Error('시도 횟수 초과. 잠시 후 다시 시도하세요.');
  if (String(p.pw) !== pw) {
    cache.put('pwfail', String(fails + 1), 300);
    throw new Error('암호가 맞지 않습니다.');
  }
  const token = Utilities.getUuid();
  cache.put('t_' + token, '1', CFG.SESSION_SEC);
  return { ok: true, token: token };
}

function requireTeacher_(p) {
  if (!p.token || !CacheService.getScriptCache().get('t_' + p.token)) throw new Error('선생님 인증이 필요합니다.');
}

/* ------------------------------ 입장 / 달력 / 수업 ------------------------------ */

function join_(p) {
  const name = String(p.name || '').trim().slice(0, 20);
  if (!name) throw new Error('이름을 입력하세요.');
  append_('입장', { date: today_(), cls: String(p.cls || '').slice(0, 10), no: p.no, name: name });
  return { ok: true };
}

function lessons_() {
  return cached_('lessons', function () {
    const map = {};
    rows_('수업').forEach(function (r) {   // 같은 날짜는 나중 행이 덮어씀
      if (r.v[COL.STATUS] === '삭제') return;
      let ex = {};
      try { ex = JSON.parse(r.v[COL.EXTRA] || '{}'); } catch (e) {}
      map[dateStr_(r.v[COL.DATE])] = { date: dateStr_(r.v[COL.DATE]), title: r.v[COL.TITLE], body: r.v[COL.BODY], links: ex };
    });
    return map;
  });
}

function getMonth_(p) {
  const ym = String(p.ym || '');   // 'yyyy-MM'
  const all = lessons_();
  const days = {};
  Object.keys(all).forEach(function (d) { if (d.indexOf(ym) === 0) days[d] = all[d].title; });
  return { ok: true, days: days, today: today_() };
}

function getLesson_(p) {
  const l = lessons_()[String(p.date)];
  return { ok: true, lesson: l || null };
}

function saveLesson_(p) {
  const L = p.links || {};
  const links = {
    intro: safeUrl_(L.intro), summary: safeUrl_(L.summary), worksheet: safeUrl_(L.worksheet), activity: safeUrl_(L.activity)
  };
  append_('수업', { date: String(p.date), title: String(p.title || '').slice(0, 100), body: String(p.body || '').slice(0, 2000), extra: links });
  return { ok: true };
}

/* ------------------------------ 방 게시판 (사전학습/심화탐구/오늘의 역사) ------------------------------ */

function getRoom_(p) {
  if (ROOMS.indexOf(p.room) < 0) throw new Error('잘못된 방');
  const items = cached_('room_' + p.room, function () {
    return rows_(p.room).filter(function (r) { return r.v[COL.STATUS] !== '삭제'; }).map(function (r) {
      return { seq: r.v[COL.SEQ], date: dateStr_(r.v[COL.DATE]), title: r.v[COL.TITLE], body: r.v[COL.BODY], link: r.v[COL.LINK] };
    }).reverse();
  });
  return { ok: true, items: items };
}

function saveItem_(p) {
  if (ROOMS.indexOf(p.room) < 0) throw new Error('잘못된 방');
  append_(p.room, { date: String(p.date || today_()), title: String(p.title || '').slice(0, 100), body: String(p.body || '').slice(0, 3000), link: safeUrl_(p.link) });
  return { ok: true };
}

function deleteItem_(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = sheet_();
    const last = sh.getLastRow();
    if (last < 2) return { ok: true };
    const seqs = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = seqs.length - 1; i >= 0; i--) {
      if (Number(seqs[i][0]) === Number(p.seq)) { sh.getRange(i + 2, COL.STATUS + 1).setValue('삭제'); break; }
    }
    CacheService.getScriptCache().removeAll(['room_사전학습', 'room_심화탐구', 'room_오늘의역사']);
  } finally { lock.releaseLock(); }
  return { ok: true };
}

/* ------------------------------ AI 보조 선생님 (Gemini) ------------------------------ */

function ask_(p) {
  const text = String(p.text || '').trim().slice(0, 2000);
  if (!text && !p.image) throw new Error('용어를 입력하거나 사진을 올려주세요.');
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('GEMINI_API_KEY가 설정되지 않았습니다.');

  const cache = CacheService.getScriptCache();
  const who = 'ai_' + String(p.name || 'anon').slice(0, 20);
  if (cache.get(who)) throw new Error('잠깐만요! ' + CFG.AI_COOLDOWN_SEC + '초 뒤에 다시 질문해 주세요.');
  cache.put(who, '1', CFG.AI_COOLDOWN_SEC);

  const parts = [];
  if (text) parts.push({ text: text });
  if (p.image && p.image.data) {
    if (String(p.image.data).length > 4000000) throw new Error('사진이 너무 큽니다.');
    parts.push({ inline_data: { mime_type: String(p.image.mime || 'image/jpeg'), data: p.image.data } });
  }
  const body = {
    system_instruction: { parts: [{ text:
      '너는 "유쾌한 소정샘의 역사 Class"의 AI 보조 선생님이야. 중·고등학생에게 한국사·세계사 용어와 사진(사료, 유물, 지도 등)을 쉽고 친절한 존댓말로 설명해 줘. ' +
      '형식: 1) 한 줄 뜻 2) 시기·배경 3) 핵심 포인트 2~3개 4) 함께 알면 좋은 용어. 전체 400자 내외. 역사와 무관한 질문은 정중히 사양하고, 확실하지 않은 내용은 추측하지 말고 모른다고 말해.' }] },
    contents: [{ role: 'user', parts: parts }],
    generationConfig: { maxOutputTokens: 1024 }
  };
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + CFG.GEMINI_MODEL + ':generateContent';
  const res = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key },
    payload: JSON.stringify(body), muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const json = JSON.parse(res.getContentText() || '{}');
  if (code !== 200) throw new Error('AI 응답 오류(' + code + '): ' + ((json.error && json.error.message) || '').slice(0, 120));
  const ans = (((json.candidates || [])[0] || {}).content || {}).parts;
  const out = ans ? ans.map(function (x) { return x.text || ''; }).join('') : '';
  if (!out) throw new Error('AI가 답변을 만들지 못했어요. 다시 시도해 주세요.');
  append_('AI질문', { date: today_(), cls: p.cls, no: p.no, name: p.name, title: text.slice(0, 100), body: out.slice(0, 3000), extra: { image: !!p.image } });
  return { ok: true, answer: out };
}

/* ------------------------------ 보고서 제출 → 노션 ------------------------------ */

function submitReport_(p) {
  const name = String(p.name || '').trim();
  const title = String(p.title || '').trim();
  const link = safeUrl_(p.link);
  if (!name || !title) throw new Error('이름과 과제명을 입력하세요.');
  if (!link) throw new Error('제출 링크는 http(s)로 시작해야 합니다.');
  const token = PropertiesService.getScriptProperties().getProperty('NOTION_TOKEN');
  if (!token) throw new Error('NOTION_TOKEN이 설정되지 않았습니다.');

  const props = {
    '학생 이름': { title: [{ text: { content: name.slice(0, 100) } }] },
    '학년·반': { rich_text: [{ text: { content: String(p.cls || '').slice(0, 50) } }] },
    '과제명': { rich_text: [{ text: { content: title.slice(0, 200) } }] },
    '제출 링크': { url: link },
    '제출일': { date: { start: today_() } },
    '확인 상태': { status: { name: '제출 완료' } }
  };
  if (p.no !== '' && p.no != null && !isNaN(Number(p.no))) props['번호'] = { number: Number(p.no) };

  const res = UrlFetchApp.fetch('https://api.notion.com/v1/pages', {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token, 'Notion-Version': '2022-06-28' },
    payload: JSON.stringify({ parent: { database_id: CFG.NOTION_DB_ID }, properties: props }),
    muteHttpExceptions: true
  });
  const ok = res.getResponseCode() === 200;
  append_('보고서', { date: today_(), cls: p.cls, no: p.no, name: name, title: title, link: link, status: ok ? '노션 등록' : '노션 실패' });
  if (!ok) throw new Error('노션 등록 실패(' + res.getResponseCode() + '). 선생님께 알려주세요.');
  return { ok: true };
}

/* ------------------------------ 손들기 / 발표자 뽑기 ------------------------------ */

function raise_(p) {
  const name = String(p.name || '').trim();
  if (!name) throw new Error('이름이 없습니다.');
  const cache = CacheService.getScriptCache();
  if (cache.get('hand_' + name)) return { ok: true };   // 연타 방지
  cache.put('hand_' + name, '1', 10);
  append_('손들기', { date: today_(), cls: p.cls, no: p.no, name: name, status: '대기' });
  return { ok: true };
}

function getHands_() {
  const list = rows_('손들기', 300).filter(function (r) {
    return r.v[COL.STATUS] === '대기' && dateStr_(r.v[COL.DATE]) === today_();
  }).map(function (r) {
    return { cls: r.v[COL.CLS], no: r.v[COL.NO], name: r.v[COL.NAME], time: String(r.v[COL.TIME]).slice(11, 19) };
  });
  return { ok: true, hands: list };
}

function clearHands_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = sheet_();
    rows_('손들기', 300).forEach(function (r) {
      if (r.v[COL.STATUS] === '대기') sh.getRange(r.row, COL.STATUS + 1).setValue('완료');
    });
  } finally { lock.releaseLock(); }
  return { ok: true };
}

/** 오늘 입장한 학생 명단(중복 제거) → 발표자 뽑기용 */
function getRoster_() {
  const seen = {}, list = [];
  rows_('입장', 500).forEach(function (r) {
    if (dateStr_(r.v[COL.DATE]) !== today_()) return;
    const k = r.v[COL.CLS] + '|' + r.v[COL.NO] + '|' + r.v[COL.NAME];
    if (seen[k]) return;
    seen[k] = 1;
    list.push({ cls: r.v[COL.CLS], no: r.v[COL.NO], name: r.v[COL.NAME] });
  });
  return { ok: true, roster: list };
}
