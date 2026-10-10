/**
 * 삼성중학교 독도 골든벨 대회 - Code.gs
 * 구글 시트 "기록" 시트 1개에 입장/답안/결과를 자동 기록합니다.
 * (게임 진행 상태는 PropertiesService + CacheService에 저장)
 */

const CFG = {
  SHEET_ID: '',                 // 비워두면 이 스크립트가 연결된 스프레드시트 사용
  SHEET_NAME: '기록',
  TEACHER_PW_DEFAULT: '1234',   // 스크립트 속성 TEACHER_PW 로 덮어쓰기 권장
  DRIVE_FOLDER_NAME: '독도 학습지 제출'   // 학습지 답안이 저장될 드라이브 폴더(없으면 자동 생성)
};

const HEADERS = ['순번', '일시', '구분', '문제번호', '이름', '제출답안', '정답여부', '생존여부', '비고'];

// phase: 0 대기 / 1 문제공개 / 2 답안제출중 / 3 답안마감 / 4 정답확인 / 5 종료
const QUESTIONS = [
  { type: 'ox', q: '독도는 우리나라 영토이다.', choices: ['O', 'X'], a: ['O'] },
  { type: 'choice', q: '독도는 크게 두 개의 큰 섬으로 이루어져 있다. 동도와 어느 섬일까?', choices: ['남도', '서도', '북도', '중도'], a: ['서도'] },
  { type: 'choice', q: '독도에서 가장 가까운 우리나라 섬은?', choices: ['제주도', '거제도', '울릉도', '강화도'], a: ['울릉도'] },
  { type: 'short', q: '조선 숙종 때 일본에 건너가 울릉도·독도가 우리 땅임을 확인받고 돌아온 인물은? (세 글자)', a: ['안용복'] },
  { type: 'short', q: '신라 장군 이사부가 512년에 복속시킨 울릉도·독도 지역의 옛 나라 이름은?', a: ['우산국'] },
  { type: 'short', q: '독도는 문화재청에서 천연기념물 제 몇 호로 지정되어 있을까? (숫자만)', a: ['336'] },
  { type: 'ox', q: '독도의 날은 10월 25일이다.', choices: ['O', 'X'], a: ['O'] },
  { type: 'short', q: '1900년 대한제국이 울릉도를 울도군으로 승격하며 독도를 관할 구역으로 명시한 법령은? (칙령 제 몇 호, 숫자만)', a: ['41'] }
];

function doGet(e) {
  if (e && e.parameter && e.parameter.page === 'worksheet') return worksheetPage_();
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('독도 골든벨 대회')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** ?page=worksheet -> 프로젝트 안의 worksheet.html */
function worksheetPage_() {
  return HtmlService.createHtmlOutputFromFile('worksheet').setTitle('독도 학습지')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function getWorksheetUrl() { return ScriptApp.getService().getUrl() + '?page=worksheet'; }

/* ---------- 저장소 헬퍼 ---------- */
function load_(key, def) {
  const cache = CacheService.getScriptCache();
  let s = cache.get('dg_' + key);
  if (!s) {
    s = PropertiesService.getScriptProperties().getProperty('dg_' + key);
    if (s) cache.put('dg_' + key, s, 21600);
  }
  return s ? JSON.parse(s) : def;
}
function save_(key, obj) {
  const s = JSON.stringify(obj);
  PropertiesService.getScriptProperties().setProperty('dg_' + key, s);
  CacheService.getScriptCache().put('dg_' + key, s, 21600);
}
function sheet_() {
  const ss = CFG.SHEET_ID ? SpreadsheetApp.openById(CFG.SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CFG.SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(CFG.SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#d9ead3');
  }
  return sh;
}
function appendRows_(rows) {
  if (!rows.length) return;
  const sh = sheet_();
  const start = sh.getLastRow();            // 헤더 포함 -> 첫 순번 = start
  const data = rows.map((r, i) => { r[0] = start + i; return r; });
  sh.getRange(start + 1, 1, data.length, HEADERS.length).setValues(data);
}
function now_() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
}
function norm_(s) {
  return String(s == null ? '' : s).replace(/[\s.,!?'"()\[\]]/g, '').toLowerCase();
}
function isCorrect_(Q, ans) {
  const n = norm_(ans);
  return Q.a.some(a => norm_(a) === n);
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function pw_() {
  return PropertiesService.getScriptProperties().getProperty('TEACHER_PW') || CFG.TEACHER_PW_DEFAULT;
}

/* ---------- 클라이언트 호출 API ---------- */
function verifyTeacher(pw) {
  return String(pw) === pw_();
}

function joinGame(name) {
  name = String(name || '').trim();
  if (!name || name.length > 12) return { ok: false, msg: '이름을 1~12자로 입력해 주세요.' };
  return withLock_(() => {
    const m = load_('meta', { phase: 0, q: -1, v: 0 });
    const players = load_('players', {});
    if (!(name in players)) {
      const late = m.q >= 0;
      players[name] = late ? 0 : 1;
      save_('players', players);
      appendRows_([['', now_(), '입장', '', name, '', '', late ? '관전' : '생존', late ? '중간입장(관전)' : '']]);
    }
    return { ok: true };
  });
}

function getState(name) {
  const m = load_('meta', { phase: 0, q: -1, v: 0, correct: null });
  const players = load_('players', {});
  const answers = load_('answers', {});
  const names = Object.keys(players);
  const alive = names.filter(n => players[n] === 1);
  const out = {
    phase: m.phase, q: m.q, total: QUESTIONS.length, v: m.v,
    joined: names.length, alive: alive.length,
    submitted: Object.keys(answers).length,
    correct: m.correct, question: null, answerText: null, champions: null
  };
  if (m.q >= 0 && m.q < QUESTIONS.length && m.phase >= 1 && m.phase <= 4) {
    const Q = QUESTIONS[m.q];
    out.question = { text: Q.q, type: Q.type, choices: Q.choices || null };
    if (m.phase === 4) out.answerText = Q.a[0];
  }
  if (m.phase === 5) out.champions = alive.slice(0, 50);
  if (name) {
    out.me = {
      joined: name in players,
      alive: players[name] === 1,
      submitted: name in answers,
      answer: name in answers ? answers[name] : null
    };
  }
  return out;
}

function submitAnswer(name, ans) {
  return withLock_(() => {
    const m = load_('meta', { phase: 0, q: -1 });
    const players = load_('players', {});
    if (m.phase !== 2) return { ok: false, msg: '지금은 답안을 낼 수 없어요.' };
    if (players[name] !== 1) return { ok: false, msg: '답안을 낼 수 있는 참가자가 아니에요.' };
    ans = String(ans || '').trim().slice(0, 100);
    if (!ans) return { ok: false, msg: '답안을 입력해 주세요.' };
    const answers = load_('answers', {});
    answers[name] = ans;
    save_('answers', answers);
    return { ok: true };
  });
}

/** 선생님 전용: show(문제보기) / open(답안제출가능) / close(답안마감) / check(정답확인) / reset */
function teacherAction(pw, act) {
  if (!verifyTeacher(pw)) return { ok: false, msg: '선생님 권한이 없습니다.' };
  return withLock_(() => {
    const m = load_('meta', { phase: 0, q: -1, v: 0, correct: null });
    const players = load_('players', {});
    const answers = load_('answers', {});

    if (act === 'reset') {
      ['meta', 'players', 'answers'].forEach(k => {
        PropertiesService.getScriptProperties().deleteProperty('dg_' + k);
        CacheService.getScriptCache().remove('dg_' + k);
      });
      appendRows_([['', now_(), '리셋', '', '', '', '', '', '새 대회 시작']]);
      return { ok: true };
    }
    if (act === 'show') {
      if (m.phase !== 0 && m.phase !== 4) return { ok: false, msg: '지금은 문제를 열 수 없어요.' };
      m.q += 1;
      m.correct = null;
      save_('answers', {});
      if (m.q >= QUESTIONS.length) {
        m.phase = 5;
        const champs = Object.keys(players).filter(n => players[n] === 1);
        appendRows_([['', now_(), '종료', '', '', '', champs.length, '', '최종 생존자: ' + champs.join(', ')]]);
      } else {
        m.phase = 1;
      }
    } else if (act === 'open') {
      if (m.phase !== 1) return { ok: false, msg: '먼저 문제보기를 눌러 주세요.' };
      m.phase = 2;
    } else if (act === 'close') {
      if (m.phase !== 2) return { ok: false, msg: '답안 제출 중이 아니에요.' };
      m.phase = 3;
    } else if (act === 'check') {
      if (m.phase !== 3) return { ok: false, msg: '먼저 답안을 마감해 주세요.' };
      const Q = QUESTIONS[m.q];
      const t = now_();
      const rows = [];
      let correct = 0;
      Object.keys(players).forEach(n => {
        if (players[n] !== 1) return;
        const has = n in answers;
        const ok = has && isCorrect_(Q, answers[n]);
        if (ok) correct++; else players[n] = 0;
        rows.push(['', t, '답안', m.q + 1, n, has ? answers[n] : '(미제출)', ok ? 'O' : 'X', ok ? '생존' : '탈락', '']);
      });
      rows.push(['', t, '결과', m.q + 1, '', Q.a[0], correct, '', '정답자 ' + correct + '명 / 문제: ' + Q.q]);
      appendRows_(rows);
      save_('players', players);
      m.correct = correct;
      m.phase = 4;
    } else {
      return { ok: false, msg: '알 수 없는 동작입니다.' };
    }
    m.v = (m.v || 0) + 1;
    save_('meta', m);
    return { ok: true };
  });
}


/* ---------- 학습지 답안 자동저장 (구글 드라이브) ---------- */
/**
 * 학생이 보낸 값(칸 번호 -> 입력값)만 받아 서버의 worksheet 원본에 끼워 넣어 HTML을 새로 만들어 저장합니다.
 * 파일명: 학번_이름_학습지.html (같은 학번·이름이면 덮어쓰기)
 */
function saveWorksheet(p) {
  const fields = (p && p.fields) || {};
  const id = cleanName_(fields['who-id']);
  const name = cleanName_(fields['who-name']);
  if (!id || !name) throw new Error('학번과 이름을 먼저 적어 주세요.');
  const clean = {};
  Object.keys(fields).forEach(function (k) {
    if (/^[A-Za-z0-9-]{1,40}$/.test(k) && typeof fields[k] === 'string') clean[k] = fields[k].slice(0, 2000);
  });
  const html = buildAnswerHtml_(clean);
  const fileName = id + '_' + name + '_학습지.html';
  return withLock_(() => {
    const it = DriveApp.getFoldersByName(CFG.DRIVE_FOLDER_NAME);
    const folder = it.hasNext() ? it.next() : DriveApp.createFolder(CFG.DRIVE_FOLDER_NAME);
    const files = folder.getFilesByName(fileName);
    if (files.hasNext()) { files.next().setContent(html); return { ok: true, fileName: fileName, updated: true }; }
    folder.createFile(fileName, html, MimeType.HTML);
    return { ok: true, fileName: fileName, updated: false };
  });
}
function cleanName_(v) { return String(v || '').replace(/[\\/:*?"<>|\r\n\t]/g, '').trim().slice(0, 20); }
function escHtml_(v) { return String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
function buildAnswerHtml_(values) {
  let html = HtmlService.createHtmlOutputFromFile('worksheet').getContent();
  html = html.replace(/<input([^>]*?)data-k="([^"]+)"([^>]*)>/g, function (m, a, k, b) {
    return values[k] ? '<input' + a + 'data-k="' + k + '"' + b + ' value="' + escHtml_(values[k]) + '">' : m;
  });
  html = html.replace(/(<textarea[^>]*?data-k="([^"]+)"[^>]*>)(<\/textarea>)/g, function (m, open, k, close) {
    return values[k] ? open + escHtml_(values[k]) + close : m;
  });
  return html;
}
