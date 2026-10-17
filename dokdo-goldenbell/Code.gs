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
  { type: 'ox', q: '독도는 동도와 서도, 단 2개의 섬으로만 이루어져 있다.', choices: ['O', 'X'], a: ['X'], exp: "독도는 동도와 서도 및 89개의 부속 도서로 이루어져 있습니다." },
  { type: 'ox', q: '독도는 우리나라 가장 동쪽 끝에 위치한 섬이다.', choices: ['O', 'X'], a: ['O'] },
  { type: 'ox', q: '울릉도에서는 날씨가 맑아도 독도를 눈으로 볼 수 없다.', choices: ['O', 'X'], a: ['X'], exp: "울릉도에서 독도까지의 거리는 약 87.4km이며, 맑은 날에는 육안으로 볼 수 있습니다." },
  { type: 'ox', q: '독도의 두 섬 중 동도가 서도보다 면적이 더 넓고 높이도 더 높다.', choices: ['O', 'X'], a: ['X'], exp: "서도(면적 88,740㎡, 높이 168.5m)가 동도(면적 73,297㎡, 높이 98.6m)보다 더 넓고 높습니다." },
  { type: 'ox', q: '독도는 해저 화산 활동으로 솟아오른 화산섬이다.', choices: ['O', 'X'], a: ['O'] },
  { type: 'ox', q: "독도의 옛 이름 중 하나인 '석도(石島)'는 '돌섬'이라는 우리말을 한자로 표기한 것이다.", choices: ['O', 'X'], a: ['O'] },
  { type: 'ox', q: '독도에 사는 대표적인 새는 고양이 울음소리를 내는 괭이갈매기이다.', choices: ['O', 'X'], a: ['O'] },
  { type: 'ox', q: '독도는 행정구역상 강원도 울릉군에 속한다.', choices: ['O', 'X'], a: ['X'], exp: "독도는 경상북도 울릉군에 속합니다." },
  { type: 'ox', q: '오늘이 독도의 날이다.', choices: ['O', 'X'], a: ['X'], exp: "독도의 날은 10월 25일입니다." },
  { type: 'ox', q: '『삼국사기』에는 512년 신라 이사부 장군이 우산국을 정벌했다는 기록이 있으며, 우산국은 오늘날의 울릉도와 독도를 포함한다.', choices: ['O', 'X'], a: ['O'] },
  { type: 'ox', q: '1454년에 편찬된 『세종실록』 「지리지」에는 우산도와 울릉도가 날씨가 맑으면 서로 보인다고 기록되어 있다.', choices: ['O', 'X'], a: ['O'] },
  { type: 'ox', q: "1900년 대한제국은 칙령 제41호를 통해 독도를 '죽도(竹島)'라고 칭하며 울도군의 관할 구역으로 정했다.", choices: ['O', 'X'], a: ['X'], exp: "대한제국 칙령 제41호는 죽도와 '석도(石島)'를 관할 구역으로 정했으며, 여기서 석도가 독도를 가리킵니다." },
  { type: 'ox', q: "1877년 일본의 최고 행정기관인 태정관은 '울릉도 외 1도(독도)는 일본과 관계없다'는 지령을 내렸다.", choices: ['O', 'X'], a: ['O'] },
  { type: 'ox', q: '일본은 1905년 러일전쟁 중에 독도를 주인 없는 땅이라며 시마네현에 불법으로 편입시켰다.', choices: ['O', 'X'], a: ['O'] },
  { type: 'ox', q: '1906년 울도 군수 심흥택은 일본 관리들의 독도 시찰에 대해 \'본군 소속 독도가...\'라는 내용의 보고서를 올리며 우리 영토임을 분명히 했다.', choices: ['O', 'X'], a: ['O'] },
  { type: 'ox', q: '조선 후기, 안용복은 일본으로 건너가 울릉도와 독도가 일본 땅이 아님을 확인받고 돌아왔다.', choices: ['O', 'X'], a: ['O'] },
  { type: 'ox', q: "2005년 일본 시마네현이 '다케시마의 날' 조례를 제정하자, 경상북도는 이에 맞서 10월을 '독도의 달'로 지정했다.", choices: ['O', 'X'], a: ['O'] },
  { type: 'ox', q: "1904년 일본 군함 '니타카호'의 행동 일지에는 한국인들이 독도를 '우산도'라고 부른다는 기록이 있다.", choices: ['O', 'X'], a: ['X'], exp: "'독도(獨島)'라고 쓴다는 기록이 있습니다." },
  { type: 'ox', q: '1951년 샌프란시스코 강화조약은 일본이 한국의 독립을 인정하고 한국에 대한 모든 권리를 포기한다고 규정하여, 독도가 국제법적으로 한국 영토임을 재확인했다.', choices: ['O', 'X'], a: ['O'] },
  { type: 'ox', q: "독도 주변 해역에는 미래 에너지 자원으로 불리는 '메탄 하이드레이트'가 다량 매장되어 있다.", choices: ['O', 'X'], a: ['O'] },
  { type: 'ox', q: '18세기 조선에서 제작된 「동국대지도」에는 울릉도와 독도가 한 쌍의 섬으로 표시되어 있다.', choices: ['O', 'X'], a: ['O'] },
  { type: 'ox', q: '독도에 대한 일본의 영유권 주장은 역사적 사실에 근거한 정당한 권리 행사이다.', choices: ['O', 'X'], a: ['X'], exp: "우리 정부는 일본의 주장이 침략의 역사에 근거한 불법적인 것이라고 보고 있습니다." },
  { type: 'ox', q: "독도의 가장 오래된 이름은 『삼국사기』에 나오는 '우산도(于山島)'이다.", choices: ['O', 'X'], a: ['O'] },
  { type: 'ox', q: "1870년 일본 외무성 관리의 보고서인 「조선국 교제시말 내탐서」에는 '울릉도와 독도가 조선의 부속이 된 경위'가 언급되어 있다.", choices: ['O', 'X'], a: ['O'] },
  { type: 'ox', q: "일본은 독도를 '다케시마(竹島)'라고 부르는데, 이는 '돌섬'이라는 뜻이다.", choices: ['O', 'X'], a: ['X'], exp: "다케시마는 '대나무 섬'이라는 뜻입니다." },
  { type: 'ox', q: '독도를 알리기 위한 노력은 정부 기관에서만 이루어지고 있다.', choices: ['O', 'X'], a: ['X'], exp: "정부뿐만 아니라 동북아역사재단 같은 연구기관, 반크, 푸른울릉독도가꾸기회 같은 시민단체 등에서도 활발히 활동하고 있습니다." },
  { type: 'ox', q: '프랑스에서 1737년에 제작된 「조선왕국전도」에도 울릉도와 독도가 한 쌍의 섬으로 나란히 표시되어 있다.', choices: ['O', 'X'], a: ['O'] },
  { type: 'ox', q: '『동국문헌비고』(1770)는 \'울릉과 우산은 모두 우산국 땅\'이라고 기록하고 있다.', choices: ['O', 'X'], a: ['O'] },
  { type: 'ox', q: '1999년 「독도 등 도서지역의 생태계 보전에 관한 특별법」에 따라 독도는 특정도서로 지정되었다.', choices: ['O', 'X'], a: ['O'] },
  { type: 'ox', q: '일본은 17세기 중반에 이미 독도의 영유권을 확립했다고 주장하며, 태정관 지령(1877년)이 이를 확인했다고 말한다.', choices: ['O', 'X'], a: ['X'], exp: "태정관 지령(1877년)은 오히려 '울릉도 외 1도(독도)는 일본과 관계없다'고 확인한 문서입니다." },
  { type: 'short', q: "1696년 안용복이 일본에 도착했을 때, 일본 관리가 그를 조사한 기록이 2005년에 발견되었습니다. 이 문서에는 안용복이 울릉도와 독도가 OOO 소속임을 밝혔다는 내용이 기록되어 있습니다. 당시 울릉도와 독도가 속해 있던 OOO는 어디일까요?", a: ["강원도"] },
  { type: 'short', q: "독도는 동도와 서도 중 어느 섬의 면적이 더 넓을까요?", a: ["서도"] },
  { type: 'short', q: "독도에 가장 많이 서식하며, 고양이 울음소리를 낸다고 하여 이름 붙여진 대표적인 새는 무엇일까요?", a: ["괭이갈매기"] },
  { type: 'short', q: "신라 지증왕 13년(서기 512년), 우산국을 정벌하여 신라 영토에 포함시킨 장군의 이름은 누구일까요?", a: ["이사부"] },
  { type: 'short', q: "조선 숙종 때 일본으로 두 번 건너가 울릉도와 독도가 조선의 땅임을 확인받은 인물은 누구일까요?", a: ["안용복"] },
  { type: 'short', q: "독도의 옛 이름 중 '돌섬'을 한자로 표기한 이름은 무엇일까요?", a: ["석도"], show: "석도(石島)" },
  { type: 'short', q: "1951년 체결된 이 조약은 일본이 한국의 독립을 인정하고 한국에 대한 모든 권리를 포기한다고 규정하여, 독도가 국제법적으로 한국 영토임을 재확인했습니다. 이 조약은 무엇일까요?", a: ["샌프란시스코강화조약", "샌프란시스코조약", "샌프란시스코평화조약"], show: "샌프란시스코 강화조약" },
  { type: 'multi', q: "1946년 연합국 최고사령관 각서 제677호는 일본의 통치 및 행정 범위에서 특정 섬들을 제외시켰습니다. 이 각서에서 일본의 영토에서 제외된 우리나라의 주요 섬 세 곳은 어디인가요? (쉼표로 구분)", must: ["울릉도", "독도", "제주도"], a: ['울릉도, 독도, 제주도'] },
  { type: 'short', q: "울릉도에서는 날씨가 맑은 날 육안으로 독도를 볼 수 있습니다. 그렇다면 울릉도에서 독도까지의 거리는 약 몇 km일까요?", a: ["87.4"], show: "87.4km" }
];
const OX_COUNT = 30;   // 앞의 30문항은 OX, 이후는 주관식

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
  const safe = v => (typeof v === 'string' && /^[=+\-@]/.test(v)) ? "'" + v : v;   // 학생 답이 수식으로 실행되지 않게
  const data = rows.map((r, i) => { r[0] = start + i; return r.map(safe); });
  sh.getRange(start + 1, 1, data.length, HEADERS.length).setValues(data);
}
function now_() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
}
function norm_(s) {
  return String(s == null ? '' : s).replace(/\(.*?\)/g, '').replace(/[\s.,!?'"\[\]]/g, '').replace(/킬로미터|km|약/gi, '').toLowerCase();
}
function isCorrect_(Q, ans) {
  const n = norm_(ans);
  if (Q.type === 'multi') return Q.must.every(w => n.indexOf(norm_(w)) > -1);
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
    const isOx = m.q < OX_COUNT;
    out.question = { text: Q.q, type: Q.type, choices: Q.choices || null,
      sec: isOx ? 'OX 퀴즈' : '주관식', no: isOx ? m.q + 1 : m.q - OX_COUNT + 1, secTotal: isOx ? OX_COUNT : QUESTIONS.length - OX_COUNT };
    if (m.phase === 4) { out.answerText = Q.show || Q.a[0]; out.explanation = Q.exp || ''; }
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
      rows.push(['', t, '결과', m.q + 1, '', Q.show || Q.a[0], correct, '', '정답자 ' + correct + '명 / 문제: ' + Q.q]);
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
