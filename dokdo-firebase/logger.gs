/**
 * 독도 골든벨 - 구글 시트 기록 + 학습지 Apps Script (Firebase 버전용)
 * 게임 진행은 Firebase 가 맡고, 이 스크립트는 ①결과를 "기록" 시트에 남기고 ②학습지를 보여주고 드라이브에 자동저장합니다.
 *
 * [설정] 프로젝트 설정 > 스크립트 속성에 LOG_SECRET 을 추가하고,
 *        seed.json 의 private > log > secret 과 같은 값을 넣으세요.
 */
const CFG = { SHEET_ID: '', SHEET_NAME: '기록', DRIVE_FOLDER_NAME: '독도 학습지 제출' };
const HEADERS = ['순번', '일시', '구분', '문제번호', '이름', '제출답안', '정답여부', '생존여부', '비고'];

function doGet(e) {
  if (e && e.parameter && e.parameter.page === 'worksheet') {
    return HtmlService.createHtmlOutputFromFile('worksheet').setTitle('독도 학습지')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return ContentService.createTextOutput('독도 골든벨 기록 서버 정상 동작 중');
}

/** 선생님 화면이 보내는 기록: { secret, rows: [[순번, 일시, 구분, 문제번호, 이름, 제출답안, 정답여부, 생존여부, 비고], ...] } */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'worksheet') {            // 학생 학습지 자동저장 (GitHub Pages 에서 호출)
      try { return jsonOut_(saveWorksheet({ fields: body.fields })); }
      catch (err) { return jsonOut_({ ok: false, msg: String(err.message || err) }); }
    }
    const secret = PropertiesService.getScriptProperties().getProperty('LOG_SECRET');
    if (!secret || body.secret !== secret) return ContentService.createTextOutput('denied');
    const rows = (body.rows || []).slice(0, 400).map(function (r) {
      const out = [];
      for (let i = 0; i < HEADERS.length; i++) out.push(safe_(r[i]));
      return out;
    });
    appendRows_(rows);
    return ContentService.createTextOutput('ok');
  } catch (err) {
    return ContentService.createTextOutput('error');
  }
}

function jsonOut_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * [처음 한 번 실행] 설정 점검 + 권한 승인
 * 편집기 위쪽 함수 선택에서 `setup` 을 고르고 ▶ 실행을 누르세요.
 * 권한 창이 뜨면 허용하고, 아래 "실행 로그"에 점검 결과(✅/❌)가 표시됩니다.
 */
function setup() {
  const log = [];
  try {
    const ss = CFG.SHEET_ID ? SpreadsheetApp.openById(CFG.SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('시트에 연결된 스크립트가 아닙니다. 구글 시트에서 [확장 프로그램 > Apps Script]로 열거나 CFG.SHEET_ID 에 시트 ID를 넣으세요.');
    sheet_();
    log.push('✅ 구글 시트 연결 + "' + CFG.SHEET_NAME + '" 시트 확인: ' + ss.getName());
  } catch (e) { log.push('❌ 구글 시트: ' + e.message); }
  try {
    const it = DriveApp.getFoldersByName(CFG.DRIVE_FOLDER_NAME);
    if (!it.hasNext()) DriveApp.createFolder(CFG.DRIVE_FOLDER_NAME);
    log.push('✅ 드라이브 폴더 확인: ' + CFG.DRIVE_FOLDER_NAME);
  } catch (e) { log.push('❌ 드라이브: ' + e.message); }
  try {
    const sec = PropertiesService.getScriptProperties().getProperty('LOG_SECRET');
    log.push(sec ? '✅ LOG_SECRET 설정됨 (길이 ' + sec.length + ')' : '❌ LOG_SECRET 이 없습니다. 프로젝트 설정 > 스크립트 속성에 추가하세요.');
  } catch (e) { log.push('❌ 스크립트 속성: ' + e.message); }
  try {
    HtmlService.createHtmlOutputFromFile('worksheet').getContent();
    log.push('✅ 학습지 파일(worksheet) 확인');
  } catch (e) { log.push('❌ HTML 파일 이름이 정확히 worksheet 인지 확인하세요: ' + e.message); }
  log.push('웹앱 주소: ' + (ScriptApp.getService().getUrl() || '(아직 배포 전)'));
  Logger.log(log.join('\n'));
  return log.join('\n');
}

/**
 * [시험] 드라이브 저장이 되는지 확인: 함수 선택에서 testSave 를 고르고 ▶ 실행
 * 성공하면 드라이브 "독도 학습지 제출" 폴더에 00000_테스트_학습지.html 파일이 생깁니다.
 */
function testSave() {
  const r = saveWorksheet({ fields: { 'who-id': '00000', 'who-name': '테스트', 'a-month': '10', 'a-day': '25' } });
  Logger.log(JSON.stringify(r));
  return r;
}

/** 웹 주소로 직접 접속했을 때(GET) 보이는 상태 확인 문구 */
/** 시트에서 수식으로 실행되지 않도록(= + - @ 로 시작하는 글자) 앞에 ' 를 붙임 */
function safe_(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number') return v;
  const s = String(v).slice(0, 500);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function sheet_() {
  const ss = CFG.SHEET_ID ? SpreadsheetApp.openById(CFG.SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CFG.SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(CFG.SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#f4cce0');
  }
  return sh;
}
function appendRows_(rows) {
  if (!rows.length) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const start = sh.getLastRow();
    rows.forEach(function (r, i) { r[0] = start + i; });
    sh.getRange(start + 1, 1, rows.length, HEADERS.length).setValues(rows);
  } finally { lock.releaseLock(); }
}

/* ---------- 학습지 답안 자동저장 (구글 드라이브) ---------- */
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
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const it = DriveApp.getFoldersByName(CFG.DRIVE_FOLDER_NAME);
    const folder = it.hasNext() ? it.next() : DriveApp.createFolder(CFG.DRIVE_FOLDER_NAME);
    const files = folder.getFilesByName(fileName);
    if (files.hasNext()) { files.next().setContent(html); return { ok: true, fileName: fileName, updated: true }; }
    folder.createFile(fileName, html, MimeType.HTML);
    return { ok: true, fileName: fileName, updated: false };
  } finally { lock.releaseLock(); }
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
