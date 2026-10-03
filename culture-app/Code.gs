/**
 * <사회> 수업 주제: 문화의 속성 - 유쾌한 소정샘
 * Code.gs (Google Apps Script 백엔드)
 */

// ===== 설정 =====
const SPREADSHEET_ID = '';      // 비워두면 스크립트가 연결된 시트(컨테이너 바인딩)를 사용
const SHEET_NAME = '기록';
const DRIVE_FOLDER_NAME = '문화의 속성 학습지 제출';   // 학습지 답안이 저장될 드라이브 폴더(없으면 자동 생성)
const HEADERS = ['순번', '제출일시', '이름', '시도회차', '정답수', '총문항', '점수(%)', '소요시간(초)', '상세결과'];

// Gemini API는 이번 버전에서 사용하지 않음 (추후 확장용 자리)
// 키는 코드에 직접 쓰지 말고 [프로젝트 설정 > 스크립트 속성]에 GEMINI_API_KEY 로 저장하세요.

function doGet(e) {
  // ?page=worksheet → 수업 학습지(worksheet.html)
  if (e && e.parameter && e.parameter.page === 'worksheet') {
    return HtmlService.createHtmlOutputFromFile('worksheet')
      .setTitle('문화의 속성 찾기 학습지')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('문화의 속성 - 유쾌한 소정샘')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** 화면에서 수업 학습지 링크 주소를 얻기 위해 호출 */
function getWorksheetUrl() {
  return ScriptApp.getService().getUrl() + '?page=worksheet';
}

function getSheet_() {
  const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#e8f0fe');
    sh.setFrozenRows(1);
  }
  return sh;
}

/**
 * 게임 결과 저장 (한 번 시도 = 한 행)
 * @param {{name:string, correct:number, total:number, seconds:number, detail:string}} r
 * @return {{ok:boolean, attempt:number, row:number}}
 */
function saveResult(r) {
  const name = String((r && r.name) || '').trim().slice(0, 30);
  if (!name) throw new Error('이름이 없습니다.');
  const total = Number(r.total) || 10;
  const correct = Math.max(0, Math.min(total, Number(r.correct) || 0));

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet_();
    const last = sh.getLastRow();
    let attempt = 1;
    if (last > 1) {
      const names = sh.getRange(2, 3, last - 1, 1).getValues();
      attempt = names.filter(function (x) { return String(x[0]).trim() === name; }).length + 1;
    }
    const row = [
      last,                                   // 순번 (헤더 제외 다음 번호)
      new Date(),
      name,
      attempt,
      correct,
      total,
      Math.round(correct / total * 100),
      Math.max(0, Number(r.seconds) || 0),
      String(r.detail || '').slice(0, 2000)
    ];
    sh.appendRow(row);
    sh.getRange(sh.getLastRow(), 2).setNumberFormat('yyyy-MM-dd HH:mm:ss');
    return { ok: true, attempt: attempt, row: sh.getLastRow() };
  } finally {
    lock.releaseLock();
  }
}


/**
 * 수업 학습지 답안을 구글 드라이브에 저장
 * - 학생이 보낸 값(칸 번호 → 입력값)만 받아서, 서버에 있는 worksheet 원본에 끼워 넣어 HTML을 새로 만듭니다.
 *   (학생이 보낸 HTML을 그대로 저장하지 않으므로 안전합니다.)
 * - 파일은 스크립트 실행 계정(교사) 드라이브의 DRIVE_FOLDER_NAME 폴더에 '학번_이름_학습지.html'로 저장되고,
 *   같은 학번·이름으로 다시 저장하면 그 파일을 덮어씁니다.
 * @param {{fields: Object<string,string>}} p
 * @return {{ok:boolean, fileName:string, updated:boolean}}
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

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);
    const files = folder.getFilesByName(fileName);
    if (files.hasNext()) {
      files.next().setContent(html);
      return { ok: true, fileName: fileName, updated: true };
    }
    folder.createFile(fileName, html, MimeType.HTML);
    return { ok: true, fileName: fileName, updated: false };
  } finally {
    lock.releaseLock();
  }
}

function cleanName_(v) {
  return String(v || '').replace(/[\\/:*?"<>|\r\n\t]/g, '').trim().slice(0, 20);
}

function escHtml_(v) {
  return String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** worksheet 원본 HTML의 입력칸(data-k)에 학생 답을 채워 넣은 HTML 문자열을 만든다 */
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
