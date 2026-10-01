/**
 * <사회> 수업 주제: 문화의 속성 - 삼성중학교 소정샘
 * Code.gs (Google Apps Script 백엔드)
 */

// ===== 설정 =====
const SPREADSHEET_ID = '';      // 비워두면 스크립트가 연결된 시트(컨테이너 바인딩)를 사용
const SHEET_NAME = '기록';
const HEADERS = ['순번', '제출일시', '이름', '시도회차', '정답수', '총문항', '점수(%)', '소요시간(초)', '상세결과'];

// Gemini API는 이번 버전에서 사용하지 않음 (추후 확장용 자리)
// 키는 코드에 직접 쓰지 말고 [프로젝트 설정 > 스크립트 속성]에 GEMINI_API_KEY 로 저장하세요.

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('문화의 속성 - 삼성중학교 소정샘')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
