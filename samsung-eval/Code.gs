/**
 * 삼성중학교 학교평가·사전 의제 설문 응답 저장용 Google Apps Script
 *
 * [설정 방법]
 * 1) 응답을 모을 구글 스프레드시트를 새로 만든다. (비공개 유지, 교육연구부 계정 소유 권장)
 * 2) 시트에서 [확장 프로그램 → Apps Script]를 열고 이 코드를 붙여넣는다.
 * 3) [배포 → 새 배포 → 유형: 웹 앱]
 *      - 실행 사용자: 나
 *      - 액세스 권한: 모든 사용자
 *    배포 후 나오는 '웹 앱 URL'을 복사한다.
 * 4) index.html 상단 CONFIG.submitUrl 에 그 URL을 붙여넣고 다시 배포한다.
 *
 * 응답은 '학교평가' / '사전의제' 시트에 한 줄씩 쌓이며, 개인 식별 정보는 수집하지 않는다.
 * 코드를 수정한 뒤에는 [배포 관리 → 새 버전]으로 다시 배포해야 반영된다.
 */

const EVAL_HEADER = ['제출시각', '구분', '1.교육과정', '2.기초학력', '3.인성', '4.학교폭력·교육활동보호', '5.학습확장',
  '6.진로', '7.디지털', '8.학습으로의평가', '9.예술체육', '10.독서', '11.만족스러운 활동', '12.추가 희망 활동'];
const AGENDA_HEADER = ['제출시각', '구분', '선택 주제', '바라는 학교', '추가 의제'];

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const now = new Date();
    if (d.type === 'eval') {
      const sh = getSheet_(ss, '학교평가', EVAL_HEADER);
      const a = (d.answers || []).slice(0, 10);
      sh.appendRow([now, d.role, ...a, clean_(d.opinion1), clean_(d.opinion2)]);
    } else if (d.type === 'agenda') {
      const sh = getSheet_(ss, '사전의제', AGENDA_HEADER);
      sh.appendRow([now, d.role, (d.topics || []).join(' / '), clean_(d.wish), clean_(d.extra)]);
    } else {
      return json_({ ok: false, error: 'unknown type' });
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet() {
  return json_({ ok: true, message: '삼성중학교 설문 응답 수집 중' });
}

function getSheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
  }
  return sh;
}

// 수식 주입 방지 (=, +, -, @ 로 시작하는 입력 무력화) 및 길이 제한
function clean_(v) {
  v = String(v || '').slice(0, 1000);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
