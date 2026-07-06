/**
 * 쿠팡 이익 계산기 — 구글시트 누적 저장용 Apps Script
 *
 * [설치 방법]
 * 1. 구글 드라이브에서 빈 스프레드시트를 하나 만듭니다.
 * 2. 상단 메뉴 [확장 프로그램] > [Apps Script] 클릭.
 * 3. 기본 코드를 모두 지우고 이 파일 내용을 통째로 붙여넣습니다.
 * 4. (선택) 아래 TOKEN 에 아무 비밀문자열을 넣으면, 앱에도 같은 값을 넣어야 저장됩니다. (비워두면 누구나 이 주소로 저장 가능)
 * 5. 우측 상단 [배포] > [새 배포] > 유형 '웹 앱' 선택.
 *      - 실행 계정: 나
 *      - 액세스 권한: '모든 사용자'
 * 6. [배포] 클릭 → 나오는 '웹 앱 URL'을 복사해서 계산기 앱의 '구글시트 주소'에 붙여넣습니다.
 *
 * 코드를 수정하면 [배포] > [배포 관리] 에서 기존 배포를 '수정'해 새 버전으로 올려야 반영됩니다.
 */

const TOKEN = ''; // 예: 'mysecret123' (비워두면 토큰 검사 안 함)

const SUMMARY_SHEET = '일별요약';
const DETAIL_SHEET = '상세';
const SUMMARY_HEADER = ['날짜', '판매수량', '인사이트매출', '정산매출', '사용광고비', '마진(광고전)', '순이익', '순이익률'];
const DETAIL_HEADER = ['날짜', '상품구분', '상품명', '등록상품ID', '판매수량', '인사이트매출', '정산매출', '사용광고비', '마진(광고전)', '순이익', '순이익률'];

function doGet() {
  return json_({ ok: true, service: 'coupang-margin' });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (TOKEN && String(body.token || '') !== TOKEN) {
      return json_({ ok: false, error: 'invalid token' });
    }
    const days = body.days || {};
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sum = getOrCreate_(ss, SUMMARY_SHEET, SUMMARY_HEADER);
    const det = getOrCreate_(ss, DETAIL_SHEET, DETAIL_HEADER);
    const dates = Object.keys(days);
    dates.forEach(function (date) {
      const d = days[date];
      upsertByDate_(sum, date, d.summary ? [d.summary] : []);
      upsertByDate_(det, date, d.rows || []);
    });
    sortByDate_(sum);
    sortByDate_(det);
    // 숫자 서식(쉼표·%) 적용
    formatSheet_(sum, { comma: [2, 3, 4, 5, 6, 7], pct: 8 });
    formatSheet_(det, { comma: [5, 6, 7, 8, 9, 10], pct: 11, id: 4 });
    return json_({ ok: true, saved: dates });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function getOrCreate_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(header);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
  }
  return sh;
}

function upsertByDate_(sh, date, rows) {
  const last = sh.getLastRow();
  if (last > 1) {
    const vals = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = vals.length - 1; i >= 0; i--) {
      if (String(vals[i][0]) === String(date)) sh.deleteRow(i + 2);
    }
  }
  if (rows && rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
}

function sortByDate_(sh) {
  const last = sh.getLastRow();
  if (last > 2) sh.getRange(2, 1, last - 1, sh.getLastColumn()).sort({ column: 1, ascending: true });
}

// 숫자 서식 적용: comma=천단위, pct=백분율, id=구분자 없는 정수(등록상품ID)
function formatSheet_(sh, spec) {
  const last = sh.getLastRow();
  if (last < 2) return;
  const n = last - 1;
  (spec.comma || []).forEach(function (c) { sh.getRange(2, c, n, 1).setNumberFormat('#,##0'); });
  if (spec.pct) sh.getRange(2, spec.pct, n, 1).setNumberFormat('0.0%');
  if (spec.id) sh.getRange(2, spec.id, n, 1).setNumberFormat('0');
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
