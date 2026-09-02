/**************************************************************
 * 품질 고객클레임 대시보드(마감자료 자동화) 연동용 Web App
 *
 * 배포 위치
 * ------------------------------------------------------------
 * 이 파일은 목적지 스프레드시트
 * (1lr-KbRyQ3zmcnT-RvRS2-krVnMl1cGndHRxEVpD6ZWM)의 Apps Script
 * 프로젝트에 추가합니다. claim-dashboard-monthly-automation.gs(마감자료
 * 자동화 코드: CONFIG, updateCurrentMonthWeeklyCounts, reshapeForNewMonth,
 * getTargetSheet_ 등)와 같은 프로젝트에 함께 있어야 합니다.
 *
 * 하는 일
 * ------------------------------------------------------------
 * - doGet: CONFIG.TARGET_SHEET_GID 시트를 PDF로 내보내 base64로
 *   돌려줍니다(읽기 전용 화면 표시용). 표가 브랜드별로 반복되는
 *   병합 셀 레이아웃이라 일반 표 형태로 옮기면 깨지기 때문에,
 *   시트를 그대로 이미지처럼 보여주는 방식을 씁니다.
 * - doPost action="updateWeeklyCounts": updateCurrentMonthWeeklyCounts(sourceWebappUrl)
 *   실행 (이번달 주간 데이터 채우기). body에 sourceWebappUrl("26년 마감자료"
 *   웹앱 URL, 대시보드가 매번 같이 보내줌)을 받아 그대로 넘겨줍니다 —
 *   이 스크립트 자신은 그 URL을 저장해두지 않습니다.
 * - doPost action="reshapeMonth": reshapeForNewMonth() 실행
 *   (지난달 접기 + 이번달 펼치기 — 열 삽입/삭제가 있어 되돌리기 어려움)
 * - doPost action="exportFull": 이 스프레드시트 파일 전체(모든 시트)를
 *   xlsx로 내보내 base64로 반환 ("누적표엑셀다운로드" 버튼)
 *
 * 배포 방법
 * ------------------------------------------------------------
 * 1. 목적지 스프레드시트 > 확장 프로그램 > Apps Script에 이 파일 추가
 *    (기존 자동화 코드 파일은 그대로 두고, 새 파일로 추가)
 * 2. 배포 > 새 배포 > 유형: 웹 앱, 실행 계정: 나, 액세스 권한: 필요 범위
 * 3. 배포 후 나오는 웹 앱 URL을 대시보드의 "클레임 대시보드" 연결
 *    정보에 입력
 * 4. UrlFetchApp을 처음 쓰는 경우 권한 재승인이 필요할 수 있습니다.
 *    함수 선택 드롭다운에서 testAuth를 선택해 한 번 실행하고 동의
 *    화면을 통과한 뒤 다시 배포하세요.
 *
 * 주의
 * ------------------------------------------------------------
 * - 토큰 검증이 없으므로 URL을 아는 사람은 누구나 이 시트를 읽고
 *   이번달 데이터 채우기/월 전환을 실행할 수 있습니다.
 * - reshapeMonth는 열을 삽입/삭제하는 구조 변경 작업입니다. 화면에서도
 *   확인창을 띄우지만, 실행 전 원본 스프레드시트 상태를 한 번
 *   확인하는 걸 권장합니다.
 **************************************************************/

/**************************************************************
 * 권한 재승인용 임시 테스트 함수
 *
 * doGet이 쓰는 UrlFetchApp(외부 요청) 권한을 승인받기 위한 함수입니다.
 * 이름에 밑줄(_)이 없어야 Apps Script 편집기의 "실행할 함수" 드롭다운에
 * 보입니다. 드롭다운에서 testAuth를 선택해 실행하면 동의 화면이
 * 뜹니다 — 승인한 뒤에는 이 함수를 지우고 다시 배포해도 되고, 그냥
 * 남겨둬도 동작에는 영향이 없습니다.
 **************************************************************/
function testAuth() {
  UrlFetchApp.fetch("https://www.google.com");
}


function doGet(e) {
  try {
    const params = (e && e.parameter) || {};

    if (params.action === "spreadsheetUrl") {
      return jsonOutput_({ url: SpreadsheetApp.getActiveSpreadsheet().getUrl() });
    }

    const sheet = getTargetSheet_();
    const pdfBase64 = exportSheetAsPdfBase64_(sheet);

    return jsonOutput_({ sheet: sheet.getName(), pdfBase64: pdfBase64 });
  } catch (error) {
    return jsonOutput_({ error: error.message });
  }
}


function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");

    if (body.action === "updateWeeklyCounts") {
      updateCurrentMonthWeeklyCounts(body.sourceWebappUrl);
      return jsonOutput_({ ok: true });
    }

    if (body.action === "reshapeMonth") {
      reshapeForNewMonth();
      return jsonOutput_({ ok: true });
    }

    if (body.action === "exportFull") {
      return jsonOutput_(exportFullWorkbookAction_());
    }

    return jsonOutput_({ error: "알 수 없는 action입니다: " + body.action });
  } catch (error) {
    return jsonOutput_({ error: error.message });
  }
}


/**************************************************************
 * "누적표엑셀다운로드" 액션
 *
 * 이 스프레드시트 파일 자체를 통째로 xlsx로 내보냅니다(모든 시트
 * 포함). base64로 인코딩해서 돌려주면 화면에서 파일로 저장합니다.
 **************************************************************/
function exportFullWorkbookAction_() {
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const base64 = exportSpreadsheetAsXlsxBase64_(spreadsheetId);
  const fileName = "월현황누적표_" +
    Utilities.formatDate(new Date(), "Asia/Seoul", "yyyyMMdd_HHmm") + ".xlsx";

  return { ok: true, fileName: fileName, base64: base64 };
}


/**************************************************************
 * 스프레드시트 ID로 xlsx 내보내기 → base64 문자열로 반환
 * (이 스크립트 자신의 OAuth 토큰으로 export 엔드포인트를 호출합니다)
 **************************************************************/
function exportSpreadsheetAsXlsxBase64_(spreadsheetId) {
  const url = "https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/export?format=xlsx";

  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error("엑셀 내보내기에 실패했습니다 (" + response.getResponseCode() + ")");
  }

  return Utilities.base64Encode(response.getBlob().getBytes());
}


/**************************************************************
 * 시트 하나를 PDF로 내보내 base64 문자열로 반환합니다
 * (이 스크립트 자신의 OAuth 토큰으로 export 엔드포인트를 호출합니다).
 * gid로 대상 시트 한 장만 지정하고, 격자선/시트 이름/페이지 번호는
 * 빼고 가로 폭에 맞춰 한 페이지에 들어가도록 요청합니다.
 **************************************************************/
function exportSheetAsPdfBase64_(sheet) {
  const spreadsheetId = sheet.getParent().getId();

  const params = [
    "format=pdf",
    "gid=" + sheet.getSheetId(),
    "portrait=false",
    "size=A4",
    "fitw=true",
    "gridlines=false",
    "printtitle=false",
    "sheetnames=false",
    "pagenumbers=false",
    "fzr=false"
  ].join("&");

  const url = "https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/export?" + params;

  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error("PDF 내보내기에 실패했습니다 (" + response.getResponseCode() + ")");
  }

  return Utilities.base64Encode(response.getBlob().getBytes());
}


function jsonOutput_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
