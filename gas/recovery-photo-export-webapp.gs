/**************************************************************
 * "회수사진" 스프레드시트 Web App
 *
 * 배포 위치
 * ------------------------------------------------------------
 * 이 파일은 "회수사진" 스프레드시트(새로 만든 전용 파일, "시트1" 탭
 * 하나만 씀)의 확장 프로그램 > Apps Script 프로젝트에 추가합니다.
 *
 * 하는 일
 * ------------------------------------------------------------
 * - doGet action="spreadsheetUrl": 이 스프레드시트의 편집 URL을
 *   반환합니다("열기" 버튼 전용).
 * - doPost action="exportToSheet1": "일일업무 > 일일회수현황 >
 *   하자사진다운로드" 버튼 전용. body에 rows(회수데이터 웹앱의
 *   action="queryRecoveryDefectPhotos" 응답과 같은 모양 — date/docNo/
 *   name/color/defect/line/detailType/customer/fileId 배열)를 담아
 *   보내면:
 *   1) "시트1" 탭의 기존 내용(텍스트+삽입된 사진)을 전부 지우고
 *   2) 헤더 + rows를 다시 쓰고, 각 행의 fileId로 드라이브에서 사진을
 *      가져와 가로 5.5cm(세로는 원본 비율)로 셀에 삽입한 뒤
 *   3) 이 스프레드시트 전체를 xlsx로 내보내 base64로 반환합니다
 *   (화면은 이 값을 그대로 파일로 저장).
 *
 * 배포 방법
 * ------------------------------------------------------------
 * 1. "회수사진" 스프레드시트 > 확장 프로그램 > Apps Script에 이 파일
 *    추가
 * 2. 배포 > 새 배포 > 유형: 웹 앱, 실행 계정: 나, 액세스 권한: 필요 범위
 * 3. 배포 후 나오는 웹 앱 URL을 대시보드의 "일일업무" 페이지 >
 *    "일일현황관리 데이터" > "상세" > "회수사진 연결"에 입력
 * 4. DriveApp/UrlFetchApp을 처음 쓰는 경우 권한 재승인이 필요할 수
 *    있습니다. 함수 선택 드롭다운에서 testAuth를 선택해 한 번 실행하고
 *    동의 화면을 통과한 뒤 다시 배포하세요.
 *
 * 주의
 * ------------------------------------------------------------
 * - 토큰 검증이 없으므로 URL을 아는 사람은 누구나 이 시트를 읽고
 *   수정할 수 있습니다.
 * - "exportToSheet1"은 매번 "시트1"의 값(텍스트+사진)만 지우고 새로
 *   씁니다. 가운데 정렬이나 열 너비처럼 시트에 직접 맞춰둔 서식은
 *   clearContents()를 써서 그대로 유지합니다(지워지는 건 값뿐). 행
 *   높이는 매번 그 행 사진의 실제 비율에 맞춰 다시 계산해서 설정합니다.
 **************************************************************/

const RECOVERY_PHOTO_SHEET_NAME_ = "시트1";
const RECOVERY_PHOTO_WIDTH_CM_ = 5.5;
const RECOVERY_PHOTO_CM_TO_PX_ = 37.7952755906; // 96dpi 기준

/**************************************************************
 * 권한 재승인용 임시 테스트 함수 — testAuth를 함수 드롭다운에서 골라
 * 한 번 실행하면 DriveApp(사진 blob 읽기)/UrlFetchApp(xlsx 내보내기)
 * 권한 동의 화면이 뜹니다.
 **************************************************************/
function testAuth() {
  UrlFetchApp.fetch("https://www.google.com");
  DriveApp.getRootFolder();
}

function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    const action = params.action || "spreadsheetUrl";

    if (action === "spreadsheetUrl") {
      return jsonOutput_({ url: SpreadsheetApp.getActiveSpreadsheet().getUrl() });
    }

    return jsonOutput_({ error: "알 수 없는 action입니다: " + action });
  } catch (error) {
    return jsonOutput_({ error: error.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const action = body.action || "exportToSheet1";

    if (action === "exportToSheet1") {
      return jsonOutput_(exportToSheet1Action_(body.rows || []));
    }

    return jsonOutput_({ error: "알 수 없는 action입니다: " + action });
  } catch (error) {
    return jsonOutput_({ error: error.message });
  }
}

function jsonOutput_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function normalizeText_(value) {
  return String(value === null || value === undefined ? "" : value).trim();
}

/**************************************************************
 * "시트1" 탭을 비우고(텍스트+삽입 사진 모두) rows를 다시 써넣은 뒤
 * 이 스프레드시트 전체를 xlsx로 내보냅니다.
 **************************************************************/
function exportToSheet1Action_(rows) {
  if (!Array.isArray(rows) || !rows.length) {
    return { ok: false, error: "내려받을 행이 없습니다." };
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RECOVERY_PHOTO_SHEET_NAME_);
  if (!sheet) {
    throw new Error("'" + RECOVERY_PHOTO_SHEET_NAME_ + "' 탭을 찾을 수 없습니다.");
  }

  // "현상상세"는 "회수누적" 탭의 실제 "현상상세" 열(item.defectDetail)
  // 내용이고, "현상분류"가 외관/포장/가공 분류(item.detailType) 열의
  // 내용입니다 — 부적합내역(item.defect)은 둘 다 아니라서 이 엑셀에는
  // 쓰지 않습니다.
  const header = ["회수 날짜", "문서번호", "품명", "색상", "현상상세", "포장라인/원인공정", "현상분류", "고객명", "사진"];

  // clearContents()는 셀 값만 지우고 가운데 정렬/열 너비 같은 서식은
  // 그대로 남깁니다(시트1에 직접 맞춰둔 서식을 유지하기 위해 clear()
  // 대신 이걸 씁니다). 삽입된(떠 있는) 사진은 clearContents()로도
  // 안 지워지므로 따로 하나씩 지웁니다. 1행(헤더) 높이/음영은 지우기
  // 전에 미리 읽어뒀다가 다 쓴 뒤에 그대로 다시 적용해서 고정합니다.
  const headerRowHeight = sheet.getRowHeight(1);
  const headerBackgrounds = sheet.getRange(1, 1, 1, header.length).getBackgrounds();

  sheet.getImages().forEach(function(img) { img.remove(); });
  sheet.clearContents();

  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.setRowHeight(1, headerRowHeight);
  sheet.getRange(1, 1, 1, header.length).setBackgrounds(headerBackgrounds);

  // "사진" 열 너비는 시트1에 직접 맞춰둔 값을 그대로 쓰고(강제로
  // 덮어쓰지 않음), 사진 자체 표시 크기만 아래에서 고정폭(5.5cm)으로
  // 맞춥니다.
  const photoColumn = header.length; // "사진" 열
  const photoWidthPx = Math.round(RECOVERY_PHOTO_WIDTH_CM_ * RECOVERY_PHOTO_CM_TO_PX_);

  rows.forEach(function(item, rowPos) {
    const rowNum = rowPos + 2;
    sheet.getRange(rowNum, 1, 1, 8).setValues([[
      normalizeText_(item.date), normalizeText_(item.docNo), normalizeText_(item.name),
      normalizeText_(item.color), normalizeText_(item.defectDetail), normalizeText_(item.line),
      normalizeText_(item.detailType), normalizeText_(item.customer)
    ]]);

    let photoHeightPx = photoWidthPx; // 비율을 못 구하면 정사각형 기본값으로
    const fileId = normalizeText_(item.fileId);
    if (fileId) {
      try {
        const blob = DriveApp.getFileById(fileId).getBlob();
        const image = sheet.insertImage(blob, photoColumn, rowNum);
        const naturalWidth = image.getWidth();
        const naturalHeight = image.getHeight();
        photoHeightPx = naturalWidth > 0 ? Math.round(photoWidthPx * naturalHeight / naturalWidth) : photoWidthPx;
        image.setWidth(photoWidthPx);
        image.setHeight(photoHeightPx);
      } catch (imgErr) {
        sheet.getRange(rowNum, photoColumn).setValue("사진을 불러오지 못했습니다: " + imgErr);
      }
    }

    sheet.setRowHeight(rowNum, photoHeightPx + 10);
  });

  // 시트1에 미리 맞춰둔 서식이 이번에 받은 행 수보다 좁은 범위에만
  // 있었을 수 있어서(건수가 매번 달라짐), 매번 실제로 쓴 범위
  // 전체(헤더+데이터)에 가로/세로 가운데 정렬을 명시적으로 다시
  // 맞춰 고정합니다.
  const totalRowCount = rows.length + 1; // 헤더 포함
  sheet.getRange(1, 1, totalRowCount, header.length)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  SpreadsheetApp.flush();

  const base64 = exportSpreadsheetAsXlsxBase64_(SpreadsheetApp.getActiveSpreadsheet().getId());
  const fileName = "하자사진다운로드_" + Utilities.formatDate(new Date(), "Asia/Seoul", "yyyyMMdd_HHmm") + ".xlsx";

  return { ok: true, matchedCount: rows.length, fileName: fileName, base64: base64 };
}

/**************************************************************
 * 이 스프레드시트 자신을 xlsx로 내보냅니다(이 스크립트 자신의 OAuth
 * 토큰으로 export 엔드포인트를 호출).
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
