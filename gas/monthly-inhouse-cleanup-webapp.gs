/**************************************************************
 * "월마감 내작건정리" 스프레드시트 조회용 Web App
 *
 * 배포 위치
 * ------------------------------------------------------------
 * 이 파일은 "월마감 내작건정리"(마감된 월마감 내작건 자료를 옮겨 담는)
 * 스프레드시트의 확장 프로그램 > Apps Script 프로젝트에 새로 추가합니다.
 *
 * 하는 일
 * ------------------------------------------------------------
 * - doGet action="sheets": 이 파일의 모든 탭 이름 + gid + 행/열 수 목록
 *   (구조 파악용).
 * - doGet action="read"(기본값): sheet/gid 파라미터로 지정한 탭을
 *   header+rows로 반환합니다. 둘 다 없으면 "시트1"을 읽습니다.
 * - doGet action="spreadsheetUrl": 이 스프레드시트의 편집 URL을
 *   반환합니다. 대시보드의 "내작월마감 자료교체" 버튼이 확인창을
 *   띄운 뒤 이 스프레드시트를 새 탭에서 열 때 씁니다(마감된 월마감
 *   자료로 시트1을 바꾸는 실제 작업은 그 스프레드시트에서 직접 함).
 * - doPost action="clean": "자료정리" 버튼 액션. "시트1"에서
 *   SOURCE_COLUMN_LABELS 순서로 열만 뽑고 맨 앞에 번호를 새로 매겨
 *   "정리" 시트에 덮어씁니다(기존 내용은 전부 지움). 포장 값 기준으로
 *   정렬하고, 포장 값이 바뀌는 경계마다 빈 행을 하나 끼워 넣습니다.
 *   자세한 내용은 cleanMonthlyInhouseListAction_ 참고.
 * - doPost action="exportResult": "정리파일다운로드" 버튼 액션.
 *   "정리" 시트 하나만 담은 xlsx를 base64로 반환합니다.
 *
 * 배포 방법
 * ------------------------------------------------------------
 * 1. 월마감 내작건정리 스프레드시트 > 확장 프로그램 > Apps Script에 이 파일 추가
 * 2. 배포 > 새 배포 > 유형: 웹 앱, 실행 계정: 나, 액세스 권한: 필요 범위
 * 3. 배포 후 나오는 웹 앱 URL을 대시보드의 "월간업무" 페이지 >
 *    "월간업무 데이터" > "상세" > "월마감 내작건정리 연결"에 입력
 * 4. UrlFetchApp/DriveApp을 처음 쓰는 경우(exportResult) 권한 재승인이
 *    필요할 수 있습니다. 함수 선택 드롭다운에서 testAuth를 선택해 한 번
 *    실행하고 동의 화면을 통과한 뒤 다시 배포하세요.
 *
 * 주의
 * ------------------------------------------------------------
 * - 토큰 검증이 없으므로 URL을 아는 사람은 누구나 이 시트를 읽을 수 있습니다.
 **************************************************************/

const SOURCE_SHEET_NAME = "시트1"; // 대시보드가 기본으로 읽는 탭
const CLEAN_RESULT_SHEET_NAME = "정리";

// "자료정리" 결과에 이 순서로 열을 남깁니다(번호는 시트1에 없는 열로,
// 정리하면서 순차적으로 새로 매깁니다).
const NUMBER_COLUMN_LABEL = "번호";
const SOURCE_COLUMN_LABELS = [
  "브랜드", "지역센터", "접수번호", "구분", "형태", "포장",
  "제품코드", "색상", "수량", "금액", "하자상세", "로트"
];

// 정렬/그룹 구분 기준 열 — 이 값이 바뀌는 경계마다 빈 행을 하나 끼워 넣습니다.
const PACKAGE_COLUMN_LABEL = "포장";


/**************************************************************
 * 권한 재승인용 임시 테스트 함수
 *
 * exportResult가 쓰는 UrlFetchApp(외부 요청)/DriveApp 권한을 승인받기
 * 위한 함수입니다. 이름에 밑줄(_)이 없어야 Apps Script 편집기의
 * "실행할 함수" 드롭다운에 보입니다. 드롭다운에서 testAuth를 선택해
 * 실행하면 동의 화면이 뜹니다 — 승인한 뒤에는 이 함수를 지우고 다시
 * 배포해도 되고, 그냥 남겨둬도 동작에는 영향이 없습니다.
 **************************************************************/
function testAuth() {
  UrlFetchApp.fetch("https://www.google.com");

  const tempFile = DriveApp.createFile("월마감내작건정리_권한테스트_임시파일.txt", "test");
  tempFile.setTrashed(true);
}


function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    const action = params.action || "read";

    if (action === "sheets") {
      return jsonOutput_({ sheets: listSheetsInfo_() });
    }

    if (action === "read") {
      const sheet = resolveSheet_(params.sheet || SOURCE_SHEET_NAME, params.gid);
      return jsonOutput_(readSheetObject_(sheet));
    }

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
    const action = body.action || "";

    if (action === "clean") {
      return jsonOutput_(cleanMonthlyInhouseListAction_());
    }

    if (action === "exportResult") {
      return jsonOutput_(exportSheetsSubsetAsXlsxBase64_([CLEAN_RESULT_SHEET_NAME], "월마감내작건정리"));
    }

    return jsonOutput_({ error: "알 수 없는 action입니다: " + action });
  } catch (error) {
    return jsonOutput_({ error: error.message });
  }
}


/**************************************************************
 * "자료정리" 액션
 *
 * "시트1"에서 SOURCE_COLUMN_LABELS 순서로 열만 뽑아, 포장 값 기준으로
 * 정렬한 뒤 맨 앞에 1부터 순차적으로 번호를 매겨 "정리" 시트에
 * 덮어씁니다(기존 내용은 전부 지우고 다시 씀). 완전히 빈 행은
 * 건너뜁니다. 포장 값이 바로 앞 행과 달라지는 경계마다(첫 행 제외)
 * 빈 행을 하나 끼워 넣어서 포장별로 시각적으로 구분되게 합니다.
 **************************************************************/
function cleanMonthlyInhouseListAction_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);

  if (!sourceSheet) {
    throw new Error("'" + SOURCE_SHEET_NAME + "' 시트를 찾을 수 없습니다.");
  }

  const source = readSheetObject_(sourceSheet);
  const header = source.header;

  const columnIndexes = SOURCE_COLUMN_LABELS.map(function(label) {
    const idx = header.indexOf(label);
    if (idx === -1) throw new Error("'" + SOURCE_SHEET_NAME + "'에서 '" + label + "' 열을 찾을 수 없습니다.");
    return idx;
  });

  const packageColPos = SOURCE_COLUMN_LABELS.indexOf(PACKAGE_COLUMN_LABEL);

  const rows = source.rows
    .map(function(row) { return columnIndexes.map(function(idx) { return row.values[idx]; }); })
    .filter(function(values) {
      return values.some(function(value) { return normalizeText_(value) !== ""; });
    });

  rows.sort(function(a, b) {
    return normalizeText_(a[packageColPos])
      .localeCompare(normalizeText_(b[packageColPos]), "ko", { numeric: true, sensitivity: "base" });
  });

  const finalHeader = [NUMBER_COLUMN_LABEL].concat(SOURCE_COLUMN_LABELS);
  const blankRow = finalHeader.map(function() { return ""; });
  const outputRows = [];
  let sequence = 0;
  let previousPackage = null;

  rows.forEach(function(values) {
    const packageValue = normalizeText_(values[packageColPos]);

    if (previousPackage !== null && packageValue !== previousPackage) {
      outputRows.push(blankRow.slice());
    }
    previousPackage = packageValue;

    sequence++;
    outputRows.push([sequence].concat(values));
  });

  const resultSheet = getOrCreateSheet_(ss, CLEAN_RESULT_SHEET_NAME);
  resultSheet.clear();
  resultSheet.getRange(1, 1, 1, finalHeader.length).setValues([finalHeader]);

  if (outputRows.length) {
    resultSheet.getRange(2, 1, outputRows.length, finalHeader.length).setValues(outputRows);
  }

  return { ok: true, resultSheet: CLEAN_RESULT_SHEET_NAME, rowCount: rows.length };
}


/**************************************************************
 * 이 스프레드시트에서 sheetNames에 해당하는 시트만 임시 스프레드시트에
 * 복사해서 xlsx로 내보낸 뒤, 임시 스프레드시트는 지웁니다
 * ("정리파일다운로드" 버튼에서 사용).
 **************************************************************/
function exportSheetsSubsetAsXlsxBase64_(sheetNames, fileNamePrefix) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempSpreadsheet = SpreadsheetApp.create(fileNamePrefix + "_임시");
  const tempId = tempSpreadsheet.getId();

  try {
    sheetNames.forEach(function(name) {
      const sourceSheet = ss.getSheetByName(name);
      if (!sourceSheet) throw new Error("'" + name + "' 시트를 찾을 수 없습니다.");
      const copied = sourceSheet.copyTo(tempSpreadsheet);
      copied.setName(name);
    });

    tempSpreadsheet.getSheets().forEach(function(sheet) {
      if (sheetNames.indexOf(sheet.getName()) === -1) {
        tempSpreadsheet.deleteSheet(sheet);
      }
    });

    SpreadsheetApp.flush();

    const base64 = exportSpreadsheetAsXlsxBase64_(tempId);
    const fileName = fileNamePrefix + "_" +
      Utilities.formatDate(new Date(), "Asia/Seoul", "yyyyMMdd_HHmm") + ".xlsx";

    return { ok: true, fileName: fileName, base64: base64 };
  } finally {
    DriveApp.getFileById(tempId).setTrashed(true);
  }
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
 * 이름의 시트를 찾아서 반환하고, 없으면 새로 만듭니다.
 **************************************************************/
function getOrCreateSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}


function normalizeText_(value) {
  return String(value === null || value === undefined ? "" : value).trim();
}


/**************************************************************
 * 이 스프레드시트의 모든 탭 이름 + gid + 행/열 수 목록
 **************************************************************/
function listSheetsInfo_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(function(sheet) {
    return {
      name: sheet.getName(),
      gid: sheet.getSheetId(),
      lastRow: sheet.getLastRow(),
      lastColumn: sheet.getLastColumn()
    };
  });
}


/**************************************************************
 * sheet 이름 또는 gid로 시트를 찾습니다. 둘 다 없으면 첫 번째 탭을
 * 씁니다. 못 찾으면 에러를 던집니다.
 **************************************************************/
function resolveSheet_(sheetName, gidParam) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error("'" + sheetName + "' 시트를 찾을 수 없습니다.");
    return sheet;
  }

  const sheets = ss.getSheets();

  if (gidParam !== undefined && gidParam !== null && gidParam !== "") {
    const gid = Number(gidParam);

    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId() === gid) return sheets[i];
    }

    throw new Error("gid " + gid + "에 해당하는 시트를 찾을 수 없습니다.");
  }

  if (!sheets.length) {
    throw new Error("이 스프레드시트에 시트가 없습니다.");
  }

  return sheets[0];
}


/**************************************************************
 * 시트 전체를 header + 행 배열(rowIndex 포함)로 반환합니다.
 * 1행을 헤더로 봅니다. 데이터가 없으면 빈 결과를 반환합니다.
 **************************************************************/
function readSheetObject_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 1 || lastColumn < 1) {
    return { sheet: sheet.getName(), gid: sheet.getSheetId(), header: [], rows: [] };
  }

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const header = values[0];
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    rows.push({ rowIndex: i + 1, values: values[i] });
  }

  return { sheet: sheet.getName(), gid: sheet.getSheetId(), header: header, rows: rows };
}


function jsonOutput_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
