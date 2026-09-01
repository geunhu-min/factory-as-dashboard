/**************************************************************
 * "서비스출고건" 스프레드시트 조회용 Web App
 *
 * 배포 위치
 * ------------------------------------------------------------
 * 이 파일은 "서비스출고건"(ERP 긴급생산 리스트를 옮겨 담는) 스프레드시트의
 * 확장 프로그램 > Apps Script 프로젝트에 새로 추가합니다.
 *
 * 하는 일
 * ------------------------------------------------------------
 * - doGet action="sheets": 이 파일의 모든 탭 이름 + gid + 행/열 수 목록
 *   (구조 파악용).
 * - doGet action="read"(기본값): sheet/gid 파라미터로 지정한 탭을
 *   header+rows로 반환합니다. 둘 다 없으면 "시트1"을 읽습니다.
 * - doGet action="spreadsheetUrl": 이 스프레드시트의 편집 URL을
 *   반환합니다. 대시보드의 "서비스출고건 리스트교체" 버튼이 확인창을
 *   띄운 뒤 이 스프레드시트를 새 탭에서 열 때 씁니다(ERP 긴급생산
 *   리스트를 변환해 시트1에 붙여넣는 실제 작업은 그 스프레드시트에서
 *   직접 함).
 * - doPost action="clean": "리스트 정리" 버튼 액션. "시트1"에서
 *   CLEAN_FIXED_COLUMN_LABELS(+값이 있는 Route 열) 순서로 열만 뽑아 "정리" 시트에
 *   덮어씁니다(기존 내용은 전부 지움). 현재공정이 빈 값인 행은
 *   포장 단계로 보고 "포장"을 채워 넣고, 최초포장일 열은 "yyyy-mm-dd"
 *   형식으로 맞춥니다. 자세한 내용은 cleanServiceShipmentListAction_ 참고.
 * - doPost action="save": 대시보드의 "정리결과" 표에서 직접 수정한
 *   행(현재공정만 편집 가능)을 "정리" 시트에 저장합니다. body에
 *   rows(각 행마다 loaded/values)를 담아 보내면, 저장 직전 시트의
 *   실제 값과 loaded를 비교해서 그 사이 다른 곳에서 먼저 바뀐 행은
 *   충돌로 보고 덮어쓰지 않습니다.
 * - doPost action="deleteRows": "삭제" 버튼 액션. body에 담긴
 *   rowIndexes에 해당하는 행을 "정리" 시트에서 지웁니다.
 * - doPost action="exportResult": "정리자료다운로드" 버튼 액션.
 *   "정리" 시트 하나만 담은 xlsx를 base64로 반환합니다.
 *
 * 배포 방법
 * ------------------------------------------------------------
 * 1. 서비스출고건 스프레드시트 > 확장 프로그램 > Apps Script에 이 파일 추가
 * 2. 배포 > 새 배포 > 유형: 웹 앱, 실행 계정: 나, 액세스 권한: 필요 범위
 * 3. 배포 후 나오는 웹 앱 URL을 대시보드의 "일일업무" 페이지 >
 *    "일일현황관리 데이터" > "상세" > "서비스출고건 연결"에 입력
 * 4. UrlFetchApp/DriveApp을 처음 쓰는 경우(exportResult) 권한 재승인이
 *    필요할 수 있습니다. 함수 선택 드롭다운에서 testAuth를 선택해 한 번
 *    실행하고 동의 화면을 통과한 뒤 다시 배포하세요.
 *
 * 주의
 * ------------------------------------------------------------
 * - 토큰 검증이 없으므로 URL을 아는 사람은 누구나 이 시트를 읽고
 *   수정할 수 있습니다.
 **************************************************************/

const SOURCE_SHEET_NAME = "시트1"; // 대시보드가 기본으로 읽는 탭
const CLEAN_RESULT_SHEET_NAME = "정리";

// "리스트 정리" 결과에 이 순서로만 열을 남깁니다.
const CLEAN_FIXED_COLUMN_LABELS = [
  "포장라인", "계획량", "최초포장일", "부품이동카드번호",
  "자재코드", "자재색상", "부품명", "현재공정", "출고지"
];

// 출고지 다음의 Route 열은 값이 있는 열까지만 남깁니다(예: 전체 행에서
// Route1~Route2까지만 값이 있으면 Route3~5는 결과에서 아예 뺌).
const ROUTE_COLUMN_LABELS = ["Route1", "Route2", "Route3", "Route4", "Route5"];

// 현재공정이 빈 값이면 아직 포장 단계인 것으로 보고 이 값을 채워 넣습니다.
const CURRENT_PROCESS_COLUMN_LABEL = "현재공정";
const CURRENT_PROCESS_DEFAULT_VALUE = "포장";

// 날짜로 표시할 열 이름과 날짜 표시 형식(시트1의 원본 표시와 맞춤 —
// 그냥 두면 "정리" 시트가 기본 서식으로 "2026. 8. 28"처럼 나옴)
const DATE_FORMAT_COLUMN_LABELS = ["최초포장일"];
const DATE_NUMBER_FORMAT = "yyyy-mm-dd";


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

  const tempFile = DriveApp.createFile("서비스출고건_권한테스트_임시파일.txt", "test");
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
      return jsonOutput_(cleanServiceShipmentListAction_());
    }

    if (action === "save") {
      const sheet = getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), CLEAN_RESULT_SHEET_NAME);
      return jsonOutput_(saveRowsAction_(sheet, body.rows || []));
    }

    if (action === "deleteRows") {
      return jsonOutput_(deleteServiceShipmentResultRowsAction_(body.rowIndexes || []));
    }

    if (action === "exportResult") {
      return jsonOutput_(exportSheetsSubsetAsXlsxBase64_([CLEAN_RESULT_SHEET_NAME], "서비스출고건_정리"));
    }

    return jsonOutput_({ error: "알 수 없는 action입니다: " + action });
  } catch (error) {
    return jsonOutput_({ error: error.message });
  }
}


/**************************************************************
 * "리스트 정리" 액션
 *
 * "시트1"(ERP 긴급생산 리스트를 변환해 붙여넣은 원본)에서
 * CLEAN_FIXED_COLUMN_LABELS(+값이 있는 Route 열) 순서로 열만 뽑아 "정리" 시트에
 * 덮어씁니다(기존 내용은 전부 지우고 다시 씀). 모든 대상 열이 빈
 * 행(완전히 빈 행)은 건너뜁니다. 현재공정 열이 빈 값인 행은 아직
 * 포장 단계인 것으로 보고 "포장"을 채워 넣습니다.
 **************************************************************/
function cleanServiceShipmentListAction_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);

  if (!sourceSheet) {
    throw new Error("'" + SOURCE_SHEET_NAME + "' 시트를 찾을 수 없습니다.");
  }

  const source = readSheetObject_(sourceSheet);
  const header = source.header;

  const fixedColumnIndexes = CLEAN_FIXED_COLUMN_LABELS.map(function(label) {
    const idx = header.indexOf(label);
    if (idx === -1) throw new Error("'" + SOURCE_SHEET_NAME + "'에서 '" + label + "' 열을 찾을 수 없습니다.");
    return idx;
  });

  // Route 열은 시트1에 아예 없어도 에러 내지 않고 그냥 건너뜁니다(있는
  // 것만 후보로 삼고, 그중에서도 값이 있는 데까지만 최종 결과에 남김).
  const routeColumnEntries = ROUTE_COLUMN_LABELS
    .map(function(label) { return { label: label, idx: header.indexOf(label) }; })
    .filter(function(entry) { return entry.idx !== -1; });

  const allColumnIndexes = fixedColumnIndexes.concat(
    routeColumnEntries.map(function(entry) { return entry.idx; })
  );

  const currentProcessPos = CLEAN_FIXED_COLUMN_LABELS.indexOf(CURRENT_PROCESS_COLUMN_LABEL);

  const allRows = source.rows
    .map(function(row) { return allColumnIndexes.map(function(idx) { return row.values[idx]; }); })
    .filter(function(values) {
      return values.some(function(value) { return normalizeText_(value) !== ""; });
    });

  allRows.forEach(function(values) {
    if (normalizeText_(values[currentProcessPos]) === "") {
      values[currentProcessPos] = CURRENT_PROCESS_DEFAULT_VALUE;
    }
  });

  // Route1부터 순서대로 보면서, 값이 있는 마지막 Route 열까지만 남깁니다
  // (예: Route1~2만 값이 있으면 Route3~5는 결과 열 자체에서 뺌).
  const fixedCount = CLEAN_FIXED_COLUMN_LABELS.length;
  let usedRouteCount = 0;
  routeColumnEntries.forEach(function(entry, i) {
    const pos = fixedCount + i;
    const hasValue = allRows.some(function(values) { return normalizeText_(values[pos]) !== ""; });
    if (hasValue) usedRouteCount = i + 1;
  });

  const finalHeader = CLEAN_FIXED_COLUMN_LABELS.concat(
    routeColumnEntries.slice(0, usedRouteCount).map(function(entry) { return entry.label; })
  );
  const rows = allRows.map(function(values) { return values.slice(0, finalHeader.length); });

  const resultSheet = getOrCreateSheet_(ss, CLEAN_RESULT_SHEET_NAME);
  resultSheet.clear();
  resultSheet.getRange(1, 1, 1, finalHeader.length).setValues([finalHeader]);

  if (rows.length) {
    resultSheet.getRange(2, 1, rows.length, finalHeader.length).setValues(rows);

    DATE_FORMAT_COLUMN_LABELS.forEach(function(label) {
      const colIndex = finalHeader.indexOf(label);
      if (colIndex !== -1) {
        resultSheet.getRange(2, colIndex + 1, rows.length, 1).setNumberFormat(DATE_NUMBER_FORMAT);
      }
    });
  }

  return { ok: true, resultSheet: CLEAN_RESULT_SHEET_NAME, rowCount: rows.length };
}


/**************************************************************
 * 화면에서 수정한 행을 저장합니다("정리" 시트, 현재공정 열만 편집 가능).
 *
 * rows: [{ rowIndex, loaded: [...], values: [...] }]
 * 저장 직전 시트의 실제 값이 loaded와 다르면(다른 곳에서 먼저 수정된
 * 경우) 그 행은 덮어쓰지 않고 충돌로 표시합니다.
 **************************************************************/
function saveRowsAction_(sheet, rows) {
  const lastColumn = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const conflicts = [];
  let updatedCount = 0;

  rows.forEach(function(row) {
    const rowIndex = Number(row.rowIndex);

    if (!rowIndex || rowIndex < 2 || rowIndex > lastRow) {
      conflicts.push(rowIndex);
      return;
    }

    const currentValues = sheet.getRange(rowIndex, 1, 1, lastColumn).getValues()[0];
    const loaded = row.loaded || [];

    const matchesLoaded = currentValues.every(function(cell, idx) {
      return normalizeText_(cell) === normalizeText_(loaded[idx]);
    });

    if (!matchesLoaded) {
      conflicts.push(rowIndex);
      return;
    }

    const newValues = normalizeRowLength_(row.values || [], lastColumn);
    sheet.getRange(rowIndex, 1, 1, lastColumn).setValues([newValues]);
    updatedCount++;
  });

  const latest = readSheetObject_(sheet);

  return {
    sheet: latest.sheet,
    gid: latest.gid,
    header: latest.header,
    rows: latest.rows,
    conflicts: conflicts,
    updatedCount: updatedCount
  };
}


/**************************************************************
 * "삭제" 버튼 액션 — rowIndexes에 해당하는 행을 "정리" 시트에서 지웁니다.
 **************************************************************/
function deleteServiceShipmentResultRowsAction_(rowIndexes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CLEAN_RESULT_SHEET_NAME);
  if (!sheet) throw new Error("'" + CLEAN_RESULT_SHEET_NAME + "' 시트를 찾을 수 없습니다.");

  const uniqueDescending = Array.from(new Set(rowIndexes.map(Number)))
    .filter(function(rowIndex) { return rowIndex >= 2; })
    .sort(function(a, b) { return b - a; });

  uniqueDescending.forEach(function(rowIndex) {
    sheet.deleteRow(rowIndex);
  });

  return { ok: true, deletedCount: uniqueDescending.length };
}


function normalizeRowLength_(row, targetColumnCount) {
  const result = row.slice(0, targetColumnCount);

  while (result.length < targetColumnCount) {
    result.push("");
  }

  return result;
}


/**************************************************************
 * 이 스프레드시트에서 sheetNames에 해당하는 시트만 임시 스프레드시트에
 * 복사해서 xlsx로 내보낸 뒤, 임시 스프레드시트는 지웁니다
 * ("정리자료다운로드" 버튼에서 사용).
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


// 날짜 값(최초포장일 등)은 Date 객체와, Date가 JSON 직렬화될 때 나오는
// ISO 형식 문자열이 서로 다르게 보여서(저장 시 충돌로 잘못 감지되는 걸
// 막기 위해) 둘 다 "yyyy-MM-dd"로 맞춰서 비교합니다.
function normalizeText_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, "Asia/Seoul", "yyyy-MM-dd");
  }
  const text = String(value === null || value === undefined ? "" : value).trim();
  const isoDateMatch = text.match(/^(\d{4}-\d{2}-\d{2})T\d{2}:\d{2}:\d{2}/);
  if (isoDateMatch) return isoDateMatch[1];
  return text;
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
