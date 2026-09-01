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
 *   CLEAN_TARGET_COLUMN_LABELS 순서로 열만 뽑아 "정리" 시트에
 *   덮어씁니다(기존 내용은 전부 지움). 현재공정이 빈 값인 행은
 *   포장 단계로 보고 "포장"을 채워 넣습니다. 자세한 내용은
 *   cleanServiceShipmentListAction_ 참고.
 *
 * [예정] "정리결과" 버튼에 쓸 doGet 액션은 아직 없습니다(요청 시 추가
 * 예정 — 지금은 대시보드 버튼만 만들어둔 상태).
 *
 * 배포 방법
 * ------------------------------------------------------------
 * 1. 서비스출고건 스프레드시트 > 확장 프로그램 > Apps Script에 이 파일 추가
 * 2. 배포 > 새 배포 > 유형: 웹 앱, 실행 계정: 나, 액세스 권한: 필요 범위
 * 3. 배포 후 나오는 웹 앱 URL을 대시보드의 "일일업무" 페이지 >
 *    "일일현황관리 데이터" > "상세" > "서비스출고건 연결"에 입력
 *
 * 주의
 * ------------------------------------------------------------
 * - 토큰 검증이 없으므로 URL을 아는 사람은 누구나 이 시트를 읽을 수 있습니다.
 **************************************************************/

const SOURCE_SHEET_NAME = "시트1"; // 대시보드가 기본으로 읽는 탭
const CLEAN_RESULT_SHEET_NAME = "정리";

// "리스트 정리" 결과에 이 순서로만 열을 남깁니다.
const CLEAN_TARGET_COLUMN_LABELS = [
  "포장라인", "계획량", "최초포장일", "부품이동카드번호",
  "자재코드", "자재색상", "부품명", "현재공정", "출고지"
];

// 현재공정이 빈 값이면 아직 포장 단계인 것으로 보고 이 값을 채워 넣습니다.
const CURRENT_PROCESS_COLUMN_LABEL = "현재공정";
const CURRENT_PROCESS_DEFAULT_VALUE = "포장";


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

    return jsonOutput_({ error: "알 수 없는 action입니다: " + action });
  } catch (error) {
    return jsonOutput_({ error: error.message });
  }
}


/**************************************************************
 * "리스트 정리" 액션
 *
 * "시트1"(ERP 긴급생산 리스트를 변환해 붙여넣은 원본)에서
 * CLEAN_TARGET_COLUMN_LABELS 순서로 열만 뽑아 "정리" 시트에
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

  const columnIndexes = CLEAN_TARGET_COLUMN_LABELS.map(function(label) {
    const idx = header.indexOf(label);
    if (idx === -1) throw new Error("'" + SOURCE_SHEET_NAME + "'에서 '" + label + "' 열을 찾을 수 없습니다.");
    return idx;
  });

  const currentProcessPos = CLEAN_TARGET_COLUMN_LABELS.indexOf(CURRENT_PROCESS_COLUMN_LABEL);

  const rows = source.rows
    .map(function(row) { return columnIndexes.map(function(idx) { return row.values[idx]; }); })
    .filter(function(values) {
      return values.some(function(value) { return normalizeText_(value) !== ""; });
    });

  rows.forEach(function(values) {
    if (normalizeText_(values[currentProcessPos]) === "") {
      values[currentProcessPos] = CURRENT_PROCESS_DEFAULT_VALUE;
    }
  });

  const resultSheet = getOrCreateSheet_(ss, CLEAN_RESULT_SHEET_NAME);
  resultSheet.clear();
  resultSheet.getRange(1, 1, 1, CLEAN_TARGET_COLUMN_LABELS.length).setValues([CLEAN_TARGET_COLUMN_LABELS]);

  if (rows.length) {
    resultSheet.getRange(2, 1, rows.length, CLEAN_TARGET_COLUMN_LABELS.length).setValues(rows);
  }

  return { ok: true, resultSheet: CLEAN_RESULT_SHEET_NAME, rowCount: rows.length };
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
