/**************************************************************
 * "회수데이터"(주간업무 회수누적 원본데이터) 스프레드시트 조회/편집용
 * Web App
 *
 * 배포 위치
 * ------------------------------------------------------------
 * 이 파일은 "회수데이터"(주간업무 회수누적 원본데이터) 스프레드시트의
 * 확장 프로그램 > Apps Script 프로젝트에 추가합니다(다른 파일과
 * 마찬가지로 그 스프레드시트에 바인딩된 스크립트라, 이 파일 안에서는
 * 스프레드시트 ID를 몰라도 SpreadsheetApp.getActiveSpreadsheet()로
 * 항상 그 파일 자신을 가리킵니다).
 *
 * 하는 일
 * ------------------------------------------------------------
 * 아직 이 시트의 실제 탭 구조/열 이름이나 "회수데이터 누적저장" 버튼이
 * 정확히 어떤 로직으로 누적해야 하는지 정해지지 않은 상태라, 우선
 * 일반적인 조회/편집/내보내기와 스프레드시트 열기만 되는 뼈대만
 * 만들어뒀습니다(daily-packaging-webapp.gs와 같은 방식). 대시보드의
 * "회수데이터 누적저장" 버튼은 지금은 action="spreadsheetUrl"만 써서
 * 이 스프레드시트를 새 탭으로 여는 데만 씁니다 — 실제 누적 저장 로직은
 * 나중에 별도로 요청받으면 채워 넣을 예정입니다.
 *
 * - doGet action="sheets": 이 파일의 모든 탭 이름 + gid + 행/열 수 목록
 *   (어떤 탭을 볼지 고르기 전에, 구조 파악용으로 먼저 호출)
 * - doGet action="read" (기본값): 탭 하나를 header + rows로 반환합니다.
 *   sheet 파라미터(탭 이름) 또는 gid 파라미터로 지정, 둘 다 없으면
 *   첫 번째 탭을 씁니다. 1행을 헤더로 보고 그 아래부터 데이터 행으로
 *   취급합니다(병합 셀이 있으면 병합된 첫 칸에만 값이 들어오고
 *   나머지는 빈 값으로 옵니다).
 * - doGet action="spreadsheetUrl": 이 스프레드시트의 편집 URL을
 *   반환합니다("회수데이터 누적저장" 버튼이 새 탭에서 열 때 씀 —
 *   화면 쪽에는 스프레드시트 ID를 저장해두지 않으므로 매번 물어봄).
 * - doPost action="save": 화면에서 수정한 행을 저장합니다. 행마다
 *   loaded(불러왔을 때 값)와 values(수정한 값)를 같이 보내면, 저장
 *   직전 시트의 실제 값과 loaded를 비교해서 그 사이에 다른 곳에서
 *   먼저 바뀐 행은 충돌로 보고 덮어쓰지 않습니다(그 행은 최신 값
 *   그대로 돌려줌). 응답에 최신 header/rows/conflicts/updatedCount를
 *   담아 돌려주므로, 화면은 그 값으로 다시 그리면 됩니다.
 * - doPost action="exportFull": 이 스프레드시트 파일 전체(모든 시트)를
 *   xlsx로 내보내 base64로 반환합니다(범용 백업용, 화면 버튼과는
 *   아직 연결되어 있지 않음).
 *
 * 배포 방법
 * ------------------------------------------------------------
 * 1. "회수데이터"(주간업무 회수누적 원본데이터) 스프레드시트 >
 *    확장 프로그램 > Apps Script에 이 파일 추가
 * 2. 배포 > 새 배포 > 유형: 웹 앱, 실행 계정: 나, 액세스 권한: 필요 범위
 * 3. 배포 후 나오는 웹 앱 URL을 대시보드의 "일일업무" 페이지 >
 *    "일일현황관리 데이터" > "상세" > "회수데이터 연결"에 입력
 * 4. UrlFetchApp/DriveApp을 처음 쓰는 경우(exportFull) 권한 재승인이
 *    필요할 수 있습니다. 함수 선택 드롭다운에서 testAuth를 선택해 한
 *    번 실행하고 동의 화면을 통과한 뒤 다시 배포하세요.
 *
 * 주의
 * ------------------------------------------------------------
 * - 토큰 검증이 없으므로 URL을 아는 사람은 누구나 이 시트를 읽고
 *   수정할 수 있습니다.
 * - 병합 셀이 많은 복잡한 표는 header/rows 구조로 옮기면 레이아웃이
 *   깨질 수 있습니다. 실제 탭 구조를 "sheets"/"read"로 먼저 확인한
 *   뒤, 필요하면 이 파일의 읽기/쓰기 로직을 그 구조에 맞게 고쳐야
 *   합니다.
 **************************************************************/

/**************************************************************
 * 권한 재승인용 임시 테스트 함수
 *
 * exportFull이 쓰는 UrlFetchApp(외부 요청), DriveApp의 파일 생성/삭제
 * (휴지통 이동) 권한을 승인받기 위한 함수입니다. 이름에 밑줄(_)이
 * 없어야 Apps Script 편집기의 "실행할 함수" 드롭다운에 보입니다.
 * 드롭다운에서 testAuth를 선택해 실행하면 동의 화면이 뜹니다 —
 * 승인한 뒤에는 이 함수를 지우고 다시 배포해도 되고, 그냥 남겨둬도
 * 동작에는 영향이 없습니다.
 **************************************************************/
function testAuth() {
  UrlFetchApp.fetch("https://www.google.com");

  const tempFile = DriveApp.createFile("회수데이터_권한테스트_임시파일.txt", "test");
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
      const sheet = resolveSheet_(params.sheet, params.gid);
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
    const action = body.action || "save";

    if (action === "save") {
      const sheet = resolveSheet_(body.sheet, body.gid);
      return jsonOutput_(saveRowsAction_(sheet, body.rows || []));
    }

    if (action === "exportFull") {
      return jsonOutput_(exportFullWorkbookAction_());
    }

    return jsonOutput_({ error: "알 수 없는 action입니다: " + action });
  } catch (error) {
    return jsonOutput_({ error: error.message });
  }
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
 * 이 스프레드시트의 모든 탭 이름 + gid + 대략적인 크기 목록
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


/**************************************************************
 * 화면에서 수정한 행을 저장합니다.
 *
 * rows: [{ rowIndex, loaded: [...], values: [...] }]
 * - loaded: 화면이 불러왔을 때의 원래 값(저장 직전 실시간 값과 비교용)
 * - values: 화면에서 수정한 최종 값
 *
 * 저장 직전 시트의 실제 값이 loaded와 다르면(다른 곳에서 먼저 수정된
 * 경우) 그 행은 덮어쓰지 않고 충돌로 표시합니다. 저장 후 최신
 * header/rows와 함께 conflicts(rowIndex 목록), updatedCount를
 * 반환합니다.
 **************************************************************/
function saveRowsAction_(sheet, rows) {
  const lastColumn = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const conflicts = [];
  let updatedCount = 0;

  rows.forEach(function(row) {
    const rowIndex = Number(row.rowIndex);

    // rowIndex가 1(헤더 행)이거나 시트 범위를 벗어나면 저장하지 않고
    // 충돌로 표시합니다(잘못된 요청이 헤더 행을 조용히 덮어쓰는 것을 방지).
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
 * 이 스프레드시트 파일 자체를 통째로 xlsx로 내보냅니다(모든 시트
 * 포함). base64로 인코딩해서 돌려주면 화면에서 파일로 저장합니다.
 **************************************************************/
function exportFullWorkbookAction_() {
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const base64 = exportSpreadsheetAsXlsxBase64_(spreadsheetId);
  const fileName = SpreadsheetApp.getActiveSpreadsheet().getName() + "_" +
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


// 날짜 값은 셀에 문자열로 써도 시트가 자동으로 Date로 바꿔버리는
// 경우가 있어서, 서버가 다시 읽은 값(Date 객체)과 화면이 처음
// 받았던 값(JSON으로 직렬화된 날짜 문자열)의 표현이 서로 달라져
// 저장 시 "충돌"로 잘못 감지되는 문제가 있었습니다. 그래서 Date
// 객체와, Date가 JSON 직렬화될 때 나오는 ISO 형식 문자열은 모두
// "yyyy-MM-dd"로 맞춰서 비교합니다.
function normalizeText_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, "Asia/Seoul", "yyyy-MM-dd");
  }
  const text = String(value === null || value === undefined ? "" : value).trim();
  const isoDateMatch = text.match(/^(\d{4}-\d{2}-\d{2})T\d{2}:\d{2}:\d{2}/);
  if (isoDateMatch) return isoDateMatch[1];
  return text;
}


function normalizeRowLength_(row, targetColumnCount) {
  const result = row.slice(0, targetColumnCount);

  while (result.length < targetColumnCount) {
    result.push("");
  }

  return result;
}


function jsonOutput_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
