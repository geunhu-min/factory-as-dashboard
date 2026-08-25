/**************************************************************
 * "반납내역서" 스프레드시트 정리용 Web App
 *
 * 배포 위치
 * ------------------------------------------------------------
 * 이 파일은 반납내역서 정리 작업을 하던 스프레드시트의 확장 프로그램 >
 * Apps Script 프로젝트에 추가합니다(기존에 쓰던 "반납내역서 정리"
 * 메뉴/returnDataClean 매크로와 같은 프로젝트에 그대로 둡니다).
 *
 * 하는 일
 * ------------------------------------------------------------
 * [기존] 스프레드시트를 열면 "반납내역서 정리" 메뉴가 생기고, "정리 실행"을
 * 누르면 현재 활성 시트의 D~L열(접수번호~...)을 뽑아 접수번호+순번을
 * "접수번호-순번" 형태로 합친 뒤 순번 열은 떼고 "정리결과" 시트에
 * 덮어씁니다. 이 로직 자체는 원본 그대로이고, returnDataCleanAction_로
 * 이름만 옮겨서 메뉴와 웹앱이 같은 함수를 재사용하게 했습니다.
 *
 * [신규] doGet action="read"(기본값): sheet/gid 파라미터로 지정한 탭을
 * header+rows로 반환합니다. 둘 다 없으면 "정리결과" 시트를 읽습니다.
 * 대시보드의 "반납내역서" 버튼은 sheet="시트1"로 원본 반납내역서를
 * 그대로 보여주는 데 씁니다.
 * [신규] doGet action="spreadsheetUrl": 이 스프레드시트의 편집 URL을
 * 반환합니다. 대시보드의 "반납내역서 정리 교체" 버튼이 확인창을 띄운
 * 뒤 이 스프레드시트를 새 탭에서 열 때 씁니다(시트1 내용을 정리할
 * 내용으로 바꾸는 실제 작업은 그 스프레드시트에서 직접 함).
 * [신규] doPost action="clean"(기본값): returnDataCleanAction_()을
 * 실행해서 "정리결과" 시트를 새로 채웁니다. 대시보드의 "반납내역서
 * 정리실행" 버튼이 씁니다.
 * [신규] doPost action="save": 대시보드의 "정리결과" 표에서 직접 수정한
 * 행을 저장합니다. body에 sheet("정리결과")와 rows(각 행마다 loaded/
 * values)를 담아 보내면, 저장 직전 시트의 실제 값과 loaded를 비교해서
 * 그 사이 다른 곳에서 먼저 바뀐 행은 충돌로 보고 덮어쓰지 않습니다.
 * [신규] doPost action="exportResult": "정리결과엑셀다운로드" 버튼
 * 액션. "정리결과" 시트 하나만 담은 xlsx를 base64로 반환합니다.
 *
 * 배포 방법
 * ------------------------------------------------------------
 * 1. 반납내역서 스프레드시트 > 확장 프로그램 > Apps Script에 이 파일 추가
 *    (기존 매크로 파일이 따로 있다면 그 파일의 onOpen/returnDataClean은
 *    지우고 이 파일로 교체하거나, 이 파일 하나만 남기고 정리하세요 —
 *    onOpen이 두 파일에 중복으로 있으면 메뉴가 두 번 생깁니다)
 * 2. 배포 > 새 배포 > 유형: 웹 앱, 실행 계정: 나, 액세스 권한: 필요 범위
 * 3. 배포 후 나오는 웹 앱 URL을 대시보드의 "일일업무" 페이지 >
 *    "일일현황관리 데이터" > "상세" > "반납내역서 정리 연결"에 입력
 * 4. UrlFetchApp/DriveApp을 처음 쓰는 경우(exportResult) 권한 재승인이
 *    필요할 수 있습니다. 함수 선택 드롭다운에서 testAuth를 선택해 한 번
 *    실행하고 동의 화면을 통과한 뒤 다시 배포하세요.
 *
 * 주의 — 꼭 확인해주세요
 * ------------------------------------------------------------
 * - returnDataCleanAction_은 원본 데이터를 SOURCE_SHEET_NAME("시트1",
 *   대시보드가 "반납내역서" 버튼으로 읽는 탭과 동일)에서 고정으로
 *   읽습니다(예전에는 SpreadsheetApp.getActiveSheet()를 써서, 웹앱으로
 *   실행하면 "마지막으로 사람이 열어뒀던 탭"이 원본이 되는 문제가
 *   있었습니다 — 실제 원본 탭 이름이 "시트1"이 아니라면 SOURCE_SHEET_NAME
 *   값을 바꿔주세요).
 * - 토큰 검증이 없으므로 URL을 아는 사람은 누구나 이 시트를 읽고
 *   정리를 실행할 수 있습니다.
 **************************************************************/

const RESULT_SHEET_NAME = "정리결과";
const SOURCE_SHEET_NAME = "시트1"; // 대시보드가 "반납내역서" 원본으로 읽는 탭과 동일(doGet 참고)


/**************************************************************
 * 권한 재승인용 임시 테스트 함수
 *
 * exportResult가 쓰는 UrlFetchApp(외부 요청) 권한을 승인받기 위한
 * 함수입니다. 이름에 밑줄(_)이 없어야 Apps Script 편집기의 "실행할
 * 함수" 드롭다운에 보입니다. 드롭다운에서 testAuth를 선택해 실행하면
 * 동의 화면이 뜹니다 — 승인한 뒤에는 이 함수를 지우고 다시 배포해도
 * 되고, 그냥 남겨둬도 동작에는 영향이 없습니다.
 **************************************************************/
function testAuth() {
  UrlFetchApp.fetch("https://www.google.com");
}


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('반납내역서 정리')
    .addItem('정리 실행', 'returnDataClean')
    .addToUi();
}


function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    const action = params.action || "read";

    if (action === "read") {
      const sheet = resolveSheet_(params.sheet || RESULT_SHEET_NAME, params.gid);
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
    const action = body.action || "clean";

    if (action === "clean") {
      return jsonOutput_(returnDataCleanAction_());
    }

    if (action === "save") {
      const sheet = resolveSheet_(body.sheet, body.gid);
      return jsonOutput_(saveRowsAction_(sheet, body.rows || []));
    }

    if (action === "exportResult") {
      return jsonOutput_(exportSheetsSubsetAsXlsxBase64_([RESULT_SHEET_NAME], "정리결과"));
    }

    return jsonOutput_({ error: "알 수 없는 action입니다: " + action });
  } catch (error) {
    return jsonOutput_({ error: error.message });
  }
}


/**************************************************************
 * "정리 실행" 메뉴 액션 — returnDataCleanAction_을 실행하고 결과를
 * 알림창으로 보여줍니다(스프레드시트에서 메뉴로 직접 실행할 때만 씀).
 **************************************************************/
function returnDataClean() {
  const result = returnDataCleanAction_();
  SpreadsheetApp.getUi().alert(
    "정리 완료\n총 " + result.processedCount + "건 처리되었습니다."
  );
}


// "정리결과" 시트가 항상 이 순서로 나와야 합니다. "시트1"의 열 순서가
// 바뀌어도(다른 컴퓨터에서 변환한 경우 등) 아래처럼 헤더 이름으로 열을
// 찾아서 이 순서대로 뽑아 쓰므로 흔들리지 않습니다. "순번"은 접수번호에
// "-순번"으로 합쳐서 쓰고 별도 출력 열로 남기지 않으므로 이 목록엔 없습니다.
const RETURN_DOC_TARGET_HEADER = [
  "접수번호", "품목코드", "칼라", "품목명칭", "제품공급업체", "고객명",
  "결과현상", "조치결과특이사항"
];

/**************************************************************
 * 헤더 셀에 정렬/필터 화살표 같은 부가 기호("제품공급업체▲"처럼)가
 * 붙어있어도 이름으로 찾을 수 있도록, 정확히 일치하는 게 없으면
 * 그 기호를 뗀 뒤에도 한 번 더 비교합니다.
 **************************************************************/
function findHeaderIndexLenient_(header, label) {
  const exact = header.indexOf(label);
  if (exact !== -1) return exact;

  return header.findIndex(function(cell) {
    return String(cell || "").replace(/[▲▼△▽]/g, "").trim() === label;
  });
}

/**************************************************************
 * "시트1"에서 헤더 이름으로 열을 찾아 접수번호+순번을 "접수번호-순번"
 * 형태로 합친 뒤 RETURN_DOC_TARGET_HEADER 순서로 뽑아 "정리결과" 시트에
 * 덮어씁니다. 메뉴("정리 실행")와 웹앱(doPost action="clean") 둘 다
 * 이 함수를 씁니다.
 **************************************************************/
function returnDataCleanAction_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = resolveSheet_(SOURCE_SHEET_NAME, null);

  let resultSheet = ss.getSheetByName(RESULT_SHEET_NAME);

  if (!resultSheet) {
    resultSheet = ss.insertSheet(RESULT_SHEET_NAME);
  }

  const data = source.getDataRange().getDisplayValues();

  if (data.length < 2) {
    throw new Error("데이터가 없습니다.");
  }

  const header = data[0];

  const seqColIndex = findHeaderIndexLenient_(header, "순번");

  if (seqColIndex === -1) {
    throw new Error("'" + SOURCE_SHEET_NAME + "'에서 '순번' 열을 찾을 수 없습니다.");
  }

  const targetColIndexes = RETURN_DOC_TARGET_HEADER.map(function(label) {
    const idx = findHeaderIndexLenient_(header, label);

    if (idx === -1) {
      throw new Error(
        "'" + SOURCE_SHEET_NAME + "'에서 '" + label + "' 열을 찾을 수 없습니다.\n\n" +
        "헤더(1행)에 이 이름이 정확히 들어있는지 확인해 주세요."
      );
    }

    return idx;
  });

  const receiptNoTargetPos = RETURN_DOC_TARGET_HEADER.indexOf("접수번호");

  const finalResult = [RETURN_DOC_TARGET_HEADER.slice()];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const seq = String(row[seqColIndex]).trim().padStart(2, "0");
    const outRow = targetColIndexes.map(function(idx) { return row[idx]; });

    const receiptNo = String(outRow[receiptNoTargetPos]).trim();

    if (receiptNo !== "") {
      outRow[receiptNoTargetPos] = receiptNo + "-" + seq;
    }

    finalResult.push(outRow);
  }

  // 데이터만 삭제 (열너비, 서식 유지)
  resultSheet.clearContents();

  // 전체 텍스트 처리
  resultSheet.getRange("A:Z").setNumberFormat("@");

  // 결과 입력
  resultSheet
    .getRange(1, 1, finalResult.length, finalResult[0].length)
    .setValues(finalResult);

  // 결과 범위 선택
  resultSheet.setActiveSelection(
    resultSheet.getRange(1, 1, finalResult.length, finalResult[0].length)
  );

  return {
    ok: true,
    processedCount: finalResult.length - 1,
    resultSheet: RESULT_SHEET_NAME
  };
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
 * 이 스프레드시트에서 sheetNames에 해당하는 시트만 임시 스프레드시트에
 * 복사해서 xlsx로 내보낸 뒤, 임시 스프레드시트는 지웁니다
 * ("정리결과엑셀다운로드" 버튼에서 사용).
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

    // 새 스프레드시트가 기본으로 만들어주는 빈 시트(Sheet1 등)는 지움
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


// 날짜 값("접수일자" 등)은 셀에 문자열로 써도 시트가 자동으로 Date로
// 바꿔버리는 경우가 있어서, 서버가 다시 읽은 값(Date 객체)과 화면이
// 처음 받았던 값(JSON으로 직렬화된 날짜 문자열)의 표현이 서로 달라져
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
