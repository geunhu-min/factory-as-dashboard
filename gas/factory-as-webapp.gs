/**************************************************************
 * 1공장추가건 / 2공장추가건 대시보드 연동용 Web App
 *
 * 배포 방법
 * ------------------------------------------------------------
 * 1. 이 파일을 대상 스프레드시트의 Apps Script 프로젝트에 추가
 *    (확장 프로그램 > Apps Script)
 * 2. 배포 > 새 배포
 *    - 유형: 웹 앱
 *    - 실행 계정: 나
 *    - 액세스 권한: 조직 내 사용자 (또는 필요 범위에 맞게)
 * 3. 배포 후 나오는 웹 앱 URL을
 *    샘플 화면(factory-as-sample/index.html)의 연결 설정에 입력
 *
 * 주의
 * ------------------------------------------------------------
 * - 토큰 검증이 없으므로 이 웹 앱 URL을 아는 사람은 누구나
 *   EDITABLE_SHEETS에 등록된 시트를 읽고 수정할 수 있습니다.
 *   URL을 외부에 공유하지 않도록 주의합니다.
 * - EDITABLE_SHEETS에 등록된 시트만 읽기/쓰기가 허용됩니다.
 * - 같은 행을 여러 사람이 동시에 열어두고 수정하면
 *   나중에 저장한 쪽이 덮어씁니다. rowIndex + keyColumn 값이
 *   저장 시점에 달라졌으면 해당 행은 충돌로 분류되어
 *   반영되지 않습니다 (응답의 conflicts 참고).
 * - "정리 실행" 기능은 factory-as-cleanup.gs의 FACTORY_CONFIG,
 *   runFactoryCore_를 그대로 사용하므로, 이 파일과
 *   factory-as-cleanup.gs가 같은 Apps Script 프로젝트(같은
 *   스프레드시트)에 함께 있어야 합니다.
 * - "회수누적/누적데이터 매칭" 기능도 factory-as-cleanup.gs의
 *   RECOVERY_SHEET_NAME, normalizeRowLength_를 그대로 사용합니다.
 **************************************************************/

/*
 * 시트 이름: 접수번호 등 고유값이 들어있는 열 인덱스(0부터 시작)
 * 저장 시 rowIndex가 가리키는 행이 불러왔을 때와 같은 행인지
 * 이 열 값으로 확인합니다.
 */
const EDITABLE_SHEETS = {
  "1공장추가건": { keyColumn: 3 },
  "2공장추가건": { keyColumn: 3 },
  "1공장삭제건": { keyColumn: 3 },
  "2공장삭제건": { keyColumn: 3 },
  "회수내역": { keyColumn: 2 }
};


function doGet(e) {
  try {
    const params = (e && e.parameter) || {};

    if (params.action === "currentWeekStatus") {
      return jsonOutput_(currentWeekStatusAction_());
    }

    if (params.action === "spreadsheetUrl") {
      return jsonOutput_({ url: SpreadsheetApp.getActiveSpreadsheet().getUrl() });
    }

    const sheetName = params.sheet || "";

    if (!EDITABLE_SHEETS[sheetName]) {
      return jsonOutput_({ error: "허용되지 않은 시트입니다: " + sheetName });
    }

    return jsonOutput_(readSheet_(sheetName));
  } catch (error) {
    return jsonOutput_({ error: error.message });
  }
}


function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");

    if (body.action === "runCleanup") {
      return jsonOutput_(runCleanupAction_(body.sheet || ""));
    }

    if (body.action === "runCleanupBoth") {
      return jsonOutput_(runCleanupBothAction_());
    }

    if (body.action === "matchRecovery") {
      return jsonOutput_(matchRecoveryAction_(body.recoveryCsvUrl || "", body.fromDate || ""));
    }

    const sheetName = body.sheet || "";

    if (!EDITABLE_SHEETS[sheetName]) {
      return jsonOutput_({ error: "허용되지 않은 시트입니다: " + sheetName });
    }

    if (!Array.isArray(body.rows)) {
      return jsonOutput_({ error: "rows 형식이 올바르지 않습니다." });
    }

    const writeResult = writeRows_(sheetName, body.rows);
    const fresh = readSheet_(sheetName);

    return jsonOutput_({
      updatedCount: writeResult.updatedCount,
      conflicts: writeResult.conflicts,
      sheet: fresh.sheet,
      header: fresh.header,
      rows: fresh.rows
    });
  } catch (error) {
    return jsonOutput_({ error: error.message });
  }
}


/**************************************************************
 * "금주현황변경" 탭 액션
 *
 * FACTORY_CONFIG(factory-as-cleanup.gs)의 originalSheet
 * (1공장/2공장 금주AS현황분석시트, 정리 전 원본)를 읽어
 * "공장" 열을 붙여 하나로 합쳐서 보여줍니다. 읽기 전용이며
 * 원본 시트는 전혀 건드리지 않습니다.
 **************************************************************/
function currentWeekStatusAction_() {
  let combinedHeader = null;
  const combinedRows = [];

  const factories = [FACTORY_CONFIG.FACTORY_1, FACTORY_CONFIG.FACTORY_2];

  factories.forEach(function(config) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.originalSheet);

    if (!sheet) {
      throw new Error("'" + config.originalSheet + "' 시트를 찾을 수 없습니다.");
    }

    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();

    if (lastRow < 1 || lastColumn < 1) {
      return;
    }

    const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();

    if (!combinedHeader) {
      combinedHeader = ["공장"].concat(values[0]);
    }

    for (let i = 1; i < values.length; i++) {
      combinedRows.push([config.name].concat(values[i]));
    }
  });

  if (!combinedHeader) {
    return { sheet: "", header: [], rows: [] };
  }

  const rows = combinedRows.map(function(row, i) {
    return { rowIndex: i + 2, values: row };
  });

  return { sheet: "금주현황", header: combinedHeader, rows: rows };
}


/**************************************************************
 * "정리 실행" 액션
 *
 * addSheet 이름(예: 1공장추가건)으로 FACTORY_CONFIG를 찾아
 * factory-as-cleanup.gs의 정리 파이프라인을 그대로 실행하고,
 * 새로 채워진 추가건/삭제건 시트를 반환합니다.
 **************************************************************/
function runCleanupAction_(addSheetName) {
  const config = findFactoryConfigByAddSheet_(addSheetName);

  if (!config) {
    return { error: "허용되지 않은 시트입니다: " + addSheetName };
  }

  runFactoryCore_(config);

  const addFresh = readSheet_(config.addSheet);
  const deleteFresh = readSheet_(config.deleteSheet);

  return {
    ok: true,
    factoryName: config.name,
    addCount: addFresh.rows.length,
    deleteCount: deleteFresh.rows.length,
    sheet: addFresh.sheet,
    header: addFresh.header,
    rows: addFresh.rows
  };
}


/**************************************************************
 * "1,2공장 정리실행" 액션
 *
 * FACTORY_CONFIG의 1공장, 2공장 정리 파이프라인을 순서대로 한 번에
 * 실행합니다. "금주현황변경" 탭에서 씁니다.
 **************************************************************/
function runCleanupBothAction_() {
  const configs = [FACTORY_CONFIG.FACTORY_1, FACTORY_CONFIG.FACTORY_2];

  const results = configs.map(function(config) {
    runFactoryCore_(config);

    const addFresh = readSheet_(config.addSheet);
    const deleteFresh = readSheet_(config.deleteSheet);

    return {
      factoryName: config.name,
      addCount: addFresh.rows.length,
      deleteCount: deleteFresh.rows.length
    };
  });

  return { ok: true, results: results };
}


/**************************************************************
 * "회수누적/누적데이터 매칭" 액션
 *
 * 1. 화면에서 고른 시작 날짜(fromDate) 이후 행만 회수누적 웹
 *    게시(CSV) 주소에서 골라
 * 2. '회수내역' 시트(RECOVERY_SHEET_NAME)는 헤더(1행)만 남기고
 *    전부 지운 뒤
 * 3. 골라낸 행을 2행부터 다시 채웁니다.
 *
 * 예: fromDate가 2026-07-14이면, 회수누적 CSV에서 2026-07-14
 * 이후(해당일 포함) 행만 가져와 채웁니다.
 **************************************************************/
function matchRecoveryAction_(recoveryCsvUrl, fromDateStr) {
  if (!recoveryCsvUrl) {
    throw new Error("회수누적 웹 게시 주소가 없습니다.");
  }

  const fromDate = toDate_(fromDateStr);

  if (!fromDate) {
    throw new Error("가져올 시작 날짜가 올바르지 않습니다.");
  }

  // recoveryCsvUrl은 클라이언트가 그대로 보내온 값을 UrlFetchApp으로
  // 가져오므로, 구글 시트 웹 게시 CSV 주소가 아닌 임의의 URL을 보내
  // 이 웹앱을 외부 요청 프록시로 악용하거나 회수내역 시트를 아무
  // 데이터로 덮어쓰는 걸 막기 위해 도메인을 제한합니다.
  if (!/^https:\/\/docs\.google\.com\/spreadsheets\//.test(recoveryCsvUrl)) {
    throw new Error("올바른 구글 시트 웹 게시 CSV 주소가 아닙니다.");
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RECOVERY_SHEET_NAME);

  if (!sheet) {
    throw new Error("'" + RECOVERY_SHEET_NAME + "' 시트를 찾을 수 없습니다.");
  }

  const lastColumn = sheet.getLastColumn();

  if (lastColumn < 1) {
    throw new Error("'" + RECOVERY_SHEET_NAME + "' 시트에 헤더가 없습니다.");
  }

  const response = UrlFetchApp.fetch(recoveryCsvUrl, { muteHttpExceptions: true });

  if (response.getResponseCode() !== 200) {
    throw new Error("회수누적 웹 게시 주소를 불러오지 못했습니다 (" + response.getResponseCode() + ")");
  }

  const csvRows = Utilities.parseCsv(response.getContentText());
  const dataRows = csvRows.length >= 2 ? csvRows.slice(1) : [];

  const matchedRows = [];
  let firstDate = null;
  let lastDate = null;

  dataRows.forEach(function(row) {
    const rowDate = toDate_(row[0]);

    if (!rowDate) return; // 날짜를 읽을 수 없는 행은 건너뜀
    if (rowDate.getTime() < fromDate.getTime()) return; // 시작 날짜보다 이전은 건너뜀

    matchedRows.push(normalizeRowLength_(row, lastColumn));

    if (!firstDate || rowDate.getTime() < firstDate.getTime()) firstDate = rowDate;
    if (!lastDate || rowDate.getTime() > lastDate.getTime()) lastDate = rowDate;
  });

  // 서식을 베낄 기준 행은 지우기 전(이전에 채워져 있던 상태)에 정해둡니다
  // — clearSheetKeepHeader_가 내용을 지우면 getLastRow()가 1로 줄어들어,
  // 지운 뒤에 기준 행을 다시 계산하면 항상 헤더(1행) 서식을 베끼게 됩니다.
  const formatReferenceRow = sheet.getLastRow() >= 2 ? 2 : 1;

  clearSheetKeepHeader_(sheet);

  if (matchedRows.length) {
    // 서식(숫자 표시 형식 포함)을 값을 쓰기 전에 먼저 적용합니다. CSV에서
    // 가져온 값은 전부 문자열인데, 셀 서식이 기본값(General)인 상태로
    // "061"처럼 숫자처럼 보이는 문자열을 쓰면 그 즉시 숫자로 바뀌어
    // 앞자리 0이 사라집니다 — 서식을 나중에 지정해도 이미 숫자로 바뀐
    // 값은 되돌아오지 않으므로 순서를 바꿨습니다.
    applyExistingFormatToRange_(sheet, formatReferenceRow, 2, matchedRows.length, lastColumn);
    sheet.getRange(2, 1, matchedRows.length, lastColumn).setValues(matchedRows);
  }

  return {
    ok: true,
    addedCount: matchedRows.length,
    firstDate: formatDate_(firstDate),
    lastDate: formatDate_(lastDate)
  };
}


/**************************************************************
 * referenceRowIndex 행의 글꼴 모양/크기와 날짜 표시 형식·천단위 콤마
 * 같은 표시 형식(numberFormat)을, startRow부터 rowCount행 동안
 * columnCount열 범위 전체에 그대로 적용합니다.
 *
 * referenceRowIndex는 호출하는 쪽에서 미리 정해서 넘겨줍니다(이
 * 함수 안에서 getLastRow()로 다시 계산하지 않음) — 값을 쓰기 전에
 * 서식부터 지정하려고 clearSheetKeepHeader_보다 먼저 호출하면,
 * 그 시점에 이미 지운 뒤라 getLastRow()가 줄어들어 기준 행이
 * 잘못(항상 헤더 행으로) 계산될 수 있기 때문입니다.
 *
 * setValues()는 셀 값만 바꾸고 서식은 그대로 두기 때문에, 기존에
 * 서식이 적용된 적 없는(이번에 새로 늘어난) 행은 구글시트 기본
 * 표시로 나옵니다. 이 함수를 값을 쓰기 전에 호출해서, 새로 채울
 * 행 전체가 기존 시트와 똑같은 모양(및 텍스트/숫자 서식)이 되도록
 * 맞춥니다.
 **************************************************************/
function applyExistingFormatToRange_(sheet, referenceRowIndex, startRow, rowCount, columnCount) {
  if (rowCount < 1 || columnCount < 1) {
    return;
  }

  const referenceRange = sheet.getRange(referenceRowIndex, 1, 1, columnCount);

  const fontFamilies = referenceRange.getFontFamilies()[0];
  const fontSizes = referenceRange.getFontSizes()[0];
  const numberFormats = referenceRange.getNumberFormats()[0];

  const fullFontFamilies = [];
  const fullFontSizes = [];
  const fullNumberFormats = [];

  for (let i = 0; i < rowCount; i++) {
    fullFontFamilies.push(fontFamilies.slice());
    fullFontSizes.push(fontSizes.slice());
    fullNumberFormats.push(numberFormats.slice());
  }

  const targetRange = sheet.getRange(startRow, 1, rowCount, columnCount);
  targetRange.setFontFamilies(fullFontFamilies);
  targetRange.setFontSizes(fullFontSizes);
  targetRange.setNumberFormats(fullNumberFormats);
}


/**************************************************************
 * 시트 헤더(1행)만 남기고 아래 내용을 전부 지웁니다.
 **************************************************************/
function clearSheetKeepHeader_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow >= 2 && lastColumn >= 1) {
    sheet.getRange(2, 1, lastRow - 1, lastColumn).clearContent();
  }
}


function toDate_(value) {
  if (value instanceof Date) return value;
  if (value === null || value === undefined || value === "") return null;

  const parsed = new Date(String(value).trim());

  return isNaN(parsed.getTime()) ? null : parsed;
}


function formatDate_(date) {
  if (!date) return "";
  return Utilities.formatDate(date, "Asia/Seoul", "yyyy-MM-dd");
}


function findFactoryConfigByAddSheet_(addSheetName) {
  const keys = Object.keys(FACTORY_CONFIG);

  for (let i = 0; i < keys.length; i++) {
    if (FACTORY_CONFIG[keys[i]].addSheet === addSheetName) {
      return FACTORY_CONFIG[keys[i]];
    }
  }

  return null;
}


/**************************************************************
 * 시트 전체를 header + 행 배열(rowIndex 포함)로 반환
 **************************************************************/
function readSheet_(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  if (!sheet) {
    throw new Error("'" + sheetName + "' 시트를 찾을 수 없습니다.");
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 1 || lastColumn < 1) {
    return { sheet: sheetName, header: [], rows: [] };
  }

  const values = sheet
    .getRange(1, 1, lastRow, lastColumn)
    .getValues();

  const header = values[0];
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    rows.push({
      rowIndex: i + 1,
      values: values[i]
    });
  }

  return { sheet: sheetName, header: header, rows: rows };
}


/**************************************************************
 * 전달받은 행을 실제 시트에 반영
 *
 * rowIndex가 가리키는 행의 keyColumn 값이
 * 요청에 담긴 key 값과 다르면 그 사이 시트가 바뀐 것이므로
 * 덮어쓰지 않고 conflicts에 담아 반환합니다.
 **************************************************************/
function writeRows_(sheetName, rows) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  if (!sheet) {
    throw new Error("'" + sheetName + "' 시트를 찾을 수 없습니다.");
  }

  const lastColumn = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const keyColumn = EDITABLE_SHEETS[sheetName].keyColumn;

  const conflicts = [];
  let updatedCount = 0;

  rows.forEach(function(row) {
    const rowIndex = Number(row.rowIndex);
    const values = row.values;

    if (
      !rowIndex ||
      rowIndex < 2 ||
      rowIndex > lastRow ||
      !Array.isArray(values)
    ) {
      conflicts.push({ rowIndex: rowIndex, reason: "잘못된 행 정보" });
      return;
    }

    const currentKey = normalizeText_(
      sheet.getRange(rowIndex, keyColumn + 1).getValue()
    );

    const expectedKey = normalizeText_(row.key);

    if (expectedKey && currentKey !== expectedKey) {
      conflicts.push({
        rowIndex: rowIndex,
        reason: "다른 곳에서 먼저 수정됨",
        currentKey: currentKey
      });
      return;
    }

    const normalized = values.slice(0, lastColumn);

    while (normalized.length < lastColumn) {
      normalized.push("");
    }

    sheet
      .getRange(rowIndex, 1, 1, lastColumn)
      .setValues([normalized]);

    updatedCount++;
  });

  SpreadsheetApp.flush();

  return { updatedCount: updatedCount, conflicts: conflicts };
}


function normalizeText_(value) {
  return String(
    value === null || value === undefined ? "" : value
  ).trim();
}


function jsonOutput_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
