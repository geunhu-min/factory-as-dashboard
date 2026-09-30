/**************************************************************
 * "월마감" 스프레드시트 연동용 Web App
 *
 * 배포 위치
 * ------------------------------------------------------------
 * 이 파일은 "월마감" 스프레드시트(종합(N)/마감(N)/VN(N)/내작(N)/
 * 설계(N)/외작(N)/구매(N) 탭이 이미 만들어져 있는 파일)의
 * 확장 프로그램 > Apps Script 프로젝트에 추가합니다.
 *
 * 하는 일
 * ------------------------------------------------------------
 * 대시보드의 "마감자료정리" 버튼을 누르면, 화면이 먼저 "월현황(주간)"
 * 웹앱에서 최신 "종합(N)"/"마감(N)" 내용을 읽어와서(header/rows) 이
 * 웹앱으로 통째로 넘겨줍니다. 이 파일은 그 데이터를 받아서:
 *
 * - 종합: 받은 종합 내용 그대로 씀
 * - 마감: 받은 마감 내용 그대로 씀
 * - VN: 종합에서 포장이 "VN"인 행만
 * - 내작: 종합에서 원인이 1라인/3라인/4라인/7라인인 행만 — 1라인
 *   (제품코드순) → 빈행 → 3라인+4라인(제품코드순) → 빈행 →
 *   7라인(제품코드순)
 * - 설계: 종합에서 원인이 "제품기술팀"인 행만
 * - 외작: 종합에서 원인이 "외주("로 시작하는 행만 — 원인별로 묶어서
 *   건수 많은 순(같으면 금액 합계 큰 순)으로 정렬하고, 그룹 내부는
 *   제품코드순, 그룹 사이에 빈 행 하나씩. 금액 열은 값 없이 헤더만 남김
 * - 구매: 원인이 "구매("로 시작하는 행만, 외작과 같은 방식
 *
 * doGet action="sheets": 이 파일의 모든 탭 이름 + gid + 행/열 수 목록
 * doGet action="read": 탭 하나를 header + rows로 반환합니다
 * doGet action="spreadsheetUrl": 이 스프레드시트의 편집 URL을 반환합니다
 *   ("열기" 버튼용)
 * doPost action="rebuildMonthlyClosingSheets": 위 정리를 전부 실행합니다.
 *   body: { summaryHeader, summaryRows, closingHeader, closingRows }
 *   (summaryRows/closingRows는 각 행이 순수 값 배열인 2차원 배열)
 *
 * 배포 방법
 * ------------------------------------------------------------
 * 1. "월마감" 스프레드시트 > 확장 프로그램 > Apps Script에 이 파일 추가
 * 2. 배포 > 새 배포 > 유형: 웹 앱, 실행 계정: 나, 액세스 권한: 누구나
 * 3. 배포 후 나오는 웹 앱 URL을 대시보드 "주간업무" 페이지 >
 *    "주간현황관리 데이터" > "상세" > "월마감" 연결에 입력
 *
 * 주의
 * ------------------------------------------------------------
 * - 토큰 검증이 없으므로 URL을 아는 사람은 누구나 이 시트를 읽고
 *   덮어쓸 수 있습니다.
 * - 종합(N)/마감(N)/VN(N)/내작(N)/설계(N)/외작(N)/구매(N) 탭이 이미
 *   있어야 합니다(이름 뒤 숫자는 매번 바뀌어도 됨). 없으면 에러를
 *   던집니다 — 새로 만들지 않습니다.
 * - VN/설계 탭은 정렬 기준을 따로 안 주셔서 종합의 원래 순서를 그대로
 *   유지했습니다. 원하시는 정렬이 있으면 말씀해주세요.
 **************************************************************/

// 이미 있는 탭을 이름 패턴으로 찾습니다(뒤에 건수가 붙어 매번 이름이
// 바뀌므로 정확한 이름 대신 "라벨(숫자)" 패턴으로 찾음). 숫자가 아직
// 안 붙은 맨 처음 상태도 찾을 수 있도록 숫자 부분은 선택 사항으로
// 둡니다 — 괄호 자체가 아예 없는 "종합"과, 괄호는 있지만 숫자가
// 비어있는 템플릿 상태인 "종합()" 둘 다 포함합니다.
function findMonthlyFinalSheet_(ss, label) {
  const pattern = new RegExp("^" + label + "(\\(\\d*\\))?$");
  return ss.getSheets().find(function(sheet) { return pattern.test(sheet.getName()); }) || null;
}

// VN/미회수 행 글자색 (그 외 행은 검정색)
const FINAL_VN_FONT_COLOR = "#1155cc";
const FINAL_UNCOLLECTED_FONT_COLOR = "#cc0000";
const FINAL_DEFAULT_FONT_COLOR = "#000000";

// "하자상세"가 "이의제기"인 칸 배경색(월현황(주간)의 강조와 동일)
const FINAL_DEFECT_DETAIL_OBJECTION_VALUE = "이의제기";
const FINAL_DEFECT_DETAIL_OBJECTION_COLOR = "#ffff00";

// 새로 쓰는 데이터의 글꼴, 날짜/금액 표시 형식(월현황(주간)과 동일)
const FINAL_DEFAULT_FONT_FAMILY = "Malgun Gothic";
const FINAL_DATE_FORMAT_COLUMN_LABELS = ["최종조치일", "반납일자"];
const FINAL_DATE_NUMBER_FORMAT = "yyyy-mm-dd";
const FINAL_AMOUNT_FORMAT_COLUMN_LABELS = ["금액"];
const FINAL_AMOUNT_NUMBER_FORMAT = "#,##0";

// 내작 탭에서 뽑을 원인 값 3그룹(1라인 / 3라인+4라인 / 7라인)
const INHOUSE_LINE_1 = ["1라인"];
const INHOUSE_LINE_34 = ["3라인", "4라인"];
const INHOUSE_LINE_7 = ["7라인"];


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
    const action = body.action || "";

    if (action === "rebuildMonthlyClosingSheets") {
      return jsonOutput_(rebuildMonthlyClosingSheetsAction_(body));
    }

    return jsonOutput_({ error: "알 수 없는 action입니다: " + action });
  } catch (error) {
    return jsonOutput_({ error: error.message });
  }
}


/**************************************************************
 * "마감자료정리" 액션 본체. 클래스 상단 설명 참고.
 **************************************************************/
function rebuildMonthlyClosingSheetsAction_(body) {
  const summaryHeader = body.summaryHeader || [];
  const summaryRows = body.summaryRows || [];
  const closingHeader = body.closingHeader || [];
  const closingRows = body.closingRows || [];

  if (!summaryHeader.length) {
    throw new Error("월현황(주간)에서 받아온 '종합' 데이터가 비어 있습니다.");
  }
  if (!closingHeader.length) {
    throw new Error("월현황(주간)에서 받아온 '마감' 데이터가 비어 있습니다.");
  }

  const causeIdx = summaryHeader.indexOf("원인");
  const packageIdx = summaryHeader.indexOf("포장");
  const productCodeIdx = summaryHeader.indexOf("제품코드");
  const amountIdx = summaryHeader.indexOf("금액");

  if (causeIdx === -1 || packageIdx === -1 || productCodeIdx === -1 || amountIdx === -1) {
    throw new Error("'종합' 데이터에서 원인/포장/제품코드/금액 열을 모두 찾을 수 없습니다.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  writeMonthlyFinalSheet_(ss, "종합", summaryHeader, summaryRows);
  writeMonthlyFinalSheet_(ss, "마감", closingHeader, closingRows);

  const vnRows = summaryRows.filter(function(row) {
    return normalizeText_(row[packageIdx]) === "VN";
  });
  writeMonthlyFinalSheet_(ss, "VN", summaryHeader, vnRows);

  const inhouseSourceRows = summaryRows.filter(function(row) {
    const cause = normalizeText_(row[causeIdx]);
    return INHOUSE_LINE_1.indexOf(cause) !== -1 ||
      INHOUSE_LINE_34.indexOf(cause) !== -1 ||
      INHOUSE_LINE_7.indexOf(cause) !== -1;
  });
  const inhouseRows = buildInhouseGroupedRows_(inhouseSourceRows, causeIdx, productCodeIdx);
  writeMonthlyFinalSheet_(ss, "내작", summaryHeader, inhouseRows);

  const designRows = summaryRows.filter(function(row) {
    return normalizeText_(row[causeIdx]) === "제품기술팀";
  });
  writeMonthlyFinalSheet_(ss, "설계", summaryHeader, designRows);

  const outsourceSourceRows = summaryRows.filter(function(row) {
    return /^외주\(/.test(normalizeText_(row[causeIdx]));
  });
  const outsourceRows = buildCauseGroupedRows_(outsourceSourceRows, causeIdx, productCodeIdx, amountIdx, true);
  writeMonthlyFinalSheet_(ss, "외작", summaryHeader, outsourceRows);

  const purchaseSourceRows = summaryRows.filter(function(row) {
    return /^구매\(/.test(normalizeText_(row[causeIdx]));
  });
  const purchaseRows = buildCauseGroupedRows_(purchaseSourceRows, causeIdx, productCodeIdx, amountIdx, true);
  writeMonthlyFinalSheet_(ss, "구매", summaryHeader, purchaseRows);

  return { ok: true };
}


/**************************************************************
 * 내작 탭용 — 1라인 / 3라인+4라인 / 7라인 순서로, 각 그룹 안에서는
 * 제품코드순 정렬. 그룹 사이(비어있지 않은 그룹끼리만)에 빈 행 하나.
 **************************************************************/
function buildInhouseGroupedRows_(rows, causeIdx, productCodeIdx) {
  const group1 = rows.filter(function(row) { return INHOUSE_LINE_1.indexOf(normalizeText_(row[causeIdx])) !== -1; });
  const group2 = rows.filter(function(row) { return INHOUSE_LINE_34.indexOf(normalizeText_(row[causeIdx])) !== -1; });
  const group3 = rows.filter(function(row) { return INHOUSE_LINE_7.indexOf(normalizeText_(row[causeIdx])) !== -1; });

  [group1, group2, group3].forEach(function(group) {
    group.sort(function(a, b) { return compareProductCode_(a[productCodeIdx], b[productCodeIdx]); });
  });

  const columnCount = rows.length ? rows[0].length : 0;
  const blankRow = new Array(columnCount).fill("");
  const nonEmptyGroups = [group1, group2, group3].filter(function(group) { return group.length; });

  const output = [];
  nonEmptyGroups.forEach(function(group, idx) {
    if (idx > 0) output.push(blankRow.slice());
    group.forEach(function(row) { output.push(row); });
  });

  return output;
}


/**************************************************************
 * 외작/구매 탭용 — 원인 값별로 묶어서 건수 많은 순(같으면 금액 합계
 * 큰 순)으로 그룹 순서를 정하고, 그룹 안에서는 제품코드순 정렬.
 * 그룹 사이에 빈 행 하나. blankAmount가 true면 금액 칸을 전부 빈
 * 값으로 만듭니다(헤더만 남기는 용도).
 **************************************************************/
function buildCauseGroupedRows_(rows, causeIdx, productCodeIdx, amountIdx, blankAmount) {
  const groups = {};
  const order = [];

  rows.forEach(function(row) {
    const cause = normalizeText_(row[causeIdx]);
    if (!groups[cause]) {
      groups[cause] = [];
      order.push(cause);
    }
    groups[cause].push(row);
  });

  order.sort(function(a, b) {
    const countDiff = groups[b].length - groups[a].length;
    if (countDiff !== 0) return countDiff;

    const sumA = groups[a].reduce(function(sum, row) { return sum + (Number(row[amountIdx]) || 0); }, 0);
    const sumB = groups[b].reduce(function(sum, row) { return sum + (Number(row[amountIdx]) || 0); }, 0);
    return sumB - sumA;
  });

  const columnCount = rows.length ? rows[0].length : 0;
  const blankRow = new Array(columnCount).fill("");
  const output = [];

  order.forEach(function(cause, idx) {
    const groupRows = groups[cause].slice().sort(function(a, b) {
      return compareProductCode_(a[productCodeIdx], b[productCodeIdx]);
    });

    if (idx > 0) output.push(blankRow.slice());

    groupRows.forEach(function(row) {
      const outputRow = row.slice();
      if (blankAmount) outputRow[amountIdx] = "";
      output.push(outputRow);
    });
  });

  return output;
}


function compareProductCode_(a, b) {
  return normalizeText_(a).localeCompare(normalizeText_(b), "ko", { numeric: true, sensitivity: "base" });
}


/**************************************************************
 * label(N) 패턴으로 탭을 찾아서 header+rows로 통째로 덮어씁니다.
 * 열 너비는 미리 기억해뒀다가 그대로 되돌리고, 날짜/금액 서식,
 * VN·미회수 글자색, 하자상세=이의제기 배경색, 헤더 필터를 적용한
 * 뒤 "label(실제 데이터 건수)"로 탭 이름을 바꿉니다.
 **************************************************************/
function writeMonthlyFinalSheet_(ss, label, header, rows) {
  const sheet = findMonthlyFinalSheet_(ss, label);
  if (!sheet) {
    throw new Error("'" + label + "(N)' 형식의 탭을 찾을 수 없습니다. 월마감 스프레드시트의 탭 이름을 확인해주세요.");
  }

  const columnCount = header.length;
  const dataRows = rows.map(function(row) { return normalizeRowLength_(row, columnCount); });

  const preservedWidths = [];
  const widthColumnCount = Math.max(sheet.getLastColumn(), columnCount);
  for (let col = 1; col <= widthColumnCount; col++) {
    preservedWidths.push(sheet.getColumnWidth(col));
  }

  sheet.clear();

  const colorColIndex = header.indexOf("색상");
  if (colorColIndex !== -1 && dataRows.length) {
    sheet.getRange(2, colorColIndex + 1, dataRows.length, 1).setNumberFormat("@");
  }

  const fullRange = sheet.getRange(1, 1, dataRows.length + 1, columnCount);
  fullRange.setValues([header].concat(dataRows));
  fullRange.setFontFamily(FINAL_DEFAULT_FONT_FAMILY);
  sheet.setFrozenRows(1);

  FINAL_DATE_FORMAT_COLUMN_LABELS.forEach(function(fieldLabel) {
    const colIndex = header.indexOf(fieldLabel);
    if (colIndex !== -1 && dataRows.length) {
      sheet.getRange(2, colIndex + 1, dataRows.length, 1).setNumberFormat(FINAL_DATE_NUMBER_FORMAT);
    }
  });

  FINAL_AMOUNT_FORMAT_COLUMN_LABELS.forEach(function(fieldLabel) {
    const colIndex = header.indexOf(fieldLabel);
    if (colIndex !== -1 && dataRows.length) {
      sheet.getRange(2, colIndex + 1, dataRows.length, 1).setNumberFormat(FINAL_AMOUNT_NUMBER_FORMAT);
    }
  });

  applyFinalPackageFontColors_(sheet, 2, dataRows, columnCount, header.indexOf("포장"));
  applyFinalDefectDetailObjectionColors_(sheet, 2, dataRows, header.indexOf("하자상세"));
  applyFinalSheetFilter_(sheet, dataRows.length + 1, columnCount);

  preservedWidths.forEach(function(width, idx) {
    sheet.setColumnWidth(idx + 1, width);
  });

  const realDataCount = dataRows.filter(function(row) {
    return row.some(function(cell) { return normalizeText_(cell) !== ""; });
  }).length;

  const newName = label + "(" + realDataCount + ")";
  if (sheet.getName() !== newName) {
    sheet.setName(newName);
  }
}


function applyFinalPackageFontColors_(sheet, startRow, rows, columnCount, packageIdx) {
  if (!rows.length || packageIdx === -1) return;

  const colors = rows.map(function(row) {
    const packageValue = normalizeText_(row[packageIdx]);
    let color = FINAL_DEFAULT_FONT_COLOR;

    if (packageValue === "VN") {
      color = FINAL_VN_FONT_COLOR;
    } else if (packageValue === "미회수") {
      color = FINAL_UNCOLLECTED_FONT_COLOR;
    }

    return new Array(columnCount).fill(color);
  });

  sheet.getRange(startRow, 1, rows.length, columnCount).setFontColors(colors);
}


function applyFinalDefectDetailObjectionColors_(sheet, startRow, rows, defectDetailIdx) {
  if (!rows.length || defectDetailIdx === -1) return;

  const colors = rows.map(function(row) {
    return [normalizeText_(row[defectDetailIdx]) === FINAL_DEFECT_DETAIL_OBJECTION_VALUE ? FINAL_DEFECT_DETAIL_OBJECTION_COLOR : null];
  });

  sheet.getRange(startRow, defectDetailIdx + 1, rows.length, 1).setBackgrounds(colors);
}


function applyFinalSheetFilter_(sheet, totalRows, columnCount) {
  const existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();

  if (totalRows > 0 && columnCount > 0) {
    sheet.getRange(1, 1, totalRows, columnCount).createFilter();
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
 * 시트 전체를 header + 행 배열(rowIndex 포함)로 반환합니다.
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


function normalizeText_(value) {
  return String(value === null || value === undefined ? "" : value).trim();
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
