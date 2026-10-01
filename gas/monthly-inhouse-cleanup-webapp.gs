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
 *   반환합니다("열기" 버튼용).
 * - doPost action="replaceSourceData": "내작월마감 자료교체" 버튼
 *   액션. "월마감" 연결 스프레드시트에서 가져온 "내작(N)" 탭의
 *   header/rows를 받아 "시트1"을 통째로 덮어씁니다(기존 내용은 전부
 *   지움). 더 이상 직접 입력/붙여넣기 하지 않습니다.
 * - doPost action="clean": "자료정리" 버튼 액션. "시트1"에서
 *   PULLED_COLUMN_LABELS 순서로 열을 뽑고 금액 뒤에 패널티(고정
 *   60,000)/클레임 계 열을 끼워 넣습니다. 포장 값으로 1라인/3·4라인
 *   (바른산업)/7라인 세 그룹(PACKAGE_GROUPS)으로 나눠 그룹별로 번호를
 *   매기고, 그룹 끝마다 "번호~수량"을 병합한 "업체명 합계(건수)" 행에
 *   금액/패널티/클레임계 합계를 채웁니다. 맨 끝에는 세 그룹 전체
 *   합계 행을 추가합니다. 헤더/업체별 합계 행은 흰색을 어둡게 한
 *   배경+굵은 글씨로 강조하고, 전체 글꼴은 맑은 고딕 11입니다. 자세한
 *   내용은 cleanMonthlyInhouseListAction_ 참고.
 * - doPost action="exportResult": "정리파일다운로드" 버튼 액션.
 *   "정리" 시트 하나만 담은 xlsx를 base64로 반환합니다(서식/글꼴
 *   그대로 유지됨).
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

// "자료정리"가 "시트1"에서 그대로 뽑아오는 열(번호/패널티/클레임 계는
// 시트1에 없고 정리하면서 새로 계산해 끼워 넣습니다).
const NUMBER_COLUMN_LABEL = "번호";
const PULLED_COLUMN_LABELS = [
  "브랜드", "지역센터", "접수번호", "구분", "형태", "포장",
  "제품코드", "색상", "수량", "금액", "하자상세", "로트"
];

// 정렬/그룹 구분 기준 열
const PACKAGE_COLUMN_LABEL = "포장";
const AMOUNT_COLUMN_LABEL = "금액";
// 구글 시트 "서식 > 숫자 > 회계"와 같은 표시 형식(₩115,000 스타일).
const AMOUNT_NUMBER_FORMAT = '_-"₩"* #,##0_-;-"₩"* #,##0_-;_-"₩"* "-"_-;_-@_-';

// 금액 바로 뒤에 끼워 넣는 계산 열 — 패널티는 건당 고정 금액, 클레임
// 계는 데이터 행에서는 비워두고 합계 행에서만 채웁니다. 값이 있는
// 클레임 계 칸(합계 행들)은 빨간 글자로 강조합니다.
const PENALTY_COLUMN_LABEL = "패널티";
const PENALTY_AMOUNT = 60000;
const CLAIM_TOTAL_COLUMN_LABEL = "클레임 계";
const CLAIM_TOTAL_FONT_COLOR = "#cc0000";

// 포장 값(1/3/4/7라인)별 담당 업체 — 그룹 순서, 묶음, 합계 행 레이블에
// 그대로 씁니다. 3라인/4라인은 같은 업체라 한 그룹으로 묶습니다.
const PACKAGE_GROUPS = [
  { packageValues: ["1라인"], companyName: "아름산업" },
  { packageValues: ["3라인", "4라인"], companyName: "바른산업" },
  { packageValues: ["7라인"], companyName: "다올산업" }
];
// 위 3개 그룹 중 어디에도 안 맞는 포장 값이 섞여 있으면 묶어서 보여줄
// 이름(합계는 구하되, 맨 아래 총합계에는 포함하지 않음 — 원래 예상된
// 값이 아니라서).
const OTHER_PACKAGE_GROUP_NAME = "기타";

// 헤더/업체별 합계 행 서식 — 흰색을 이 비율만큼 어둡게 한 배경 +
// 굵은 글씨. 전체 글꼴은 맑은 고딕 11(정리파일다운로드에도 그대로 반영됨).
const SUMMARY_ROW_BACKGROUND_DARKEN_PERCENT = 25;
const RESULT_FONT_FAMILY = "Malgun Gothic";
const RESULT_FONT_SIZE = 11;


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

    if (action === "replaceSourceData") {
      return jsonOutput_(replaceSourceDataAction_(body.header || [], body.rows || []));
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
 * "내작월마감 자료교체" 버튼 액션 — "월마감" 스프레드시트에서 가져온
 * "내작(N)" 탭의 header/rows를 그대로 "시트1"에 덮어씁니다(기존
 * 내용은 전부 지움). 화면에서 직접 입력/붙여넣기 하던 걸 대신합니다.
 **************************************************************/
function replaceSourceDataAction_(header, rows) {
  if (!header.length) {
    throw new Error("가져올 데이터의 헤더가 비어 있습니다.");
  }

  const sheet = getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), SOURCE_SHEET_NAME);
  const columnCount = header.length;
  const dataRows = rows.map(function(row) { return normalizeRowLength_(row, columnCount); });

  sheet.clear();
  sheet.getRange(1, 1, dataRows.length + 1, columnCount).setValues([header].concat(dataRows));

  // 내작(N)은 1라인/3·4라인/7라인 그룹 사이에 구분용 빈 행이 끼워져
  // 있어서, 전체 행 수를 그대로 쓰면 실제 값이 있는 건수보다 많게
  // 보입니다(빈 행도 포함됨). 값이 하나라도 있는 행만 셉니다.
  const realRowCount = dataRows.filter(function(row) {
    return row.some(function(cell) { return normalizeText_(cell) !== ""; });
  }).length;

  return { ok: true, rowCount: realRowCount };
}


function normalizeRowLength_(row, targetColumnCount) {
  const result = row.slice(0, targetColumnCount);

  while (result.length < targetColumnCount) {
    result.push("");
  }

  return result;
}


/**************************************************************
 * "자료정리" 액션
 *
 * "시트1"에서 PULLED_COLUMN_LABELS 순서로 열을 뽑고, 금액 뒤에
 * 패널티(고정 60,000)/클레임 계(데이터 행은 비움) 열을 끼워 넣습니다.
 * 포장 값으로 1라인/3·4라인/7라인 세 그룹(PACKAGE_GROUPS)으로 나눠
 * 그룹 순서대로 번호를 새로 매기고, 각 그룹 끝에는 "번호~수량"을
 * 병합한 "업체명 합계(건수)" 행을 넣어 금액/패널티/클레임계 합계를
 * 채웁니다. 맨 끝에는 빈 행 하나를 두고 세 그룹 전체 합계 행을
 * 추가합니다. 헤더와 업체별 합계 행은 흰색을 어둡게 한 배경 + 굵은
 * 글씨로 강조하고, 전체 글꼴은 맑은 고딕 11로 맞춥니다.
 **************************************************************/
function cleanMonthlyInhouseListAction_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);

  if (!sourceSheet) {
    throw new Error("'" + SOURCE_SHEET_NAME + "' 시트를 찾을 수 없습니다.");
  }

  const source = readSheetObject_(sourceSheet);
  const header = source.header;

  const pulledColumnIndexes = PULLED_COLUMN_LABELS.map(function(label) {
    const idx = header.indexOf(label);
    if (idx === -1) throw new Error("'" + SOURCE_SHEET_NAME + "'에서 '" + label + "' 열을 찾을 수 없습니다.");
    return idx;
  });

  const packagePulledPos = PULLED_COLUMN_LABELS.indexOf(PACKAGE_COLUMN_LABEL);
  const amountPulledPos = PULLED_COLUMN_LABELS.indexOf(AMOUNT_COLUMN_LABEL);

  const pulledRows = source.rows
    .map(function(row) { return pulledColumnIndexes.map(function(idx) { return row.values[idx]; }); })
    .filter(function(values) {
      return values.some(function(value) { return normalizeText_(value) !== ""; });
    });

  // 번호 + (브랜드~금액) + 패널티 + 클레임 계 + (하자상세~로트)
  const finalHeader = [NUMBER_COLUMN_LABEL]
    .concat(PULLED_COLUMN_LABELS.slice(0, amountPulledPos + 1))
    .concat([PENALTY_COLUMN_LABEL, CLAIM_TOTAL_COLUMN_LABEL])
    .concat(PULLED_COLUMN_LABELS.slice(amountPulledPos + 1));

  const columnCount = finalHeader.length;
  const amountColIndex = finalHeader.indexOf(AMOUNT_COLUMN_LABEL);
  const penaltyColIndex = finalHeader.indexOf(PENALTY_COLUMN_LABEL);
  const claimTotalColIndex = finalHeader.indexOf(CLAIM_TOTAL_COLUMN_LABEL);
  const summaryMergeColumnCount = finalHeader.indexOf("수량") + 1; // 번호~수량

  function buildDataRow_(sequence, pulledValues) {
    return [sequence]
      .concat(pulledValues.slice(0, amountPulledPos + 1))
      .concat([PENALTY_AMOUNT, ""])
      .concat(pulledValues.slice(amountPulledPos + 1));
  }

  // 포장 값 기준으로 1라인/3·4라인(바른산업)/7라인 세 그룹으로 나눕니다.
  // 세 그룹 어디에도 안 맞는 값은 "기타"로 따로 모읍니다(원래 월마감
  // 내작(N)에는 이 4개 값만 들어오지만, 혹시 다른 값이 섞여도 자료가
  // 누락되지 않도록 — 다만 기타 그룹 합계는 맨 아래 총합계에는
  // 포함하지 않습니다).
  const matchedGroups = PACKAGE_GROUPS.map(function() { return []; });
  const otherGroupRows = [];

  pulledRows.forEach(function(pulledValues) {
    const packageValue = normalizeText_(pulledValues[packagePulledPos]);
    const groupIdx = PACKAGE_GROUPS.findIndex(function(group) {
      return group.packageValues.indexOf(packageValue) !== -1;
    });

    if (groupIdx === -1) {
      otherGroupRows.push(pulledValues);
    } else {
      matchedGroups[groupIdx].push(pulledValues);
    }
  });

  const outputRows = [];
  const summaryRowNumbers = []; // 1-based 시트 행 번호(헤더 포함) — 서식/병합용
  const summaryRowNumberByCompany = {}; // 업체명 → 그 합계 행 번호(특정 업체만 따로 서식 줄 때 씀)
  let sequence = 0;

  // groupRows를 데이터 행으로 이어 붙이고, 끝에 "업체명 합계(건수)"
  // 병합 행을 추가합니다. 반환값은 이 그룹의 금액/패널티 합계(총합계
  // 계산용) — groupRows가 비어 있으면 아무 것도 하지 않고 null을 반환.
  function appendGroup_(groupRows, companyName) {
    if (!groupRows.length) return null;

    groupRows.forEach(function(pulledValues) {
      sequence++;
      outputRows.push(buildDataRow_(sequence, pulledValues));
    });

    const amountSum = groupRows.reduce(function(sum, values) {
      return sum + (Number(values[amountPulledPos]) || 0);
    }, 0);
    const penaltySum = PENALTY_AMOUNT * groupRows.length;

    const summaryRow = new Array(columnCount).fill("");
    summaryRow[0] = companyName + " 합계(" + groupRows.length + ")";
    summaryRow[amountColIndex] = amountSum;
    summaryRow[penaltyColIndex] = penaltySum;
    summaryRow[claimTotalColIndex] = amountSum + penaltySum;

    outputRows.push(summaryRow);
    const summaryRowNumber = outputRows.length + 1; // +1은 헤더 행만큼의 오프셋
    summaryRowNumbers.push(summaryRowNumber);
    summaryRowNumberByCompany[companyName] = summaryRowNumber;

    return { amountSum: amountSum, penaltySum: penaltySum };
  }

  const groupTotals = [];
  PACKAGE_GROUPS.forEach(function(group, idx) {
    const totals = appendGroup_(matchedGroups[idx], group.companyName);
    if (totals) groupTotals.push(totals);
  });

  appendGroup_(otherGroupRows, OTHER_PACKAGE_GROUP_NAME);

  let grandTotalRowNumber = null;

  if (groupTotals.length) {
    outputRows.push(new Array(columnCount).fill("")); // 총합계 앞 빈 행

    const grandAmount = groupTotals.reduce(function(sum, t) { return sum + t.amountSum; }, 0);
    const grandPenalty = groupTotals.reduce(function(sum, t) { return sum + t.penaltySum; }, 0);

    const grandTotalRow = new Array(columnCount).fill("");
    grandTotalRow[amountColIndex] = grandAmount;
    grandTotalRow[penaltyColIndex] = grandPenalty;
    grandTotalRow[claimTotalColIndex] = grandAmount + grandPenalty;
    outputRows.push(grandTotalRow);
    grandTotalRowNumber = outputRows.length + 1; // +1은 헤더 행만큼의 오프셋
  }

  const resultSheet = getOrCreateSheet_(ss, CLEAN_RESULT_SHEET_NAME);

  // clear()는 셀 서식은 지워도 열 너비는 그대로 두지만, 혹시 모를 경우에
  // 대비해 직접 맞춰둔 열 너비를 미리 기억해뒀다가 다시 쓴 뒤 그대로
  // 되돌려서, 수동으로 조절한 폭이 "자료정리"를 다시 실행해도 항상
  // 고정되게 합니다.
  const columnCountForWidths = Math.max(resultSheet.getLastColumn(), columnCount);
  const preservedWidths = [];
  for (let col = 1; col <= columnCountForWidths; col++) {
    preservedWidths.push(resultSheet.getColumnWidth(col));
  }

  resultSheet.clear();
  resultSheet.getRange(1, 1, 1, columnCount).setValues([finalHeader]);

  if (outputRows.length) {
    resultSheet.getRange(2, 1, outputRows.length, columnCount).setValues(outputRows);
  }

  preservedWidths.forEach(function(width, idx) {
    resultSheet.setColumnWidth(idx + 1, width);
  });

  const totalRowCount = outputRows.length + 1; // 헤더 포함
  resultSheet.getRange(1, 1, totalRowCount, columnCount)
    .setHorizontalAlignment("center")
    .setFontFamily(RESULT_FONT_FAMILY)
    .setFontSize(RESULT_FONT_SIZE);

  [amountColIndex, penaltyColIndex, claimTotalColIndex].forEach(function(colIndex) {
    resultSheet.getRange(1, colIndex + 1, totalRowCount, 1).setHorizontalAlignment("right");
    if (outputRows.length) {
      resultSheet.getRange(2, colIndex + 1, outputRows.length, 1).setNumberFormat(AMOUNT_NUMBER_FORMAT);
    }
  });

  // 헤더 + 업체별 합계 행: 로트 열까지 배경(흰색을 어둡게) + 굵게,
  // 업체별 합계 행은 "번호~수량"을 하나로 병합합니다.
  const summaryBackgroundHex = getDarkenedWhiteBackgroundHex_(ss, SUMMARY_ROW_BACKGROUND_DARKEN_PERCENT);
  [1].concat(summaryRowNumbers).forEach(function(rowNumber) {
    resultSheet.getRange(rowNumber, 1, 1, columnCount)
      .setBackground(summaryBackgroundHex)
      .setFontWeight("bold");
  });

  summaryRowNumbers.forEach(function(rowNumber) {
    resultSheet.getRange(rowNumber, 1, 1, summaryMergeColumnCount).merge();
  });

  // "다올산업 합계" 행은 로트 열까지 모든 테두리를 그려서, 마지막
  // 그룹과 그 아래 총합계 사이를 시각적으로 구분합니다.
  const daolCompanyName = PACKAGE_GROUPS[PACKAGE_GROUPS.length - 1].companyName;
  const daolSummaryRowNumber = summaryRowNumberByCompany[daolCompanyName];
  if (daolSummaryRowNumber) {
    resultSheet.getRange(daolSummaryRowNumber, 1, 1, columnCount)
      .setBorder(true, true, true, true, true, true);
  }

  // 맨 아래 총합계 행은 배경은 그대로 두고 굵게만 적용합니다.
  if (grandTotalRowNumber !== null) {
    resultSheet.getRange(grandTotalRowNumber, 1, 1, columnCount).setFontWeight("bold");
  }

  // 클레임 계 값이 있는 칸(업체별 합계 행 + 총합계 행)은 빨간 글자로.
  const claimTotalHighlightRows = summaryRowNumbers.slice();
  if (grandTotalRowNumber !== null) claimTotalHighlightRows.push(grandTotalRowNumber);
  claimTotalHighlightRows.forEach(function(rowNumber) {
    resultSheet.getRange(rowNumber, claimTotalColIndex + 1).setFontColor(CLAIM_TOTAL_FONT_COLOR);
  });

  return { ok: true, resultSheet: CLEAN_RESULT_SHEET_NAME, rowCount: pulledRows.length };
}


/**************************************************************
 * "흰색"(배경 1 테마 색) 기준으로 darkenPercent(%)만큼 검정 쪽으로
 * 섞은 hex 색을 돌려줍니다("흰색, N% 더 어둡게"와 같은 계산 방식 —
 * daily-service-shipment-webapp.gs의 강조2 라이터 계산과 반대 방향).
 **************************************************************/
function getDarkenedWhiteBackgroundHex_(ss, darkenPercent) {
  const rgb = ss.getSpreadsheetTheme()
    .getConcreteColor(SpreadsheetApp.ThemeColorType.BACKGROUND)
    .asRgbColor();

  const blend = function(channel) {
    return Math.round(channel * (1 - darkenPercent / 100));
  };
  const toHex = function(n) {
    return ("0" + n.toString(16)).slice(-2);
  };

  return "#" + toHex(blend(rgb.getRed())) + toHex(blend(rgb.getGreen())) + toHex(blend(rgb.getBlue()));
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
