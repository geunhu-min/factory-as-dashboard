/**************************************************************
 * 신규 스프레드시트 조회/편집용 Web App
 * (https://docs.google.com/spreadsheets/d/1havLRI8jserbdVfcETda0a_4y7UdVpaV5knviKIanQg/...)
 *
 * 배포 위치
 * ------------------------------------------------------------
 * 이 파일은 위 스프레드시트(1havLRI8jserbdVfcETda0a_4y7UdVpaV5knviKIanQg)의
 * 확장 프로그램 > Apps Script 프로젝트에 추가합니다.
 *
 * 하는 일
 * ------------------------------------------------------------
 * - doGet action="sheets": 이 파일의 모든 탭 이름 + gid + 행/열 수 목록
 *   (어떤 탭을 볼지 고르기 전에, 구조 파악용으로 먼저 호출)
 * - doGet action="read" (기본값): 탭 하나를 header + rows로 반환합니다.
 *   sheet 파라미터(탭 이름) 또는 gid 파라미터로 지정, 둘 다 없으면
 *   DEFAULT_GID(처음 요청한 gid=1014469752)를 씁니다. 1행을 헤더로
 *   보고 그 아래부터 데이터 행으로 취급합니다(병합 셀이 있으면 병합된
 *   첫 칸에만 값이 들어오고 나머지는 빈 값으로 옵니다).
 * - doPost action="save": 화면에서 수정한 행을 저장합니다. 행마다
 *   loaded(불러왔을 때 값)와 values(수정한 값)를 같이 보내면, 저장
 *   직전 시트의 실제 값과 loaded를 비교해서 그 사이에 다른 곳에서
 *   먼저 바뀐 행은 충돌로 보고 덮어쓰지 않습니다(그 행은 최신 값
 *   그대로 돌려줌). 응답에 최신 header/rows/conflicts/updatedCount를
 *   담아 돌려주므로, 화면은 그 값으로 다시 그리면 됩니다.
 * - doPost action="exportFull": 이 스프레드시트 파일 전체(모든 시트)를
 *   xlsx로 내보내 base64로 반환합니다(오프라인 비교/백업용).
 * - doPost action="carryOverPreviousMonth": "전월데이터 이월하기" 버튼
 *   액션. "이번달" 시트의 현재 값+서식을 "전월" 시트에 그대로
 *   덮어씁니다(전월 기존 내용은 전부 지움). 자세한 내용은
 *   carryOverPreviousMonthAction_ 참고.
 * - doPost action="cleanupAddedItems": "추가건정리" 버튼 액션.
 *   "이번달"/"전월" 시트를 접수번호+순번+부품코드+색상+금액 기준으로
 *   비교해 이번달에 새로 추가된 행만 골라, 발주일자/접수번호가 빈
 *   행은 버리고, 값이 하나도 없는 열은 헤더 포함 지운 뒤 "정리"
 *   시트에 덮어씁니다. 여기까지만 하고, 브랜드별로 나누는 건 하지
 *   않습니다(아래 cleanupByBrand가 이어서 처리). 자세한 내용은
 *   cleanupAddedItemsAction_ 참고.
 * - doPost action="cleanupByBrand": "브랜드별정리" 버튼 액션.
 *   cleanupAddedItems가 이미 만들어둔 "정리" 시트 내용을 접수번호
 *   앞글자(F/P → 퍼시스, I → 일룸) 기준으로 나눠 "퍼시스"/"일룸"
 *   시트에 각각 덮어씁니다(맨 아래에 금액 합계 행 추가). 이어서
 *   퍼시스/일룸 시트에서 번호·발주일자·제품구분·분류·접수번호·
 *   부품코드·부품명·색상·확정·매입단가·금액 순으로 열만 뽑아
 *   "퍼시스내역서"/"일룸내역서" 시트에도 덮어씁니다(역시 맨 아래에
 *   금액 합계). "정리" 시트가 아직 없으면(추가건정리를 먼저 실행하지
 *   않았으면) 에러가 납니다. "퍼시스"/"일룸"/내역서 시트 모두 열
 *   너비·숫자 표시 형식은 "전월" 시트 기준으로 맞추고, 글꼴은 전월
 *   기준(글자색은 항상 검정 고정)으로, 행 높이는 고정값으로,
 *   발주일자·입고요청일·입고확정일·출고예정일은 날짜 형식으로,
 *   매입단가·금액은 천단위 콤마로 표시합니다. 자세한 내용은
 *   cleanupByBrandAction_ 참고.
 * - doPost action="exportPaidClosing": "퍼시스"/"일룸" 시트 2개만 담은
 *   xlsx를 base64로 반환합니다("유상마감다운로드" 버튼).
 * - doPost action="exportPaidClosingDetail": "퍼시스내역서"/"일룸내역서"
 *   시트 2개만 담은 xlsx를 base64로 반환합니다("유상마감내역서다운로드"
 *   버튼).
 *
 * 배포 방법
 * ------------------------------------------------------------
 * 1. 대상 스프레드시트 > 확장 프로그램 > Apps Script에 이 파일 추가
 * 2. 배포 > 새 배포 > 유형: 웹 앱, 실행 계정: 나, 액세스 권한: 필요 범위
 * 3. 배포 후 나오는 웹 앱 URL을 대시보드의 연결 정보에 입력
 * 4. UrlFetchApp을 처음 쓰는 경우 권한 재승인이 필요할 수 있습니다.
 *    함수 선택 드롭다운에서 testAuth를 선택해 한 번 실행하고 동의
 *    화면을 통과한 뒤 다시 배포하세요.
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

const DEFAULT_GID = 1014469752;

// "유상건정리" 액션에서 쓰는 시트 이름/비교 기준 열
const CURRENT_MONTH_SHEET_NAME = "이번달";
const PREVIOUS_MONTH_SHEET_NAME = "전월";
const CLEANUP_SHEET_NAME = "정리";
const COMPARE_KEY_COLUMN_LABELS = ["접수번호", "순번", "부품코드", "색상", "금액"];
const REQUIRED_NONBLANK_COLUMN_LABELS = ["발주일자", "접수번호"];
const ACCESSION_COLUMN_LABEL = "접수번호";
const VENDOR_SHEETS_BY_PREFIX = [
  { prefixes: ["F", "P"], sheetName: "퍼시스", detailSheetName: "퍼시스내역서" },
  { prefixes: ["I"], sheetName: "일룸", detailSheetName: "일룸내역서" }
];

// 퍼시스/일룸, 퍼시스내역서/일룸내역서 시트 맨 아래에 붙일 금액 합계
const AMOUNT_COLUMN_LABEL = "금액";
const AMOUNT_TOTAL_LABEL = "합계";
const AMOUNT_TOTAL_HIGHLIGHT_COLOR = "#ffff00";

// 글자색이 안 보이게 되는 문제를 막기 위해 항상 이 색으로 강제
const DEFAULT_FONT_COLOR = "#000000";

// 행 높이를 항상 이 값(구글시트 기본 행 높이)으로 고정
// ("전월" 시트의 실제 행 높이를 그대로 복사했더니 비정상적으로 커지는
// 문제가 있어서, 고정값으로 맞춰서 항상 일정하게 나오도록 함)
const DEFAULT_ROW_HEIGHT = 21;

// 퍼시스/일룸/퍼시스내역서/일룸내역서 안에서 1번부터 다시 매기는 번호 열
const NUMBER_COLUMN_LABEL = "번호";

// 날짜로 표시할 열 이름과 날짜 표시 형식
const DATE_FORMAT_COLUMN_LABELS = ["발주일자", "입고요청일", "입고확정일", "출고예정일"];
const DATE_NUMBER_FORMAT = "yyyy-mm-dd";

// 천단위 콤마로 표시할 열 이름과 숫자 표시 형식
const AMOUNT_FORMAT_COLUMN_LABELS = ["매입단가", "금액"];
const AMOUNT_NUMBER_FORMAT = "#,##0";

// "내역서" 시트는 퍼시스/일룸 시트에서 이 순서로 열만 뽑아 만듭니다.
const DETAIL_COLUMN_LABELS = [
  "번호", "발주일자", "제품구분", "분류", "접수번호",
  "부품코드", "부품명", "색상", "확정", "매입단가", "금액"
];

/**************************************************************
 * 권한 재승인용 임시 테스트 함수
 *
 * exportFull이 쓰는 UrlFetchApp(외부 요청), exportPaidClosing/
 * exportPaidClosingDetail이 쓰는 DriveApp의 파일 생성/삭제(휴지통
 * 이동) 권한을 승인받기 위한 함수입니다. DriveApp.getRootFolder()
 * 같은 읽기 동작은 "drive" 전체 권한을 요청하지 않아서, 실제로
 * exportSheetsSubsetAsXlsxBase64_가 하는 것과 똑같이 파일을 하나
 * 만들었다가 지워봅니다. 이름에 밑줄(_)이 없어야 Apps Script
 * 편집기의 "실행할 함수" 드롭다운에 보입니다. 드롭다운에서
 * testAuth를 선택해 실행하면 동의 화면이 뜹니다 — 승인한 뒤에는
 * 이 함수를 지우고 다시 배포해도 되고, 그냥 남겨둬도 동작에는
 * 영향이 없습니다.
 **************************************************************/
function testAuth() {
  UrlFetchApp.fetch("https://www.google.com");

  const tempFile = DriveApp.createFile("유상건정리_권한테스트_임시파일.txt", "test");
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

    if (action === "carryOverPreviousMonth") {
      return jsonOutput_(carryOverPreviousMonthAction_());
    }

    if (action === "cleanupAddedItems") {
      return jsonOutput_(cleanupAddedItemsAction_());
    }

    if (action === "cleanupByBrand") {
      return jsonOutput_(cleanupByBrandAction_());
    }

    if (action === "exportPaidClosing") {
      return jsonOutput_(exportSheetsSubsetAsXlsxBase64_(["퍼시스", "일룸"], "유상마감"));
    }

    if (action === "exportPaidClosingDetail") {
      return jsonOutput_(exportSheetsSubsetAsXlsxBase64_(["퍼시스내역서", "일룸내역서"], "유상마감내역서"));
    }

    return jsonOutput_({ error: "알 수 없는 action입니다: " + action });
  } catch (error) {
    return jsonOutput_({ error: error.message });
  }
}


/**************************************************************
 * "추가건정리" 액션
 *
 * 1. "이번달"/"전월" 시트를 접수번호+순번+부품코드+색상+금액 조합으로
 *    비교해서, 이번달에만 있고 전월에는 없는(=이번달에 새로 추가된)
 *    행만 골라냅니다.
 * 2. 그 중 발주일자 또는 접수번호가 빈 행은 제외합니다.
 * 3. 남은 행 전체를 기준으로, 모든 행에서 값이 하나도 없는 열은
 *    헤더까지 통째로 지웁니다(예: 긴급/도착지/미확정사유/생산사업장/
 *    하자구분처럼 이번 배치에서 전혀 안 쓰인 열).
 * 4. 그 결과(정리된 헤더 + 행)를 "정리" 시트에 덮어씁니다
 *    (기존 내용은 전부 지우고 다시 씀).
 *
 * 예전에는 이어서 퍼시스/일룸 등 브랜드별로 나누는 것까지 한 번에
 * 했지만, 그 부분은 "브랜드별정리" 버튼(cleanupByBrandAction_)으로
 * 분리했습니다 — 이 액션은 "정리" 시트를 만드는 데까지만 합니다.
 **************************************************************/
/**************************************************************
 * "전월데이터 이월하기" 액션
 *
 * "이번달" 시트의 현재 값+서식을 "전월" 시트에 그대로 덮어씁니다
 * (전월 시트 기존 내용은 전부 지우고 씀). 매달 마감 후 다음 달로
 * 넘어갈 때, 지난달 마감한 "이번달" 시트를 "전월" 시트로 옮기는
 * 수동 작업(복사해서 전월에 붙여넣기)을 대신합니다.
 **************************************************************/
function carryOverPreviousMonthAction_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getSheetByName(CURRENT_MONTH_SHEET_NAME);
  const previousSheet = ss.getSheetByName(PREVIOUS_MONTH_SHEET_NAME);

  if (!currentSheet) throw new Error("'" + CURRENT_MONTH_SHEET_NAME + "' 시트를 찾을 수 없습니다.");
  if (!previousSheet) throw new Error("'" + PREVIOUS_MONTH_SHEET_NAME + "' 시트를 찾을 수 없습니다.");

  const sourceRange = currentSheet.getDataRange();
  const rowCount = sourceRange.getNumRows();
  const columnCount = sourceRange.getNumColumns();

  previousSheet.clear();
  sourceRange.copyTo(previousSheet.getRange(1, 1, rowCount, columnCount));

  for (let col = 1; col <= columnCount; col++) {
    previousSheet.setColumnWidth(col, currentSheet.getColumnWidth(col));
  }

  return { ok: true, rowCount: Math.max(rowCount - 1, 0) };
}


function cleanupAddedItemsAction_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getSheetByName(CURRENT_MONTH_SHEET_NAME);
  const previousSheet = ss.getSheetByName(PREVIOUS_MONTH_SHEET_NAME);

  if (!currentSheet) throw new Error("'" + CURRENT_MONTH_SHEET_NAME + "' 시트를 찾을 수 없습니다.");
  if (!previousSheet) throw new Error("'" + PREVIOUS_MONTH_SHEET_NAME + "' 시트를 찾을 수 없습니다.");

  const current = readSheetObject_(currentSheet);
  const previous = readSheetObject_(previousSheet);
  const header = current.header;

  // "이번달"과 "전월" 시트의 열 순서가 같다고 가정하지 않도록, 각 시트
  // 자신의 헤더에서 키 열 인덱스를 따로 구합니다(둘의 열 순서가 어긋나면
  // 엉뚱한 열끼리 비교하게 되어 중복/신규 판정이 잘못될 수 있었습니다).
  function buildKeyColumnIndexes_(sheetHeader, sheetLabelForError) {
    return COMPARE_KEY_COLUMN_LABELS.map(function(label) {
      const idx = sheetHeader.indexOf(label);
      if (idx === -1) throw new Error("'" + sheetLabelForError + "'에서 '" + label + "' 열을 찾을 수 없습니다.");
      return idx;
    });
  }

  const currentKeyColumnIndexes = buildKeyColumnIndexes_(header, CURRENT_MONTH_SHEET_NAME);
  const previousKeyColumnIndexes = buildKeyColumnIndexes_(previous.header, PREVIOUS_MONTH_SHEET_NAME);

  function rowCompareKey_(values, keyColumnIndexes) {
    return keyColumnIndexes.map(function(idx) { return normalizeText_(values[idx]); }).join("||");
  }

  const previousKeys = new Set(previous.rows.map(function(r) {
    return rowCompareKey_(r.values, previousKeyColumnIndexes);
  }));

  // 1) 이번달에 새로 추가된 행만
  let newRows = current.rows
    .map(function(r) { return r.values; })
    .filter(function(values) { return !previousKeys.has(rowCompareKey_(values, currentKeyColumnIndexes)); });

  // 2) 발주일자/접수번호 없는 행 제거
  const requiredIndexes = REQUIRED_NONBLANK_COLUMN_LABELS.map(function(label) {
    const idx = header.indexOf(label);
    if (idx === -1) throw new Error("'" + CURRENT_MONTH_SHEET_NAME + "'에서 '" + label + "' 열을 찾을 수 없습니다.");
    return idx;
  });

  newRows = newRows.filter(function(values) {
    return requiredIndexes.every(function(idx) { return normalizeText_(values[idx]) !== ""; });
  });

  // 3) 값이 하나도 없는 열은 헤더 포함 삭제 (남은 행이 없으면 원본 열 그대로 둠)
  let keepColIndexes = header.map(function(_, idx) { return idx; });

  if (newRows.length) {
    keepColIndexes = keepColIndexes.filter(function(idx) {
      return newRows.some(function(values) { return normalizeText_(values[idx]) !== ""; });
    });
  }

  const filteredHeader = keepColIndexes.map(function(idx) { return header[idx]; });
  const filteredRows = newRows.map(function(values) {
    return keepColIndexes.map(function(idx) { return values[idx]; });
  });

  // "정리" 시트도 열 너비·글꼴을 "전월" 시트 기준으로 맞춥니다
  // (열 이름으로 매칭 — 전월에 없는 열은 건너뜀).
  const referenceFont = getReferenceFontFromSheet_(previousSheet);
  const formatOptions = {
    fontFamily: referenceFont.family,
    fontSize: referenceFont.size,
    widthSourceSheet: previousSheet,
    widthSourceHeader: previous.header
  };

  // 4) "정리" 시트에 덮어쓰기
  const cleanupSheet = getOrCreateSheet_(ss, CLEANUP_SHEET_NAME);
  writeSheetFullReplace_(cleanupSheet, filteredHeader, filteredRows, formatOptions);

  return {
    ok: true,
    newRowCount: filteredRows.length,
    removedColumnCount: header.length - filteredHeader.length
  };
}


/**************************************************************
 * "브랜드별정리" 액션
 *
 * "추가건정리"로 이미 만들어져 있는 "정리" 시트 내용을 접수번호
 * 앞글자 기준으로 나눠, F/P로 시작하면 "퍼시스" 시트에, I로
 * 시작하면 "일룸" 시트에 각각 덮어씁니다(맨 아래에 금액 합계 행을
 * 추가). 이어서 퍼시스/일룸 시트에서 DETAIL_COLUMN_LABELS 순서로
 * 열만 뽑아 "퍼시스내역서"/"일룸내역서" 시트에도 각각 덮어씁니다
 * (역시 맨 아래에 금액 합계 행을 추가). "정리" 시트가 아직 없으면
 * (추가건정리를 먼저 실행하지 않았으면) 에러를 던집니다.
 **************************************************************/
function cleanupByBrandAction_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cleanupSheet = ss.getSheetByName(CLEANUP_SHEET_NAME);
  const previousSheet = ss.getSheetByName(PREVIOUS_MONTH_SHEET_NAME);

  if (!cleanupSheet) {
    throw new Error("'" + CLEANUP_SHEET_NAME + "' 시트를 찾을 수 없습니다. 먼저 추가건정리를 실행하세요.");
  }
  if (!previousSheet) throw new Error("'" + PREVIOUS_MONTH_SHEET_NAME + "' 시트를 찾을 수 없습니다.");

  const cleaned = readSheetObject_(cleanupSheet);
  const filteredHeader = cleaned.header;
  const filteredRows = cleaned.rows.map(function(r) { return r.values; });
  const previous = readSheetObject_(previousSheet);

  const referenceFont = getReferenceFontFromSheet_(previousSheet);
  const formatOptions = {
    fontFamily: referenceFont.family,
    fontSize: referenceFont.size,
    widthSourceSheet: previousSheet,
    widthSourceHeader: previous.header
  };

  const accessionIdx = filteredHeader.indexOf(ACCESSION_COLUMN_LABEL);

  if (accessionIdx === -1) {
    throw new Error("'" + CLEANUP_SHEET_NAME + "'에서 '" + ACCESSION_COLUMN_LABEL + "' 열을 찾을 수 없습니다.");
  }

  const detailColIndexes = DETAIL_COLUMN_LABELS.map(function(label) {
    const idx = filteredHeader.indexOf(label);
    if (idx === -1) throw new Error("내역서용 '" + label + "' 열을 '" + CLEANUP_SHEET_NAME + "'에서 찾을 수 없습니다.");
    return idx;
  });

  const numberColIdx = filteredHeader.indexOf(NUMBER_COLUMN_LABEL);
  const vendorCounts = {};

  VENDOR_SHEETS_BY_PREFIX.forEach(function(vendor) {
    const vendorRows = filteredRows
      .filter(function(values) {
        const accession = normalizeText_(values[accessionIdx]).toUpperCase();
        return vendor.prefixes.some(function(prefix) { return accession.indexOf(prefix) === 0; });
      })
      .map(function(values) { return values.slice(); }); // 번호 재부여를 위해 복사본 사용

    // 퍼시스/일룸/내역서 시트 안에서는 "번호"를 1번부터 다시 순서대로 매김
    if (numberColIdx !== -1) {
      vendorRows.forEach(function(values, i) { values[numberColIdx] = i + 1; });
    }

    const vendorSheet = getOrCreateSheet_(ss, vendor.sheetName);
    writeSheetFullReplace_(
      vendorSheet, filteredHeader, vendorRows,
      Object.assign({ appendAmountTotal: true }, formatOptions)
    );
    vendorCounts[vendor.sheetName] = vendorRows.length;

    // "내역서" 시트: 위 vendorRows에서 DETAIL_COLUMN_LABELS 순서로 열만 뽑아서 씀
    const detailRows = vendorRows.map(function(values) {
      return detailColIndexes.map(function(idx) { return values[idx]; });
    });

    const detailSheet = getOrCreateSheet_(ss, vendor.detailSheetName);
    writeSheetFullReplace_(
      detailSheet, DETAIL_COLUMN_LABELS, detailRows,
      Object.assign({ appendAmountTotal: true }, formatOptions)
    );
  });

  return { ok: true, vendorCounts: vendorCounts };
}


/**************************************************************
 * name 시트가 있으면 그대로, 없으면 새로 만들어서 반환합니다.
 **************************************************************/
function getOrCreateSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}


/**************************************************************
 * sheet의 기존 내용을 전부 지우고 header + rows로 다시 채웁니다.
 *
 * 조건부 서식 규칙을 먼저 전부 지웁니다(sheet.clear()로는 안 지워짐 —
 * 예전에 남은 규칙이 글자를 안 보이게 만드는 원인이 될 수 있음).
 *
 * 순서가 중요합니다: 서식(날짜/금액/일반 텍스트)과 열 너비는 값을
 * 쓰기 "전에" 지정합니다. 값부터 쓰고 서식을 나중에 지정하면,
 * "01"처럼 0으로 시작하는 텍스트가 setValues 시점에 이미 숫자 1로
 * 자동 변환되어 버려서 나중에 서식을 맞춰도 되돌릴 수 없습니다.
 * 그래서 DATE_FORMAT_COLUMN_LABELS/AMOUNT_FORMAT_COLUMN_LABELS에 해당
 * 하는 열은 미리 날짜 형식/천단위 콤마로, 그 외 모든 열은 "일반
 * 텍스트(@)"로 값을 쓰기 전에 지정해둡니다("전월" 서식을 그대로
 * 복사하는 방식은 전월 쪽이 항상 텍스트로 되어 있다는 보장이 없어서
 * 불안정했습니다). 이렇게 미리 지정해두면 값을 쓸 때 그 서식에 맞춰
 * 바로 파싱되므로, 나중에 또 열별로 서식을 지정할 필요가 없어
 * 속도도 더 빠릅니다(열 하나하나마다 API를 따로 호출하는 대신 전체
 * 범위를 setNumberFormats로 한 번에 지정합니다. 열 너비도 같은 값이
 * 연속된 열은 setColumnWidths로 묶어서 호출 횟수를 줄입니다).
 *
 * 행 높이는 반드시 setRowHeightsForced로 고정합니다. 일반
 * setRowHeight/setRowHeights는 "초기값"만 지정할 뿐이라, 변경사유/
 * 특이사항처럼 줄바꿈이 든 긴 텍스트가 있으면 시트가 내용에 맞춰
 * 나중에 다시 자동으로 행을 늘려버립니다(값을 쓰기 전/후 어느
 * 시점에 호출하든 마찬가지). setRowHeightsForced를 써야 셀 내용과
 * 무관하게 항상 지정한 높이로 고정됩니다.
 *
 * options.appendAmountTotal이 true면, header에서 "금액" 열을 찾아
 * 맨 아래에 그 열의 합계 행("합계" 라벨 + 합계 값)을 하나 추가하고,
 * 그 합계 셀은 노란색으로 채웁니다(header에 "금액" 열이 없으면 합계
 * 행 없이 그대로 씁니다).
 *
 * options.fontFamily/fontSize를 주면 전체 범위에 그 글꼴을 지정하고,
 * 글자색은 항상 검정(DEFAULT_FONT_COLOR)으로 강제합니다. 줄바꿈이
 * 들어간 긴 텍스트(예: 변경사유/특이사항) 때문에 행이 늘어나지
 * 않도록 넘치는 내용은 자르는 서식(CLIP)도 강제합니다 — 행 높이는
 * 항상 DEFAULT_ROW_HEIGHT로 고정되고, 셀을 클릭하면 전체 내용을
 * 볼 수 있습니다.
 *
 * options.widthSourceSheet/widthSourceHeader를 주면, 열 이름이 같은
 * 열끼리 그 시트의 열 너비만 그대로 복사합니다.
 * sheet.clear()는 값과 서식을 모두 지우기 때문에, 매번 다시 채운 뒤
 * 이 서식들을 명시적으로 다시 지정합니다.
 **************************************************************/
function writeSheetFullReplace_(sheet, header, rows, options) {
  sheet.clear();

  // 예전에 만들어진 조건부 서식 규칙이 남아있으면(예: 특정 값일 때
  // 글자를 흰색으로 숨기는 규칙) API로 글자색을 검정으로 지정해도
  // 화면에서는 그 규칙이 덮어써서 계속 안 보이는 것처럼 보입니다.
  // 이 시트들은 전부 스크립트가 관리하는 결과물이라 조건부 서식이
  // 남아있을 이유가 없으므로 매번 깨끗하게 지웁니다.
  sheet.setConditionalFormatRules([]);

  const columnCount = header.length;

  if (!columnCount) {
    return;
  }

  const data = [header].concat(rows.map(function(row) {
    return normalizeRowLength_(row, columnCount);
  }));

  let totalRowIndex = -1;
  let totalAmountColIndex = -1;

  if (options && options.appendAmountTotal) {
    const amountIdx = header.indexOf(AMOUNT_COLUMN_LABEL);

    if (amountIdx !== -1) {
      const total = rows.reduce(function(sum, row) {
        const num = Number(String(row[amountIdx] === undefined || row[amountIdx] === null ? "" : row[amountIdx]).replace(/,/g, ""));
        return sum + (isNaN(num) ? 0 : num);
      }, 0);

      const totalRow = new Array(columnCount).fill("");
      totalRow[0] = AMOUNT_TOTAL_LABEL;
      totalRow[amountIdx] = total;
      data.push(totalRow);

      totalRowIndex = data.length; // 헤더가 1행이므로 이 값이 곧 시트 행 번호
      totalAmountColIndex = amountIdx;
    }
  }

  const bodyRowCount = data.length - 1;

  // 값을 쓰기 "전에" 먼저 서식을 지정합니다(열마다 API를 따로 호출하면
  // 열 개수만큼 요청이 늘어나 느려지므로, 전체 범위를 한 번에
  // setNumberFormats로 지정합니다).
  // - 날짜/금액 열은 각각 날짜 형식/천단위 콤마로, 그 외 열은 "일반
  //   텍스트(@)"로 미리 정해서 값을 씁니다. "01"처럼 0으로 시작하는
  //   텍스트가 setValues 시점에 숫자 1로 자동 변환되어 앞의 0이
  //   사라지는 문제를 막으려면 텍스트 서식을 값보다 먼저 지정해야
  //   합니다("전월" 서식을 그대로 복사하는 방식은 전월 쪽 서식이 항상
  //   텍스트로 되어 있다는 보장이 없어서 불안정했습니다). 날짜/금액
  //   서식도 미리 지정해두면, 값을 쓸 때 그 서식에 맞춰 바로
  //   숫자/날짜로 파싱되어 나중에 또 지정할 필요가 없습니다.
  // - 열 너비는 "전월"(widthSourceSheet) 기준으로 맞추되, 같은 너비가
  //   연속된 열은 setColumnWidths로 한 번에 묶어서 호출 횟수를
  //   줄입니다.
  const numberFormatRow = header.map(function(label) {
    if (DATE_FORMAT_COLUMN_LABELS.indexOf(label) !== -1) return DATE_NUMBER_FORMAT;
    if (AMOUNT_FORMAT_COLUMN_LABELS.indexOf(label) !== -1) return AMOUNT_NUMBER_FORMAT;
    return "@";
  });
  const numberFormats = data.map(function() { return numberFormatRow.slice(); });

  sheet.getRange(1, 1, data.length, columnCount).setNumberFormats(numberFormats);

  if (options && options.widthSourceSheet && options.widthSourceHeader) {
    let runStartIdx = -1;
    let runWidth = null;

    for (let idx = 0; idx <= columnCount; idx++) {
      const sourceIdx = idx < columnCount ? options.widthSourceHeader.indexOf(header[idx]) : -1;
      const width = idx < columnCount && sourceIdx !== -1
        ? options.widthSourceSheet.getColumnWidth(sourceIdx + 1)
        : null;

      if (width !== runWidth) {
        if (runStartIdx !== -1 && runWidth !== null) {
          sheet.setColumnWidths(runStartIdx + 1, idx - runStartIdx, runWidth);
        }
        runStartIdx = idx;
        runWidth = width;
      }
    }
  }

  const fullRange = sheet.getRange(1, 1, data.length, columnCount);
  fullRange.setValues(data);

  if (options && options.fontFamily) {
    fullRange.setFontFamily(options.fontFamily);
  }

  if (options && options.fontSize) {
    fullRange.setFontSize(options.fontSize);
  }

  // 글자색이 안 보이게 남아있던 문제를 막기 위해 항상 검정으로 강제
  fullRange.setFontColor(DEFAULT_FONT_COLOR);

  // 변경사유/특이사항처럼 줄바꿈이 들어간 긴 텍스트가 있어도 행이
  // 늘어나지 않도록, 넘치는 부분은 자르고(CLIP) 셀을 클릭해야 전체
  // 내용이 보이게 합니다.
  fullRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  // 행 높이는 반드시 setRowHeightsForced로 고정합니다. setRowHeight/
  // setRowHeights는 "초기값"만 지정할 뿐이라, 셀 내용(특히 줄바꿈이
  // 든 긴 텍스트)에 맞춰 시트가 나중에 다시 자동으로 늘려버립니다.
  // setRowHeightsForced를 써야 내용과 무관하게 항상 이 높이로
  // 고정되고, 넘치는 내용은 위에서 지정한 CLIP 서식대로 잘려서
  // 보입니다(셀을 클릭하면 전체 내용 확인 가능).
  sheet.setRowHeightsForced(1, 1, DEFAULT_ROW_HEIGHT);

  if (bodyRowCount > 0) {
    sheet.setRowHeightsForced(2, bodyRowCount, DEFAULT_ROW_HEIGHT);
  }

  sheet.setFrozenRows(1);

  // 날짜/금액 서식은 값을 쓰기 전에 이미 setNumberFormats로 지정해둬서
  // (위 참고) 여기서 또 열별로 지정할 필요가 없습니다.

  if (totalRowIndex !== -1 && totalAmountColIndex !== -1) {
    sheet.getRange(totalRowIndex, totalAmountColIndex + 1).setBackground(AMOUNT_TOTAL_HIGHLIGHT_COLOR);
  }
}


/**************************************************************
 * sheet의 기준 행(2행이 있으면 2행, 없으면 1행)의 글꼴 모양/크기를
 * 반환합니다. writeSheetFullReplace_에 그대로 넘겨서, 새로 채운
 * 다른 시트가 이 시트와 같은 글꼴로 보이게 맞출 때 씁니다.
 **************************************************************/
function getReferenceFontFromSheet_(sheet) {
  const referenceRow = sheet.getLastRow() >= 2 ? 2 : 1;
  const cell = sheet.getRange(referenceRow, 1);

  return { family: cell.getFontFamily(), size: cell.getFontSize() };
}


/**************************************************************
 * sheet 이름 또는 gid로 시트를 찾습니다. 둘 다 없으면 DEFAULT_GID를
 * 씁니다. 못 찾으면 에러를 던집니다.
 **************************************************************/
function resolveSheet_(sheetName, gidParam) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error("'" + sheetName + "' 시트를 찾을 수 없습니다.");
    return sheet;
  }

  const gid = gidParam !== undefined && gidParam !== null && gidParam !== ""
    ? Number(gidParam)
    : DEFAULT_GID;

  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === gid) return sheets[i];
  }

  throw new Error("gid " + gid + "에 해당하는 시트를 찾을 수 없습니다.");
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

    // rowIndex가 1(헤더 행)이거나 시트 범위를 벗어나면(예: 다른 곳에서
    // 그 사이 행이 삭제됨) 저장하지 않고 충돌로 표시합니다. 이 체크가
    // 없으면 rowIndex=1인 잘못된 요청이 헤더 행을 조용히 덮어쓸 수
    // 있었습니다.
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
 * 이 스프레드시트에서 sheetNames에 해당하는 시트만 임시 스프레드시트에
 * 복사해서 xlsx로 내보낸 뒤, 임시 스프레드시트는 지웁니다
 * ("유상마감다운로드"/"유상마감내역서다운로드" 버튼에서 사용).
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


// 발주일자/입고요청일/입고확정일/출고예정일처럼 날짜 서식이 적용된 셀은
// getValues()로 읽으면 실제 Date 객체로 돌아오는데, 화면이 처음 받았던
// 값(JSON 직렬화된 날짜 문자열, 예: "2026-07-23T15:00:00.000Z")과 그냥
// String()으로 비교하면 표현이 달라서(Date는 "Thu Jul 23 2026 ..." 형태로
// 나옴) 실제로는 안 바뀐 행도 저장할 때마다 "충돌"로 잘못 감지되는
// 문제가 있었습니다. Date 객체와 그 JSON 직렬화 문자열을 모두
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
