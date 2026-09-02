/**************************************************************
 * 월현황(주간) 대시보드 연동용 Web App
 *
 * 배포 위치
 * ------------------------------------------------------------
 * 이 파일은 "월현황" 스프레드시트
 * (1bpeFlrwVzpr9WSLJJDSOJq_7KJLJ2wZFt3cs-8am8Vk)의
 * Apps Script 프로젝트에 추가합니다. AS현황 스프레드시트
 * (1ojKHospTSYSfKWlidRwvqyWhuHGfnsrLWRqVpn2SFBU)는
 * openById로 읽기만 하므로, 배포 실행 계정("나")이 두 시트
 * 모두에 접근 권한을 가지고 있어야 합니다.
 *
 * 하는 일
 * ------------------------------------------------------------
 * 화면에서 "이번 주 저장"을 누르면 1~5주 중 저장할 주차를 고른 뒤
 * week 값을 포함해 저장 요청을 보냅니다. 요청을 받으면
 * 1. AS현황 스프레드시트의 1공장추가건 / 2공장추가건을 읽어 그대로
 *    합치고, Q열(등록일, index 16)은 선택한 주차 문자열로 통일합니다
 *    (AS현황 원본은 그대로 두고 복사본만 바꿈).
 * 2. 선택한 주차의 탭("N주(숫자)" 패턴, 예: 3주(45))을 갱신합니다:
 *    포장(W열)이 "미회수"인 기존 행은 지우고, 남은 기존 행 + 새로
 *    들어온 행을 D열(접수번호)+G열(고객명) 기준으로 합쳐서 기존
 *    정렬 규칙으로 다시 정렬한 뒤 탭 전체를 다시 씁니다. 없으면
 *    새로 만듭니다.
 * 3. 종합(N) 시트(추가건과 같은 26열 구조)를 갱신합니다:
 *    - 포장이 "미회수"인 행은 주차와 상관없이 전부 삭제
 *    - 제조자(X열)별로, 이번 주차로 이미 저장된 행(미회수 제외) +
 *      새로 들어온 행을 합쳐서 같은 규칙으로 다시 정렬한 뒤,
 *      "충주1"/"충주2" 블록 마지막 자리에 다시 삽입
 *      (충주2 블록이 아예 없으면 1공장 삽입분 아래로 4행을 띄우고 삽입)
 *
 * 배포 방법
 * ------------------------------------------------------------
 * 1. 월현황 스프레드시트 > 확장 프로그램 > Apps Script에 이 파일 추가
 * 2. 배포 > 새 배포 > 유형: 웹 앱, 실행 계정: 나, 액세스 권한: 필요 범위
 * 3. 배포 후 나오는 웹 앱 URL을 대시보드의 "월현황(주간)" 연결 정보에 입력
 *
 * 주의
 * ------------------------------------------------------------
 * - 토큰 검증이 없으므로 URL을 아는 사람은 누구나 읽기/쓰기가 가능합니다.
 * - 유형 분류(일반/감성/취급/VN/미회수) 규칙은
 *   factory-as-cleanup.gs의 sortAndFormat_ 그룹 분류와 동일하게
 *   맞춰뒀습니다. 그쪽 로직이 바뀌면 classifyGroup_도 같이 맞춰야 합니다.
 **************************************************************/

const AS_SPREADSHEET_ID = "1ojKHospTSYSfKWlidRwvqyWhuHGfnsrLWRqVpn2SFBU";

const FACTORY_SHEETS = [
  { label: "1공장", addSheet: "1공장추가건" },
  { label: "2공장", addSheet: "2공장추가건" }
];

// 종합 시트 이름은 "종합(76)"처럼 뒤에 전체 건수가 붙어서 매번 바뀝니다.
// 정확한 이름 대신 이 패턴으로 찾습니다.
const SUMMARY_SHEET_PATTERN = /^종합\(\d+\)$/;
const SUMMARY_SHEET_PLACEHOLDER_NAME = "종합(0)";

// 종합(N) 시트는 추가건과 같은 26열 구조입니다(공장 열 없음).
const SUMMARY_ROW_HEADER = [
  "최종조치일", "브랜드", "지역센터", "접수번호", "구분", "형태",
  "고객명", "부품명", "제품코드", "색상", "수량", "금액", "조치결과",
  "회수구분", "서비스요구내역", "반납일자", "등록일", "유형", "세부유형",
  "하자상세", "로트", "원인", "포장", "제조자", "유형분류", "유형분류"
];

const MANUFACTURER_BY_FACTORY_LABEL = {
  "1공장": "충주1",
  "2공장": "충주2"
};

// 종합(N) 시트에 충주2 블록이 없을 때, 충주1 삽입분 아래 두는 빈 행 수
const SUMMARY_GAP_ROWS = 4;

// 1주~5주 탭도 "1주(23)"처럼 뒤에 숫자가 붙어서 매번 바뀝니다.
const WEEK_NUMBERS = [1, 2, 3, 4, 5];

// VN/미회수 행 글자색 (그 외 행은 검정색)
const VN_FONT_COLOR = "#1155cc";
const UNCOLLECTED_FONT_COLOR = "#cc0000";
const DEFAULT_FONT_COLOR = "#000000";

// 종합(N) 원본 26열 기준 W열(포장) 인덱스 (classifyGroup_과 동일)
const NATIVE_PACKAGE_COLUMN_INDEX = 22;

// 종합(N) 원본 26열 기준 J열(색상) 인덱스. "061"처럼 앞자리 0이 있는
// 값을 텍스트로 유지하려면 새로 쓰는 범위는 항상 값을 쓰기 전에
// 이 열만 먼저 "@"(텍스트) 서식으로 지정해야 합니다 — 서식이
// "General"인 채로 쓰면 그 순간 숫자로 재해석되어 0이 사라집니다.
const NATIVE_COLOR_COLUMN_INDEX = SUMMARY_ROW_HEADER.indexOf("색상");

// 새로 쓰는/삽입하는 데이터의 글꼴
const DEFAULT_FONT_FAMILY = "Malgun Gothic";
const DEFAULT_FONT_SIZE = 10;

// 날짜로 표시할 열 이름과, 기존 시트와 맞출 날짜 표시 형식
const DATE_FORMAT_COLUMN_LABELS = ["최종조치일", "반납일자"];
const DATE_NUMBER_FORMAT = "yyyy-mm-dd";

// 천단위 콤마로 표시할 열 이름과 숫자 표시 형식
const AMOUNT_FORMAT_COLUMN_LABELS = ["금액"];
const AMOUNT_NUMBER_FORMAT = "#,##0";

// "로데이터" 시트: 종합(N)에서 이 순서의 열만 뽑고, 맨 뒤에 주차 열을
// 새로 추가합니다(주차 값 = 등록일 값 그대로).
const RAW_DATA_SHEET_NAME = "로데이터";
const RAW_DATA_SOURCE_COLUMN_LABELS = [
  "최종조치일", "브랜드", "접수번호", "형태", "제품코드", "색상",
  "조치결과", "서비스요구내역", "등록일", "유형", "세부유형",
  "하자상세", "로트", "원인", "포장"
];
const RAW_DATA_HEADER = RAW_DATA_SOURCE_COLUMN_LABELS.concat(["주차"]);

// "마감" 시트: 종합(N)과 같은 26열 그대로, 빈 행과 W열(포장) VN/미회수
// 행만 제거해서 옮겨 놓는 시트(예전 "월마감다운로드" 버튼이 만들던 자료).
// 종합(N)/N주(N)처럼 이름에 건수를 붙여 "마감(N)"으로 관리합니다.
const CLOSING_SHEET_NAME = "마감";
const CLOSING_SHEET_PATTERN = /^마감(\(\d+\))?$/;


/**************************************************************
 * 권한 재승인용 임시 테스트 함수
 *
 * exportFull이 쓰는 UrlFetchApp(외부 요청) 권한을 승인받기 위한
 * 함수입니다. 이름에 밑줄(_)이 없어야 Apps Script 편집기의 "실행할
 * 함수" 드롭다운에 보입니다. 드롭다운에서 testAuth를 선택해 실행하면
 * 동의 화면이 뜹니다 — 승인한 뒤에는 이 함수를 지우고 다시 배포해도
 * 되고, 그냥 남겨둬도 동작에는 영향이 없습니다.
 **************************************************************/
function testAuth() {
  UrlFetchApp.fetch("https://www.google.com");
}


function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    const action = params.action || "summary";

    if (action === "summary") {
      return jsonOutput_(readSheetObject_(findSummarySheet_()));
    }

    if (action === "weekByNumber") {
      const weekNumber = Number(params.week);

      if (WEEK_NUMBERS.indexOf(weekNumber) === -1) {
        return jsonOutput_({ error: "week는 1~5 중 하나여야 합니다." });
      }

      return jsonOutput_(readSheetObject_(findWeekSheet_(weekNumber)));
    }

    if (action === "closing") {
      return jsonOutput_(readSheetObject_(findClosingSheet_()));
    }

    if (action === "rawData") {
      return jsonOutput_(readSheetSafe_(RAW_DATA_SHEET_NAME));
    }

    if (action === "weeks") {
      return jsonOutput_({ weeks: listWeekSheetNames_() });
    }

    if (action === "week") {
      return jsonOutput_(readSheetSafe_(params.name || ""));
    }

    return jsonOutput_({ error: "알 수 없는 action입니다: " + action });
  } catch (error) {
    return jsonOutput_({ error: error.message });
  }
}


function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");

    if (body.action === "archiveWeek") {
      const weekNumber = Number(body.week);

      if (WEEK_NUMBERS.indexOf(weekNumber) === -1) {
        return jsonOutput_({ error: "week는 1~5 중 하나여야 합니다." });
      }

      return jsonOutput_(archiveWeekAction_(weekNumber));
    }

    if (body.action === "removeEntries") {
      return jsonOutput_(removeEntriesAction_(body.entries || []));
    }

    if (body.action === "matchAccumulated") {
      return jsonOutput_(matchAccumulatedDataAction_());
    }

    if (body.action === "clearMonthlyData") {
      return jsonOutput_(clearMonthlyDataAction_());
    }

    if (body.action === "exportFull") {
      return jsonOutput_(exportFullWorkbookAction_());
    }

    return jsonOutput_({ error: "알 수 없는 action입니다: " + body.action });
  } catch (error) {
    return jsonOutput_({ error: error.message });
  }
}


/**************************************************************
 * "이번 주 저장" 액션
 *
 * 화면에서 선택한 주차(1~5) 기준으로 저장합니다.
 * 원본 Q열(등록일, index 16) 값은 선택한 주차 문자열로
 * 통일해서 씁니다(AS현황 원본 시트는 건드리지 않고,
 * 월현황 쪽 복사본만 바꿉니다).
 *
 * 같은 주차를 다시 저장해도 이미 저장된(미회수 제외) 행은 지우지
 * 않고 그대로 유지한 채, 새로 들어온 행만 추가합니다(N주 시트,
 * 종합(N) 모두 D열 접수번호 + G열 고객명 기준으로 비교). 다만
 * 미회수 행은 해소 여부가 계속 바뀔 수 있어서 주차와 상관없이
 * 매번 지우고 이번 저장 시점의 최신 상태로 다시 채웁니다. 실제
 * 정렬/병합 로직은 writeWeeklySheet_, updateSummarySheetForWeek_
 * 안에서 처리합니다.
 **************************************************************/
function archiveWeekAction_(weekNumber) {
  const weekLabel = weekLabelFor_(weekNumber);
  const asSpreadsheet = SpreadsheetApp.openById(AS_SPREADSHEET_ID);

  let combinedHeader = null;
  const counts = {};
  const factoryRawRows = {};
  const combinedRows = [];

  FACTORY_SHEETS.forEach(function(factory) {
    const sheet = asSpreadsheet.getSheetByName(factory.addSheet);

    if (!sheet) {
      throw new Error(
        "AS현황 스프레드시트에서 '" + factory.addSheet + "' 시트를 찾을 수 없습니다."
      );
    }

    counts[factory.label] = { "일반": 0, "감성": 0, "취급": 0, "VN": 0, "미회수": 0 };
    factoryRawRows[factory.label] = [];

    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();

    if (lastRow < 1 || lastColumn < 1) {
      return;
    }

    const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();

    if (!combinedHeader) {
      combinedHeader = values[0];
    }

    for (let i = 1; i < values.length; i++) {
      const row = values[i].slice();
      row[16] = weekLabel; // Q열 등록일 → 선택한 주차로 통일

      const group = classifyGroup_(row);

      counts[factory.label][group]++;
      factoryRawRows[factory.label].push(row);
      combinedRows.push(row);
    }
  });

  if (!combinedHeader) {
    throw new Error("추가건 시트에서 데이터를 찾지 못했습니다.");
  }

  const weeklyResult = writeWeeklySheet_(weekNumber, combinedHeader, combinedRows);

  let summary;

  try {
    const summaryResult = updateSummarySheetForWeek_(weekLabel, factoryRawRows);

    // 마감/로데이터 둘 다 방금 갱신된 종합(N) 전체를 그대로 읽어서 쓰므로,
    // 한 번만 읽어서 같이 넘겨줍니다(둘 다 종합(N)을 고치지 않으니 안전).
    const summarySheetAfterUpdate = findSummarySheet_();
    const preloadedSummaryValues = summarySheetAfterUpdate &&
      summarySheetAfterUpdate.getLastRow() >= 1 &&
      summarySheetAfterUpdate.getLastColumn() >= 1
      ? summarySheetAfterUpdate.getRange(
          1, 1, summarySheetAfterUpdate.getLastRow(), summarySheetAfterUpdate.getLastColumn()
        ).getValues()
      : null;

    updateClosingSheet_(preloadedSummaryValues);
    updateRawDataSheet_(preloadedSummaryValues);
    summary = { ok: true, addedCount: summaryResult.addedCount };
  } catch (error) {
    summary = { ok: false, message: error.message };
  }

  return {
    ok: true,
    weekLabel: weekLabel,
    sheetName: weeklyResult.sheetName,
    rowCount: weeklyResult.totalCount,
    addedCount: weeklyResult.addedCount,
    counts: counts,
    summary: summary
  };
}


/**************************************************************
 * D열(접수번호, index 3) + G열(고객명, index 6) 조합으로 행을
 * 식별하는 키. N주 시트와 종합(N) 모두 공장 열이 없는 동일한
 * 26열 구조라 그대로 재사용합니다.
 **************************************************************/
function rowMatchKey_(row) {
  return normalizeText_(row[3]) + "||" + normalizeText_(row[6]);
}


/**************************************************************
 * rows를 factory-as-cleanup.gs의 sortAndFormat_과 같은 규칙(일반→
 * 감성→취급→VN→미회수, 그룹 내 다중 키 정렬)으로 묶고 정렬합니다.
 * 기존에 저장된 행 + 새로 들어온 행을 합친 뒤 다시 정렬할 때 씁니다.
 **************************************************************/
function groupAndSortRows_(rows) {
  const buckets = { "일반": [], "감성": [], "취급": [], "VN": [], "미회수": [] };

  rows.forEach(function(row) {
    buckets[classifyGroup_(row)].push(row);
  });

  ["일반", "감성", "취급", "VN", "미회수"].forEach(function(group) {
    sortRowsWithinGroup_(buckets[group]);
  });

  return buckets["일반"].concat(buckets["감성"], buckets["취급"], buckets["VN"], buckets["미회수"]);
}


/**************************************************************
 * "삭제건 제거" 액션
 *
 * 화면(1공장삭제건/2공장삭제건 탭)에서 보낸 접수번호+제품코드+색상
 * 조합과 일치하는 행을, 이 스프레드시트(월현황)의 모든 시트에서
 * 찾아 전부 삭제합니다. 헤더에 접수번호/제품코드/색상 열이 없는
 * 시트는 건너뜁니다(공장 열 유무에 상관없이 헤더 이름으로 찾음).
 * (원본 로데이터 시트에는 고객명 열이 없어 이전에는 삭제가 안 됐음)
 *
 * 종합(N), N주(N)처럼 이름에 건수가 들어있는 시트는 삭제 후
 * 그 숫자를 새 행 수로 갱신합니다.
 **************************************************************/
const MATCH_COLUMN_LABELS = { accession: "접수번호", productCode: "제품코드", color: "색상" };

function removeEntriesAction_(entries) {
  if (!Array.isArray(entries) || !entries.length) {
    throw new Error("삭제할 대상이 없습니다.");
  }

  const entryKeys = new Set(entries.map(function(entry) {
    return normalizeText_(entry.accessionNumber) + "||" +
      normalizeText_(entry.productCode) + "||" +
      normalizeText_(entry.colorName);
  }));

  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  const deletedBySheet = {};
  let totalDeleted = 0;

  sheets.forEach(function(sheet) {
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();

    if (lastRow < 2 || lastColumn < 1) {
      return;
    }

    const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
    const header = values[0];

    const accessionIdx = header.indexOf(MATCH_COLUMN_LABELS.accession);
    const productIdx = header.indexOf(MATCH_COLUMN_LABELS.productCode);
    const colorIdx = header.indexOf(MATCH_COLUMN_LABELS.color);

    if (accessionIdx === -1 || productIdx === -1 || colorIdx === -1) {
      return; // 매칭 열이 없는 시트는 건너뜀
    }

    function isMatchedRow(row) {
      const key = normalizeText_(row[accessionIdx]) + "||" +
        normalizeText_(row[productIdx]) + "||" +
        normalizeText_(row[colorIdx]);
      return entryKeys.has(key);
    }

    let deletedInSheet = 0;

    // 지울 행을 뒤에서부터 훑으면서 연속 구간으로 묶어 deleteRows()로
    // 한 번에 지웁니다(흩어진 행마다 deleteRow()를 따로 부르면 지울 행
    // 수만큼 API 호출이 들었음 — removeMatchingRows_와 같은 방식).
    let i = values.length - 1;

    while (i >= 1) {
      if (!isMatchedRow(values[i])) {
        i--;
        continue;
      }

      const runEndIndex = i;

      while (i >= 1 && isMatchedRow(values[i])) {
        i--;
      }

      const runStartRow = i + 2; // i는 구간 바로 위(0-based values 인덱스), 시트 행은 +1, 구간 시작은 그 다음 행
      const runLength = runEndIndex - i;

      sheet.deleteRows(runStartRow, runLength);
      deletedInSheet += runLength;
    }

    if (deletedInSheet > 0) {
      deletedBySheet[sheet.getName()] = deletedInSheet;
      totalDeleted += deletedInSheet;
      renumberSheetNameIfNeeded_(sheet);
    }
  });

  return { ok: true, deletedBySheet: deletedBySheet, totalDeleted: totalDeleted };
}


/**************************************************************
 * 종합(N), N주(N) 이름 패턴이면 삭제 후 행 수에 맞춰 이름 갱신
 **************************************************************/
function renumberSheetNameIfNeeded_(sheet) {
  const name = sheet.getName();

  if (SUMMARY_SHEET_PATTERN.test(name)) {
    const newName = "종합(" + countNonBlankDataRows_(sheet) + ")";
    if (name !== newName) sheet.setName(newName);
    return;
  }

  if (CLOSING_SHEET_PATTERN.test(name)) {
    const newName = "마감(" + countNonBlankDataRows_(sheet) + ")";
    if (name !== newName) sheet.setName(newName);
    return;
  }

  const weekMatch = name.match(/^(\d+)주\(\d+\)$/);

  if (weekMatch) {
    // N주(N)는 writeWeeklySheet_와 같은 이유로 VN/미회수 행을 뺀 건수를 씁니다.
    const newName = weekMatch[1] + "주(" + countWeekSheetNamedRows_(sheet) + ")";
    if (name !== newName) sheet.setName(newName);
  }
}


/**************************************************************
 * N주(N) 시트의 실제 현재 내용을 다시 읽어서, 포장이 VN/미회수인 행을
 * 뺀 건수를 셉니다(writeWeeklySheet_ 안에서는 이미 메모리에 있는
 * combinedRows로 직접 계산하므로 이 헬퍼가 필요 없고, "삭제건 제거"처럼
 * 시트를 직접 수정한 뒤 다시 세야 하는 경우에만 씁니다).
 **************************************************************/
function countWeekSheetNamedRows_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 2 || lastColumn < 1) {
    return 0;
  }

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const header = values[0];
  const packageIdx = header.indexOf("포장");

  return values.slice(1).filter(function(row) {
    const isBlank = row.every(function(cell) { return normalizeText_(cell) === ""; });
    if (isBlank) return false;
    return packageIdx === -1 || !isVnOrUncollectedPackage_(row[packageIdx]);
  }).length;
}


/**************************************************************
 * "1,2공장 누적데이터 매칭" 액션
 *
 * 종합(N) 시트(이 스프레드시트)를 읽어서, AS현황 스프레드시트의
 * '1공장 누적데이터' / '2공장 누적데이터' 시트를 다시 채웁니다.
 *
 * 1. 대상 누적데이터 시트는 헤더만 남기고 기존 내용을 전부 지움
 * 2. 종합(N)에서 제조자(X열)가 충주1/충주2와 일치하고,
 *    유형 또는 포장이 "미회수"가 아닌 행만 골라
 * 3. A~P열(16개 열)만 잘라서 붙여넣고, 글꼴(맑은 고딕)/글자크기(10)/
 *    글자색(검정)과 최종조치일·반납일자 열의 날짜 표시 형식,
 *    금액 열의 천단위 콤마 표시 형식을 매번 명시적으로 다시 지정함
 *    (기존 셀 서식을 그대로 복사하지 않음 — 예전에 한 번 깨진 서식이
 *    있으면 계속 복제되는 문제가 있어서 고침. 1공장/2공장 시트마다
 *    크기가 제각각이던 것도 이제 항상 10으로 통일됨)
 **************************************************************/
const ACCUMULATED_SHEET_BY_FACTORY_LABEL = {
  "1공장": "1공장 누적데이터",
  "2공장": "2공장 누적데이터"
};
const ACCUMULATED_COPY_COLUMN_COUNT = 16; // A~P열

function matchAccumulatedDataAction_() {
  const summarySheet = findSummarySheet_();

  if (!summarySheet) {
    throw new Error("종합(N) 시트를 찾을 수 없습니다.");
  }

  const lastRow = summarySheet.getLastRow();
  const lastColumn = summarySheet.getLastColumn();

  // 월이 갓 바뀌어 종합(N)에 아직 데이터 행이 하나도 없는 것(헤더만
  // 있는 lastRow === 1)은 정상적인 상태입니다 — 에러 없이 진행해서
  // 1공장/2공장 누적데이터 시트를 그냥 빈 채로(헤더만) 맞춰줍니다.
  // 헤더조차 없는 완전히 빈 시트일 때만 에러로 알립니다.
  if (lastRow < 1 || lastColumn < 1) {
    throw new Error("종합(N) 시트에 헤더가 없습니다.");
  }

  const values = summarySheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const header = values[0];

  const manufacturerIdx = header.indexOf("제조자");
  const typeIdx = header.indexOf("유형");
  const packageIdx = header.indexOf("포장");

  if (manufacturerIdx === -1) {
    throw new Error("종합(N) 시트에서 '제조자' 열을 찾을 수 없습니다.");
  }

  const asSpreadsheet = SpreadsheetApp.openById(AS_SPREADSHEET_ID);
  const counts = {};

  Object.keys(ACCUMULATED_SHEET_BY_FACTORY_LABEL).forEach(function(factoryLabel) {
    const manufacturerValue = MANUFACTURER_BY_FACTORY_LABEL[factoryLabel];
    const targetSheetName = ACCUMULATED_SHEET_BY_FACTORY_LABEL[factoryLabel];
    const targetSheet = asSpreadsheet.getSheetByName(targetSheetName);

    if (!targetSheet) {
      throw new Error("'" + targetSheetName + "' 시트를 찾을 수 없습니다.");
    }

    const filteredRows = [];

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const isManufacturerMatch = normalizeText_(row[manufacturerIdx]) === manufacturerValue;

      const isUncollected =
        (typeIdx !== -1 && normalizeText_(row[typeIdx]) === "미회수") ||
        (packageIdx !== -1 && normalizeText_(row[packageIdx]) === "미회수");

      if (isManufacturerMatch && !isUncollected) {
        filteredRows.push(row.slice(0, ACCUMULATED_COPY_COLUMN_COUNT));
      }
    }

    clearSheetKeepHeader_(targetSheet);

    if (filteredRows.length) {
      // 색상 열은 값을 쓰기 전에 텍스트("@") 서식부터 지정합니다. 종합(N)에서
      // 읽어온 "061" 같은 값이 서식이 General인 채로 쓰이면 그 즉시 숫자로
      // 바뀌어 앞자리 0이 사라집니다(writeWeeklySheet_ 등과 같은 이유의
      // 같은 수정 — 이 지점은 빠져 있어서 "N공장삭제건"에 색상이 61로
      // 나오는 문제로 이어졌습니다).
      if (NATIVE_COLOR_COLUMN_INDEX !== -1 && NATIVE_COLOR_COLUMN_INDEX < ACCUMULATED_COPY_COLUMN_COUNT) {
        targetSheet
          .getRange(2, NATIVE_COLOR_COLUMN_INDEX + 1, filteredRows.length, 1)
          .setNumberFormat("@");
      }

      const dataRange = targetSheet.getRange(2, 1, filteredRows.length, ACCUMULATED_COPY_COLUMN_COUNT);
      dataRange.setValues(filteredRows);
      dataRange.setFontFamily(DEFAULT_FONT_FAMILY);
      dataRange.setFontSize(DEFAULT_FONT_SIZE);
      dataRange.setFontColor(DEFAULT_FONT_COLOR);

      const accumulatedHeader = SUMMARY_ROW_HEADER.slice(0, ACCUMULATED_COPY_COLUMN_COUNT);

      DATE_FORMAT_COLUMN_LABELS.forEach(function(label) {
        const colIndex = accumulatedHeader.indexOf(label);

        if (colIndex !== -1) {
          targetSheet.getRange(2, colIndex + 1, filteredRows.length, 1).setNumberFormat(DATE_NUMBER_FORMAT);
        }
      });

      AMOUNT_FORMAT_COLUMN_LABELS.forEach(function(label) {
        const colIndex = accumulatedHeader.indexOf(label);

        if (colIndex !== -1) {
          targetSheet.getRange(2, colIndex + 1, filteredRows.length, 1).setNumberFormat(AMOUNT_NUMBER_FORMAT);
        }
      });
    }

    counts[factoryLabel] = filteredRows.length;
  });

  return { ok: true, counts: counts };
}


/**************************************************************
 * "월자료 초기화" 액션
 *
 * 이 스프레드시트(월현황 주간)의 종합(N), 마감(N), 로데이터, N주(N)
 * 시트를 전부 찾아서 헤더만 남기고 데이터 행을 지웁니다. 지운 뒤에는
 * 이름에 건수가 들어있는 시트(종합/마감/N주)는 renumberSheetNameIfNeeded_로
 * 이름도 0건으로 맞춥니다.
 **************************************************************/
function clearMonthlyDataAction_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const weekSheetPatterns = WEEK_NUMBERS.map(function(n) {
    return new RegExp("^" + n + "주\\(\\d+\\)$");
  });

  const clearedSheets = [];

  ss.getSheets().forEach(function(sheet) {
    const name = sheet.getName();
    const isWeekSheet = weekSheetPatterns.some(function(pattern) { return pattern.test(name); });

    if (!SUMMARY_SHEET_PATTERN.test(name) && !CLOSING_SHEET_PATTERN.test(name) &&
        name !== RAW_DATA_SHEET_NAME && !isWeekSheet) {
      return;
    }

    clearSheetKeepHeader_(sheet);
    renumberSheetNameIfNeeded_(sheet);
    clearedSheets.push(name);
  });

  return { ok: true, clearedSheets: clearedSheets };
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


/**************************************************************
 * "엑셀다운로드(전체)" 액션
 *
 * 이 월현황(주간) 스프레드시트 파일 자체를 통째로 xlsx로 내보냅니다
 * (모든 시트 포함). base64로 인코딩해서 돌려주면 화면에서 파일로
 * 저장합니다.
 **************************************************************/
function exportFullWorkbookAction_() {
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const base64 = exportSpreadsheetAsXlsxBase64_(spreadsheetId);
  const fileName = "월현황(주간)_" +
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


/**************************************************************
 * 유형 그룹 분류 (factory-as-cleanup.gs의 sortAndFormat_과 동일 규칙)
 *
 * row는 1공장추가건/2공장추가건의 원본 행(26열 기준)
 * R열(index 17) 유형, V열(index 21) 원인, W열(index 22) 포장
 **************************************************************/
function classifyGroup_(row) {
  const typeValue = normalizeText_(row[17]);
  const causeValue = normalizeText_(row[21]);
  const packageValue = normalizeText_(row[22]);

  if (typeValue === "미회수") return "미회수";
  if (typeValue === "감성") return "감성";
  if (typeValue === "취급") return "취급";
  if (causeValue === "VN" || packageValue === "VN") return "VN";

  return "일반";
}


function compareKoreanText_(valueA, valueB) {
  const textA = normalizeText_(valueA);
  const textB = normalizeText_(valueB);

  return textA.localeCompare(textB, "ko", { numeric: true, sensitivity: "base" });
}


/**************************************************************
 * 그룹 내부 정렬 (종합(N) 원본 26열 기준, 공장 열 없음)
 *
 * 1차 포장(W,22) 2차 세부유형(S,18) 3차 원인(V,21)
 * 4차 유형(R,17) 5차 접수번호(D,3)
 **************************************************************/
function sortRowsWithinGroup_(rows) {
  rows.sort(function(a, b) {
    let result = compareKoreanText_(a[22], b[22]);
    if (result !== 0) return result;

    result = compareKoreanText_(a[18], b[18]);
    if (result !== 0) return result;

    result = compareKoreanText_(a[21], b[21]);
    if (result !== 0) return result;

    result = compareKoreanText_(a[17], b[17]);
    if (result !== 0) return result;

    return compareKoreanText_(a[3], b[3]);
  });
}


/**************************************************************
 * 주차 번호 → 표시용 문자열 (예: 3 → "3주")
 **************************************************************/
function weekLabelFor_(weekNumber) {
  return weekNumber + "주";
}


/**************************************************************
 * "N주" 형태의 라벨에서 숫자만 뽑아냅니다. 형식이 다르면 null.
 **************************************************************/
function parseWeekNumber_(label) {
  const match = /^(\d+)주$/.exec(label);
  return match ? Number(match[1]) : null;
}


/**************************************************************
 * rows(startRow부터 순서대로 들어갈 데이터 행들)에 대해,
 * packageIdx 열(포장) 값이 "VN"/"미회수"인 행만 파란색/빨간색으로,
 * 나머지는 전부 검정색으로 글자색을 한 번에 지정합니다.
 * (유형/원인으로 그룹만 VN·미회수로 분류되고 실제 포장 값은
 * 다른 행이 잘못 색칠되는 것을 막기 위해, 그룹이 아니라 포장 열
 * 값만 기준으로 색을 정합니다.)
 **************************************************************/
function applyPackageFontColors_(sheet, startRow, rows, columnCount, packageIdx) {
  if (!rows.length) {
    return;
  }

  const colors = rows.map(function(row) {
    const packageValue = packageIdx !== -1 ? normalizeText_(row[packageIdx]) : "";
    let color = DEFAULT_FONT_COLOR;

    if (packageValue === "VN") {
      color = VN_FONT_COLOR;
    } else if (packageValue === "미회수") {
      color = UNCOLLECTED_FONT_COLOR;
    }

    return new Array(columnCount).fill(color);
  });

  sheet.getRange(startRow, 1, rows.length, columnCount).setFontColors(colors);
}


/**************************************************************
 * 시트 1행(헤더) 기준으로 필터를 새로 겁니다. 이미 필터가 있으면
 * 지우고 다시 걸어서, 늘어난/줄어든 행 범위에 맞춥니다.
 * totalRows가 0이면(데이터가 전혀 없으면) 아무 것도 하지 않습니다.
 **************************************************************/
function applySheetFilter_(sheet, totalRows, columnCount) {
  const existingFilter = sheet.getFilter();

  if (existingFilter) {
    existingFilter.remove();
  }

  if (totalRows > 0 && columnCount > 0) {
    sheet.getRange(1, 1, totalRows, columnCount).createFilter();
  }
}


/**************************************************************
 * sourceSheet의 열 너비를 targetSheet에 그대로 복사합니다
 * (엑셀 다운로드 시 종합(N)과 열 너비를 맞추기 위함).
 **************************************************************/
function copyColumnWidths_(targetSheet, sourceSheet, columnCount) {
  const widthColumnCount = Math.min(columnCount, sourceSheet.getLastColumn());

  for (let col = 1; col <= widthColumnCount; col++) {
    targetSheet.setColumnWidth(col, sourceSheet.getColumnWidth(col));
  }
}


/**************************************************************
 * 주간 탭 생성/갱신 ("N주(숫자)" 패턴으로 찾아서 내용을 다시 씀)
 *
 * incomingRows는 1공장추가건/2공장추가건에서 이번에 새로 읽어온
 * 전체 행(등록일은 이미 이번 주차로 맞춰진 상태)입니다.
 * 1. 기존 탭에 이미 있던 행 중 포장(W열)이 "미회수"인 행은 버립니다
 *    (해소 여부가 계속 바뀔 수 있어서 매번 최신 상태로 다시 채움).
 * 2. 남은 기존 행(미회수 제외) + incomingRows 중 D열(접수번호)+
 *    G열(고객명) 기준으로 기존에 없던 행만 합쳐서
 * 3. factory-as-cleanup.gs와 같은 규칙으로 다시 정렬한 뒤
 * 4. 탭 전체를 다시 씁니다(기존 저장분을 유지한 채 합쳐서 쓰는
 *    것이라 데이터가 사라지지 않습니다).
 * 글꼴(맑은 고딕), 날짜 표시 형식, VN/미회수 글자색, 헤더 필터,
 * 종합(N) 기준 열 너비를 매번 다시 지정합니다. 탭이 없으면 새로
 * 만듭니다. 반환값은 { sheetName, addedCount, totalCount }입니다.
 **************************************************************/
function writeWeeklySheet_(weekNumber, header, incomingRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = findWeekSheet_(weekNumber);
  const isNewSheet = !sheet;
  const columnCount = header.length;
  const packageIdx = header.indexOf("포장");
  const colorColIndex = header.indexOf("색상");

  let existingKept = [];

  if (sheet) {
    const lastRow = sheet.getLastRow();

    if (lastRow >= 2) {
      const existingRows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

      existingKept = existingRows.filter(function(row) {
        return packageIdx === -1 || normalizeText_(row[packageIdx]) !== "미회수";
      });
    }
  } else {
    sheet = ss.insertSheet(weekNumber + "주(0)", ss.getNumSheets());
  }

  const incomingByKey = new Map();
  incomingRows.forEach(function(row) { incomingByKey.set(rowMatchKey_(row), row); });

  const existingKeptKeys = new Set(existingKept.map(rowMatchKey_));

  const newOnly = incomingRows.filter(function(row) {
    return !existingKeptKeys.has(rowMatchKey_(row));
  });

  // 이미 저장돼 있던 행도 이번에 같은 키(접수번호+고객명)로 다시 들어온
  // 값이 있으면 최신 값으로 바꿔서, 재실행 시 그 사이 수정된 내용(금액/
  // 조치결과 등)이 반영되게 합니다(예전에는 기존 행을 그대로 두고
  // 새 값을 버려서, 저장 후 원본을 고치고 다시 저장해도 반영이 안 됐음).
  const existingUpdated = existingKept.map(function(row) {
    return incomingByKey.get(rowMatchKey_(row)) || row;
  });

  const combinedRows = groupAndSortRows_(existingUpdated.concat(newOnly));

  sheet.clear();

  const data = [header].concat(combinedRows).map(function(row) {
    return normalizeRowLength_(row, columnCount);
  });

  // 색상 열은 값을 쓰기 전에 먼저 텍스트("@") 서식으로 고정합니다.
  // sheet.clear()로 서식이 기본값(General)으로 돌아간 상태에서 "061"처럼
  // 숫자처럼 보이는 문자열을 그대로 쓰면, 쓰는 순간 시트가 숫자 61로
  // 재해석해버려서 앞자리 0이 사라집니다(factory-as-cleanup.gs에서
  // 겪은 것과 같은 문제).
  if (colorColIndex !== -1 && combinedRows.length) {
    sheet.getRange(2, colorColIndex + 1, combinedRows.length, 1).setNumberFormat("@");
  }

  const fullRange = sheet.getRange(1, 1, data.length, columnCount);
  fullRange.setValues(data);
  fullRange.setFontFamily(DEFAULT_FONT_FAMILY);
  sheet.setFrozenRows(1);

  DATE_FORMAT_COLUMN_LABELS.forEach(function(label) {
    const colIndex = header.indexOf(label);

    if (colIndex !== -1 && combinedRows.length) {
      sheet.getRange(2, colIndex + 1, combinedRows.length, 1).setNumberFormat(DATE_NUMBER_FORMAT);
    }
  });

  AMOUNT_FORMAT_COLUMN_LABELS.forEach(function(label) {
    const colIndex = header.indexOf(label);

    if (colIndex !== -1 && combinedRows.length) {
      sheet.getRange(2, colIndex + 1, combinedRows.length, 1).setNumberFormat(AMOUNT_NUMBER_FORMAT);
    }
  });

  applyPackageFontColors_(sheet, 2, combinedRows, columnCount, packageIdx);

  applySheetFilter_(sheet, data.length, columnCount);

  // 열 너비 복사는 API 호출이 열 수만큼(최대 26번) 드는데, 이미 한 번
  // 맞춰둔 시트는 다시 실행해도 너비가 바뀔 이유가 없어서 새로 만든
  // 시트일 때만 복사합니다(매번 반복 호출하던 걸 없애 재실행 속도를
  // 크게 줄임).
  if (isNewSheet) {
    const referenceSheet = findSummarySheet_();

    if (referenceSheet) {
      copyColumnWidths_(sheet, referenceSheet, columnCount);
    }
  }

  // 탭 이름의 건수는 VN/미회수 행을 빼고 셉니다(그 행들은 시트에는 그대로
  // 남아있고, 이름에만 실제로 "정상 처리된" 건수가 보이도록).
  const namedRowCount = combinedRows.filter(function(row) {
    return packageIdx === -1 || !isVnOrUncollectedPackage_(row[packageIdx]);
  }).length;

  const newName = weekNumber + "주(" + namedRowCount + ")";

  if (sheet.getName() !== newName) {
    sheet.setName(newName);
  }

  // newOnly에는 "이미 저장돼 있었지만 미회수라서 existingKept에서 매번
  // 제외되는" 행도 매번 새로 섞여 들어옵니다(미회수는 해소 여부가 계속
  // 바뀔 수 있어서 위에서 항상 제외 후 다시 채우기 때문). 그런 행까지
  // "추가건"으로 세면 재실행할 때마다 미회수 건수만큼 addedCount가
  // 부풀어 보이므로, 화면에 보여주는 건수에서는 뺍니다(시트에는 그대로
  // 들어갑니다 — 이건 집계용 숫자만 고치는 것).
  const genuinelyAddedCount = newOnly.filter(function(row) {
    return packageIdx === -1 || normalizeText_(row[packageIdx]) !== "미회수";
  }).length;

  return { sheetName: sheet.getName(), addedCount: genuinelyAddedCount, totalCount: combinedRows.length };
}


/**************************************************************
 * 포장 값이 VN 또는 미회수인지 확인합니다.
 **************************************************************/
function isVnOrUncollectedPackage_(packageValue) {
  const text = normalizeText_(packageValue);
  return text === "VN" || text === "미회수";
}


/**************************************************************
 * 종합(N) 시트 갱신
 *
 * 1. 포장(W열)이 "미회수"인 행은 주차와 상관없이 전부 삭제합니다
 *    (해소 여부가 계속 바뀔 수 있어서, 매번 최신 상태로 다시 채움).
 * 2. 제조자(1공장=충주1, 2공장=충주2)별로:
 *    - 이 제조자 + 이번 주차(등록일 === weekLabel)로 이미 저장돼
 *      있던 행(1번에서 미회수는 이미 지워졌으므로 미회수 아님)을
 *      모아두고, 그 블록을 시트에서 지웁니다.
 *    - D열(접수번호)+G열(고객명) 기준으로 그 기존 행과 겹치지 않는
 *      새 행만 골라, 기존 행과 합쳐서 factory-as-cleanup.gs와 같은
 *      규칙으로 다시 정렬합니다(같은 주차끼리 재정렬).
 *    - X열(제조자) 블록 마지막 자리(충주2가 아예 없었으면 4행 띄운
 *      뒤)에 정렬된 결과를 다시 삽입합니다.
 * 3. 시트 이름을 "종합(전체 행 수)"로 갱신합니다.
 *
 * factoryRawRows = { "1공장": [[...26열],...], "2공장": [...] }
 * (헤더 제외, Q열은 이미 이번 주 라벨로 맞춰진 상태)
 *
 * 반환값은 { addedCount }입니다.
 **************************************************************/
function updateSummarySheetForWeek_(weekLabel, factoryRawRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = findSummarySheet_();

  if (!sheet) {
    sheet = ss.insertSheet(SUMMARY_SHEET_PLACEHOLDER_NAME, ss.getNumSheets());
    sheet.getRange(1, 1, 1, SUMMARY_ROW_HEADER.length)
      .setValues([SUMMARY_ROW_HEADER])
      .setFontFamily(DEFAULT_FONT_FAMILY);
    sheet.setFrozenRows(1);
  }

  const columnCount = SUMMARY_ROW_HEADER.length;

  // 미회수는 주차와 상관없이 항상 지우고 이번 저장의 최신 상태로 다시 채움
  removeMatchingRows_(sheet, function(row) {
    return normalizeText_(row[22]) === "미회수";
  });

  let addedCount = 0;
  let block1End = null;

  ["1공장", "2공장"].forEach(function(factoryLabel) {
    const manufacturer = MANUFACTURER_BY_FACTORY_LABEL[factoryLabel];
    const incoming = factoryRawRows[factoryLabel] || [];

    const incomingByKey = new Map();
    incoming.forEach(function(row) { incomingByKey.set(rowMatchKey_(row), row); });

    const existingSameWeek = collectManufacturerWeekRows_(sheet, manufacturer, weekLabel);
    const existingKeys = new Set(existingSameWeek.map(rowMatchKey_));

    const newOnly = incoming.filter(function(row) {
      return !existingKeys.has(rowMatchKey_(row));
    });

    // 미회수 행은 위에서 항상 먼저 지우고 다시 채우기 때문에(944-947행),
    // 매번 "새로 추가된 것"으로 잡혀 addedCount가 미회수 건수만큼 부풀어
    // 보이는 걸 막기 위해 집계에서는 뺍니다(시트에는 그대로 들어갑니다).
    addedCount += newOnly.filter(function(row) {
      return normalizeText_(row[NATIVE_PACKAGE_COLUMN_INDEX]) !== "미회수";
    }).length;

    if (existingSameWeek.length) {
      removeMatchingRows_(sheet, function(row) {
        return normalizeText_(row[16]) === weekLabel && normalizeText_(row[23]) === manufacturer;
      });
    }

    // 이미 저장돼 있던 행도 이번에 같은 키로 다시 들어온 값이 있으면
    // 최신 값으로 바꿔서, 재실행 시 수정사항이 반영되게 합니다
    // (writeWeeklySheet_와 같은 이유의 같은 수정).
    const existingUpdated = existingSameWeek.map(function(row) {
      return incomingByKey.get(rowMatchKey_(row)) || row;
    });

    const combined = groupAndSortRows_(existingUpdated.concat(newOnly));

    if (combined.length) {
      const anchor = findLastRowIndexByManufacturer_(sheet, manufacturer);
      let insertAfter;

      if (anchor !== -1) {
        insertAfter = anchor;
      } else if (factoryLabel === "2공장") {
        const gapBase = block1End !== null ? block1End : sheet.getLastRow();
        sheet.insertRowsAfter(gapBase, SUMMARY_GAP_ROWS);
        insertAfter = gapBase + SUMMARY_GAP_ROWS;
      } else {
        insertAfter = sheet.getLastRow();
      }

      const blockEnd = insertRowsAt_(sheet, insertAfter, combined, columnCount);

      if (factoryLabel === "1공장") {
        block1End = blockEnd;
      }
    }
  });

  applySheetFilter_(sheet, sheet.getLastRow(), columnCount);

  // 시트 이름의 건수는 실제 데이터 행만 셉니다(충주1/충주2 블록 사이의
  // 빈 여백 행(SUMMARY_GAP_ROWS)은 건수에 포함하지 않음).
  const finalRowCount = countNonBlankDataRows_(sheet);
  const newSheetName = "종합(" + finalRowCount + ")";

  if (sheet.getName() !== newSheetName) {
    sheet.setName(newSheetName);
  }

  return { addedCount: addedCount };
}


/**************************************************************
 * 종합(N) 시트에서 제조자 === manufacturer 이고 등록일 === weekLabel인
 * 행들을 그대로 가져옵니다(재정렬을 위해 실제 값이 필요해서 키만이
 * 아니라 행 전체를 반환).
 **************************************************************/
function collectManufacturerWeekRows_(sheet, manufacturer, weekLabel) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  return values.filter(function(row) {
    return normalizeText_(row[23]) === manufacturer && normalizeText_(row[16]) === weekLabel;
  });
}


/**************************************************************
 * 시트의 데이터 행(헤더 제외) 중 완전히 빈 행을 뺀 실제 건수를 셉니다.
 **************************************************************/
function countNonBlankDataRows_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 2 || lastColumn < 1) {
    return 0;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();

  return values.filter(function(row) {
    return !row.every(function(cell) { return normalizeText_(cell) === ""; });
  }).length;
}


/**************************************************************
 * "마감" 시트 갱신 (종합(N) 바로 오른쪽에 있는, 예전 "월마감다운로드"
 * 버튼이 만들던 것과 같은 자료)
 *
 * 종합(N) 내용을 옮기되:
 * 1. 모든 열이 빈 값인 행(블록 사이 여백 행)은 제외
 * 2. W열(포장)이 "VN" 또는 "미회수"인 행은 제외
 * 남은 행에 sortRowsWithinGroup_과 같은 기준(1차 포장, 2차 세부유형,
 * 3차 원인, 4차 유형, 5차 접수번호)으로 한 번 더 정렬을 적용합니다
 * (마감 눌러서 봤을 때 순서가 뒤섞여 보인다는 요청으로 추가함 — 등록일
 * 순서를 유지하려던 게 아니라, 종합(N)에 주차별로 쌓이면서 이 기준으로
 * 완전히 정렬되지 않은 상태였던 것). 정렬 후 헤더 필터와 종합(N) 기준
 * 열 너비를 적용합니다.
 **************************************************************/
function updateClosingSheet_(preloadedSummaryValues) {
  const summarySheet = findSummarySheet_();

  if (!summarySheet) {
    return;
  }

  const lastRow = summarySheet.getLastRow();
  const lastColumn = summarySheet.getLastColumn();

  if (lastRow < 2 || lastColumn < 1) {
    return;
  }

  // archiveWeekAction_가 updateRawDataSheet_와 같이 쓸 값을 이미 한 번
  // 읽어서 넘겨주면(preloadedSummaryValues) 종합(N)을 다시 읽지 않고
  // 재사용합니다(같은 실행 안에서 종합(N)이 바뀌지 않았으므로 안전).
  const values = preloadedSummaryValues || summarySheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const header = values[0];
  const dataRows = values.slice(1);

  const packageIdx = header.indexOf("포장");

  const filteredRows = dataRows.filter(function(row) {
    const isBlankRow = row.every(function(cell) { return normalizeText_(cell) === ""; });

    if (isBlankRow) {
      return false;
    }

    const packageValue = packageIdx !== -1 ? normalizeText_(row[packageIdx]) : "";
    return packageValue !== "VN" && packageValue !== "미회수";
  });

  // 종합(N)과 같은 열 구성이므로 sortRowsWithinGroup_의 고정 열 인덱스를
  // 그대로 재사용해서 한 번 더 정렬합니다.
  sortRowsWithinGroup_(filteredRows);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = findClosingSheet_();
  const isNewClosingSheet = !sheet;

  if (!sheet) {
    sheet = ss.insertSheet(CLOSING_SHEET_NAME, summarySheet.getIndex());
  } else {
    sheet.clear();
  }

  const columnCount = header.length;
  const outputData = [header].concat(filteredRows);

  // 색상 열은 값을 쓰기 전에 텍스트("@") 서식부터 지정합니다(위
  // writeWeeklySheet_와 같은 이유 — sheet.clear() 직후 서식이
  // General인 채로 "061"을 쓰면 숫자로 바뀌어 앞자리 0이 사라짐).
  const closingColorColIndex = header.indexOf("색상");

  if (closingColorColIndex !== -1 && filteredRows.length) {
    sheet.getRange(2, closingColorColIndex + 1, filteredRows.length, 1).setNumberFormat("@");
  }

  const fullRange = sheet.getRange(1, 1, outputData.length, columnCount);
  fullRange.setValues(outputData);
  fullRange.setFontFamily(DEFAULT_FONT_FAMILY);
  sheet.setFrozenRows(1);

  DATE_FORMAT_COLUMN_LABELS.forEach(function(label) {
    const colIndex = header.indexOf(label);

    if (colIndex !== -1 && filteredRows.length) {
      sheet.getRange(2, colIndex + 1, filteredRows.length, 1).setNumberFormat(DATE_NUMBER_FORMAT);
    }
  });

  AMOUNT_FORMAT_COLUMN_LABELS.forEach(function(label) {
    const colIndex = header.indexOf(label);

    if (colIndex !== -1 && filteredRows.length) {
      sheet.getRange(2, colIndex + 1, filteredRows.length, 1).setNumberFormat(AMOUNT_NUMBER_FORMAT);
    }
  });

  applySheetFilter_(sheet, outputData.length, columnCount);

  // writeWeeklySheet_와 같은 이유로 새로 만든 시트일 때만 너비를 복사합니다.
  if (isNewClosingSheet) {
    copyColumnWidths_(sheet, summarySheet, columnCount);
  }

  // 종합(N), N주(N)와 같은 방식으로 이름에 건수를 붙입니다(VN/미회수는
  // 이미 위에서 filteredRows를 만들 때 제외했으므로 그대로 씀).
  const newClosingName = "마감(" + filteredRows.length + ")";

  if (sheet.getName() !== newClosingName) {
    sheet.setName(newClosingName);
  }
}


/**************************************************************
 * "로데이터" 시트 갱신 (마감 시트 바로 오른쪽에 있는, 재조합된 원본 데이터 시트)
 *
 * 종합(N)의 모든 데이터 행을 가져와서:
 * 1. 완전히 빈 행은 제외
 * 2. Q열(등록일) 값("1주","2주",...) 기준으로 그룹화하고,
 *    그룹 순서는 주차 숫자 오름차순으로 정렬(1주→2주→3주...).
 *    제조자(X열, 충주1/충주2) 구분 없이 같은 주차면 한 그룹으로 합칩니다.
 * 3. 그룹 내부는 W열(포장) 값만 기준으로 정렬
 * 4. RAW_DATA_SOURCE_COLUMN_LABELS 순서로 열을 다시 뽑고, 맨 뒤에
 *    "주차" 열을 추가(값은 등록일 값 그대로)
 * 5. 그룹 사이에 빈 행을 하나씩 넣어서 출력
 *
 * 종합(N)이 없으면 아무 것도 하지 않습니다(archiveWeekAction_에서
 * updateSummarySheetForWeek_ 다음에 호출되므로 이 시점엔 항상 있음).
 **************************************************************/
function updateRawDataSheet_(preloadedSummaryValues) {
  const summarySheet = findSummarySheet_();

  if (!summarySheet) {
    return;
  }

  const lastRow = summarySheet.getLastRow();
  const lastColumn = summarySheet.getLastColumn();

  if (lastRow < 2 || lastColumn < 1) {
    return;
  }

  const values = preloadedSummaryValues || summarySheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const header = values[0];
  const dataRows = values.slice(1);

  const registerIdx = header.indexOf("등록일");
  const packageIdx = header.indexOf("포장");
  const sourceColumnIndexes = RAW_DATA_SOURCE_COLUMN_LABELS.map(function(label) {
    return header.indexOf(label);
  });

  const nonBlankRows = dataRows.filter(function(row) {
    return !row.every(function(cell) { return normalizeText_(cell) === ""; });
  });

  const groups = {};
  const groupLabels = [];

  nonBlankRows.forEach(function(row) {
    const label = normalizeText_(registerIdx !== -1 ? row[registerIdx] : "");

    if (!groups[label]) {
      groups[label] = [];
      groupLabels.push(label);
    }

    groups[label].push(row);
  });

  groupLabels.sort(function(a, b) {
    const weekA = parseWeekNumber_(a);
    const weekB = parseWeekNumber_(b);

    if (weekA !== null && weekB !== null) return weekA - weekB;
    if (weekA !== null) return -1;
    if (weekB !== null) return 1;
    return 0;
  });

  groupLabels.forEach(function(label) {
    groups[label].sort(function(a, b) {
      const packageA = packageIdx !== -1 ? normalizeText_(a[packageIdx]) : "";
      const packageB = packageIdx !== -1 ? normalizeText_(b[packageIdx]) : "";
      const isUncollectedA = packageA === "미회수" ? 1 : 0;
      const isUncollectedB = packageB === "미회수" ? 1 : 0;

      // 미회수는 포장 순서와 상관없이 같은 등록일 그룹의 맨 아래로
      if (isUncollectedA !== isUncollectedB) {
        return isUncollectedA - isUncollectedB;
      }

      return compareKoreanText_(packageA, packageB);
    });
  });

  const columnCount = RAW_DATA_HEADER.length;
  const outputRows = [];

  groupLabels.forEach(function(label, groupIndex) {
    if (groupIndex > 0) {
      outputRows.push(new Array(columnCount).fill(""));
    }

    groups[label].forEach(function(row) {
      const mappedRow = sourceColumnIndexes.map(function(idx) {
        return idx !== -1 ? row[idx] : "";
      });
      mappedRow.push(label); // 주차 = 등록일 값 그대로
      outputRows.push(mappedRow);
    });
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(RAW_DATA_SHEET_NAME);

  if (!sheet) {
    const closingSheet = findClosingSheet_();
    const anchorSheet = closingSheet || summarySheet;
    sheet = ss.insertSheet(RAW_DATA_SHEET_NAME, anchorSheet.getIndex());
  } else {
    sheet.clear();
  }

  const outputData = [RAW_DATA_HEADER].concat(outputRows);

  // 색상 열은 값을 쓰기 전에 텍스트("@") 서식부터 지정합니다(같은
  // 이유로 writeWeeklySheet_/updateClosingSheet_와 동일한 수정).
  const rawDataColorColIndex = RAW_DATA_HEADER.indexOf("색상");

  if (rawDataColorColIndex !== -1 && outputRows.length) {
    sheet.getRange(2, rawDataColorColIndex + 1, outputRows.length, 1).setNumberFormat("@");
  }

  const fullRange = sheet.getRange(1, 1, outputData.length, columnCount);
  fullRange.setValues(outputData);
  fullRange.setFontFamily(DEFAULT_FONT_FAMILY);
  sheet.setFrozenRows(1);

  const dateColIndex = RAW_DATA_HEADER.indexOf("최종조치일");

  if (dateColIndex !== -1 && outputRows.length) {
    sheet.getRange(2, dateColIndex + 1, outputRows.length, 1).setNumberFormat(DATE_NUMBER_FORMAT);
  }

  const outputPackageIdx = RAW_DATA_HEADER.indexOf("포장");
  applyPackageFontColors_(sheet, 2, outputRows, columnCount, outputPackageIdx);
}

/**************************************************************
 * X열(제조자, 1-indexed 24번째 열)이 manufacturerValue인 행 중
 * 가장 아래(마지막) 행 번호(1-based)를 찾습니다. 없으면 -1.
 **************************************************************/
function findLastRowIndexByManufacturer_(sheet, manufacturerValue) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return -1;
  }

  const values = sheet.getRange(2, 24, lastRow - 1, 1).getValues();
  let lastIndex = -1;

  for (let i = 0; i < values.length; i++) {
    if (normalizeText_(values[i][0]) === manufacturerValue) {
      lastIndex = i + 2; // 1-based 시트 행 번호로 변환
    }
  }

  return lastIndex;
}

/**************************************************************
 * afterRow(1-based) 바로 다음에 새 행을 삽입하고 rows를 채웁니다.
 * rows가 비어있으면 아무 것도 하지 않고 afterRow를 그대로 반환합니다.
 * 삽입한 행 중 W열(포장) 값이 VN/미회수인 행만 글자색을 입히고,
 * 나머지는 검정색으로 맞춥니다.
 * 반환값은 삽입된 블록의 마지막 행 번호입니다.
 **************************************************************/
function insertRowsAt_(sheet, afterRow, rows, columnCount) {
  if (!rows.length) {
    return afterRow;
  }

  sheet.insertRowsAfter(afterRow, rows.length);

  const data = rows.map(function(row) {
    return normalizeRowLength_(row, columnCount);
  });

  if (NATIVE_COLOR_COLUMN_INDEX !== -1 && NATIVE_COLOR_COLUMN_INDEX < columnCount) {
    sheet.getRange(afterRow + 1, NATIVE_COLOR_COLUMN_INDEX + 1, data.length, 1).setNumberFormat("@");
  }

  sheet.getRange(afterRow + 1, 1, data.length, columnCount).setValues(data);
  applyPackageFontColors_(sheet, afterRow + 1, rows, columnCount, NATIVE_PACKAGE_COLUMN_INDEX);

  return afterRow + rows.length;
}

/**************************************************************
 * predicate(row)가 true인 데이터 행(헤더 제외)을 전부 삭제합니다.
 * 아래에서 위로 지워서 인덱스가 밀리는 문제를 피합니다.
 **************************************************************/
function removeMatchingRows_(sheet, predicate) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  // 지울 행을 뒤에서부터 훑으면서 "연속된 구간"으로 묶어 deleteRows()로
  // 한 번에 지웁니다(흩어진 행마다 deleteRow()를 따로 부르면 지울 행
  // 수만큼 API 호출이 들었음 — 미회수처럼 여러 건이 한꺼번에 지워지는
  // 경우 이 구간묶기로 호출 수를 크게 줄일 수 있습니다. 결과는 기존
  // 한 행씩 지우던 것과 완전히 동일합니다).
  let i = values.length - 1;

  while (i >= 0) {
    if (!predicate(values[i])) {
      i--;
      continue;
    }

    const runEndIndex = i;

    while (i >= 0 && predicate(values[i])) {
      i--;
    }

    const runStartRow = i + 3; // (i+1)이 구간 첫 행의 0-based values 인덱스, 시트 행은 +2
    const runLength = runEndIndex - i;

    sheet.deleteRows(runStartRow, runLength);
  }
}


/**************************************************************
 * 패턴("종합(숫자)")으로 종합 시트 찾기
 **************************************************************/
function findSummarySheet_() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();

  for (let i = 0; i < sheets.length; i++) {
    if (SUMMARY_SHEET_PATTERN.test(sheets[i].getName())) {
      return sheets[i];
    }
  }

  return null;
}


/**************************************************************
 * 패턴("마감" 또는 "마감(숫자)")으로 마감 시트 찾기
 **************************************************************/
function findClosingSheet_() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();

  for (let i = 0; i < sheets.length; i++) {
    if (CLOSING_SHEET_PATTERN.test(sheets[i].getName())) {
      return sheets[i];
    }
  }

  return null;
}


/**************************************************************
 * 패턴("N주(숫자)")으로 해당 주차 시트 찾기
 *
 * 해당 주차 탭이 없으면 null을 반환합니다(에러 아님) —
 * 대시보드에서 "데이터 없음"으로 표시하기 위함입니다.
 **************************************************************/
function findWeekSheet_(weekNumber) {
  const pattern = new RegExp("^" + weekNumber + "주\\(\\d+\\)$");
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();

  for (let i = 0; i < sheets.length; i++) {
    if (pattern.test(sheets[i].getName())) {
      return sheets[i];
    }
  }

  return null;
}


/**************************************************************
 * 주간 탭 이름 목록 (종합 시트 제외)
 **************************************************************/
function listWeekSheetNames_() {
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .map(function(sheet) { return sheet.getName(); })
    .filter(function(name) { return !SUMMARY_SHEET_PATTERN.test(name); });
}


/**************************************************************
 * 시트 전체를 header + 행 배열(rowIndex 포함)로 반환 (이름으로 조회)
 **************************************************************/
function readSheetSafe_(sheetName) {
  return readSheetObject_(
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
  );
}


/**************************************************************
 * 시트 전체를 header + 행 배열(rowIndex 포함)로 반환 (Sheet 객체로 조회)
 *
 * 1행이 실제 헤더입니다(안내문 행은 제거됨). 종합(N), N주(N),
 * 이름으로 찾는 시트 전부 이 함수로 읽습니다.
 *
 * sheet가 null이면(해당 탭이 없으면) 빈 결과를 반환합니다 —
 * 대시보드에서 이 경우 "데이터 없음"으로 표시합니다.
 **************************************************************/
function readSheetObject_(sheet) {
  if (!sheet) {
    return { sheet: "", header: [], rows: [] };
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 1 || lastColumn < 1) {
    return { sheet: sheet.getName(), header: [], rows: [] };
  }

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const header = values[0];
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    rows.push({ rowIndex: i + 1, values: values[i] });
  }

  return { sheet: sheet.getName(), header: header, rows: rows };
}


function normalizeText_(value) {
  return String(
    value === null || value === undefined ? "" : value
  ).trim();
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
