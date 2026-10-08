/**************************************************************
 * 품질 고객클레임 대시보드 "1. 기존데이터" (25년/26년 마감자료) 연동용 Web App
 *
 * 배포 위치
 * ------------------------------------------------------------
 * 이 파일은 "N년 마감자료" 스프레드시트(25년용, 26년용 각각 별도 파일)의
 * Apps Script 프로젝트에 추가합니다. 같은 코드를 두 스프레드시트에
 * 각각 복사해 넣고, 스프레드시트별로 따로 배포합니다 → 웹 앱 URL이
 * 두 개(25년용, 26년용) 나옵니다.
 *
 * 하는 일
 * ------------------------------------------------------------
 * doGet: 이 스프레드시트에서 시트 이름이 "숫자+월"로 시작하는 탭만
 * 찾아서(예: "1월마감(320)", "2월마감(15)"도 인식), 각 탭의 전체 셀
 * 값(표시되는 그대로의 문자열)을
 * { ok: true, months: { "1월": [[...행...], ...], "2월": [...] } }
 * 형태의 JSON으로 돌려줍니다(뒤에 붙은 "마감(N)" 같은 글자는 떼고
 * "1월", "2월"처럼 정리해서 돌려줍니다). 대시보드가 매달 이 URL로
 * 다시 요청해서 새로 추가된 월 탭을 자동으로 읽어가므로, 매달 새
 * CSV 링크를 다시 붙여넣지 않아도 됩니다.
 *
 * 배포 방법
 * ------------------------------------------------------------
 * 1. 해당 연도 마감자료 스프레드시트 열기
 * 2. 확장 프로그램 > Apps Script
 * 3. 기본 Code.gs 내용을 전부 지우고 이 파일 내용 붙여넣기
 * 4. 저장(Ctrl+S)
 * 5. 우측 상단 "배포" > "새 배포"
 * 6. 유형 선택(톱니바퀴 아이콘) > "웹 앱"
 * 7. "실행 계정": 나, "액세스 권한이 있는 사용자": 전체 → "배포"
 * 8. 처음 배포 시 승인 화면이 뜨면 "고급" > "(프로젝트명)(안전하지 않음)로 이동" > 허용
 * 9. 생성된 웹 앱 URL(.../exec로 끝남)을 복사해서 대시보드
 *    "데이터 삽입" 모달의 "25년마감자료 웹앱 URL"(또는 26년) 입력칸에
 *    붙여넣고 "넣기" 클릭
 * 10. 시트 구조를 바꿔서 코드를 다시 수정한 경우: "배포" > "배포 관리"에서
 *     기존 배포를 "편집"(연필 아이콘) > 버전 "새 버전"으로 다시 배포해야
 *     URL을 바꾸지 않고 최신 코드가 반영됩니다.
 *
 * doPost action="updateClosingStatus": "주간업무 > 마감 > 마감현황 저장"
 * 버튼 전용. body에 header(월현황(주간) "마감(N)" 시트의 실제 헤더
 * 배열)와 rows(그 시트의 데이터 행 배열, 헤더 제외)를 담아 보내면:
 * 1) 이번달(Asia/Seoul 기준) "{N}월마감" 또는 "{N}월마감(숫자)" 탭을
 *    찾고, 없으면 직전달 탭(예: 10월이면 9월마감(N))을 통째로 복사해서
 *    "{N}월마감"이라는 이름으로 새로 만듭니다(1~2행 헤더/서식을 그대로
 *    물려받기 위해 — 직전달 탭도 없으면 에러)
 * 2) 그 탭의 2행(실제 헤더 — 1행은 상위 제목행)과 이름이 같은 열끼리
 *    맞춰서(열 순서가 달라도 안전하도록) 3행부터 기존 내용을 지우고
 *    새 데이터를 씁니다(색상 열은 "061"처럼 0으로 시작해도 숫자로
 *    안 바뀌도록 쓰기 전에 텍스트 서식부터 지정합니다)
 * 3) 이 프로젝트의 다른 "(N)" 탭들과 같은 방식으로, 실제로 쓴 행 수에
 *    맞춰 탭 이름의 "(N)"을 갱신합니다.
 *
 * 주의
 * ------------------------------------------------------------
 * - 토큰 검증이 없으므로 URL을 아는 사람은 누구나 이 시트 내용을 읽고
 *   "마감현황 저장" 기능으로 쓸 수도 있습니다.
 * - 탭 이름이 "숫자+월"로 시작하기만 하면 인식합니다("1월마감(320)",
 *   "1월", "1월 마감" 등 모두 인식). 다만 같은 스프레드시트에 "1월"로
 *   시작하는 탭이 두 개 이상 있으면 나중 탭이 앞 탭을 덮어씁니다(doGet
 *   읽기 기준). "updateClosingStatus"는 이름이 정확히
 *   "{이번달}월마감" 또는 "{이번달}월마감(숫자)"인 탭만 찾습니다.
 **************************************************************/

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const params = (e && e.parameter) || {};

    if (params.action === "spreadsheetUrl") {
      return jsonOutput_({ url: ss.getUrl() });
    }

    const months = {};
    ss.getSheets().forEach((sheet) => {
      const name = sheet.getName().trim();
      const match = name.match(/^(\d{1,2})\s*월/);
      if (!match) return;
      const monthLabel = match[1] + "월";
      months[monthLabel] = sheet.getDataRange().getDisplayValues();
    });
    return jsonOutput_({ ok: true, months: months });
  } catch (error) {
    return jsonOutput_({ ok: false, error: String(error) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const action = body.action || "";

    if (action === "updateClosingStatus") {
      return jsonOutput_(updateClosingStatusAction_(body.header || [], body.rows || []));
    }

    return jsonOutput_({ ok: false, error: "알 수 없는 action입니다: " + action });
  } catch (error) {
    return jsonOutput_({ ok: false, error: String(error) });
  }
}

function jsonOutput_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function normalizeText_(value) {
  return String(value === null || value === undefined ? "" : value).trim();
}

// 2행(실제 헤더 — 1행은 상위 제목행)과 3행부터 시작하는 데이터 구성은
// 이 스프레드시트의 "{N}월마감(N)" 탭들이 공통으로 쓰는 고정 레이아웃.
const CLOSING_ARCHIVE_HEADER_ROW_ = 2;
const CLOSING_ARCHIVE_DATA_START_ROW_ = 3;

/**************************************************************
 * Asia/Seoul 기준 이번달 숫자로 "{N}월마감" 또는 "{N}월마감(숫자)"와
 * 이름이 정확히 일치하는 탭을 찾습니다("월"과 "마감" 사이 공백은
 * 있어도/없어도 인식 — doGet의 "1월 마감"도 인식한다는 것과 같은 이유).
 **************************************************************/
function findCurrentMonthClosingArchiveSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const monthNum = Number(Utilities.formatDate(new Date(), "Asia/Seoul", "M"));
  const pattern = new RegExp("^" + monthNum + "\\s*월\\s*마감(\\s*\\(\\d+\\))?$");
  const sheets = ss.getSheets();

  for (let i = 0; i < sheets.length; i++) {
    if (pattern.test(sheets[i].getName().trim())) {
      return { sheet: sheets[i], monthNum: monthNum };
    }
  }

  return { sheet: null, monthNum: monthNum };
}

/**************************************************************
 * 이번달 "{N}월마감" 탭이 아직 없을 때, 직전달 탭(예: 10월이면
 * "9월마감(N)")을 통째로 복사해서 "{N}월마감"이라는 이름으로 만듭니다
 * (1~2행 헤더/서식을 그대로 물려받기 위해 — 데이터 행은 이후
 * updateClosingStatusAction_에서 지우고 새로 씁니다). 직전달 탭도 없으면
 * 복사할 대상이 없으므로 에러를 던집니다.
 **************************************************************/
function createCurrentMonthClosingArchiveSheet_(monthNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prevMonthNum = monthNum === 1 ? 12 : monthNum - 1;
  const prevPattern = new RegExp("^" + prevMonthNum + "\\s*월\\s*마감");
  const sheets = ss.getSheets();
  let prevSheet = null;

  for (let i = 0; i < sheets.length; i++) {
    if (prevPattern.test(sheets[i].getName().trim())) {
      prevSheet = sheets[i]; // 같은 접두사 탭이 여러 개면 나중 탭을 우선(doGet과 같은 기준)
    }
  }

  if (!prevSheet) {
    throw new Error(
      "'" + monthNum + "월마감' 탭이 없고, 복사할 '" + prevMonthNum + "월마감(N)' 탭도 찾을 수 없습니다. 먼저 수동으로 탭을 만들어주세요."
    );
  }

  const newSheet = prevSheet.copyTo(ss);
  newSheet.setName(monthNum + "월마감");
  ss.setActiveSheet(newSheet);
  ss.moveActiveSheet(prevSheet.getIndex() + 1);

  return newSheet;
}

/**************************************************************
 * "월현황(주간)" 웹앱의 "마감(N)" 시트 자료(header/rows, 헤더 제외한
 * 데이터 행만)를 이번달 "{N}월마감(N)" 탭의 3행부터 덮어씁니다.
 **************************************************************/
function updateClosingStatusAction_(sourceHeader, sourceRows) {
  if (!Array.isArray(sourceRows) || !sourceRows.length) {
    return { ok: false, error: "저장할 마감 데이터가 없습니다." };
  }

  const found = findCurrentMonthClosingArchiveSheet_();
  let sheet = found.sheet;

  if (!sheet) {
    try {
      sheet = createCurrentMonthClosingArchiveSheet_(found.monthNum);
    } catch (error) {
      return { ok: false, error: String(error.message || error) };
    }
  }

  const lastColumn = sheet.getLastColumn();

  if (lastColumn < 1) {
    return { ok: false, error: "'" + sheet.getName() + "' 탭에 헤더가 없습니다." };
  }

  const destHeader = sheet.getRange(CLOSING_ARCHIVE_HEADER_ROW_, 1, 1, lastColumn).getValues()[0]
    .map(normalizeText_);

  // 마감(N)과 이번달마감(N)의 열 순서가 다를 수 있으므로, 위치가 아니라
  // 열 이름으로 맞춰서 옮겨 씁니다.
  const sourceIndexByName = {};
  (sourceHeader || []).forEach(function(name, idx) {
    const key = normalizeText_(name);
    if (key && !Object.prototype.hasOwnProperty.call(sourceIndexByName, key)) {
      sourceIndexByName[key] = idx;
    }
  });

  const unmatchedDestColumns = [];
  const columnMap = destHeader.map(function(name) {
    const key = normalizeText_(name);
    if (key && Object.prototype.hasOwnProperty.call(sourceIndexByName, key)) {
      return sourceIndexByName[key];
    }
    if (key) unmatchedDestColumns.push(key);
    return -1;
  });

  const outputRows = sourceRows.map(function(row) {
    return columnMap.map(function(srcIdx) {
      if (srcIdx === -1) return "";
      const value = row[srcIdx];
      return value === undefined || value === null ? "" : value;
    });
  });

  const existingLastRow = sheet.getLastRow();
  const clearRowCount = Math.max(existingLastRow - CLOSING_ARCHIVE_DATA_START_ROW_ + 1, outputRows.length);

  if (clearRowCount > 0) {
    sheet.getRange(CLOSING_ARCHIVE_DATA_START_ROW_, 1, clearRowCount, lastColumn).clearContent();
  }

  // 색상 열은 값을 쓰기 전에 텍스트("@") 서식부터 지정합니다 — 이
  // 프로젝트의 다른 "(N)" 탭들(마감(N), 정리(N) 등)과 같은 이유로,
  // 쓴 뒤에 서식을 바꾸면 이미 "061" 같은 값이 61로 바뀐 뒤라 늦습니다.
  const colorColIndex = destHeader.indexOf("색상");
  if (colorColIndex !== -1 && outputRows.length) {
    sheet.getRange(CLOSING_ARCHIVE_DATA_START_ROW_, colorColIndex + 1, outputRows.length, 1).setNumberFormat("@");
  }

  if (outputRows.length) {
    sheet.getRange(CLOSING_ARCHIVE_DATA_START_ROW_, 1, outputRows.length, lastColumn).setValues(outputRows);
  }

  // "(N)"은 건수를 뜻하므로, 이 프로젝트의 다른 "(N)" 탭들과 같은
  // 방식으로 실제 저장한 행 수에 맞춰 탭 이름을 갱신합니다.
  const newName = found.monthNum + "월마감(" + outputRows.length + ")";
  if (sheet.getName() !== newName) {
    sheet.setName(newName);
  }

  return {
    ok: true,
    sheet: newName,
    writtenRows: outputRows.length,
    unmatchedDestColumns: unmatchedDestColumns
  };
}
