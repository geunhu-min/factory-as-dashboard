/**************************************************************
 * "포장실적상세" 스프레드시트 조회/편집용 Web App
 *
 * 배포 위치
 * ------------------------------------------------------------
 * 이 파일은 "포장실적상세" 스프레드시트의 확장 프로그램 > Apps Script
 * 프로젝트에 추가합니다(다른 파일과 마찬가지로 그 스프레드시트에
 * 바인딩된 스크립트라, 이 파일 안에서는 스프레드시트 ID를 몰라도
 * SpreadsheetApp.getActiveSpreadsheet()로 항상 그 파일 자신을
 * 가리킵니다).
 *
 * 하는 일
 * ------------------------------------------------------------
 * 아직 "포장실적상세" 시트의 실제 탭 구조/열 이름을 모르는 상태라,
 * 우선 일반적인 조회/편집/내보내기만 되는 뼈대만 만들어뒀습니다
 * (comparison-edit-webapp.gs와 같은 방식). "포장실적상세 데이터
 * 교체"/"포장일확인" 같은 실제 업무 로직은 시트 구조를 확인한 뒤
 * 별도로 채워 넣어야 합니다.
 *
 * - doGet action="sheets": 이 파일의 모든 탭 이름 + gid + 행/열 수 목록
 *   (어떤 탭을 볼지 고르기 전에, 구조 파악용으로 먼저 호출)
 * - doGet action="read" (기본값): 탭 하나를 header + rows로 반환합니다.
 *   sheet 파라미터(탭 이름) 또는 gid 파라미터로 지정, 둘 다 없으면
 *   첫 번째 탭을 씁니다. 1행을 헤더로 보고 그 아래부터 데이터 행으로
 *   취급합니다(병합 셀이 있으면 병합된 첫 칸에만 값이 들어오고
 *   나머지는 빈 값으로 옵니다).
 * - doGet action="spreadsheetUrl": 이 스프레드시트의 편집 URL을
 *   반환합니다("포장실적상세 데이터 교체" 버튼이 새 탭에서 열 때 씀 —
 *   화면 쪽에는 스프레드시트 ID를 저장해두지 않으므로 매번 물어봄).
 * - doPost action="save": 화면에서 수정한 행을 저장합니다. 행마다
 *   loaded(불러왔을 때 값)와 values(수정한 값)를 같이 보내면, 저장
 *   직전 시트의 실제 값과 loaded를 비교해서 그 사이에 다른 곳에서
 *   먼저 바뀐 행은 충돌로 보고 덮어쓰지 않습니다(그 행은 최신 값
 *   그대로 돌려줌). 응답에 최신 header/rows/conflicts/updatedCount를
 *   담아 돌려주므로, 화면은 그 값으로 다시 그리면 됩니다.
 * - doPost action="exportFull": 이 스프레드시트 파일 전체(모든 시트)를
 *   xlsx로 내보내 base64로 반환합니다(범용 백업용, 화면 버튼과는
 *   연결되어 있지 않음).
 * - doPost action="checkPackagingDate": "포장일확인" 버튼 액션(현재
 *   화면이 쓰는 방식). 화면은 접수데이터 웹앱에서 그 날짜에 해당하는
 *   부품코드+색상+생산로트+고객명+하자내역 목록(keys)만 가볍게 받아와서
 *   그대로 이 액션에 넘겨줍니다. 이 스프레드시트의 시트1을 여기서
 *   직접 읽어(화면으로 왕복하지 않음) J열 제품코드+K열 색상+H열 Lot번호
 *   기준으로만 대조합니다(접수일과 생산일은 다를 수 있으므로 생산일로는
 *   거르지 않음). N열(생산량)이 0인 행은 결과에서 제외하고, 제품코드
 *   기준으로 먼저 묶어서 그 안에서 생산일 기준 오름차순으로 정렬한 뒤
 *   "시트2"에 씁니다.
 * - doPost action="writeSheet2": rows(이미 골라낸 행 배열)를 그대로
 *   "시트2"에 덮어씁니다(있으면 지우고 다시 씀, 없으면 새로 만듦).
 *   checkPackagingDateAction_이 내부적으로 재사용하는 더 낮은 단계의
 *   액션입니다. 열 구성은 DAILY_SHEET2_HEADER 고정(생산자/생산라인/
 *   Lot번호/생산일/제품코드/색상/계획량/생산량/미포장량/고객명/
 *   하자내역). 열 너비는 DAILY_SHEET2_COLUMN_WIDTHS 고정값으로
 *   맞춥니다. 응답에 최신 header/rows를 돌려주므로 화면은 그 값으로
 *   표를 그리면 됩니다.
 * - doPost action="exportSheet2": "포장일엑셀다운로드" 버튼 액션.
 *   "시트2" 하나만 담은 xlsx를 base64로 반환합니다(열 너비가 이미
 *   writeSheet2Action_에서 고정값으로 맞춰져 있어 바로 가로로 출력할
 *   수 있습니다).
 * - 스프레드시트를 열면 메뉴 막대에 "포장실적상세 도구" 메뉴가 추가되고,
 *   그 안의 "마지막 행 A열로 이동"을 누르면 현재 보고 있는 탭의 맨
 *   마지막 데이터 행 A열로 화면이 이동합니다(onOpen 참고).
 *
 * 배포 방법
 * ------------------------------------------------------------
 * 1. "포장실적상세" 스프레드시트 > 확장 프로그램 > Apps Script에 이
 *    파일 추가
 * 2. 배포 > 새 배포 > 유형: 웹 앱, 실행 계정: 나, 액세스 권한: 필요 범위
 * 3. 배포 후 나오는 웹 앱 URL을 대시보드의 "일일업무" 페이지 >
 *    "일일현황관리 데이터" > "상세" > "포장실적상세 연결"에 입력
 * 4. UrlFetchApp을 처음 쓰는 경우 권한 재승인이 필요할 수 있습니다.
 *    함수 선택 드롭다운에서 testAuth를 선택해 한 번 실행하고 동의
 *    화면을 통과한 뒤 다시 배포하세요.
 * 5. "포장실적상세 도구" 메뉴는 onOpen(단순 트리거)이라 별도 권한
 *    승인이나 배포 없이, 스프레드시트 탭을 새로고침만 하면 바로
 *    메뉴 막대에 나타납니다.
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

// "포장일확인" 버튼이 만드는 정리 시트 이름과 고정 열 구성
// (포장실적상세 시트1에서 D/G/H/I/J/K/M/N/O열을 그대로 옮기고, 맨
// 뒤에 접수데이터 시트1 I열(고객명)/M열(하자내역)을 붙인 것).
const DAILY_SHEET2_NAME = "시트2";
const DAILY_SHEET2_HEADER = [
  "생산자", "생산라인", "Lot번호", "생산일", "제품코드",
  "색상", "계획량", "생산량", "미포장량", "고객명", "하자내역"
];

// DAILY_SHEET2_HEADER와 같은 순서의 열 너비(픽셀). autoResizeColumns는
// 한글처럼 폭이 넓은 글자의 너비를 실제보다 좁게 계산하는 경우가 있어서,
// 매번 확실하게 잘리지 않도록 고정값으로 직접 지정합니다.
const DAILY_SHEET2_COLUMN_WIDTHS = [90, 110, 90, 100, 120, 60, 70, 70, 80, 100, 260];

const DEFAULT_FONT_FAMILY = "Malgun Gothic";

/**************************************************************
 * 스프레드시트를 열 때 메뉴 막대(도움말 오른쪽)에 "포장실적상세 도구"
 * 메뉴를 추가합니다. 단순 트리거(onOpen)라 별도 권한 승인 없이 항상
 * 자동으로 실행됩니다.
 **************************************************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("포장실적상세 도구")
    .addItem("마지막 행 A열로 이동", "goToLastRowColumnA")
    .addToUi();
}


/**************************************************************
 * 현재 보고 있는 탭의 맨 마지막 데이터 행 A열로 화면을 이동합니다.
 * 데이터가 전혀 없는 빈 시트라면 A1으로 이동합니다.
 **************************************************************/
function goToLastRowColumnA() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = Math.max(sheet.getLastRow(), 1);

  sheet.getRange(lastRow, 1).activate();
}


/**************************************************************
 * 권한 재승인용 임시 테스트 함수
 *
 * exportFull이 쓰는 UrlFetchApp(외부 요청), exportSheet2("포장일
 * 엑셀다운로드" 버튼)이 쓰는 DriveApp의 파일 생성/삭제(휴지통 이동)
 * 권한을 승인받기 위한 함수입니다. DriveApp.getRootFolder() 같은
 * 읽기 동작은 "drive" 전체 권한을 요청하지 않아서, 실제로
 * exportSheetsSubsetAsXlsxBase64_가 하는 것과 똑같이 파일을 하나
 * 만들었다가 지워봅니다. 이름에 밑줄(_)이 없어야 Apps Script
 * 편집기의 "실행할 함수" 드롭다운에 보입니다. 드롭다운에서
 * testAuth를 선택해 실행하면 동의 화면이 뜹니다 — 승인한 뒤에는
 * 이 함수를 지우고 다시 배포해도 되고, 그냥 남겨둬도 동작에는
 * 영향이 없습니다.
 **************************************************************/
function testAuth() {
  UrlFetchApp.fetch("https://www.google.com");

  const tempFile = DriveApp.createFile("포장실적상세_권한테스트_임시파일.txt", "test");
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

    if (action === "writeSheet2") {
      return jsonOutput_(writeSheet2Action_(body.rows || []));
    }

    if (action === "checkPackagingDate") {
      const sheet = resolveSheet_(body.sheet, body.gid);
      return jsonOutput_(checkPackagingDateAction_(sheet, body.keys || []));
    }

    if (action === "exportSheet2") {
      return jsonOutput_(exportSheetsSubsetAsXlsxBase64_([DAILY_SHEET2_NAME], "포장일확인"));
    }

    return jsonOutput_({ error: "알 수 없는 action입니다: " + action });
  } catch (error) {
    return jsonOutput_({ error: error.message });
  }
}


/**************************************************************
 * "포장일확인" 실제 대조 로직. keys는 접수데이터 웹앱의 packagingKeys
 * 액션이 그날 접수분만 걸러서 보내준 { partCode, color, lot,
 * customerName, defect } 배열입니다. 접수일과 생산일은 서로 다를 수
 * 있으므로(생산이 접수일 전후로 이루어질 수 있음), 생산일로는 절대
 * 거르지 않고 오직 J열 제품코드+K열 색상+H열 Lot번호 조합(키)만으로
 * 대조합니다. N열(생산량)이 0인 행은 결과에서 제외하고, 제품코드
 * 기준으로 먼저 묶어서 그 안에서 생산일 기준 오름차순으로 정렬한 뒤
 * writeSheet2Action_로 "시트2"에 씁니다.
 *
 * 시트1이 실제 대조/출력에 쓰는 열은 O열(15번째)까지뿐이므로, 시트에
 * 그보다 더 많은 열이 있어도 딱 그만큼만 읽어 전송량을 줄입니다
 * (결과에는 영향 없는, 대조 조건과 무관한 안전한 최적화입니다).
 **************************************************************/
function checkPackagingDateAction_(sheet, keys) {
  const infoByKey = new Map();

  keys.forEach(function(k) {
    const key = [
      normalizeText_(k.partCode),
      normalizeText_(k.color),
      normalizeText_(k.lot)
    ].join("||");

    infoByKey.set(key, {
      customerName: normalizeText_(k.customerName),
      defect: normalizeText_(k.defect)
    });
  });

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const matchedRows = [];

  if (lastRow >= 2 && lastColumn >= 1 && infoByKey.size) {
    const columnsToRead = Math.min(lastColumn, 15); // O열까지만 있으면 됨
    const values = sheet.getRange(2, 1, lastRow - 1, columnsToRead).getValues();

    values.forEach(function(row) {
      const key = [
        normalizeText_(row[9]),  // J열 제품코드
        normalizeText_(row[10]), // K열 색상
        normalizeText_(row[7])   // H열 Lot번호
      ].join("||");

      if (!infoByKey.has(key)) return;

      const qty = Number(String(row[13]).replace(/,/g, "")) || 0; // N열 생산량
      if (qty === 0) return; // 생산량 0건은 정리 결과에서 제외

      const info = infoByKey.get(key);

      matchedRows.push([
        row[3],  // D열 생산자
        row[6],  // G열 생산라인
        row[7],  // H열 Lot번호
        row[8],  // I열 생산일
        row[9],  // J열 제품코드
        row[10], // K열 색상
        row[12], // M열 계획량
        row[13], // N열 생산량
        row[14], // O열 미포장량
        info.customerName, // 접수데이터 고객명
        info.defect         // 접수데이터 하자내역
      ]);
    });
  }

  // 제품코드 기준으로 먼저 묶고, 같은 제품코드 안에서는 생산일 기준
  // 오름차순으로 정렬합니다.
  matchedRows.sort(function(a, b) {
    const codeA = normalizeText_(a[4]);
    const codeB = normalizeText_(b[4]);
    if (codeA !== codeB) return codeA < codeB ? -1 : 1;

    const dateA = formatDateOnly_(a[3]);
    const dateB = formatDateOnly_(b[3]);
    return dateA < dateB ? -1 : dateA > dateB ? 1 : 0;
  });

  return writeSheet2Action_(matchedRows);
}


/**************************************************************
 * 날짜 셀 값을 "yyyy-MM-dd" 문자열로 맞춥니다(Date 타입/텍스트 모두
 * 처리) — 생산일 기준 정렬에 씁니다.
 **************************************************************/
function formatDateOnly_(value) {
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, "Asia/Seoul", "yyyy-MM-dd");
  }

  const text = String(value === null || value === undefined ? "" : value).trim();
  const match = /^(\d{4}-\d{2}-\d{2})/.exec(text);
  return match ? match[1] : text;
}


/**************************************************************
 * "포장일확인" 결과를 "시트2"에 덮어씁니다(있으면 지우고 다시 씀).
 * rows는 DAILY_SHEET2_HEADER 순서(11열)에 맞춰 이미 골라둔 값들입니다.
 * 매번 새로 정리해서 보여줘야 하므로 호출할 때마다 통째로 다시
 * 씁니다. 헤더/내용 전체를 가운데 정렬하고, 열 너비는
 * DAILY_SHEET2_COLUMN_WIDTHS 고정값으로 맞춰서("포장일엑셀다운로드"로
 * 내려받았을 때 바로 가로로 출력할 수 있게) 잘리는 열이 없도록
 * 합니다(autoResizeColumns는 한글 너비를 실제보다 좁게 계산하는
 * 경우가 있어 쓰지 않습니다).
 **************************************************************/
function writeSheet2Action_(rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, DAILY_SHEET2_NAME);
  const columnCount = DAILY_SHEET2_HEADER.length;

  sheet.clear();

  const data = [DAILY_SHEET2_HEADER].concat(rows.map(function(row) {
    return normalizeRowLength_(row, columnCount);
  }));

  // Lot번호/제품코드/색상 열은 값을 쓰기 전에 텍스트("@") 서식부터
  // 지정합니다. sheet.clear() 직후 서식이 General인 채로 "061"처럼
  // 숫자처럼 보이는 값을 쓰면 그 즉시 숫자로 바뀌어 앞자리 0이 사라집니다.
  ["Lot번호", "제품코드", "색상"].forEach(function(label) {
    const colIndex = DAILY_SHEET2_HEADER.indexOf(label);

    if (colIndex !== -1 && rows.length) {
      sheet.getRange(2, colIndex + 1, rows.length, 1).setNumberFormat("@");
    }
  });

  const fullRange = sheet.getRange(1, 1, data.length, columnCount);
  fullRange.setValues(data);
  fullRange.setFontFamily(DEFAULT_FONT_FAMILY);
  fullRange.setHorizontalAlignment("center");
  sheet.setFrozenRows(1);

  DAILY_SHEET2_COLUMN_WIDTHS.forEach(function(width, idx) {
    sheet.setColumnWidth(idx + 1, width);
  });

  return readSheetObject_(sheet);
}


/**************************************************************
 * 이 스프레드시트에서 sheetNames에 해당하는 시트만 임시 스프레드시트에
 * 복사해서 xlsx로 내보낸 뒤, 임시 스프레드시트는 지웁니다
 * ("포장일엑셀다운로드" 버튼에서 사용).
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
 * name 시트가 있으면 그대로, 없으면 새로 만들어서 반환합니다.
 **************************************************************/
function getOrCreateSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
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


// 날짜 값("생산일" 등)은 셀에 문자열로 써도 시트가 자동으로 Date로
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
