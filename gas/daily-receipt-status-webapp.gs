/**************************************************************
 * "일일접수현황" 스프레드시트 조회/편집용 Web App
 *
 * 배포 위치
 * ------------------------------------------------------------
 * 이 파일은 "일일접수현황" 스프레드시트의 확장 프로그램 > Apps Script
 * 프로젝트에 추가합니다(다른 파일과 마찬가지로 그 스프레드시트에
 * 바인딩된 스크립트라, 이 파일 안에서는 스프레드시트 ID를 몰라도
 * SpreadsheetApp.getActiveSpreadsheet()로 항상 그 파일 자신을
 * 가리킵니다).
 *
 * 하는 일
 * ------------------------------------------------------------
 * 아직 "일일접수현황" 시트의 실제 탭 구조/열 이름을 모르는 상태라,
 * 우선 일반적인 조회/편집/내보내기만 되는 뼈대만 만들어뒀습니다
 * (daily-packaging-webapp.gs와 같은 방식). 이 연결로 실제 어떤 버튼/
 * 동작을 만들지는 시트 구조를 확인한 뒤 별도로 채워 넣어야 합니다.
 *
 * - doGet action="sheets": 이 파일의 모든 탭 이름 + gid + 행/열 수 목록
 *   (어떤 탭을 볼지 고르기 전에, 구조 파악용으로 먼저 호출)
 * - doGet action="read" (기본값): 탭 하나를 header + rows로 반환합니다.
 *   sheet 파라미터(탭 이름) 또는 gid 파라미터로 지정, 둘 다 없으면
 *   첫 번째 탭을 씁니다. 1행을 헤더로 보고 그 아래부터 데이터 행으로
 *   취급합니다(병합 셀이 있으면 병합된 첫 칸에만 값이 들어오고
 *   나머지는 빈 값으로 옵니다).
 * - doGet action="spreadsheetUrl": 이 스프레드시트의 편집 URL을
 *   반환합니다(나중에 "데이터 교체" 같은 버튼을 새 탭에서 열 때 쓸 수
 *   있도록 미리 준비해둔 것 — 화면 쪽에는 스프레드시트 ID를 저장해두지
 *   않으므로 매번 물어봄).
 * - doPost action="save": 화면에서 수정한 행을 저장합니다. 행마다
 *   loaded(불러왔을 때 값)와 values(수정한 값)를 같이 보내면, 저장
 *   직전 시트의 실제 값과 loaded를 비교해서 그 사이에 다른 곳에서
 *   먼저 바뀐 행은 충돌로 보고 덮어쓰지 않습니다(그 행은 최신 값
 *   그대로 돌려줌). 응답에 최신 header/rows/conflicts/updatedCount를
 *   담아 돌려주므로, 화면은 그 값으로 다시 그리면 됩니다.
 * - doPost action="exportFull": 이 스프레드시트 파일 전체(모든 시트)를
 *   xlsx로 내보내 base64로 반환합니다(범용 백업용, 화면 버튼과는
 *   아직 연결되어 있지 않음).
 * - doPost action="checkAddedEntries": "추가건 확인" 버튼 액션.
 *   FACTORY_GROUPS(현재는 퍼시스 하나만 등록됨 — 원래 퍼시스/일룸 두
 *   회사였는데 일룸 쪽 시트는 없앴음)별로 "유형별자료" 시트와
 *   "이전데이터" 시트를 E열(접수번호)+F열(순번)+J열(고객명)+
 *   L열(부품코드)+N열(부품명) 기준으로 비교해서(rowKey_ 참고 —
 *   접수번호+순번+고객명만으로는 한 접수건에 딸린 여러 부품 행이 같은
 *   키로 묶일 수 있어 부품코드/부품명까지 같이 봅니다), 유형별자료에는
 *   있는데 이전데이터에는 없는(=새로 추가된) 행만 골라 "추가건" 시트에
 *   지우고 새로 씁니다. 둘 다 처리한 뒤에는, 다음 비교 기준이
 *   되도록 이전데이터 시트를 그 시점의 유형별자료 시트 전체 내용으로
 *   다시 덮어씁니다. 응답에 회사별 추가 건수를 담아 돌려줍니다.
 * - doPost action="deleteAddedEntries": "삭제" 버튼 액션. 대시보드에서
 *   체크박스로 고른 추가건 행들을 "추가건" 시트에서 지웁니다. body에
 *   factory("퍼시스", FACTORY_GROUPS의 label과 맞춰야 함)와 rows(선택한
 *   행 값 배열)를 담아 보냅니다.
 * - doPost action="generateMeetingMaterial": 추가건 탭의 "회의자료 저장"
 *   버튼 액션(예전에는 "접수저장건 시트"를 거쳐야 했지만, 그 중간
 *   시트가 필요 없어져서 이제 추가건에서 회의자료로 바로 갑니다 —
 *   추가건과 접수저장건은 열 구성이 같아서 매핑 로직은
 *   그대로 재사용). body에 rows(체크박스로 고른 추가건 행 값 배열)를
 *   담아 보내면, 그 행들만 "회의자료" 형식으로 바꿔서 회의자료에 아직
 *   없는(접수번호+순번+부품코드+고객명 기준) 건만 뒤에 추가합니다. 이미
 *   있는 행(직접 입력해둔 생산로트/원인처 등 포함)은 그대로 두고, 맨
 *   아래 "일일합계" 요약행(금액/패널티 합계, 노란색 배경)만 매번 다시
 *   계산해서 항상 마지막 줄에 오도록 합니다. 원본 추가건 시트는
 *   지우지 않습니다(새로고침해도 방금 처리한 행이 그대로 보여야
 *   헷갈리지 않는다는 요청).
 * - doPost action="exportMeetingMaterial": "회의자료 엑셀 다운로드"
 *   버튼 액션. "회의자료" 시트 하나만 담은 xlsx를 base64로
 *   반환합니다.
 * - doPost action="deleteMeetingMaterialEntries": 회의자료 보기에서
 *   체크박스로 고른 행들을 "회의자료" 시트에서 지웁니다. body에
 *   rowIndexes(지울 실제 시트 행 번호 배열)를 담아 보냅니다. 지운 뒤
 *   "일일합계" 요약행의 금액/패널티 합계를 다시 계산합니다.
 * - doPost action="deletePreviousDaysMeetingMaterial": "전일자료 삭제"
 *   버튼 액션. "회의자료" 시트에서 접수일자가 오늘보다 이전인 행을
 *   모두 지우고, "일일합계" 요약행의 금액/패널티 합계를 다시
 *   계산합니다.
 *
 * 배포 방법
 * ------------------------------------------------------------
 * 1. "일일접수현황" 스프레드시트 > 확장 프로그램 > Apps Script에 이
 *    파일 추가
 * 2. 배포 > 새 배포 > 유형: 웹 앱, 실행 계정: 나, 액세스 권한: 필요 범위
 * 3. 배포 후 나오는 웹 앱 URL을 대시보드의 "일일업무" 페이지 >
 *    "일일현황관리 데이터" > "상세" > "일일접수현황 연결"에 입력
 * 4. UrlFetchApp/DriveApp을 처음 쓰는 경우 권한 재승인이 필요할 수
 *    있습니다. 함수 선택 드롭다운에서 testAuth를 선택해 한 번 실행하고
 *    동의 화면을 통과한 뒤 다시 배포하세요.
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
 * exportFull이 쓰는 UrlFetchApp(외부 요청), exportMeetingMaterial
 * ("회의자료 엑셀 다운로드" 버튼)이 쓰는 DriveApp의 파일 생성/삭제
 * (휴지통 이동) 권한을 승인받기 위한 함수입니다. 이름에 밑줄(_)이
 * 없어야 Apps Script 편집기의 "실행할 함수" 드롭다운에 보입니다.
 * 드롭다운에서 testAuth를 선택해 실행하면 동의 화면이 뜹니다 —
 * 승인한 뒤에는 이 함수를 지우고 다시 배포해도 되고, 그냥 남겨둬도
 * 동작에는 영향이 없습니다.
 **************************************************************/
function testAuth() {
  UrlFetchApp.fetch("https://www.google.com");

  const tempFile = DriveApp.createFile("일일접수현황_권한테스트_임시파일.txt", "test");
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

    if (action === "checkAddedEntries") {
      return jsonOutput_(checkAddedEntriesAction_());
    }

    if (action === "deleteAddedEntries") {
      return jsonOutput_(deleteAddedEntriesAction_(body.factory, body.rows || []));
    }

    if (action === "generateMeetingMaterial") {
      return jsonOutput_(generateMeetingMaterialAction_(body.rows || []));
    }

    if (action === "exportMeetingMaterial") {
      return jsonOutput_(exportSheetsSubsetAsXlsxBase64_([MEETING_MATERIAL_SHEET_NAME], "회의자료"));
    }

    if (action === "deleteMeetingMaterialEntries") {
      return jsonOutput_(deleteMeetingMaterialEntriesAction_(body.rowIndexes || []));
    }

    if (action === "deletePreviousDaysMeetingMaterial") {
      return jsonOutput_(deletePreviousDaysMeetingMaterialAction_());
    }

    return jsonOutput_({ error: "알 수 없는 action입니다: " + action });
  } catch (error) {
    return jsonOutput_({ error: error.message });
  }
}


/**************************************************************
 * "추가건 확인" 버튼 액션
 *
 * FACTORY_GROUPS에 등록된 회사(현재는 퍼시스 하나)별로 "유형별자료"
 * 시트를 "이전데이터" 시트와 비교해서 rowKey_(E열 접수번호+F열 순번+
 * J열 고객명+L열 부품코드+N열 부품명) 조합이 이전데이터에는 없고
 * 유형별자료에만 있는(=새로 추가된) 행만 "추가건" 시트에 지우고 새로
 * 씁니다. 둘 다 처리한 뒤에는, 다음 비교 기준이 되도록 이전데이터
 * 시트를 그 시점의 유형별자료 시트 전체 내용으로 다시 덮어씁니다.
 **************************************************************/
const FACTORY_GROUPS = [
  { label: "퍼시스", typeSheet: "유형별자료", compareSheet: "이전데이터", addedSheet: "추가건" }
];

// F열(순번) — 원본 표시값을 그대로 가져오고, 텍스트 서식을 강제할 열 번호(1-based)
const SEQUENCE_COLUMN_NUMBER = 6;

// "추가건 확인" 원본(유형별자료)을 조회할 때 사람이 직접 제품공급업체를
// 이 두 값으로 필터링해서 붙여넣도록 안내하고 있는데, 간혹 필터가 안 된
// 자료가 섞여 들어올 수 있어서 코드에서도 한 번 더 걸러냅니다.
const ALLOWED_PRODUCT_SUPPLIERS = ["퍼시스충주1", "퍼시스충주2"];
const PRODUCT_SUPPLIER_COLUMN_LABEL = "제품공급업체";

/**************************************************************
 * 헤더 셀에 정렬/필터 화살표 같은 부가 기호가 붙어있어도 이름으로
 * 찾을 수 있도록, 정확히 일치하는 게 없으면 그 기호를 뗀 뒤에도
 * 한 번 더 비교합니다.
 **************************************************************/
function findHeaderIndexLenient_(header, label) {
  const exact = header.indexOf(label);
  if (exact !== -1) return exact;

  return header.findIndex(function(cell) {
    return String(cell || "").replace(/[▲▼△▽]/g, "").trim() === label;
  });
}

function checkAddedEntriesAction_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const results = [];

  FACTORY_GROUPS.forEach(function(group) {
    const typeSheet = ss.getSheetByName(group.typeSheet);
    if (!typeSheet) throw new Error("'" + group.typeSheet + "' 시트를 찾을 수 없습니다.");

    const compareSheet = getOrCreateSheet_(ss, group.compareSheet);
    const addedSheet = getOrCreateSheet_(ss, group.addedSheet);

    const typeValues = readAllValues_(typeSheet);

    // 1행(헤더)이 통째로 비어있으면, 유형별자료 시트가 한 줄 밀려
    // 붙여넣어진 것으로 보고(진짜 헤더가 2행에 있는 상태) 이후 로직이
    // 그대로 진행되지 않게 여기서 막습니다. 그대로 진행하면 진짜
    // 헤더 텍스트가 데이터 행으로 잘못 섞여 추가건/이전데이터 시트가
    // 전부 틀어집니다.
    const typeHeaderRow = typeValues.length ? typeValues[0] : [];
    const isTypeHeaderBlank = !typeHeaderRow.length ||
      typeHeaderRow.every(function(cell) { return normalizeText_(cell) === ""; });
    if (isTypeHeaderBlank) {
      throw new Error("1행이 비어있으면 심각한 오류가 발생합니다. 확인해주세요.");
    }

    const compareValues = readAllValues_(compareSheet);
    const compareKeys = buildKeySet_(compareValues);

    const supplierColIndex = findHeaderIndexLenient_(typeHeaderRow, PRODUCT_SUPPLIER_COLUMN_LABEL);

    if (supplierColIndex === -1) {
      throw new Error(
        "'" + group.typeSheet + "'에서 '" + PRODUCT_SUPPLIER_COLUMN_LABEL + "' 열을 찾을 수 없습니다.\n\n" +
        "헤더(1행)에 이 이름이 정확히 들어있는지 확인해 주세요."
      );
    }

    const addedRows = [];
    for (let i = 1; i < typeValues.length; i++) {
      const row = typeValues[i];
      if (compareKeys.has(rowKey_(row))) continue;

      // 조회할 때 사람이 직접 제품공급업체를 퍼시스충주1/2로 필터링해서
      // 붙여넣도록 안내하고 있는데, 필터가 안 된 자료가 섞여 들어와도
      // 여기서 한 번 더 걸러서 추가건에는 그 두 곳 것만 들어가게 합니다.
      if (ALLOWED_PRODUCT_SUPPLIERS.indexOf(normalizeText_(row[supplierColIndex])) === -1) continue;

      addedRows.push(row);
    }

    const header = typeValues.length ? typeValues[0] : [];
    writeSheetFullReplace_(addedSheet, header, addedRows);

    results.push({ factory: group.label, addedCount: addedRows.length });

    // 다음 비교 기준으로 쓰일 "이전데이터"도 유형별자료 시트가 이번에
    // 바뀌지 않았으니 위에서 이미 읽어둔 typeValues를 그대로 재사용합니다
    // (유형별자료를 다시 읽는 두 번째 getValues() 호출을 없앰).
    const compareDataRows = typeValues.slice(1);
    writeSheetFullReplace_(compareSheet, header, compareDataRows);
  });

  return { ok: true, results: results };
}


/**************************************************************
 * E열(접수번호)+F열(순번)+J열(고객명)+L열(부품코드)+N열(부품명)을
 * 합친 비교 키. 접수번호+순번+고객명만으로는 한 접수건에 여러 부품이
 * 딸려 순번이 비어있거나 같은 값으로 중복될 수 있어서(실제로 이
 * 조합만으로 서로 다른 두 행이 같은 키로 묶여 "추가건 확인"에서
 * 중복 집계되거나 "삭제"에서 엉뚱한 행까지 함께 지워지는 문제가
 * 있었음), 부품을 구분하는 L열/N열까지 같이 넣어 키를 더 좁힙니다.
 **************************************************************/
function rowKey_(row) {
  return [
    normalizeText_(row[4]),  // E열 접수번호
    normalizeText_(row[5]),  // F열 순번
    normalizeText_(row[9]),  // J열 고객명
    normalizeText_(row[11]), // L열 부품코드
    normalizeText_(row[13])  // N열 부품명
  ].join("||");
}


function buildKeySet_(values) {
  const set = new Set();
  for (let i = 1; i < values.length; i++) {
    set.add(rowKey_(values[i]));
  }
  return set;
}


/**************************************************************
 * 시트 전체 값을 읽어옵니다. F열(순번)은 getValues() 대신
 * getDisplayValues()로 읽습니다 — 원본 시트에 그 열이 실제로는
 * 숫자(1, 2...)로 저장되어 있고 "00" 같은 서식으로 화면에만 0을
 * 채워 보여주는 경우, getValues()는 서식을 무시한 숫자만 주지만
 * getDisplayValues()는 화면에 보이는 그대로("02")를 문자열로 주기
 * 때문에, 행마다 서식이 다르더라도 항상 화면에 보이는 값 그대로
 * 비교/복사할 수 있습니다.
 **************************************************************/
function readAllValues_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 1 || lastColumn < 1) return [];

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();

  if (lastColumn >= SEQUENCE_COLUMN_NUMBER) {
    const sequenceDisplayValues = sheet
      .getRange(1, SEQUENCE_COLUMN_NUMBER, lastRow, 1)
      .getDisplayValues();

    for (let i = 0; i < values.length; i++) {
      values[i][SEQUENCE_COLUMN_NUMBER - 1] = sequenceDisplayValues[i][0];
    }
  }

  return values;
}


/**************************************************************
 * 시트 내용을 지우고 header + dataRows로 완전히 새로 씁니다.
 **************************************************************/
function writeSheetFullReplace_(sheet, header, dataRows) {
  sheet.clear();

  const allRows = header.length ? [header].concat(dataRows) : dataRows;
  if (!allRows.length) return;

  const columnCount = allRows.reduce(function(max, row) {
    return Math.max(max, row.length);
  }, 0);

  const hasHeader = !!header.length;
  const normalized = allRows.map(function(row, idx) {
    const fixedRow = normalizeRowLength_(row, columnCount);
    const isDataRow = !hasHeader || idx > 0;
    if (isDataRow && fixedRow.length >= SEQUENCE_COLUMN_NUMBER) {
      fixedRow[SEQUENCE_COLUMN_NUMBER - 1] = padSequenceValue_(fixedRow[SEQUENCE_COLUMN_NUMBER - 1]);
    }
    return fixedRow;
  });

  // F열(순번)은 "01"처럼 앞자리 0이 있는 값이 숫자로 바뀌지 않도록,
  // 값을 쓰기 전에 그 열을 텍스트 서식으로 먼저 고정해둡니다.
  if (columnCount >= SEQUENCE_COLUMN_NUMBER) {
    sheet.getRange(1, SEQUENCE_COLUMN_NUMBER, normalized.length, 1).setNumberFormat("@");
  }

  // 색상 열도 같은 이유로 값을 쓰기 전에 텍스트 서식으로 고정합니다
  // ("061" 같은 값이 서식이 General인 채로 쓰이면 숫자로 바뀜).
  const colorColIndex = header.indexOf("색상");
  const dataRowCount = normalized.length - (hasHeader ? 1 : 0);

  if (colorColIndex !== -1 && colorColIndex < columnCount && dataRowCount > 0) {
    sheet.getRange(hasHeader ? 2 : 1, colorColIndex + 1, dataRowCount, 1).setNumberFormat("@");
  }

  sheet.getRange(1, 1, normalized.length, columnCount).setValues(normalized);
}


/**************************************************************
 * 순번 값을 2자리 0채움 문자열로 맞춥니다. 원본 시트의 순번 열이
 * 실제로는 숫자(1, 2, 3...)로 저장되어 있고 화면에만 "00" 서식으로
 * 0을 채워 보여주는 경우, getValues()로는 숫자 그대로(1)만 받아오게
 * 되므로 여기서 직접 0을 채워줍니다. 이미 "01"처럼 텍스트로 들어온
 * 값이나 숫자가 아닌 값은 그대로 둡니다.
 **************************************************************/
function padSequenceValue_(value) {
  if (value === null || value === undefined || value === "") return value;
  const text = String(value).trim();
  if (!/^\d+$/.test(text)) return value;
  return text.length < 2 ? ("0" + text) : text;
}


/**************************************************************
 * "삭제" 버튼 액션
 *
 * 체크박스로 고른 추가건 행들을 "추가건" 시트에서 지웁니다.
 **************************************************************/
function deleteAddedEntriesAction_(factoryLabel, rows) {
  const group = FACTORY_GROUPS.filter(function(g) { return g.label === factoryLabel; })[0];
  if (!group) throw new Error("알 수 없는 공장입니다: " + factoryLabel);
  if (!rows.length) return { ok: true, factory: factoryLabel, deletedCount: 0 };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const addedSheet = ss.getSheetByName(group.addedSheet);
  if (!addedSheet) throw new Error("'" + group.addedSheet + "' 시트를 찾을 수 없습니다.");

  const deletedCount = removeRowsByKey_(addedSheet, rows);

  return { ok: true, factory: factoryLabel, deletedCount: deletedCount };
}


/**************************************************************
 * rowKey_(E열 접수번호+F열 순번+J열 고객명+L열 부품코드+N열 부품명)로
 * sheet에서 rows에 해당하는 행을 찾아 지우고, 나머지 내용으로 시트를
 * 다시 씁니다. "삭제" 버튼 액션이 씁니다. 지운 행 수를 반환합니다.
 **************************************************************/
function removeRowsByKey_(sheet, rows) {
  const keysToRemove = new Set(rows.map(rowKey_));
  const allValues = readAllValues_(sheet);
  const header = allValues.length ? allValues[0] : [];

  const remainingRows = [];
  let removedCount = 0;

  for (let i = 1; i < allValues.length; i++) {
    const row = allValues[i];
    if (keysToRemove.has(rowKey_(row))) {
      removedCount++;
    } else {
      remainingRows.push(row);
    }
  }

  writeSheetFullReplace_(sheet, header, remainingRows);

  return removedCount;
}


// "회의자료" 시트 — 접수저장건에서 만들어지는 회의 준비용 시트.
// 헤더는 시트에 이미 들어가 있으므로 이 파일에서는 절대 지우지 않고,
// 데이터 행만 다룹니다.
const MEETING_MATERIAL_SHEET_NAME = "회의자료";
const MEETING_MATERIAL_HEADER = [
  "접수일자", "번호", "구분", "접수번호", "순번", "부품코드", "색상",
  "생산로트", "고객명", "원인처", "포장처", "유형", "하자내역",
  "미결구분", "수량", "금액", "패널티", "일일합계"
];
const MEETING_MATERIAL_PENALTY_AMOUNT = 60000;
const MEETING_MATERIAL_SEQUENCE_COLUMN = 5; // E열(순번) — 1-based
const MEETING_MATERIAL_AMOUNT_COLUMN_INDEX = 15; // 금액(0-based)
const MEETING_MATERIAL_PENALTY_COLUMN_INDEX = 16; // 패널티(0-based)
const MEETING_MATERIAL_DAILY_TOTAL_COLUMN_INDEX = 17; // 일일합계(0-based, R열)
const MEETING_MATERIAL_SUMMARY_LABEL = "일일합계";
const MEETING_MATERIAL_SUMMARY_COLOR = "#fff4d6";


/**************************************************************
 * 추가건 탭의 "회의자료 저장" 버튼 액션.
 *
 * 대시보드에서 체크박스로 고른 추가건 행들(rows)만 회의자료 형식으로
 * 바꿔서, 회의자료에 아직 없는(접수번호+순번+부품코드+고객명 기준)
 * 건만 뒤에 추가합니다. 이미 있는 행(직접 입력해둔 생산로트/원인처/
 * 포장처/유형/하자내역/미결구분 포함)은 그대로 두고, 맨 아래 "일일합계"
 * 요약행(금액/패널티 합계)만 매번 다시 계산해서 항상 마지막 줄에
 * 오도록 다시 씁니다. 원본 추가건 시트는 지우지 않습니다(새로고침해도
 * 방금 처리한 행이 그대로 보여야 헷갈리지 않는다는 요청).
 *
 * 예전에는 추가건 → 접수저장건 → 회의자료 순으로 한 단계 거쳐야
 * 했지만, 그 중간 시트가 필요 없어져서 이제 추가건에서 바로 옵니다.
 * 추가건은 유형별자료 시트를 그대로 옮긴 것이라 열 구성이 접수저장건과
 * 같아서, 아래 매핑은 그대로 재사용합니다.
 *
 * 열 매핑(추가건 → 회의자료):
 * D열(브랜드)→구분, E열(접수번호)→접수번호, F열(순번)→순번,
 * L열(부품코드)→부품코드, O열(칼라)→색상, J열(고객명)→고객명,
 * Q열(조치)→수량, AI열(금액)→금액. 생산로트/원인처/포장처/유형/
 * 하자내역/미결구분은 직접 입력하도록 빈칸으로 둡니다. 패널티는
 * 새로 추가되는 모든 행에 60,000을 자동으로 채웁니다. R열(일일합계)은
 * 각 행마다 그 행의 금액+패널티를 자동으로 채우고(기존 행도 매번
 * 다시 계산), 맨 아래 요약행의 일일합계 칸에는 금액 합계+패널티
 * 합계(=일일합계 열 전체 합계)를 채웁니다.
 **************************************************************/
function generateMeetingMaterialAction_(rows) {
  if (!rows.length) return { ok: true, addedCount: 0 };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const meetingSheet = getOrCreateSheet_(ss, MEETING_MATERIAL_SHEET_NAME);
  const columnCount = MEETING_MATERIAL_HEADER.length;

  if (meetingSheet.getLastRow() < 1) {
    meetingSheet.getRange(1, 1, 1, columnCount).setValues([MEETING_MATERIAL_HEADER]);
  }

  // 헤더 테두리가 R열(일일합계)을 추가하기 전 기준(Q열까지)으로 남아있을
  // 수 있어서, 매번 헤더 전체 폭(R열 포함)에 테두리를 다시 그려둡니다
  // (내용/다른 서식은 건드리지 않음).
  meetingSheet.getRange(1, 1, 1, columnCount).setBorder(true, true, true, true, true, true);

  // 헤더 배경색(초록)과 필터(▼)도 R열을 추가하기 전 기준(Q열까지)으로
  // 남아있을 수 있어서, 바로 왼쪽(Q열) 헤더 칸의 서식을 R열 헤더 칸에
  // 그대로 복사해 맞추고, 필터 범위도 R열까지 포함하도록 다시 만듭니다.
  if (columnCount >= 2) {
    meetingSheet.getRange(1, columnCount - 1, 1, 1).copyTo(
      meetingSheet.getRange(1, columnCount, 1, 1),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false
    );
  }

  const existingFilter = meetingSheet.getFilter();
  if (existingFilter) {
    const filterRange = existingFilter.getRange();
    if (filterRange.getLastColumn() < columnCount) {
      const filterFirstRow = filterRange.getRow();
      const filterNumRows = filterRange.getNumRows();
      existingFilter.remove();
      meetingSheet.getRange(filterFirstRow, 1, filterNumRows, columnCount).createFilter();
    }
  }

  const lastRow = meetingSheet.getLastRow();
  const existingDataRows = lastRow >= 2
    ? meetingSheet.getRange(2, 1, lastRow - 1, columnCount).getValues()
    : [];

  // 맨 마지막 "일일합계" 요약행은 실제 건이 아니므로 떼어놓고 계산
  let realRows = existingDataRows;
  if (realRows.length && normalizeText_(realRows[realRows.length - 1][0]) === MEETING_MATERIAL_SUMMARY_LABEL) {
    realRows = realRows.slice(0, realRows.length - 1);
  }

  // 접수번호+순번+부품코드+고객명 기준 중복 검사(회의자료 쪽 열 위치)
  const existingKeys = new Set(realRows.map(function(row) {
    return [
      normalizeText_(row[3]), // 접수번호
      normalizeText_(row[4]), // 순번
      normalizeText_(row[5]), // 부품코드
      normalizeText_(row[8])  // 고객명
    ].join("||");
  }));

  let nextNumber = realRows.reduce(function(max, row) {
    return Math.max(max, Number(row[1]) || 0);
  }, 0) + 1;

  const today = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd");
  const newRows = [];

  rows.forEach(function(row) {
    // 접수저장건 쪽 열 위치(E접수번호+F순번+L부품코드+J고객명)로 키를 만들어 비교
    const key = [
      normalizeText_(row[4]),  // E열 접수번호
      normalizeText_(row[5]),  // F열 순번
      normalizeText_(row[11]), // L열 부품코드
      normalizeText_(row[9])   // J열 고객명
    ].join("||");
    if (existingKeys.has(key)) return;
    existingKeys.add(key);

    const amount = Number(String(row[34]).replace(/,/g, "")) || 0; // AI열 금액

    newRows.push([
      today,
      nextNumber,
      row[3] || "",  // D열 브랜드 → 구분
      row[4] || "",  // E열 접수번호
      padSequenceValue_(row[5] || ""), // F열 순번
      row[11] || "", // L열 부품코드
      row[14] || "", // O열 칼라 → 색상
      "",            // 생산로트(직접 입력)
      row[9] || "",  // J열 고객명
      "",            // 원인처(직접 입력)
      "",            // 포장처(직접 입력)
      "",            // 유형(직접 입력)
      "",            // 하자내역(직접 입력)
      "",            // 미결구분(직접 입력)
      row[16] || "", // Q열 조치 → 수량
      amount,
      MEETING_MATERIAL_PENALTY_AMOUNT,
      amount + MEETING_MATERIAL_PENALTY_AMOUNT // 일일합계(R열) = 이 행의 금액+패널티
    ]);
    nextNumber++;
  });

  // 기존 행도 열 개수를 최신 헤더 기준으로 맞추고, 일일합계(R열)를
  // 그 행의 현재 금액+패널티 기준으로 다시 계산해둡니다(예전에 이
  // 열이 없을 때 추가된 행이거나, 그 사이 금액/패널티가 수정된 행도
  // 항상 최신 값으로 맞추기 위함).
  const allRealRows = realRows.concat(newRows).map(function(row) {
    const fixedRow = normalizeRowLength_(row, columnCount);
    const amount = Number(String(fixedRow[MEETING_MATERIAL_AMOUNT_COLUMN_INDEX]).replace(/,/g, "")) || 0;
    const penalty = Number(String(fixedRow[MEETING_MATERIAL_PENALTY_COLUMN_INDEX]).replace(/,/g, "")) || 0;
    fixedRow[MEETING_MATERIAL_DAILY_TOTAL_COLUMN_INDEX] = amount + penalty;
    return fixedRow;
  });

  const amountTotal = allRealRows.reduce(function(sum, row) {
    return sum + (Number(String(row[MEETING_MATERIAL_AMOUNT_COLUMN_INDEX]).replace(/,/g, "")) || 0);
  }, 0);
  const penaltyTotal = allRealRows.reduce(function(sum, row) {
    return sum + (Number(String(row[MEETING_MATERIAL_PENALTY_COLUMN_INDEX]).replace(/,/g, "")) || 0);
  }, 0);

  const summaryRow = normalizeRowLength_([MEETING_MATERIAL_SUMMARY_LABEL], columnCount);
  summaryRow[MEETING_MATERIAL_AMOUNT_COLUMN_INDEX] = amountTotal;
  summaryRow[MEETING_MATERIAL_PENALTY_COLUMN_INDEX] = penaltyTotal;
  summaryRow[MEETING_MATERIAL_DAILY_TOTAL_COLUMN_INDEX] = amountTotal + penaltyTotal; // 일일합계 열 전체 합계

  const finalRows = allRealRows.concat([summaryRow]);

  if (lastRow >= 2) {
    // 이전 요약행의 노란색 배경이 새로 채워지는 실제 데이터 행에
    // 남아있지 않도록, 내용과 배경을 함께 지운 뒤 다시 씁니다.
    meetingSheet.getRange(2, 1, lastRow - 1, columnCount).clearContent();
    meetingSheet.getRange(2, 1, lastRow - 1, columnCount).setBackground(null);
  }

  if (finalRows.length) {
    // A열(접수일자)은 "2026-07-23" 문자열을 그대로 쓰면 시트가 자동으로
    // 실제 Date 값으로 바꿔버려서, 나중에 화면에서 불러온 값과 비교할 때
    // 표현이 달라져 저장이 "충돌"로 잘못 처리되는 문제가 있었습니다.
    // 그래서 텍스트 서식으로 먼저 고정해 순수 문자열로 남게 합니다.
    // E열(순번)은 "01"처럼 앞자리 0이 있는 값이 숫자로 바뀌지 않도록
    // 값을 쓰기 전에 텍스트 서식으로 먼저 고정해둡니다. 금액/패널티/
    // 일일합계는 반대로, 이 시트에 남아있을 수 있는 날짜 서식이
    // 숫자에 씌워져 날짜처럼 표시되는 일이 없도록 일반 숫자 서식으로
    // 고정해둡니다.
    meetingSheet.getRange(2, 1, finalRows.length, 1).setNumberFormat("@");
    meetingSheet.getRange(2, MEETING_MATERIAL_SEQUENCE_COLUMN, finalRows.length, 1).setNumberFormat("@");

    // G열(색상)도 같은 이유로 값을 쓰기 전에 텍스트 서식으로 고정합니다.
    const meetingColorColIndex = MEETING_MATERIAL_HEADER.indexOf("색상");

    if (meetingColorColIndex !== -1) {
      meetingSheet.getRange(2, meetingColorColIndex + 1, finalRows.length, 1).setNumberFormat("@");
    }

    meetingSheet.getRange(2, MEETING_MATERIAL_AMOUNT_COLUMN_INDEX + 1, finalRows.length, 1).setNumberFormat("#,##0");
    meetingSheet.getRange(2, MEETING_MATERIAL_PENALTY_COLUMN_INDEX + 1, finalRows.length, 1).setNumberFormat("#,##0");
    meetingSheet.getRange(2, MEETING_MATERIAL_DAILY_TOTAL_COLUMN_INDEX + 1, finalRows.length, 1).setNumberFormat("#,##0");
    meetingSheet.getRange(2, 1, finalRows.length, columnCount).setValues(finalRows);

    const summaryRowNumber = finalRows.length + 1; // 1행은 헤더, 데이터는 2행부터
    meetingSheet.getRange(summaryRowNumber, 1, 1, columnCount).setBackground(MEETING_MATERIAL_SUMMARY_COLOR);
  }

  return { ok: true, addedCount: newRows.length, totalRows: finalRows.length };
}


/**************************************************************
 * 회의자료 보기에서 체크박스로 고른 행을 "회의자료" 시트에서
 * 지웁니다. rowIndex(실제 시트 행 번호)로 정확히 지정해서 지우고,
 * 지운 뒤에는 "일일합계" 요약행의 금액/패널티 합계를 남은 행 기준으로
 * 다시 계산합니다.
 **************************************************************/
function deleteMeetingMaterialEntriesAction_(rowIndexes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MEETING_MATERIAL_SHEET_NAME);
  if (!sheet) throw new Error("'" + MEETING_MATERIAL_SHEET_NAME + "' 시트를 찾을 수 없습니다.");

  const uniqueDescending = Array.from(new Set(rowIndexes)).sort(function(a, b) { return b - a; });
  uniqueDescending.forEach(function(rowIndex) {
    sheet.deleteRow(rowIndex);
  });

  recomputeMeetingMaterialSummaryRow_(sheet);

  return { ok: true, deletedCount: uniqueDescending.length };
}


/**************************************************************
 * "전일자료 삭제" 버튼 액션. "회의자료" 시트에서 접수일자가 오늘보다
 * 이전인 행을 모두 지운 뒤, "일일합계" 요약행의 금액/패널티 합계를
 * 남은 행 기준으로 다시 계산합니다.
 **************************************************************/
function deletePreviousDaysMeetingMaterialAction_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MEETING_MATERIAL_SHEET_NAME);
  if (!sheet) throw new Error("'" + MEETING_MATERIAL_SHEET_NAME + "' 시트를 찾을 수 없습니다.");

  const lastRow = sheet.getLastRow();
  const columnCount = MEETING_MATERIAL_HEADER.length;
  if (lastRow < 2) return { ok: true, deletedCount: 0 };

  const today = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd");
  const dataRows = sheet.getRange(2, 1, lastRow - 1, columnCount).getValues();

  let deletedCount = 0;
  for (let i = dataRows.length - 1; i >= 0; i--) {
    const row = dataRows[i];
    if (normalizeText_(row[0]) === MEETING_MATERIAL_SUMMARY_LABEL) continue; // 요약행은 남겨둠

    const receivedDate = formatDateOnly_(row[0]);
    if (receivedDate && receivedDate < today) {
      sheet.deleteRow(2 + i);
      deletedCount++;
    }
  }

  recomputeMeetingMaterialSummaryRow_(sheet);

  return { ok: true, deletedCount: deletedCount };
}


/**************************************************************
 * "회의자료" 시트 맨 아래 "일일합계" 요약행의 금액/패널티 합계를,
 * 그 위 실제 데이터 행들(요약행 자신은 제외) 기준으로 다시 계산해서
 * 반영합니다. 행을 지운 뒤(체크박스 삭제, 전일자료 삭제)에는 반드시
 * 이 함수를 불러줘야 합계가 최신 상태로 유지됩니다.
 **************************************************************/
function recomputeMeetingMaterialSummaryRow_(sheet) {
  const lastRow = sheet.getLastRow();
  const columnCount = MEETING_MATERIAL_HEADER.length;
  if (lastRow < 2) return;

  const dataRows = sheet.getRange(2, 1, lastRow - 1, columnCount).getValues();
  const lastIndex = dataRows.length - 1;

  if (normalizeText_(dataRows[lastIndex][0]) !== MEETING_MATERIAL_SUMMARY_LABEL) return;

  const realRows = dataRows.slice(0, lastIndex);
  const amountTotal = realRows.reduce(function(sum, row) {
    return sum + (Number(String(row[MEETING_MATERIAL_AMOUNT_COLUMN_INDEX]).replace(/,/g, "")) || 0);
  }, 0);
  const penaltyTotal = realRows.reduce(function(sum, row) {
    return sum + (Number(String(row[MEETING_MATERIAL_PENALTY_COLUMN_INDEX]).replace(/,/g, "")) || 0);
  }, 0);

  const summaryRowNumber = 2 + lastIndex;
  sheet.getRange(summaryRowNumber, MEETING_MATERIAL_AMOUNT_COLUMN_INDEX + 1)
    .setNumberFormat("#,##0").setValue(amountTotal);
  sheet.getRange(summaryRowNumber, MEETING_MATERIAL_PENALTY_COLUMN_INDEX + 1)
    .setNumberFormat("#,##0").setValue(penaltyTotal);
  sheet.getRange(summaryRowNumber, MEETING_MATERIAL_DAILY_TOTAL_COLUMN_INDEX + 1)
    .setNumberFormat("#,##0").setValue(amountTotal + penaltyTotal); // 일일합계 열 전체 합계
}


/**************************************************************
 * 날짜 셀 값을 "yyyy-MM-dd" 문자열로 맞춥니다(Date 타입/텍스트 모두
 * 처리) — 전일자료 삭제에서 접수일자와 오늘 날짜를 비교할 때 씁니다.
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
 * 이 스프레드시트에서 sheetNames에 해당하는 시트만 임시 스프레드시트에
 * 복사해서 xlsx로 내보낸 뒤, 임시 스프레드시트는 지웁니다
 * ("회의자료 엑셀 다운로드" 버튼에서 사용).
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


// 정렬 화살표(▲▼△▽) 같은 부가 기호가 이름/값 끝에 붙어있어도 실제
// 내용이 같으면 같은 값으로 보도록, 비교/키 생성 전에 그 기호를
// 지우고 앞뒤 공백을 정리합니다.
//
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
  const text = String(value === null || value === undefined ? "" : value)
    .replace(/[▲▼△▽]/g, "")
    .trim();
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
