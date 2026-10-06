/**************************************************************
 * "월마감 외작건정리" 스프레드시트 조회용 Web App
 *
 * 배포 위치
 * ------------------------------------------------------------
 * 이 파일은 "월마감 외작건정리"(마감된 월마감 외작건 자료를 옮겨 담는)
 * 스프레드시트의 확장 프로그램 > Apps Script 프로젝트에 새로 추가합니다.
 *
 * 하는 일
 * ------------------------------------------------------------
 * - doGet action="sheets": 이 파일의 모든 탭 이름 + gid + 행/열 수 목록
 *   (구조 파악용).
 * - doGet action="read"(기본값): sheet/gid 파라미터로 지정한 탭을
 *   header+rows로 반환합니다. 둘 다 없으면 "시트1"을 읽습니다.
 * - doGet action="spreadsheetUrl": 이 스프레드시트의 편집 URL을
 *   반환합니다. 대시보드의 "외작월마감 자료교체" 버튼이 확인창을
 *   띄운 뒤 이 스프레드시트를 새 탭에서 열 때 씁니다(마감된 외작
 *   월마감 자료로 시트1을 바꾸는 실제 작업은 그 스프레드시트에서
 *   직접 함).
 * - doPost action="clean": "자료정리" 버튼 액션. "시트1"에서 "종합"
 *   시트로 자료를 정리해 옮기고, 원인별 상세 시트("시트2" 양식을
 *   복사해 원인 이름으로 만듦)도 함께 채웁니다. 자세한 내용은
 *   cleanMonthlyOutsourceListAction_, updateOutsourceCauseSheets_ 참고.
 *   [신규] body에 recoveryAccumulateUrl("회수데이터" 연결 웹앱 URL,
 *   화면이 일일업무 페이지에 저장해둔 값을 그대로 같이 보냄)이 있으면,
 *   "회수누적" 탭에서 문서번호+품명이 이 시트의 접수번호+제품코드와
 *   똑같은 건의 사진(Q열 링크)을 찾아 원인별 상세 시트 데이터 끝에서
 *   두 줄 아래부터 가로 6장씩 붙입니다(insertCauseSheetImages_). 칸
 *   위치는 그 건의 실제 순번 기준(압축하지 않음 — 9번만 있으면 9번
 *   자리에 옴)이고, 사진 폭은 1~X(24)열 너비 합을 6등분해서 X열을
 *   넘지 않고 고르게 간격을 두도록 맞춥니다(원본 비율 고정). 번호는
 *   사진의 오른쪽 아래 모서리에 숫자 뱃지 이미지로 겹쳐 올립니다. URL이
 *   없거나 불러오기에 실패해도 자료정리 자체는 그대로 진행합니다
 *   (사진만 생략, 응답의 recoveryImageNote에 사유가 담김).
 * - doPost action="exportResult": "정리파일다운로드" 버튼 액션.
 *   "시트1"과 "시트2"(양식 시트)을 뺀 나머지 모든 시트(종합 + 원인별
 *   상세 시트)를 담은 xlsx를 base64로 반환합니다.
 *
 * 배포 방법
 * ------------------------------------------------------------
 * 1. 월마감 외작건정리 스프레드시트 > 확장 프로그램 > Apps Script에 이 파일 추가
 * 2. 배포 > 새 배포 > 유형: 웹 앱, 실행 계정: 나, 액세스 권한: 필요 범위
 * 3. 배포 후 나오는 웹 앱 URL을 대시보드의 "월간업무" 페이지 >
 *    "월간업무 데이터" > "상세" > "월마감 외작건정리 연결"에 입력
 * 4. UrlFetchApp/DriveApp을 처음 쓰는 경우(exportResult) 권한 재승인이
 *    필요할 수 있습니다. 함수 선택 드롭다운에서 testAuth를 선택해 한 번
 *    실행하고 동의 화면을 통과한 뒤 다시 배포하세요.
 *
 * 주의
 * ------------------------------------------------------------
 * - 토큰 검증이 없으므로 URL을 아는 사람은 누구나 이 시트를 읽고
 *   수정할 수 있습니다.
 * - "종합" 시트의 제목/기준 안내/헤더(순번~조치결과, "패널티 금액(원)"
 *   묶음 헤더 포함)는 이 스크립트가 만들지 않습니다 — 이미 만들어진
 *   양식이 있다고 가정하고, "순번" 글자가 있는 헤더 행을 찾아 그
 *   바로 아래부터만 데이터를 새로 씁니다.
 * - 필요한 행 수가 기존 데이터 영역보다 많으면 부족한 만큼 실제로
 *   행을 삽입하고, 데이터 행/구분행(그룹 소계)/합계행 각각 알맞은
 *   기존 행의 서식(열 너비·글자 크기·정렬·숫자 서식 등)을 그대로
 *   복사해서 입힙니다 — 그냥 빈 칸에 값만 채우면 서식이 없어서 글자가
 *   안 굵어지거나 내용이 옆 칸으로 넘치는 문제가 있어서 이렇게 처리함.
 **************************************************************/

const SOURCE_SHEET_NAME = "시트1"; // 대시보드가 기본으로 읽는 탭
const SUMMARY_SHEET_NAME = "종합";

// "종합" 시트에서 이 글자가 있는 행을 헤더 행으로 보고, 그 열 이름들로
// 실제 열 위치를 찾습니다(양식의 열 순서가 바뀌어도 이름으로 찾으므로
// 안전합니다). 헤더 바로 다음 행부터 데이터를 씁니다.
const SUMMARY_HEADER_MARKER = "순번";

// "시트1"에서 그대로 옮기는 열(이름이 "종합"과 동일)
const DIRECT_COLUMN_LABELS = [
  "브랜드", "지역센터", "접수번호", "구분", "형태",
  "고객명", "부품명", "제품코드", "색상", "수량"
];

// "시트1" 금액 → "종합" 제품가
const AMOUNT_SOURCE_LABEL = "금액";
// "시트1"에서 그대로 옮기는 열(제품가 다음 순서)
const AFTER_AMOUNT_SOURCE_LABELS = ["유형", "세부유형"];
// "시트1" 원인 → 그룹 기준이면서 동시에 "종합" 업체 열 값
const CAUSE_SOURCE_LABEL = "원인";
const ACTION_RESULT_SOURCE_LABEL = "조치결과";

// 패널티는 건당 고정 금액, 합계 = 제품가 + 패널티
const PENALTY_AMOUNT = 60000;

// 원인별 상세 시트 양식(이미 만들어져 있는 시트 이름 — 헤더 1행 + 예시
// 데이터 2행짜리 구조). 원인 값과 이름이 같은 시트가 없으면 이 시트를
// 복사해서 그 원인 이름으로 새로 만듭니다.
const CAUSE_SHEET_TEMPLATE_NAME = "시트2";
// 원인별 상세 시트에서 순번 역할을 하는 열 이름(시트1의 "구분"과는
// 별개 — 이 열만 1부터 순차적으로 새로 채우고, 나머지 열은 이름이
// 같은 시트1 열 값을 그대로 복사합니다).
const CAUSE_SHEET_SEQUENCE_COLUMN_LABEL = "구분";
// 원인별 상세 시트의 "매입금액" 열은 이름이 달라서 이름 매칭이 안 되므로,
// "시트1"의 금액(AMOUNT_SOURCE_LABEL) 열 값을 명시적으로 채웁니다.
const CAUSE_SHEET_AMOUNT_COLUMN_LABEL = "매입금액";
// 날짜 열은 값을 쓰면 시트 기본 로캘 서식("2026. 8. 13")으로 보이므로,
// 시트1과 같은 "yyyy-mm-dd" 형식으로 명시적으로 맞춰줍니다.
const CAUSE_SHEET_DATE_COLUMN_LABELS = ["최종조치일", "반납일자"];
const CAUSE_SHEET_DATE_NUMBER_FORMAT = "yyyy-mm-dd";
// 위 날짜 열이 이 너비(px)보다 좁으면 "yyyy-mm-dd"가 다 안 보이고
// "###..."으로 깨지므로, 그럴 때만 이 너비로 넓혀줍니다(이미 이보다
// 넓으면 그대로 둠).
const CAUSE_SHEET_DATE_MIN_COLUMN_WIDTH = 80;
const GRAND_TOTAL_LABEL = "합계";

// "회수현황"(회수데이터 연결) 사진 첨부 — 원인별 상세 시트 데이터 끝에서
// 두 줄 아래부터 가로 6장씩 사진을 붙입니다. 회수누적 탭의 "문서번호"+
// "품명" 값이 이 시트(시트1 기준)의 "접수번호"+"제품코드"와 글자 그대로
// 같은 건만 매칭합니다("품명" 칸에 실제로는 제품코드 값이 들어있다고
// 확인받음).
const RECOVERY_DOC_NO_COLUMN_LABEL = "문서번호"; // 회수누적 탭 기준
const RECOVERY_PRODUCT_COLUMN_LABEL = "품명"; // 회수누적 탭 기준
const CAUSE_SHEET_ACCESSION_SOURCE_LABEL = "접수번호"; // 시트1 기준
const CAUSE_SHEET_PRODUCT_CODE_SOURCE_LABEL = "제품코드"; // 시트1 기준
const CAUSE_SHEET_IMAGES_PER_ROW = 6;
// 사진 영역이 이 열(X열, 24번째)을 넘지 않도록, 1~24열의 실제 너비
// 합을 6등분해서 사진 폭을 정합니다 — 시트마다 열 너비가 달라도
// 6장이 항상 고르게 간격을 두고 X열 끝까지만 차지합니다.
const CAUSE_SHEET_IMAGE_AREA_LAST_COLUMN = 24;
const CAUSE_SHEET_IMAGE_GAP_PX = 10; // 사진 사이 가로 간격
const CAUSE_SHEET_IMAGE_START_GAP_ROWS = 2; // "마지막 행에서 두 줄 아래"
const CAUSE_SHEET_IMAGE_BLOCK_GAP_ROWS = 1; // 사진 줄 사이 빈 줄
// 번호를 셀 텍스트가 아니라 사진 위에 겹친 "뱃지" 이미지로 표시합니다
// (흰 바탕 둥근 네모 + 굵은 검정 숫자, 0~9 10장을 PNG로 미리 만들어
// base64로 이 파일에 내장해뒀습니다 — Apps Script에는 사진 픽셀 위에
// 직접 글자를 그리는 기능이 없어서, 숫자 자체를 작은 이미지로 준비해
// 사진 위에 또 하나의 이미지로 겹쳐 올리는 방식입니다). 두 자리 이상인
// 번호는 숫자 이미지를 옆으로 이어 붙입니다.
const RECOVERY_BADGE_SIZE_PX = 20; // 뱃지(숫자 이미지) 한 변 크기
const RECOVERY_BADGE_MARGIN_PX = 2; // 사진 오른쪽 아래 모서리에서 뱃지까지 여백

// "정리파일다운로드"에서 제외할 시트(원본 데이터/양식 시트)
const EXPORT_EXCLUDED_SHEET_NAMES = [SOURCE_SHEET_NAME, CAUSE_SHEET_TEMPLATE_NAME];


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

  const tempFile = DriveApp.createFile("월마감외작건정리_권한테스트_임시파일.txt", "test");
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
      return jsonOutput_(cleanMonthlyOutsourceListAction_(body.recoveryAccumulateUrl || ""));
    }

    if (action === "replaceSourceData") {
      return jsonOutput_(replaceSourceDataAction_(body.header || [], body.rows || []));
    }

    if (action === "exportResult") {
      return jsonOutput_(exportAllExceptAsXlsxBase64_(EXPORT_EXCLUDED_SHEET_NAMES, "월마감외작건정리"));
    }

    return jsonOutput_({ error: "알 수 없는 action입니다: " + action });
  } catch (error) {
    return jsonOutput_({ error: error.message });
  }
}


/**************************************************************
 * "자료정리" 액션
 *
 * "시트1"에서 완전히 빈 행을 뺀 나머지 행을, 원인(→업체) 열 값
 * 기준으로 묶어서 행이 많은 그룹부터(내림차순), 개수가 같으면 금액
 * 합이 높은 그룹부터 정렬해 "종합" 시트에 씁니다. 같은 그룹 안에서는
 * 시트1의 원래 순서를 유지합니다.
 *
 * 각 행에는 1부터 순차적인 순번을 매기고, 제품가(시트1 금액)+패널티
 * (건당 60,000원 고정)=합계를 계산합니다. 그룹이 끝날 때마다 순번~
 * 수량 열을 하나로 병합한 구분행을 넣어 그 안에 원인 값을 표시하고,
 * 합계 열에는 그 그룹의 합계 소계를 넣습니다. 맨 마지막에는 "합계"
 * 라벨과 전체 총합을 넣은 행을 하나 더 추가합니다.
 *
 * "종합" 시트의 제목/기준 안내/헤더 행은 그대로 두고, "순번" 글자가
 * 있는 헤더 행 바로 아래 데이터 영역만 지우고 새로 씁니다(서식은
 * clearContent만 사용해 그대로 유지 — 열 너비/글자 크기/정렬 불변).
 **************************************************************/
function cleanMonthlyOutsourceListAction_(recoveryAccumulateUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);

  if (!sourceSheet) {
    throw new Error("'" + SOURCE_SHEET_NAME + "' 시트를 찾을 수 없습니다.");
  }

  const summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);

  if (!summarySheet) {
    throw new Error("'" + SUMMARY_SHEET_NAME + "' 시트를 찾을 수 없습니다.");
  }

  const source = readSheetObject_(sourceSheet);
  const header = source.header;

  function sourceIdx(label) {
    const idx = header.indexOf(label);
    if (idx === -1) throw new Error("'" + SOURCE_SHEET_NAME + "'에서 '" + label + "' 열을 찾을 수 없습니다.");
    return idx;
  }

  const directIdx = DIRECT_COLUMN_LABELS.map(sourceIdx);
  const amountIdx = sourceIdx(AMOUNT_SOURCE_LABEL);
  const afterAmountIdx = AFTER_AMOUNT_SOURCE_LABELS.map(sourceIdx);
  const causeIdx = sourceIdx(CAUSE_SOURCE_LABEL);
  const actionResultIdx = sourceIdx(ACTION_RESULT_SOURCE_LABEL);

  const parsedRows = source.rows
    .map(function(row) {
      const values = row.values;
      return {
        raw: values, // 원인별 시트용 — 이름 매칭으로 아무 열이나 그대로 옮길 때 씀
        direct: directIdx.map(function(idx) { return values[idx]; }),
        amount: values[amountIdx],
        afterAmount: afterAmountIdx.map(function(idx) { return values[idx]; }),
        cause: values[causeIdx],
        actionResult: values[actionResultIdx]
      };
    })
    .filter(function(item) {
      return item.direct.some(function(value) { return normalizeText_(value) !== ""; }) ||
        normalizeText_(item.amount) !== "" || normalizeText_(item.cause) !== "";
    });

  // 원인 값 기준으로 그룹핑(먼저 나온 순서 기억 — 개수가 같을 때 유지용)
  const groups = {};
  const groupOrder = [];

  parsedRows.forEach(function(item) {
    const key = normalizeText_(item.cause);

    if (!groups[key]) {
      groups[key] = { causeValue: item.cause, rows: [] };
      groupOrder.push(key);
    }

    groups[key].rows.push(item);
  });

  // 행이 많은 그룹부터(내림차순). 개수가 같으면 시트1 금액 합이 높은
  // 그룹부터.
  function groupAmountSum_(key) {
    return groups[key].rows.reduce(function(sum, item) {
      return sum + (Number(item.amount) || 0);
    }, 0);
  }

  groupOrder.sort(function(a, b) {
    const rowCountDiff = groups[b].rows.length - groups[a].rows.length;
    if (rowCountDiff !== 0) return rowCountDiff;
    return groupAmountSum_(b) - groupAmountSum_(a);
  });

  // "종합" 시트의 실제 헤더 행을 찾아서, 그 열 이름으로 각 열 위치를 파악합니다.
  const summaryLastColumn = Math.max(summarySheet.getLastColumn(), 1);
  const headerScanRowCount = Math.min(30, summarySheet.getLastRow() || 30);
  const headerScanValues = headerScanRowCount > 0
    ? summarySheet.getRange(1, 1, headerScanRowCount, summaryLastColumn).getValues()
    : [];

  let headerRowNumber = -1;
  let summaryHeader = null;

  for (let i = 0; i < headerScanValues.length; i++) {
    if (normalizeText_(headerScanValues[i][0]) === SUMMARY_HEADER_MARKER) {
      headerRowNumber = i + 1;

      // "순번" 같은 단일 열 헤더는 이 행과 바로 다음 행이 세로로 병합되어
      // 있어서 실제 값은 이 행에만 있지만, "패널티 금액(원)"처럼 묶음
      // 헤더 아래 제품가/패널티/합계 같은 하위 열 이름은 다음 행에만
      // 있습니다. 그래서 두 행을 합쳐서(다음 행 값이 있으면 그걸 우선)
      // 하나의 헤더로 씁니다.
      const nextRow = headerScanValues[i + 1] || [];
      summaryHeader = headerScanValues[i].map(function(value, colIndex) {
        const nextValue = nextRow[colIndex];
        return normalizeText_(nextValue) !== "" ? nextValue : value;
      });

      // 다음 행에 새 열 이름이 하나라도 있으면 그 행도 헤더의 일부이므로
      // 데이터는 그 다음 행부터 시작합니다.
      const nextRowHasOwnLabel = nextRow.some(function(value) { return normalizeText_(value) !== ""; });
      headerRowNumber = nextRowHasOwnLabel ? headerRowNumber + 1 : headerRowNumber;
      break;
    }
  }

  if (headerRowNumber === -1) {
    throw new Error("'" + SUMMARY_SHEET_NAME + "' 시트에서 '" + SUMMARY_HEADER_MARKER + "' 헤더 행을 찾을 수 없습니다.");
  }

  // 제목("OO년 O월 협력업체 사외하자 클레임 현황")과 기간
  // ("YYYY.MM.26~YYYY.MM.25") 안내 문구를, 실행 시점(오늘) 기준
  // 전월/전전월로 자동으로 맞춥니다(마감은 항상 전월 자료 기준).
  updateSummaryDateLabels_(summarySheet, headerRowNumber);

  function summaryIdx(label) {
    const idx = summaryHeader.indexOf(label);
    if (idx === -1) throw new Error("'" + SUMMARY_SHEET_NAME + "' 헤더에서 '" + label + "' 열을 찾을 수 없습니다.");
    return idx;
  }

  const outPos = {
    brand: summaryIdx("브랜드"),
    region: summaryIdx("지역센터"),
    accession: summaryIdx("접수번호"),
    category: summaryIdx("구분"),
    type: summaryIdx("형태"),
    customer: summaryIdx("고객명"),
    partName: summaryIdx("부품명"),
    productCode: summaryIdx("제품코드"),
    color: summaryIdx("색상"),
    qty: summaryIdx("수량"),
    price: summaryIdx("제품가"),
    penalty: summaryIdx("패널티"),
    total: summaryIdx("합계"),
    kind: summaryIdx("유형"),
    subKind: summaryIdx("세부유형"),
    vendor: summaryIdx("업체"),
    actionResult: summaryIdx("조치결과")
  };

  const totalColumnCount = summaryHeader.length;
  const mergeColumnCount = outPos.qty + 1; // 순번(A열)부터 수량 열까지

  const outputRows = [];
  const mergeRowOffsets = []; // 데이터 시작 행 기준 상대 위치(0-based) — 나중에 실제 시트 행으로 변환
  let sequence = 0;
  let grandTotal = 0;

  groupOrder.forEach(function(key) {
    const group = groups[key];
    let groupTotal = 0;

    group.rows.forEach(function(item) {
      sequence++;

      const priceValue = item.amount;
      const priceNumber = Number(priceValue) || 0;
      const total = priceNumber + PENALTY_AMOUNT;
      groupTotal += total;

      const row = new Array(totalColumnCount).fill("");
      row[0] = sequence; // 순번은 항상 첫 열
      row[outPos.brand] = item.direct[0];
      row[outPos.region] = item.direct[1];
      row[outPos.accession] = item.direct[2];
      row[outPos.category] = item.direct[3];
      row[outPos.type] = item.direct[4];
      row[outPos.customer] = item.direct[5];
      row[outPos.partName] = item.direct[6];
      row[outPos.productCode] = item.direct[7];
      row[outPos.color] = item.direct[8];
      row[outPos.qty] = item.direct[9];
      row[outPos.price] = priceValue;
      row[outPos.penalty] = PENALTY_AMOUNT;
      row[outPos.total] = total;
      row[outPos.kind] = item.afterAmount[0];
      row[outPos.subKind] = item.afterAmount[1];
      row[outPos.vendor] = item.cause;
      row[outPos.actionResult] = item.actionResult;

      outputRows.push(row);
    });

    const dividerRow = new Array(totalColumnCount).fill("");
    dividerRow[0] = group.causeValue;
    dividerRow[outPos.total] = groupTotal;
    outputRows.push(dividerRow);
    mergeRowOffsets.push(outputRows.length - 1);

    grandTotal += groupTotal;
  });

  const grandTotalRow = new Array(totalColumnCount).fill("");
  grandTotalRow[0] = GRAND_TOTAL_LABEL;
  grandTotalRow[outPos.total] = grandTotal;
  outputRows.push(grandTotalRow);
  mergeRowOffsets.push(outputRows.length - 1);

  const dataStartRow = headerRowNumber + 1;
  const neededRowCount = outputRows.length;
  const existingLastRow = summarySheet.getLastRow();
  const existingDataRowCount = Math.max(existingLastRow - dataStartRow + 1, 0);

  // 지우기 전에 서식 기준으로 삼을 행을 먼저 찾아둡니다 — 일반 데이터
  // 행은 dataStartRow(첫 데이터 행) 서식을, 구분행/합계행은 그 아래에서
  // 처음 만나는 병합 행의 서식을 기준으로 삼습니다(전에 만든 데이터가
  // 없으면 서식 기준 없이 진행 — 최초 1회만 해당).
  const dataRowFormatSource = existingDataRowCount > 0
    ? summarySheet.getRange(dataStartRow, 1, 1, totalColumnCount)
    : null;
  const dividerRowNumber = existingDataRowCount > 0
    ? findFirstMergedRowNumber_(summarySheet, dataStartRow, existingLastRow)
    : -1;
  const dividerRowFormatSource = dividerRowNumber !== -1
    ? summarySheet.getRange(dividerRowNumber, 1, 1, totalColumnCount)
    : null;

  // 필요한 행 수가 기존보다 많으면 부족한 만큼 실제로 행을 삽입합니다.
  // (그냥 그 아래 빈 행에 값만 써넣으면 서식이 없는 기본 칸이라
  // 글자가 안 굵어지거나 내용이 옆 칸으로 넘치는 문제가 생겼음)
  if (neededRowCount > existingDataRowCount) {
    const extraRowCount = neededRowCount - existingDataRowCount;
    const insertAfterRow = existingDataRowCount > 0 ? existingLastRow : dataStartRow - 1;
    summarySheet.insertRowsAfter(insertAfterRow, extraRowCount);
  }

  // 기존 병합은 새 데이터와 어긋날 수 있으므로 전부 풀고 값만 지웁니다.
  const wipeRowCount = Math.max(neededRowCount, existingDataRowCount);

  if (wipeRowCount > 0) {
    const wipeRange = summarySheet.getRange(dataStartRow, 1, wipeRowCount, totalColumnCount);
    wipeRange.breakApart();
    wipeRange.clearContent();
  }

  if (outputRows.length) {
    summarySheet.getRange(dataStartRow, 1, outputRows.length, totalColumnCount).setValues(outputRows);
  }

  // 행마다 종류(데이터 행 / 구분·합계 행)에 맞는 서식을 다시 입혀서,
  // 새로 늘어난 행도 표 서식(열 너비 제외 — 열 너비는 열 단위라 항상
  // 유지됨)이 그대로 유지되게 합니다.
  const mergeOffsetSet = {};
  mergeRowOffsets.forEach(function(offset) { mergeOffsetSet[offset] = true; });

  outputRows.forEach(function(row, offset) {
    const isDividerRow = !!mergeOffsetSet[offset];
    const formatSource = isDividerRow ? (dividerRowFormatSource || dataRowFormatSource) : dataRowFormatSource;
    const targetRange = summarySheet.getRange(dataStartRow + offset, 1, 1, totalColumnCount);

    if (formatSource) {
      formatSource.copyTo(targetRange, { formatOnly: true });
    }

    if (isDividerRow) {
      targetRange.setFontWeight("bold");
    }

    // 서식 복사가 테두리까지 완전히 물려주지 못하는 경우(특히 맨 아래
    // 합계 행)를 대비해, 모든 행에 얇은 테두리를 확실하게 다시 그립니다.
    targetRange.setBorder(true, true, true, true, true, true);
  });

  mergeRowOffsets.forEach(function(offset) {
    const sheetRow = dataStartRow + offset;
    summarySheet.getRange(sheetRow, 1, 1, mergeColumnCount).merge();
  });

  // 이번에 필요한 행 수가 지난번보다 적으면(건수가 줄어든 경우), 남는
  // 아래쪽 행은 값만 지워서는 테두리 등 서식이 그대로 남아 빈 칸에
  // 테두리만 보이는 문제가 생기므로, 그 행들을 아예 삭제합니다.
  if (existingDataRowCount > neededRowCount) {
    const removeRowCount = existingDataRowCount - neededRowCount;
    summarySheet.deleteRows(dataStartRow + neededRowCount, removeRowCount);
  }

  // getLastRow()는 내용이 있는 행까지만 세기 때문에, 맨 처음 수동으로
  // 만든 양식에 내용 없이 테두리만 미리 그어둔 행이 있으면
  // existingDataRowCount에 안 잡혀서 위 삭제 로직을 피해갑니다. 그래서
  // 합계 바로 아래 일정 범위는 테두리를 확실하게 지워서 빈 칸에
  // 테두리만 남아있는 문제를 없앱니다.
  //
  // 반드시 이 다음(아래)에 데이터 범위 테두리를 "마지막으로" 다시
  // 그려야 합니다 — 합계 행의 아래쪽 테두리와 그 바로 아래 빈 행의
  // 위쪽 테두리는 시트에서 같은 경계선이라, 순서가 바뀌면 방금 지운
  // "빈 행 위쪽 테두리 없음"이 합계 행 테두리까지 같이 지워버립니다.
  clearStrayBordersBelow_(summarySheet, dataStartRow + neededRowCount, totalColumnCount);

  // merge()가 방금 합친 셀의 테두리를 다시 정리해버리는 경우가 있고,
  // 위 테두리 지우기와 경계선을 공유하므로, 모든 테두리 처리의 맨
  // 마지막에 전체 데이터 범위 테두리를 다시 그려서 확실하게 맞춥니다.
  if (outputRows.length) {
    summarySheet.getRange(dataStartRow, 1, outputRows.length, totalColumnCount)
      .setBorder(true, true, true, true, true, true);
  }

  // 원인별로 "시트2" 양식 시트를 복사(또는 이름이 같은 기존 시트를 재사용)해서
  // 원인 이름의 상세 시트를 만들고, 시트1에서 해당 원인 행만 옮겨 채웁니다.
  // "종합"은 이미 위에서 다 갱신된 뒤라, 여기서 실패해도 종합만 갱신되고
  // 원인별 시트는 예전 상태로 남을 수 있습니다 — 그 사실을 에러 메시지에
  // 명확히 남겨서 사용자가 상태가 서로 안 맞는 걸 바로 알 수 있게 합니다.
  // "회수데이터" 연결 웹앱 URL이 전달되면 "회수누적" 탭에서 문서번호+품명
  // 기준 이미지 링크 맵을 미리 한 번만 가져와 모든 원인별 시트가 같이
  // 씁니다. URL이 없거나 불러오기에 실패해도 자료정리 자체는 그대로
  // 진행하고(사진만 못 붙임), 그 사실을 응답 메시지에 남깁니다.
  let recoveryImageMap = {};
  let recoveryImageNote = "";

  if (recoveryAccumulateUrl) {
    try {
      recoveryImageMap = fetchRecoveryImageMap_(recoveryAccumulateUrl);
    } catch (recoveryError) {
      recoveryImageNote = "회수현황 사진 불러오기 실패: " + recoveryError.message;
    }
  } else {
    recoveryImageNote = "회수데이터 연결 URL이 없어 사진을 붙이지 않았습니다.";
  }

  let imageStats = { insertedCount: 0, failedCount: 0 };

  try {
    imageStats = updateOutsourceCauseSheets_(ss, header, groups, groupOrder, recoveryImageMap);
  } catch (causeSheetError) {
    throw new Error(
      "'" + SUMMARY_SHEET_NAME + "'은(는) 갱신됐지만, 원인별 상세 시트 갱신에 실패했습니다: " +
      causeSheetError.message
    );
  }

  return {
    ok: true,
    resultSheet: SUMMARY_SHEET_NAME,
    rowCount: parsedRows.length,
    groupCount: groupOrder.length,
    imageInsertedCount: imageStats.insertedCount,
    imageFailedCount: imageStats.failedCount,
    recoveryImageNote: recoveryImageNote
  };
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


/**************************************************************
 * 원인별 상세 시트를 채웁니다.
 *
 * 먼저 "시트1"/"종합"/양식 시트("시트2") 이 셋을 뺀 나머지 시트를
 * 전부 지웁니다 — 지난 실행 때 만들어졌던, 이번엔 없어진 원인의
 * 시트가 계속 남아있지 않게 하기 위함입니다. 그 다음 이번 원인
 * 목록으로 새로 만듭니다: 원인 값과 이름이 같은 시트가 있으면(양식
 * 시트 자신이 실제 원인 값과 같은 경우) 그 시트를, 없으면
 * CAUSE_SHEET_TEMPLATE_NAME("시트2") 시트를 복사해서 그 원인 이름으로
 * 만듭니다(양식은 헤더 1행 + 예시 데이터 2행 구조라고 가정).
 *
 * 대상 시트의 헤더(1행)를 그대로 읽어서, CAUSE_SHEET_SEQUENCE_COLUMN_LABEL
 * ("구분") 열은 1부터 순차 번호를 새로 채우고, 나머지 열은 이름이 같은
 * "시트1" 열 값을 그대로 복사합니다. 필요한 행 수가 기존(2행 기준
 * 1건)보다 많으면 부족한 만큼 2행 서식을 복사해서 행을 늘립니다.
 * 열 너비 등은 건드리지 않고 값만 채웁니다.
 **************************************************************/
function updateOutsourceCauseSheets_(ss, sourceHeader, groups, groupOrder, recoveryImageMap) {
  const templateSheet = ss.getSheetByName(CAUSE_SHEET_TEMPLATE_NAME);

  if (!templateSheet) {
    throw new Error("'" + CAUSE_SHEET_TEMPLATE_NAME + "' 시트(원인별 상세 시트 양식)를 찾을 수 없습니다.");
  }

  // 이전 실행에서 만들어졌던 원인별 시트는 전부 지우고, 이번 정리
  // 결과로만 새로 채웁니다("시트1"/"종합"/양식 시트("시트2")는 항상
  // 남겨둠). 그래서 이번에 없어진 원인의 옛날 시트가 계속 남아있지
  // 않습니다.
  const permanentSheetNames = [SOURCE_SHEET_NAME, SUMMARY_SHEET_NAME, CAUSE_SHEET_TEMPLATE_NAME];
  ss.getSheets().forEach(function(sheet) {
    if (permanentSheetNames.indexOf(sheet.getName()) === -1) {
      ss.deleteSheet(sheet);
    }
  });

  const dataStartRow = 2;
  const imageMap = recoveryImageMap || {};
  const accessionSourceIdx = sourceHeader.indexOf(CAUSE_SHEET_ACCESSION_SOURCE_LABEL);
  const productCodeSourceIdx = sourceHeader.indexOf(CAUSE_SHEET_PRODUCT_CODE_SOURCE_LABEL);
  let totalImageInsertedCount = 0;
  let totalImageFailedCount = 0;

  groupOrder.forEach(function(key) {
    const group = groups[key];
    const causeLabel = normalizeText_(group.causeValue);

    if (!causeLabel) return; // 원인 값이 없는 행은 상세 시트를 만들 대상이 없음

    let targetSheet = ss.getSheetByName(causeLabel);
    if (!targetSheet) {
      targetSheet = templateSheet.copyTo(ss);
      targetSheet.setName(causeLabel);
    }

    const lastColumn = targetSheet.getLastColumn();
    const header = targetSheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    const sequenceColIndex = header.indexOf(CAUSE_SHEET_SEQUENCE_COLUMN_LABEL);

    const sourceColByTargetCol = header.map(function(label, idx) {
      if (idx === sequenceColIndex) return -1;
      if (label === CAUSE_SHEET_AMOUNT_COLUMN_LABEL) return sourceHeader.indexOf(AMOUNT_SOURCE_LABEL);
      return sourceHeader.indexOf(label);
    });

    const rowsForCause = group.rows;
    const neededRowCount = rowsForCause.length;
    const existingLastRow = targetSheet.getLastRow();
    const existingDataRowCount = Math.max(existingLastRow - dataStartRow + 1, 0);

    // 지우기/삽입 전에 서식 기준(2행)을 미리 잡아둡니다.
    const dataRowFormatSource = existingDataRowCount > 0
      ? targetSheet.getRange(dataStartRow, 1, 1, lastColumn)
      : null;

    if (neededRowCount > existingDataRowCount) {
      const extraRowCount = neededRowCount - existingDataRowCount;
      const insertAfterRow = existingDataRowCount > 0 ? existingLastRow : dataStartRow - 1;
      targetSheet.insertRowsAfter(insertAfterRow, extraRowCount);

      if (dataRowFormatSource) {
        dataRowFormatSource.copyTo(
          targetSheet.getRange(dataStartRow + existingDataRowCount, 1, extraRowCount, lastColumn),
          { formatOnly: true }
        );
      }
    }

    const wipeRowCount = Math.max(neededRowCount, existingDataRowCount);
    if (wipeRowCount > 0) {
      targetSheet.getRange(dataStartRow, 1, wipeRowCount, lastColumn).clearContent();
    }

    const outputRows = rowsForCause.map(function(item, i) {
      return header.map(function(label, colIndex) {
        if (colIndex === sequenceColIndex) return i + 1;
        const srcIdx = sourceColByTargetCol[colIndex];
        return srcIdx === -1 ? "" : item.raw[srcIdx];
      });
    });

    if (outputRows.length) {
      targetSheet.getRange(dataStartRow, 1, outputRows.length, lastColumn).setValues(outputRows);

      CAUSE_SHEET_DATE_COLUMN_LABELS.forEach(function(label) {
        const colIndex = header.indexOf(label);
        if (colIndex === -1) return;

        targetSheet.getRange(dataStartRow, colIndex + 1, outputRows.length, 1)
          .setNumberFormat(CAUSE_SHEET_DATE_NUMBER_FORMAT);

        if (targetSheet.getColumnWidth(colIndex + 1) < CAUSE_SHEET_DATE_MIN_COLUMN_WIDTH) {
          targetSheet.setColumnWidth(colIndex + 1, CAUSE_SHEET_DATE_MIN_COLUMN_WIDTH);
        }
      });
    }

    // 이번 건수가 지난번(이 시트가 "시트2"처럼 재사용된 경우)보다 적으면
    // 남는 아래쪽 행은 값만 지워선 서식(테두리 등)이 남으므로 아예 삭제합니다.
    if (existingDataRowCount > neededRowCount) {
      const removeRowCount = existingDataRowCount - neededRowCount;
      targetSheet.deleteRows(dataStartRow + neededRowCount, removeRowCount);
    }

    clearStrayBordersBelow_(targetSheet, dataStartRow + neededRowCount, lastColumn);

    // 이 원인 그룹의 행 중 "회수현황"에 사진이 있는 건만 골라, 순번
    // (출력 행 기준 1부터)과 함께 모아서 사진 삽입 함수에 넘깁니다.
    if (accessionSourceIdx !== -1 && productCodeSourceIdx !== -1 && neededRowCount > 0) {
      const matchedImages = [];

      rowsForCause.forEach(function(item, i) {
        const accession = normalizeText_(item.raw[accessionSourceIdx]);
        const productCode = normalizeText_(item.raw[productCodeSourceIdx]);
        if (!accession || !productCode) return;

        const imageUrl = imageMap[accession + "||" + productCode];
        if (imageUrl) {
          matchedImages.push({ seq: i + 1, url: imageUrl });
        }
      });

      if (matchedImages.length) {
        const lastDataRow = dataStartRow + neededRowCount - 1;
        const imageResult = insertCauseSheetImages_(targetSheet, lastDataRow, matchedImages);
        totalImageInsertedCount += imageResult.insertedCount;
        totalImageFailedCount += imageResult.failedCount;
      }
    }
  });

  return { insertedCount: totalImageInsertedCount, failedCount: totalImageFailedCount };
}


/**************************************************************
 * fromRow부터 아래로 일정 범위(최대 CLEAR_STRAY_BORDER_ROW_COUNT행)의
 * 테두리를 전부 지웁니다. getLastRow()는 내용이 있는 행까지만 세서,
 * 맨 처음 수동으로 만든 양식에 내용 없이 테두리만 그어둔 행은 그
 * 기준으로는 안 잡히는데, 그런 빈 칸이 표 바로 아래 그대로 남아있는
 * 것처럼 보이는 문제를 막기 위한 안전장치입니다.
 **************************************************************/
const CLEAR_STRAY_BORDER_ROW_COUNT = 50;

function clearStrayBordersBelow_(sheet, fromRow, columnCount) {
  const maxRows = sheet.getMaxRows();
  const rowCount = Math.min(CLEAR_STRAY_BORDER_ROW_COUNT, maxRows - fromRow + 1);

  if (rowCount <= 0) return;

  // fromRow의 위쪽 테두리는 바로 위 실제 마지막 데이터 행(합계 등)의
  // 아래쪽 테두리와 같은 경계선을 공유합니다. 여기서 top을 false로
  // 지우면 그 경계선 자체가 없어져서 합계 행 테두리까지 같이
  // 사라지므로, top만 null(그대로 둠)로 남겨두고 나머지만 지웁니다.
  sheet.getRange(fromRow, 1, 1, columnCount).setBorder(null, false, false, false, false, false);

  if (rowCount > 1) {
    sheet.getRange(fromRow + 1, 1, rowCount - 1, columnCount)
      .setBorder(false, false, false, false, false, false);
  }
}


/**************************************************************
 * "회수데이터" 연결 웹앱의 action="recoveryStatus"를 호출해, "회수누적"
 * 탭의 문서번호+품명 → 이미지 링크 맵을 만듭니다. 문서번호/품명/
 * 이미지 중 하나라도 비어 있는 행은 맵에 넣지 않습니다. 같은 키가
 * 여러 번 나오면 먼저 찾은 값을 그대로 씁니다.
 **************************************************************/
function fetchRecoveryImageMap_(recoveryAccumulateUrl) {
  const requestUrl = recoveryAccumulateUrl +
    (recoveryAccumulateUrl.indexOf("?") === -1 ? "?" : "&") + "action=recoveryStatus";

  const response = UrlFetchApp.fetch(requestUrl, { muteHttpExceptions: true });
  const data = JSON.parse(response.getContentText());

  if (data.error) {
    throw new Error(data.error);
  }

  const header = data.header || [];
  const docNoIdx = header.indexOf(RECOVERY_DOC_NO_COLUMN_LABEL);
  const productIdx = header.indexOf(RECOVERY_PRODUCT_COLUMN_LABEL);
  const imageIdx = header.length - 1; // recoveryStatusAction_의 마지막 열이 항상 "이미지"

  if (docNoIdx === -1 || productIdx === -1) {
    throw new Error(
      "'회수누적' 헤더에서 '" + RECOVERY_DOC_NO_COLUMN_LABEL + "' 또는 '" +
      RECOVERY_PRODUCT_COLUMN_LABEL + "' 열을 찾을 수 없습니다."
    );
  }

  const map = {};

  (data.rows || []).forEach(function(row) {
    const values = row.values || [];
    const imageUrl = normalizeText_(values[imageIdx]);
    if (!imageUrl) return;

    const docNo = normalizeText_(values[docNoIdx]);
    const product = normalizeText_(values[productIdx]);
    if (!docNo || !product) return;

    const key = docNo + "||" + product;
    if (!map[key]) map[key] = imageUrl;
  });

  return map;
}


/**************************************************************
 * 회수현황 사진 링크(드라이브 "보기" URL, 예: .../file/d/ID/view)에서
 * 파일 ID만 뽑습니다 — daily-recovery-accumulate-webapp.gs의 같은
 * 이름 함수와 패턴이 같습니다(서로 다른 스크립트 프로젝트라 공유는
 * 안 되고, 이 파일에도 그대로 복사해뒀습니다).
 **************************************************************/
function driveFileIdFromRecoveryLinkUrl_(url) {
  const text = String(url || "");
  const match = text.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) ||
    text.match(/[?&]id=([a-zA-Z0-9_-]+)/) ||
    text.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : "";
}


/**************************************************************
 * startRow부터 아래로 실제 행 높이(sheet.getRowHeight)를 누적해서,
 * pixelHeight를 담는 데 몇 개 행이 필요한지 셉니다. 템플릿마다 행
 * 높이가 다를 수 있어 고정값(기본 21px) 대신 실제 높이를 씁니다.
 **************************************************************/
function rowsNeededForPixelHeight_(sheet, startRow, pixelHeight) {
  let consumed = 0;
  let rows = 0;

  while (consumed < pixelHeight) {
    consumed += sheet.getRowHeight(startRow + rows);
    rows++;
  }

  return rows;
}


/**************************************************************
 * sheet의 fromCol~toCol(둘 다 포함) 열 너비(px) 합.
 **************************************************************/
function totalColumnWidth_(sheet, fromCol, toCol) {
  let total = 0;
  for (let col = fromCol; col <= toCol; col++) {
    total += sheet.getColumnWidth(col);
  }
  return total;
}


/**************************************************************
 * 1열(A) 왼쪽 끝에서 pixelX만큼 떨어진 위치가 몇 번째 열의, 그 열
 * 기준으로 얼마만큼(offsetX) 떨어진 자리인지 계산합니다. insertImage는
 * (열, 행, 열 안에서의 픽셀 오프셋)으로 위치를 지정해야 해서, "A열
 * 왼쪽 끝에서부터 누적 픽셀 위치" 기준으로 계산한 자리를 다시 열+
 * 오프셋 쌍으로 바꿔주는 역할입니다.
 **************************************************************/
function pixelXToColumnOffset_(sheet, pixelX) {
  let remaining = pixelX;
  let col = 1;

  while (true) {
    const width = sheet.getColumnWidth(col);
    if (remaining < width || col >= sheet.getMaxColumns()) {
      return { column: col, offsetX: remaining };
    }
    remaining -= width;
    col++;
  }
}


// 0~9 숫자 뱃지 PNG(흰 바탕 둥근 네모 + 굵은 검정 숫자, 40x40) 10장을
// base64로 미리 인코딩해서 내장해뒀습니다. Apps Script에는 이미지를
// 직접 그리는 기능이 없어서, 이 프로젝트 밖에서 미리 만든 PNG를 그대로
// 가져왔습니다(RECOVERY_BADGE_SIZE_PX 크기로 줄여서 사진 위에 올림).
const RECOVERY_BADGE_BASE64_BY_DIGIT = [
  "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"
];


/**************************************************************
 * 숫자 하나(0~9)의 뱃지 이미지 Blob을 내장된 base64 PNG에서 만듭니다.
 **************************************************************/
function recoveryBadgeBlobForDigit_(digit) {
  const base64 = RECOVERY_BADGE_BASE64_BY_DIGIT[digit];
  return Utilities.newBlob(Utilities.base64Decode(base64), "image/png", "badge_" + digit + ".png");
}


/**************************************************************
 * 사진 하나의 오른쪽 아래 모서리에 seq(순번) 숫자 뱃지 이미지를
 * 겹쳐 올립니다. 두 자리 이상이면 숫자 이미지를 옆으로 이어
 * 붙입니다. anchorCol/anchorRow/photoOffsetX/photoOffsetY는 방금
 * 넣은 사진과 똑같은 칸 기준이라, 그 사진의 실제 가로/세로(px)만
 * 더하면 사진 안에서의 좌표를 그대로 계산할 수 있습니다.
 **************************************************************/
function overlaySeqBadge_(targetSheet, seq, anchorCol, anchorRow, photoOffsetX, photoOffsetY, photoWidth, photoHeight) {
  const digits = String(seq).split("").map(Number);
  const groupWidth = digits.length * RECOVERY_BADGE_SIZE_PX;

  const groupRight = photoOffsetX + photoWidth - RECOVERY_BADGE_MARGIN_PX;
  const groupLeft = Math.max(photoOffsetX, groupRight - groupWidth);
  const badgeTop = Math.max(photoOffsetY, photoOffsetY + photoHeight - RECOVERY_BADGE_MARGIN_PX - RECOVERY_BADGE_SIZE_PX);

  digits.forEach(function(digit, i) {
    const badge = targetSheet.insertImage(
      recoveryBadgeBlobForDigit_(digit), anchorCol, anchorRow, groupLeft + i * RECOVERY_BADGE_SIZE_PX, badgeTop
    );
    badge.setWidth(RECOVERY_BADGE_SIZE_PX).setHeight(RECOVERY_BADGE_SIZE_PX);
  });
}


/**************************************************************
 * 원인별 상세 시트의 lastDataRow(마지막 데이터 행)에서 두 줄 아래부터
 * 사진을 가로 6장씩 붙입니다.
 *
 * - 가로 폭: 1~CAUSE_SHEET_IMAGE_AREA_LAST_COLUMN(X)열의 실제 너비
 *   합을 6등분(간격 CAUSE_SHEET_IMAGE_GAP_PX 포함)해서 칸 폭을 정하고,
 *   사진은 그 칸 폭에 맞춰 원본 가로세로 비율 그대로 줄입니다(비율
 *   고정) — 그래서 6번째 사진도 X열 끝을 넘지 않고, 시트의 실제 열
 *   너비와 무관하게 6장이 항상 고르게 간격을 둡니다.
 * - 칸 위치는 사진이 있는 것끼리 압축해서 채우지 않고, 각 건의 실제
 *   순번(seq) 기준으로 "(seq-1)/6이 몇 번째 줄, (seq-1)%6이 몇 번째
 *   칸"을 그대로 씁니다 — 예를 들어 9번 행에만 사진이 있으면 1번
 *   칸이 아니라 두 번째 줄의 세 번째 칸에 옵니다. 사진이 전혀 없는
 *   줄은 아예 건너뛰어 공간을 낭비하지 않습니다.
 * - 번호는 셀 텍스트가 아니라, 사진의 오른쪽 아래 모서리에 숫자
 *   뱃지 이미지를 겹쳐 올리는 방식으로 표시합니다(overlaySeqBadge_).
 **************************************************************/
function insertCauseSheetImages_(targetSheet, lastDataRow, matchedImages) {
  // seq 기준으로 블록(줄)별로 묶습니다. 사진이 있는 블록만 처리하고,
  // 빈 블록은 건너뛰어 위에서부터 압축해서 보여줍니다.
  const blocksByIndex = {};
  const blockIndexes = [];

  matchedImages.forEach(function(item) {
    const blockIndex = Math.floor((item.seq - 1) / CAUSE_SHEET_IMAGES_PER_ROW);
    const slotIndex = (item.seq - 1) % CAUSE_SHEET_IMAGES_PER_ROW;

    if (!blocksByIndex[blockIndex]) {
      blocksByIndex[blockIndex] = [];
      blockIndexes.push(blockIndex);
    }

    blocksByIndex[blockIndex].push({ seq: item.seq, url: item.url, slotIndex: slotIndex });
  });

  blockIndexes.sort(function(a, b) { return a - b; });

  // 1~X열 너비 합을 6등분해서 칸(슬롯) 폭을 구합니다. 전부 A열(1)
  // 기준 누적 픽셀 위치로 계산하므로, 실제 열 경계와 무관하게 6장이
  // 항상 똑같은 간격으로 놓입니다.
  const availableWidth = totalColumnWidth_(targetSheet, 1, CAUSE_SHEET_IMAGE_AREA_LAST_COLUMN);
  const slotWidth = (availableWidth - CAUSE_SHEET_IMAGE_GAP_PX * (CAUSE_SHEET_IMAGES_PER_ROW - 1)) / CAUSE_SHEET_IMAGES_PER_ROW;

  let blockStartRow = lastDataRow + CAUSE_SHEET_IMAGE_START_GAP_ROWS;
  let insertedCount = 0;
  let failedCount = 0;

  blockIndexes.forEach(function(blockIndex) {
    const block = blocksByIndex[blockIndex];
    let maxHeight = 0;

    block.forEach(function(item) {
      try {
        const fileId = driveFileIdFromRecoveryLinkUrl_(item.url);
        if (!fileId) throw new Error("드라이브 링크에서 파일 ID를 찾을 수 없습니다.");

        const pixelX = item.slotIndex * (slotWidth + CAUSE_SHEET_IMAGE_GAP_PX);
        const anchor = pixelXToColumnOffset_(targetSheet, pixelX);

        const blob = DriveApp.getFileById(fileId).getBlob();
        const image = targetSheet.insertImage(blob, anchor.column, blockStartRow, anchor.offsetX, 4);

        const nativeWidth = image.getWidth() || slotWidth;
        const nativeHeight = image.getHeight() || slotWidth;
        const targetHeight = Math.round(nativeHeight * (slotWidth / nativeWidth));
        image.setWidth(slotWidth).setHeight(targetHeight);

        overlaySeqBadge_(targetSheet, item.seq, anchor.column, blockStartRow, anchor.offsetX, 4, slotWidth, targetHeight);

        if (targetHeight > maxHeight) maxHeight = targetHeight;
        insertedCount++;
      } catch (error) {
        failedCount++;
      }
    });

    if (maxHeight === 0) maxHeight = slotWidth; // 전부 실패했을 때 대비
    blockStartRow = blockStartRow + rowsNeededForPixelHeight_(targetSheet, blockStartRow, maxHeight) + CAUSE_SHEET_IMAGE_BLOCK_GAP_ROWS;
  });

  return { insertedCount: insertedCount, failedCount: failedCount };
}


/**************************************************************
 * "종합" 시트 위쪽(제목/안내 영역, 헤더 행 이전)에서 "OO년 O월" 형태의
 * 제목 문구와 "YYYY.M.D~YYYY.M.D" 형태의 기간 문구를 찾아, 오늘 기준
 * 전월/전전월 값으로 바꿔 씁니다.
 *
 * 마감은 항상 "전월" 자료 기준이라(예: 9월에 실행하면 8월 마감),
 * 제목은 전월 연/월, 기간은 "전전월 26일~전월 25일"로 계산합니다.
 * (예: 오늘 2026-09-03 실행 → 제목 "26년 8월", 기간 "2026.07.26~2026.08.25")
 **************************************************************/
function updateSummaryDateLabels_(summarySheet, headerRowNumber) {
  const scanRowCount = headerRowNumber - 1;
  if (scanRowCount <= 0) return;

  const range = summarySheet.getRange(1, 1, scanRowCount, 1); // 제목/기간 문구는 A열에 있음
  const values = range.getValues();

  const now = new Date();
  const prevMonthDate = new Date(now.getFullYear(), now.getMonth() - 1, 1); // 전월
  const prevPrevMonthDate = new Date(now.getFullYear(), now.getMonth() - 2, 1); // 전전월

  const titleReplacement = String(prevMonthDate.getFullYear()).slice(-2) + "년 " + (prevMonthDate.getMonth() + 1) + "월";
  const periodReplacement =
    prevPrevMonthDate.getFullYear() + "." + pad2_(prevPrevMonthDate.getMonth() + 1) + ".26~" +
    prevMonthDate.getFullYear() + "." + pad2_(prevMonthDate.getMonth() + 1) + ".25";

  const titlePattern = /\d+\s*년\s*\d+\s*월/;
  const periodPattern = /\d{4}\.\d{1,2}\.\d{1,2}\s*~\s*\d{4}\.\d{1,2}\.\d{1,2}/;

  let changed = false;

  const newValues = values.map(function(row) {
    const text = String(row[0] === null || row[0] === undefined ? "" : row[0]);

    if (titlePattern.test(text)) {
      changed = true;
      return [text.replace(titlePattern, titleReplacement)];
    }

    if (periodPattern.test(text)) {
      changed = true;
      return [text.replace(periodPattern, periodReplacement)];
    }

    return [text];
  });

  if (changed) {
    range.setValues(newValues);
  }
}


function pad2_(number) {
  return number < 10 ? "0" + number : String(number);
}


/**************************************************************
 * fromRow~toRow 범위에서 A열이 병합의 일부인 첫 번째 행 번호를
 * 찾습니다(구분행/합계행은 항상 A열부터 병합돼 있음). 없으면 -1.
 **************************************************************/
function findFirstMergedRowNumber_(sheet, fromRow, toRow) {
  for (let row = fromRow; row <= toRow; row++) {
    if (sheet.getRange(row, 1).getMergedRanges().length > 0) return row;
  }
  return -1;
}


/**************************************************************
 * 이 스프레드시트의 시트 중 excludeNames에 없는 시트만 전부 골라
 * xlsx로 내보냅니다("정리파일다운로드" 버튼에서 사용).
 **************************************************************/
function exportAllExceptAsXlsxBase64_(excludeNames, fileNamePrefix) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = ss.getSheets()
    .map(function(sheet) { return sheet.getName(); })
    .filter(function(name) { return excludeNames.indexOf(name) === -1; });

  if (!sheetNames.length) {
    throw new Error("내보낼 시트가 없습니다.");
  }

  return exportSheetsSubsetAsXlsxBase64_(sheetNames, fileNamePrefix);
}


/**************************************************************
 * 이 스프레드시트에서 sheetNames에 해당하는 시트만 임시 스프레드시트에
 * 복사해서 xlsx로 내보낸 뒤, 임시 스프레드시트는 지웁니다.
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
 * "외작월마감 자료교체"/"구매월마감 자료교체" 버튼 액션 — "월마감"
 * 스프레드시트에서 가져온 "외작(N)" 또는 "구매(N)" 탭의 header/rows를
 * 그대로 "시트1"에 덮어씁니다(기존 내용은 전부 지움). 외작/구매 둘 다
 * 양식이 같아서 이 웹앱을 같이 씁니다.
 **************************************************************/
function replaceSourceDataAction_(header, rows) {
  if (!header.length) {
    throw new Error("가져올 데이터의 헤더가 비어 있습니다.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!sheet) throw new Error("'" + SOURCE_SHEET_NAME + "' 시트를 찾을 수 없습니다.");

  const columnCount = header.length;
  const dataRows = rows.map(function(row) { return normalizeRowLength_(row, columnCount); });

  // summaryRows처럼 JSON으로 건너온 값이라, 월현황(주간)/월마감에서는
  // 진짜 Date였던 값이 "2026-08-27T07:00:00.000Z" 같은 ISO 문자열로
  // 바뀌어 있습니다. 문자열인 채로 쓰면 숫자 서식을 걸어도 텍스트라서
  // 전혀 안 먹히므로, 날짜 열은 값을 쓰기 전에 시간대 영향 없는 순수
  // "yyyy-mm-dd" 문자열로 바꿔둡니다.
  CAUSE_SHEET_DATE_COLUMN_LABELS.forEach(function(label) {
    const colIndex = header.indexOf(label);
    if (colIndex === -1) return;

    dataRows.forEach(function(row) {
      row[colIndex] = parseDateOnlyValue_(row[colIndex]);
    });
  });

  sheet.clear();

  // "061"처럼 앞자리 0이 있는 색상 값은, 서식이 기본값(General)인 채로
  // 쓰면 그 즉시 숫자로 재해석되어 0이 사라집니다. 값을 쓰기 전에 이
  // 열만 먼저 "@"(텍스트) 서식으로 지정해야 합니다.
  const colorColIndex = header.indexOf("색상");
  if (colorColIndex !== -1 && dataRows.length) {
    sheet.getRange(2, colorColIndex + 1, dataRows.length, 1).setNumberFormat("@");
  }

  sheet.getRange(1, 1, dataRows.length + 1, columnCount).setValues([header].concat(dataRows));

  CAUSE_SHEET_DATE_COLUMN_LABELS.forEach(function(label) {
    const colIndex = header.indexOf(label);
    if (colIndex !== -1 && dataRows.length) {
      sheet.getRange(2, colIndex + 1, dataRows.length, 1).setNumberFormat(CAUSE_SHEET_DATE_NUMBER_FORMAT);
    }
  });

  const realRowCount = dataRows.filter(function(row) {
    return row.some(function(cell) { return normalizeText_(cell) !== ""; });
  }).length;

  return { ok: true, rowCount: realRowCount };
}


/**************************************************************
 * value(Date 객체 또는 JSON 직렬화된 ISO 날짜 문자열)에서 "yyyy-MM-dd"
 * 문자열만 뽑아 돌려줍니다. 시각(Z)이 있는 문자열은 실제 Date로
 * 파싱해서 한국 시간 기준 날짜로 바꾸고, 시각이 없는 "yyyy-mm-dd"는
 * 그대로 날짜 부분만 뽑습니다(new Date()로 그냥 파싱하면 UTC 자정으로
 * 해석되어 한국 시간으로는 전날이 되어버리는 문제를 피하기 위함).
 * 날짜로 못 읽으면 원래 값을 그대로 돌려줍니다.
 **************************************************************/
function parseDateOnlyValue_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, "Asia/Seoul", "yyyy-MM-dd");
  }

  const text = String(value === null || value === undefined ? "" : value).trim();
  if (!text) return value;

  if (/T\d{2}:\d{2}/.test(text)) {
    const parsed = new Date(text);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, "Asia/Seoul", "yyyy-MM-dd");
    }
  }

  const match = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) return value;

  return match[1] + "-" + match[2] + "-" + match[3];
}


function normalizeRowLength_(row, targetColumnCount) {
  const result = row.slice(0, targetColumnCount);

  while (result.length < targetColumnCount) {
    result.push("");
  }

  return result;
}


function normalizeText_(value) {
  return String(value === null || value === undefined ? "" : value).trim();
}


function jsonOutput_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
