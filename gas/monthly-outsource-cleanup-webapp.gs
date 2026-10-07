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
// 합을 6등분해서 "칸 폭"(슬롯)을 정합니다 — 시트마다 열 너비가 달라도
// 6장이 항상 고르게 간격을 두고 X열 끝까지만 차지합니다.
const CAUSE_SHEET_IMAGE_AREA_LAST_COLUMN = 24;
const CAUSE_SHEET_IMAGE_GAP_PX = 10; // 사진 사이 가로 간격
// 사진 표시 세로 높이(px) — "5.5cm(원본크기)" 요청을 화면 96dpi 기준으로
// 환산(5.5 / 2.54 * 96 ≈ 208px). 가로폭은 원본 가로세로 비율대로
// 계산하되, 슬롯 폭(칸 폭)을 넘으면 슬롯 폭에 맞춰 다시 줄입니다
// (세로를 우선하되 가로가 옆 칸을 침범하지는 않게).
const CAUSE_SHEET_IMAGE_TARGET_HEIGHT = 208;
const CAUSE_SHEET_IMAGE_START_GAP_ROWS = 2; // "마지막 행에서 두 줄 아래"
const CAUSE_SHEET_IMAGE_BLOCK_GAP_ROWS = 1; // 사진 줄 사이 빈 줄
// 번호를 셀 텍스트가 아니라 사진 위에 겹친 "뱃지" 이미지로 표시합니다
// (흰 바탕 둥근 네모/알약 모양 + 굵은 검정 숫자를 PNG로 미리 만들어
// base64로 이 파일에 내장해뒀습니다 — Apps Script에는 사진 픽셀 위에
// 직접 글자를 그리는 기능이 없어서, 숫자 자체를 이미지로 준비해 사진
// 위에 또 하나의 이미지로 겹쳐 올리는 방식입니다). 한 자리(0~9)는
// 정사각 뱃지, 두 자리(10~99)는 숫자 두 개가 한 이미지 안에 같이
// 들어있는 가로로 넓은 뱃지를 따로 준비해서 칸을 넘치거나 다른 사진에
// 가려지는 일이 없게 했습니다.
const RECOVERY_BADGE_SIZE_PX = 40; // 뱃지 세로 크기(= 한 자리 뱃지의 가로 크기)
const RECOVERY_BADGE_WIDE_WIDTH_PX = 52; // 두 자리 뱃지 가로 크기(세로는 동일)
// 사진 아래쪽 모서리에서 뱃지까지 여백(약간 띄움). 오른쪽은 여백 없이
// 사진 자신의 실제 오른쪽 끝에 딱 맞춥니다(overlaySeqBadge_ 참고).
const RECOVERY_BADGE_MARGIN_Y_PX = 2;

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
 * startCol 왼쪽 끝에서 pixelX만큼 떨어진 위치가 몇 번째 열의, 그 열
 * 기준으로 얼마만큼(offsetX) 떨어진 자리인지 계산합니다. insertImage는
 * (열, 행, 열 안에서의 픽셀 오프셋)으로 위치를 지정해야 해서, "어떤
 * 열 왼쪽 끝에서부터 누적 픽셀 위치" 기준으로 계산한 자리를 다시 열+
 * 오프셋 쌍으로 바꿔주는 역할입니다.
 **************************************************************/
function pixelXToColumnOffset_(sheet, startCol, pixelX) {
  let remaining = pixelX;
  let col = startCol;

  while (true) {
    const width = sheet.getColumnWidth(col);
    if (remaining < width || col >= sheet.getMaxColumns()) {
      return { column: col, offsetX: remaining };
    }
    remaining -= width;
    col++;
  }
}


/**************************************************************
 * pixelXToColumnOffset_와 같은 역할을 세로(행) 방향으로 합니다.
 * startRow 위쪽 끝에서 pixelY만큼 떨어진 위치가 몇 번째 행의, 그 행
 * 기준으로 얼마만큼(offsetY) 떨어진 자리인지 계산합니다 — 번호 뱃지를
 * 사진 "아래쪽"(한 행 높이보다 훨씬 큰 오프셋)에 놓을 때, offsetY를
 * 그대로 큰 값으로 넘기면 한 행 안으로만 해석돼 엉뚱한 자리(사진
 * 위쪽 근처)에 놓이므로, 실제 몇 번째 행인지 미리 계산해서 넘겨야
 * 합니다.
 **************************************************************/
function pixelYToRowOffset_(sheet, startRow, pixelY) {
  let remaining = pixelY;
  let row = startRow;

  while (true) {
    const height = sheet.getRowHeight(row);
    if (remaining < height || row >= sheet.getMaxRows()) {
      return { row: row, offsetY: remaining };
    }
    remaining -= height;
    row++;
  }
}


// 0~9 숫자 뱃지(정사각) + 10~99 숫자 뱃지(가로로 넓음) PNG를 base64로
// 미리 인코딩해서 내장해뒀습니다. Apps Script에는 이미지를 직접 그리는
// 기능이 없어서, 이 프로젝트 밖에서 미리 만든 PNG를 그대로 가져왔습니다.
// 두 자리 숫자는 "1"+"0" 이미지를 옆으로 이어 붙이지 않고, "10"이 통째로
// 들어있는 뱃지 하나를 따로 준비해서 옆 칸 사진에 가려지거나 숫자가
// 겹쳐 보이는 문제를 없앴습니다(100 이상은 드물어서 생략 — 뱃지 없이
// 사진만 들어갑니다).
const RECOVERY_BADGE_BASE64_BY_DIGIT = [
  "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"
];

const RECOVERY_BADGE_BASE64_BY_NUMBER = {
  10: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAS8SURBVGhD5ZlZSHVVFMf/OeFEDmApIWGSPghl5ABqmlo5gA+iZE8aKPjg0IOCij6YL+YACmFo+KaSODw4z+P34ESkkA+KA4ZoDpSGU6auWBfP4Zz93Xu8mnqv9IP94F5rn7P/Z9+91tpb4H+MG4APAHwC4DMA0c/Q+D38vg8BvCVO6CE4AMgCMAbgEACZsP0BYBJALoA3xYkaw1cAVvlhgYGBVFhYSM3NzdTT00MjIyM0PDz85I3f09vbSy0tLVRcXEwhISGSuHUAX4sT1qKcB4aGhtLAwACZE+Pj4xQZGSkJ+x7AG+LkRb5j55ycHLq+vhafZzYUFRVJon4QBSj5UhLzEigpKZFEpYtCGEcAvwUFBYnjzJqoqCgWtAfARRSUzWonJibEMWbNwsICWVpasqgCUdCrgIAA0f9erK6uUnt7O3V2dtLBwYFofo3d3V2anp6msbExWl9fF81GcxskflEGCE5af/Fv8r/g7+8v5w2ORoY4Pj6mzMxMcnZ2lv1tbW0pMTGRtra2RPc7qays5GdcAnhXEvQRP7S1tVX0NZqCggJ5ctympqZEFx0cOaOjo1W+yubt7W3U6irp7++XxnNVoeNT7hgcHBR97+Ts7EwXFcWJGRLU2Nio8uOV4hDs6Ogo92VlZYnDNJmcnJTGfiEJ4ppJl5mN5fz8nDo6OlQ/s7sE3dzcEO9TyScjI0O2NTQ0yP1OTk50eHioGqvF6OioNPZzSRAXgrpyw1h44ysFeHh4kJ2dnaagzc1NsrKykn2UEXV/f1+1St3d3aqxWjyKoKamJvnl6enptLi4qJqQPkFDQ0Oy3dramjY2NmQbr56vr69sr6ioUI3V4lEEtbW1UUREhG5DMjs7O6qvr0+Qcv+4uLjoVkVJcHCwbM/OzlbZtHgUQRcXF6q/OY/wV9cSVFNTI9vd3d3p6OhIZecPJNlTU1NVNi0eRZCIMYLKy8s1BYWHh8v2lJQUlU0Lkwm6TYAGBSlXKC0tTWXTwmSC6uvrZTvvob29PZVduYfuk4tMJqirq0u2cwARo5yPj49sr6qqUo3VwmSCVlZWyMLCQq8Pr5Yy7N/npGwyQVdXV+Tn5yf7cNkjUVdXJ/e7urq+tr+0MJkgpra2VvbhxuVPfn4+OTg4yH25ubniME2eRNDa2ppqooYOipeXlxQXF6fyVTbeR/ep45gnEcTnGC8vL/L09NS12dlZ0UWGk3JeXh65ubnJQuzt7Sk5OZm2t7dF9zvRJ+je1bYIR6nT01M6OTnRNWNujHglZmZmdD/PhxzsJPQJiuSOh5yHzAF956GPuYMLzpcIL8StoAhJkDuAk7KyMtH3RXAbOf8B8J4kiJkNCwsTfV8EsbGxLOhXAJZKQXmcxefm5kR/s2Z5eZlsbGxY0LdKMYwzgN/5juslkZCQwGL+NPT/I74jptLSUnGcWcIF7G0w+EYUouRHdjL3AKEQ0yoKEOGN1cjOvJzz8/Pis0zK0tISJSUlSWJ+AmArCjAEX97vcqCIiYmh6upq6uvr02V1zszP1fjum48SfB8RHx8vXcTscxATJ2wMbwMoAfAzgL+lustEjfPMEoAyAO+IE30InLQ4E3N5wTXTczV+H5dm74sT0se/STF7QHu0qRsAAAAASUVORK5CYII=",
  11: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAPUSURBVGhD7ZlNSFVbFMf/qeiDJ1qDUlARgxQcFJGKqM/UPhTBgQjZSINmljVokKIDnwN9qKCQhGSCAyVR8QM/UfOrSSoRoiIGgTRQk0eBX+DzYz3WwX05rnu9na7kPZfuD/bgnr32PuvnvWevfbbAb8xFAFcB/AXgNoBbZ9D4Pny/awAuyYQc4U8AjwC8BfAvAHJi+wZgHMATAH4yUSPcB/CJJ4uOjqaCggJqamqinp4eGh4epqGhoV/e+D69vb3U3NxMRUVFFBcXp+Q+A3ggE7ZHOQ+Mj4+ngYEBMhOjo6OUnJysxF4AOCeTl/zDwfn5+XRwcCDnMw2FhYVK6qUU0HNPybgCxcXFSuqhFGF8AXyJiYmR40xNSkoKC30FcEEKPWbbsbExOcbUzMzMkKenJ0s9l0LvoqKiZPxPww9ta2urtkoZYXl5mdra2qi9vZ1WV1dltyGOFomP+gWCi9YG/yZPw9zcnKVuBAUFyW6bJCYmWsZ0dnbKbkNUVFTw+P8AhCqh6zxhS0uLjDXM5uYm8TeskgsPD5chVpSXl1viuXGNc4T+/n41B+8qNJL4wuDgoIw1xOLiIsXGxh5LLiIiQoZZ2Nvb0y+7pxYaHx9Xc9xVQrxn0irzz7C2tkZlZWXk5+dnlZwtod3dXS1pKX9aoZGRETXHHSXEG0HDD7IiNzf3WEJhYWF2hXh+fXxAQAD5+vqaRygzM1ObiJNqaGigxsZGu0Ld3d2W/uzsbJqfn6fAwEDzCOXl5VFOTg4tLS1pn+vr6+0K8b4wISGBOjo6tM8bGxvHfq5OF+JnQk9dXZ1dIRm/srJC/v7+5hGS/EhIwoXULWQDt5BR3EJuIdu4hYxSW1trSS40NFR2W8F1yMfHxzKmq6tLhhjilwnxMVdwcLDWkpKSZLcV6+vrFBkZSSEhIdoYTswRbAk5tNuW8GvB1taW1nZ2dmS3FYeHh7S9vW0Zs7+/L0MMYUsomS84+j7kbGy9D93gC3wW4IrwF3EkdFMJBQLYKi0tlbEuQU1NDcvsAbishJj3vLV3RdLS0lhoHoCnXuiZh4cHTU1NyXhTs7CwQN7e3iz0t16GOQ9gjc+4XImMjAyW+X7S/4/4jJhKSkrkOFNSWVmpFoOnUkTPKw4y+wKhk2mRAhJ+sF5zMH+d09PTci6nMjs7S1lZWUrmDYA/pMBJ8OH9Ki8UqampVFVVRX19fTQxMaFV5rNqk5OT2uFKdXU1paenk5eXF4us8yImEzZCAIBiAB8A7B79VZzVuM7MAigFECQTdQQuWlyJeXvBe6azanw/3ppdkQnZ4n+BmQ2eyqaBUwAAAABJRU5ErkJggg==",
  12: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAASESURBVGhD5ZltKH53GMe/HvszadR/m0Zj8pBk1vACy9gDKUXyz0vlHTPJC0Qxb7ZQFHmx9kZo8vDK88M8jTw1Q1MIRbKZ2bQ8beJa18m5u8/1v8/tZv7uW/vU9eb8ruuc38c59+93nQP4H/McQBiADwF8AuDjRwi+Dl/vPQBvyAndh9cA5AL4HsDvAMiK8QeAcQBfAHCXE7WETAAbfLLIyEgqLi6mlpYW6u7upuHhYRoaGnrlwdfp6emh1tZWKi0tpejoaFVuC0CWnLA5vuLCmJgY6u/vJ1tidHSU4uPjVbF6AHZy8pKvOTkvL4+urq7k+WyGkpISVapRChjzQpV5CpSVlalS2VKEcQOwGxUVJetsmoSEBBY6AOAhhT5n27GxMVlj0ywsLJCDgwNLFUmhHyIiImT+ndjY2KCOjg7q6uqiw8NDOfwSOzs7ND4+rqxoi4uLdHFxIVMs4maR+Ml4geBN6y9+Jv8L4eHhhn2DVyM9pqenKTExkVxdXY33GQoKCqL6+nqZfitVVVVc/w+Ad1Sh9/mEbW1tMtdiioqKNJObmJiQKQp8N5ycnDS5MgoLC2WZWfr6+tRa7ioUPuIDAwMDMvdWzs7OlFVRTsqUED9SAQEBhhwPDw9l8uXl5RQSEqKpn5qakuW68GN7U/eZKsQ9k7IzW8r5+Tl1dnZqHrPbhHiTVsft7Ow0C9DR0RH5+PgYxgsKCjS15hgZGVHrPlWFuBFUHgdL4R++sYCXlxe5uLiYFeIWxt/fnzw9PSksLEwOU0ZGhqE+MzNTDuvyIELNzc2Gi2dnZ9PS0hK5ubmZFbq+vqbLy0s6ODig/f19OUyxsbGG+vz8fDmsy4MItbe3U1xcnPKDZHiCjo6OZoXMMTMzo6nnJthSHkRI7hlbW1ua1esuQru7u5rFgh9HvpOW8iBCkvsKsUxwcLChjuMuKxxjM0Kbm5sUGBiokWlsbJRpt2ITQuvr6+Tr66uRaWhokGkWYXWh7e1tjYy9vT01NTXJNIuxqtDx8TGFhoYa8tzd3WlwcFCm3QmrCmVlZRlyOLgj2NvbUx7BtbU1Q5jap/SwmtDKyopGhoM7bm6D5PFH7xQkvGIZT8jUi2JOTs5LE9eL1NRUWa7LKxHilzU/Pz+lweSYnZ2VKZSWlkbe3t6GHL3gnNzcXFmuiymhO3fbEu7TTk9P6eTkRAlTX4zUMc4zF5wjOxFzmBKK5wP3eR+yBUy9D33AB7jhfIrwjbgRilOF3gJwUllZKXOfBHV1dSxzCeBdVYiZ5feRp0hSUhIL/QzAwViokFuQubk5mW/TrK6ukrOzMwt9aSzDvA7gV/7G9ZRISUlhmT/1/n/E34ipoqJC1tkk1dXV6mKQL0WM+YaTbH2BMJJpkwIS/mF9y8l8O+fn5+W5rMry8jKlp6erMt8BeCYF9OCP97/wQsGfbWtqaqi3t1dpOnlnfqyYnJxUvufV1tZScnKy+iHlN17E5IQt4U0AZQB+BPD3zV/FWsH7zDKASgBvy4neB960eCfm9oJ7pscKvh63ZgFyQqb4Fx0ZmL/Kzn0TAAAAAElFTkSuQmCC",
  13: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAStSURBVGhD5ZlZSHVVFMf/ziNaD2kSEgY96EMZqQ9qmmYpiiAoofjgBz6WKfrgDA4PhQMKQQ+RTyo5vzii5ojiiAqJ4KyIZUpBDp8lumIdvnO4d3Xv9V5zuNIP9oN7r3XO+p+7z97/fQT+x7wB4D0AHwGIAfDJIzS+D9/vfQBesqC74AbgCwA/ATgBQE/YfgcwBuArAB6yUHNIBbDOFwsODqaCggJqamqi7u5uGhoaosHBwQdvfJ+enh5qbm6m4uJiCg0NVcVtAXghCzbF15wYFhZG/f39ZE2MjIxQVFSUKuxbADayeMk3HJyVlUXX19fyelZDYWGhKuo7KUCXz1Uxz4GSkhJVVKYUwrgD2A8JCZF5Vk10dDQLOgLwuhT0JasdHR2VOVbN/Pw82dnZsah8KWgyKChIxlvE+vo6tbe3U2dnJx0fH8vhf7GxsaG85PwQd3d35bDZvFoklnQXCN60/uQ5+V8IDAzU9g0u1BgLCwsUGxtLTk5OWryrqyulpKTQ9va2DL+VqqoqvsbfAN5WBX3AF21paZGxZpOfn68Vx218fFyGKCwuLpK7u7terG7z9fWlnZ0dmWaSvr4+NZ9dhcLH3DEwMCBjb+Xi4kJZFWVhhgTxNsB7m27x5eXllJeXRy4uLlp/UlKSTDXJ2NiYmvuZKog9k7Izm8vLly+po6NDb5rdJmh5eVkvZnV1VRurra3V+j08POjk5EQv1xTDw8Nq7qeqIDaCit0wF37xdYvz8fHRe8qGBB0dHSn2iadnaWmp3hgvDGqum5ubEmsu9yKosbFRKyAzM1N5+rrvhiFBpsjIyNByIyIi6ObmRoYY5V4EtbW1UWRkpPJCMoeHh2Rvb2+RIJ62lZWVut6MvLy8aGlpSYaa5F4EXV5e6v29tbVFDg4OFglaW1vT4rl5e3vT9PS0DLuVexEkuYsgdvK8/6hTlX/hgIAA6urqkqEmsRpBvJLt7+/TwcEBFRUVabk2NjY0Nzcnw41iNYIkOgc4Sk9Pl8NGsQpBV1dXsotycnK0fEt85ZMJam1tpYSEBPL396fU1FQ5TGlpaVp+eHi4HDbKkwlqaGjQxtnyT01NaWPsGjw9PbXx3NxcvVxTPJmg09NT8vPz02JYQHZ2tuIc2Gmo/ew4+GhhLg8iaHNzUyuIm7GD4uTkpEm3zUs3n6ks4UEE7e3tKU+fHTS3mZkZGaLBTiAxMVHxbKoQZ2dniomJoYmJCRl+K4YEWey2Jey9zs/P6ezsTGnmfDHih8AC2P7zlL0rhgRFccddzkPWgKHz0IfcwYbzOcI/xCtBkaqgNwGcVVRUyNhnQX19PYu5AvCOKoiZsWQzsybi4uJY0M8A7HQF5dna2tLs7KyMt2p4Q3Z0dGRB5bpimNcA/MqHrecEL/8A/jD2/yP+RkxlZWUyzyqprq5WF4NsKUSX7znI2hcIHTEtUoCEX6wfOJh/TksOWo/BysoKJScnq2J+BOAsBRiDP97/wgsFf7atqamh3t5exXTyzvxYjZ0EH9Xr6uooPj5e/RDzGy9ismBz8AZQAmARwF+q73qixvvMCoAKAG/JQu8Cb1q8E7O9YM/0WI3vx9bsXVmQIf4BMxGxK4mrpVsAAAAASUVORK5CYII=",
  14: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAARASURBVGhD5ZlJSCVXFIb/OEtEDZhBgqgBNy5sQ1REDUYzKIorIWQZcCFiTBaCA4oaXURUUAlkEbIQHCJOG0fUOGXjHIV2oShoFibRYDQ4ptHTnMIq6p1+lvVs2qpHPrgLq86pOp+v3rn31gP+x7wJIAbAhwA+AfDxIwy+D9/vCYC3ZEEP4XUAhQB+AfA3ALJwHAGYAfA1gEBZqBm+ALDFF4uPj6eysjLq6OigwcFBmpiYoPHx8Vc++D5DQ0PU2dlJFRUVlJSUpMrtAPhSFmzEd5yYnJxMo6OjZCempqYoLS1NFfsewGuyeEk9BxcVFdH19bW8nm0oLy9XpX6QAno+V2XcgcrKSlUqT4owAQB+T0hIkHm2Jj09nYX+AvCGFPqKbaenp2WOrVlaWiJPT0+WKpVCv8bFxcl4l9ja2qLe3l7q7++nw8NDedqQhYUFJXdgYICOj4/laUNum8Rv+gbBk9a//Ey+DLGxsdq8wd3ILBsbG+Tr66vlrqysyBBDGhoaOO8/AOGq0Pt8oe7ubhlrmtLSUq0gHrOzszLEKdxJExMTHXLX1tZkmCEjIyNqLq8qFD7iA2NjYzL2Xs7Pz5WuqC/IFaHa2toXcl0VmpmZUXM/U4V4zaTMzGa5uLigvr4+h8fMVaHl5WXy8vJ6IddVocnJSTX3U1WIF4LKcsMs/MXXFxEaGkr+/v6mha6urigmJkaL14tZItTe3q4VkJeXpxQREBBgWqikpESLLSwspPDwcGuFenp6KDU1VflCMvv7+w7/ZSOhubk5LY6vcXR0RCEhIdYKXV5eOvy9s7ND3t7e9wqdnp5SVFSUEuPj40Obm5tKYwkMDLRWSGJWqKCgQIupr69XjvGnGxQU5H5CuvmCsrOztePcIIKDg7Vzu7u7Dnn3YYnQyckJRUREaOd5w8iNpa2tjVpbWx06ZF1dnXLcrJglQlycsznHaHR1dTlc4y4sE5IF3zdsLcQraH6UqqqqHEZ1dTUVFxeTn5+flpufn6+c40WrGSwRMoKnANs1he3tba0gHq5sFLlt67cPq6urMsSQVyK0t7dHkZGRFBYWpoz5+XkZcicHBwcUHR2t5Zp91FScCbm82pbc3NzQ2dmZshLg4cobo5fJZZwJpfGBh+yH7ICz/dAHfIAXnO4IfxC3Qqmq0DsATnn36I60tLSwzDMA76lCzHxKSoqMdQsyMzNZ6CkAT71QsYeHh/I6yZ3gjsjbEADf6mWYYAB/8jsudyInJ4dl/rnr9yN+R0w1NTUyz5Y0NjaqzeAbKaLnRw6ye4PQyXRLAQl/sX7iYP44FxcX5bUsZX19nXJzc1WZnwH4SYG74Jf3f3CjyMjIoKamJhoeHlYWnTwzP9bgFyr8o1tzczNlZWWpe6oDbmKyYDO8DaASwAqAq9v/ilWD55l1ALUA3pWFPgSetHgm5uUFr5kea/D9eGkWJQtyxnOLap9wo5rTPwAAAABJRU5ErkJggg==",
  15: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAASsSURBVGhD5ZlrKP5nGMe/cxzTWNhoaRn/V2qzhhxz2IHIK7UoL1ZeeLHZaArxwhRbCLVaWbyQw+TwyjHnQ+SQRg05lhU2tGX4O8W1rl9+vx7XnoeH4Xm0T12l331dv/v+un/3fV33/QD/Y1wBvAcgDMDHAD56AuN+uL/3AbwpB3QfXgPwBYB+AHsAyIT2J4AhAF8BeF0O1BgSASzzy/z9/Sk7O5vq6uqora2Nent7qaen59GN+2lvb6f6+nrKzc2l4OBgVdwagM/lgG/iOw4MCQmhrq4uMicGBgYoMjJSFfYDgFfk4CXfs3NaWhpdXFzI95kNOTk5qqgfpQBdPlPFPAfy8vJUUSlSCOMA4LeAgAAZZ9ZERUWxoD8AvCEFfclqBwcHZYxZMz09TZaWliwqSwoa9fPzk/53Ynl5mZqbm6m1tZV2d3dls8b5+bnSvre3p9d2dnbo5OREhhnkapP4RXeD4KT1N3+T/wVfX18tb/BuZIjq6mpydnYmFxcXvcZtDQ0NMswgxcXF3OcZgHdUQR/wIBobG6Wv0WRlZWli2IaHh6WLRkpKyjVffVZZWSnDDNLZ2anGcVWhEMEPuru7pe+tvHz5UtkV5YBuEhQUFKT52djYkL29PdnZ2Wlma2tLNTU1MswgQ0ND6vs+VQVxzaRkZmM5Pj6mlpaWa5+ZMYL29/fJ1dVVEzM6OkobGxvXbH19nQ4ODmSoQfr6+tR+P1EFcSGolBvGwgtfV4C7u7vy371N0MzMjObj7e0tm+/Fgwiqra3VBsZrYnZ2lhwcHG4VpBvHOYS/isLCQioqKrpT/7o8iKCmpiYKDw9XFiSztbVFVlZWtwrKzMzUfHitqH+rFhYWRqurqzLsRh5EkMwTa2trZG1tfaugiIiIf4mQ5uXlpeQiY3kQQRJjBJ2enpKPj4/mk5ycTGNjY7SwsEAFBQVkYWGhtaWnp8twg5hM0OXlJW1ubtLk5KTevlJTU7V4Dw8PJSUYg8kE3QYf5tR4nq2VlRXpohezFdTf36/Fsy0uLkoXvZhM0NzcHGVkZFBSUhIlJCTQ2dnZtXauDtR4JycnpVA1BpMJ4qOJ7gxUVVVpbby++MivtsXFxV2LvQmTCeIjfWBgoObD/jxjXIheHdQ0Gx8fl+EGeRRBnAx1B2TooDg/P09ubm7XfKWVlZXJsBt5FEFcVHp6eirbLdvExIR00eDZTExMJEdHR00Enzr5uowL3ruiT9Cdq20Jr4GjoyM6PDxUzJgbo+3tbeXT4tlcWlqSzUajT1AkP7jPecgc0Hce+pAfcMH5HOGJuBIUrgpyA3DI9dRzpKKigsWcA3hXFcRMhIaGSt9nQUxMDAv6FYClrqBvuH7iwvE5wWmAj/IAvtUVwzgB+J3vuJ4T8fHxLOYvQ78f8R0x5efnyzizpKSkRN0MvpZCdPmJncx9g9AR0ygFSHhhVbEzT+fU1JR8l0nhip0r9SsxPwN4VQowBF/eb/NGER0dTaWlpdTR0aEUnZyZn8pGRkaUH93Ky8spNjZWvYjZ4U1MDtgY3gKQB2AGwKlad5nIOM/MASgA8LYc6H3gpMWZmMsLrpmeyrg/Ls1eyAHp4x/VnI5+Ka8hUQAAAABJRU5ErkJggg==",
  16: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAS9SURBVGhD5ZlZSD11FMePK2kSCrZIahikIFhmbrhgWqkoPiki+BKIL5aF+uCKmC/lAoppQvQgbonLk/u+/FFc0BQVQfTBQLNcwnBPPXGGO8PM6d75zzX1XukD58E558ycr86c3/n9BPgf8zoAvA8A4QDwKQB88gRGz6HnfQAAb/CC7sOrAPAFAIwCwCEAoAntGAAmAOArAHiNF6qFFADYpJsFBARgXl4eNjc3Y3d3Nw4PD+PQ0NCjGz2np6cHW1pasLCwEENCQkRx2wDwOS9YjW8pMTQ0FPv7+9GcGBsbw8jISFHY9wBgwYvnfEfBmZmZeHt7y+9nNuTn54uifuAC5CSLYp4DRUVFoqg0LoRwAIBfAwMDeZ5ZExUVRYJ+BwAnLuhLUjs+Ps5zzJqFhQW0srIiUblc0At/f38ebxSbm5vY0dGBXV1deHBwwN16WV9fx5GREZyYmMD9/X3u1oSuSfwibxC0aP1F7+R/wdfXV1o3qBupQe0/KCgILS0tpRxHR0fMysrCq6srHq5KeXk55V8DwDuioA/phm1tbTxWM7m5uVJhZJOTkzxEor6+XhHLLTU1laeo0tfXJ+bSVCHwMV0YGBjgsS/l/Pxc6Iq8KEOClpaW0MLCQooLDw/HsrIyTEpKUuQPDg7yVIPQ66rLixYF0cwkrMxaubi4wM7OTsVrpkVQcnKyFBMcHIw3NzeSLywsTPLl5OQo8tSgb1CX95koiAZBYdzQCn34cgEuLi5oZ2enKuj4+BidnJykmKamJoV/a2tL+KVubGzg4eGhwqfGgwiiYsTC0tLScHl5GR0cHFQFzczMSH5qBru7u3h3d4dra2u4urrKwzXzIILa29sxIiJC+CCJvb09tLa2VhVEw63od3V1xdbWVvTz85OueXh4YEVFBU97KQ8i6PLyUvHz9vY22tjYqAqqq6uT/PLXk1tGRgZPVeVBBHG0CNKtFwrLzs4Wtgi87dM6pRWTCaqpqVEUTd+enPT0dMkXExOj8KlhMkENDQ0KQeL3J0J7L9Hn5uYmLA1aMJkgGnzlgkZHRxX+6elpyefs7IxHR0cKvyFMJoiGT3t7eymmtrZW4e/t7ZV87u7u/2o8hjCZICI6OlqK8fLywpOTE8mXkpIi+RISEhR5aphUEE0CYgyZj48PFhQUYHx8vOK6MbPcowiisUVekNpGsbi4WBHLzZg5jngUQTs7O8JKT92JbHZ2locoaGxsRG9vb4UQT09PYfE1Fn2CjJ62OTSTnZ2d4enpqWBaToyur6+FGZC63eLiouYmwNEnKJIu3Gc/ZA7o2w99RBdo4HyO0B9CJyhCFPQWAJyWlpby2GdBdXU1ifkbAN4VBRGztGN8jsTGxpKgNQCwkgvKoU3X3Nwcjzdr6BjM1taWBH0jF0M4AsA+nXE9J2iiAIA/Df3/iM6IsaSkhOeZJbS71TWDr7kQOT9SkLk3CJmYNi6AQx/WTxRMf875+Xl+L5OysrKCiYmJopifAeAVLsAQdHj/GzUK2j1WVlYK4z0NnbQyP5VNTU0JG7+qqiqMi4sTD2L+oCbGC9bCmwBQBACLAHAln7tMYLTOrABAKQC8zQu9D7Ro0UpM4wXNTE9l9Dwazd7jBenjHz/jcOoSOGimAAAAAElFTkSuQmCC",
  17: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAARSSURBVGhD5ZlLLHx3FMe/3qXSVOhDGkSlYuWvigUapS1CrCRNN5KKjaTVLrrwXKhNBQlSGmm6QyoeK+9HvWrhESlJkQgSFvyLtIhni9OcG/dmnMzjznjMnfSTnMXMPefe32dm7u93fneA/zFvAIgG8CGATwB8/AzB1+HrvQDwphyQI7wK4EsAvwI4BEBOjL8ATAL4GsBrcqB6+BzAOp8sPj6eSkpKqK2tjXp7e2l0dJRGRkaePPg6fX191N7eTuXl5ZSYmKjKbQL4Qg7YGt9zYVJSEg0ODpKRGB8fp9TUVFXsBwBucvCSak4uKiqim5sbeT7DUFpaqkr9KAVM+UyVcQUqKipUqQIpwvgD2ElISJB1hiYtLY2F/gQQIIW+YtuJiQlZY2gWFhbIw8ODpYql0G9xcXEy3y7W19epq6uLenp66ODgQB6m8/Nz2t/fp8PDQ6vBtUdHR7LcIneTxO+mEwQvWif8m3wIMTEx2rrBs5GksbGRAgMDKSgoyGpwTnZ2tiy3SE1NDV/zHwBhqtD7PIiOjg6Zq5vi4mJNhmNqakqmUFVV1b0caxEdHS3LLTIwMKDWcVeh8BG/MTQ0JHNtwj8jnhXlgMwJ1dXVkY+PD/n6+t4LPz8/CggIIDc3N60+Pz9flltkcnJSrUtXhbhnUlZmvVxcXFB3d/e9n5ktoePjY9ra2qLt7e17sbe3R83NzeTu7q7UxsbG0unpqSy3yNjYmHrdT1UhbgSVdkMvfOObCgQHByuftjUhS5ycnFBYWJhSx+dYW1uTKVZ5FKHW1lZt8AUFBbS0tET+/v4OCeXl5Wl11dXV8rBNHkWos7OTUlJSlBuS2d3dJU9PT7uFZmZmtJqIiAi6urqSKTZ5FKHLy8t7rzc3N8nLy8tuofT0dK2mpaVFHtbFowhJHBFaXFzU8vkeOjs7kym6MIxQYWGhll9WViYP68YQQvxt8Myo5s/NzckU3RhCaHh4WMuNioqi6+trmaIbQwjxll7NtacrMIchhEy20dTU1CQP24XThXjKVzsDW7l6eBKhjY0NbYAc1jaKOzs7Wt/Gsbq6KlPs4kmEuMkMDw+nkJAQJWZnZ2WKBsuHhoYqeZGRkUqX8RDMCdndbUtub2+VqZi7ZA5rT4x4RlPzePvxUMwJpfIbjuyHjIC5/dAH/AY3nK4IfxF3Qimq0NsATnmL7Io0NDSwzL8A3lWFmNnk5GSZ6xJkZmay0B8APEyFvuWp9CE9lTNYWVkhb29vFvrOVIZ5HcBLXsFdiZycHJb529L/R/yMmCorK2WdIamtrVUng2+kiCk/cZLRJwgTmQ4pIOEb62dO5q9zfn5ensupLC8vU25urirzC4BXpIAl+OH9Hk8UGRkZykPC/v5+pZHklfm5Ynp6WvnTrb6+nrKystQHMfs8ickB6+EtABUAFgFc3X0qzgpeZ5YBVAF4Rw7UEXjR4pWY2wvumZ4r+Hrcmr0nB2SO/wBUZ/LXo3CWjgAAAABJRU5ErkJggg==",
  18: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAATMSURBVGhD5ZlJSDRXEMf/7ivRKFk0SjCgR6NGPbhviSLoRQi5CB8IeogmhygqenA5JC6gEshBBMGFiMvJFfcFxC0uqBfRg1GMxmhQ3JKoFaqxm+n3Oe2MUWckP6jDdFX1e//ufvWqe4D/Me8B8AcQCSABQPwLGI/D430K4H1xQo/BCcDXAEYA/AGATGjHAMYBfAPgHXGihvAVgA0+WUhICBUUFFBLSwt1d3fT0NAQDQ4OPrvxOD09PdTa2kpFRUUUFhYmi9sC8EacsBbfc2J4eDj19/eTOTE6OkqxsbGysB8BWIiTF/mBg3Nycujm5kY8n9lQWFgoi/pJFKDLl7KY10BxcbEsKkMUwjgD+DU0NFTMM2vi4uJY0AGAd0VB2ax2bGxMzDFr5ufnycrKikXli4KmgoODxXij2NjYoI6ODurq6qLDw0PR/RYcPzIyIlWztbU1ur29FUMM4q5ILOkWCN60TvmZ/C8EBAQo+wZXI31MTU1RVFQU2djYKPEWFhYUGBgoXQxjqays5HP8DeBjWVAgn7StrU2MNZj8/HxlcmwTExNiiMTs7CzZ29urYkVrb28X0zTp6+uTc7mrkIjhAwMDA2Lsg1xcXEhVUZyUPkExMTFKjKenJ9XU1FB9fT35+/srx729vens7ExM1cv4+Lic+4UsiHsm6Vk2lMvLS+rs7FQ9Zg8JOjg4IGdnZ8nPj9jw8LDi29nZIVdXVyV/ZmZGlasFn+cu73NZEDeCUrthKPys6wrw8PAgBwcHTUF7e3tKjJ2dHR0dHan8fn5+Sv709LTKp8WTCGpublYGz8jIoOXlZeXq6xN0fX1NQUFBSkxjY6PiYwHy2nJzc6Pj42NVrhZPIogXbnR0tLQgGb761tbWmoIYfqwdHR2lGFtbW0pPT6esrCxJhJzb0NAgpmnyJIKurq5Uv7e2tlRlWJ8ghsu2u7u7EisbXxBjxTBPIkjEUEG8+BMSElR3UzYuFFxkjO1WTCaIu/fIyEglJj4+XtqXFhcXKTMzUznu5OREm5ubYrpeTCZIZ78gFxcX2t/fV/l5Tcr+3NxclU8LkwmqqKhQ/Nz6iJSXlyt+vnuGYjJBZWVlip/XikheXp7i54bTUEwmiDtx2c9FQXe809NT8vX1VfzZ2dmqXC1MJujk5IS8vLyUGF5HfNfq6ureaqEWFhbEdL08iyCuSroT0ld6+bWCq5hurGilpaVimibPImh7e5t8fHykTplNq7lcWlqi1NRUVavExh13U1OTGP4g9wkyutsW4bfN8/Nzqe1nM+SL0e7uLk1OTkrjrq6uGpRzH/cJiuUDj3kfMgfuex/6jA8Y+6ZoLvCNuBMULQv6EMAZV5zXSG1tLYv5B8AnsiBmJiIiQox9FSQlJbGgNQBWuoK+s7S0lJrF18T6+rr0XgWgVFcM4wpg35iWwxxISUlhMX/q+/+IvxFTSUmJmGeWVFVVycXgW1GILvUcZO4FQkdMmyhAhBdWAwfz7ZybmxPPZVJWVlYoLS1NFvMzAHtRgD744/1vXCgSExOpurqaent7paaTd+aXMu4k+E83/iCZnJwsv7r/zkVMnLAhfACgGMAvAP7S7btMYLzPrAAoA/CRONHHwJsW78TcXnDP9FLG43Fr5itO6D7+BUzFamzRNylPAAAAAElFTkSuQmCC",
  19: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAS+SURBVGhD5ZlZSEVlEMfHXVHCrUxC07QXkTRSH9QwLXNBfRFD0IdAH4SyEMEFfVBfChUU0pDsTSVxA3FfcgvBBUkhEUURezCXSMm1RCfm4PdxzuQ9Xk2999IPBvTOzPm+/7nfmW/OdwH+x7wKAO8AwPsA8BEAfPgCRuPQeCEA8Bqf0GNwBoDPAOBHAPgdANCE9gcATAHAFwDwCp+oMWQCwAZdLDw8HEtKSrC1tRX7+vpwbGwMR0dHn91onP7+fmxra8OysjKMjIwU4rYA4FM+YT2+osSoqCgcGhpCc2JiYgJjY2OFsG8AwIpPnvM1Befn5+P19TW/ntlQWloqRH3LBaj5RIixBMrLy4WoHC6EcAGAXyMiInieWRMXF0eC9gHAjQv6nNROTk7yHLNmcXERbWxsSFQxF/RTWFgYj38QGxsb2NnZid3d3Xh4eMjd/2J1dVWpZLOzs0bFG+K2SPysLhC0af1Ja/K/EBoaKvcNqkaGIBFUgq2trWW8p6cnFhQU4Pn5OQ+/l+rqarrG3wDwphD0Ll20vb2dxxpNcXGxnBzZ9PQ0D1Ho6elBKysrTazaYmJi8OzsjKfpMjg4KPKpq1D4gD4YHh7msfdCd5SqIp/YXYIODg7Qzc1NxgQHB2NFRQVmZGRocgsLC3mqLlNTUyL3YyGIeiZlKRjLxcUFdnV1aZbZfYIaGxul38/PD4+Pj6UvNzdX+pydnXF/f1+Tq8f4+LjIjReCqBFU2g1joQdfLcDb2xudnJx0BWVnZ0t/Xl6exre+vo62trbS/5Dl/ySCWlpa5OA5OTm4vLyMLi4uuoJSU1Olnxegk5MT5aYIP/WOxvIkgjo6OpQHmB5IYnd3V3OH7xKUlZUl/fwbouXn7u4u/ZmZmRq/Hk8i6PLyUvP/1tYW2tnZ6Qqqq6uT/sDAQLy6upI+ukHCR5aSkqLJ1eNJBHGMEbS3t4eurq4yJikpCXt7e7GhoUFT/cji4+N5ukFMJoigYmJoH6LqJv5OS0vjqQYxqSCCxqF3LQcHByXW19dXKekkQuRTRTQWkwsSbG9v49ramlLhiJCQEJlfWVnJww1iMkE0cdpvyMe7eqqSjo6OMn9gYEDj18NkglQDKy3/zs6O9BUVFUmfl5cXnp6eanL1MJmgo6Mj9PDwkDFBQUFYVVWl6SDIqLw/hGcRtLm5qZkUX1KC5uZmTRw32nxvbm54mi7PIoiWj7+/P/r4+Cg2NzfHQyRNTU1KrFpIQEAA1tbW8lCjuEvQg7ttDt1Veo+htU9234kRxczPz+PIyAguLS096sVOcJegWPrgMe9D5sBd70Pv0QfUT1ki9EXcCooRgl4HgFOqOJZIfX09ibkCgLeEIGIuOjqax1oEiYmJJOgXALBRCyqkUxh6UC0JOgqzt7cnQZVqMYQrAOzRGZclcfsWfGTo9yM6I1ZOYiyBmpoaUQy+5ELUfEdB5l4gVGLauQAOPVjfUzB9nQsLC/xaJmVlZQXT09OFmB8AwJELMAQd3v9GhSIhIUFpS6ilp6aTduaXspmZGeVHN2pYk5OTxUHMARUxPmFj8AKAcgBYAoC/1L2XCYz2mRUAqAKAN/hEHwNtWrQTU3tBPdNLGY1HrdnbfEJ38Q9/TnBIe8vlqAAAAABJRU5ErkJggg==",
  20: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAVeSURBVGhD5ZlbSF1XEIbHS0SiNCp4gRqrhsYHoaXUJJBYrbGtRsmDKEmfvFXwwUsDBpqgD1YUm0RQKBYtPvhgqOhTvCTxbixoVUoRGomCSopivaSaqom3OOXfuBZ7r5xzPCeaeKQfDMieNXuv/6y9ZmZtif7HeBPRR0T0GRF9QUTR78DwHDzvYyLyUSf0JrgRURYRdRHRIhHxIdo/RNRLRLlE9J46UWv4mojGcbMzZ87wjRs3uK6ujpubm7mjo4Pb29vfuuE5LS0tfPfuXc7Pz+fz588LcRNElKpO2BKlCLxw4QI/ePCA7Ynu7m6OiooSwn4kIgd18io/YHBOTg6/evVKvZ/dcPPmTSHqJ1WAnitCzFGgoKBAiPpGFQLcieivs2fPqnF2zcWLFyFojog8VUHZUNvT06PG2DXDw8Ps5OQEUd+pgn4NCwtTx1vF/Pw89/f3c1tbGw8MDPCzZ8/UISaZnZ3lvr4+7urq4omJCdVtNbtJ4g99gkDR+hfvpC1MT09zamoqe3l5GWqGj48PX7t2jVdXV9UQjefPn3NmZiZ7eHjIGFdXV05ISOCnT5+qw/fk9u3buMcmEX0gBH2Cm9bX16tjzTIzM8NBQUEGIapFRkbyixcvDHHInNHR0a+NFXbq1CleWFgwxOzF/fv3RTy6Co3PceHhw4fqWLOkp6fLSTg4OGgrVVJSwrGxsYYJ3rp1yxBXU1Nj8GOlkILd3d3ltaysLEPMXvT29orYr4Qg9ExaZbYG/IIuLi5yAsXFxQZ/fHy89KHLEOzs7DD2qfBlZGRIX3V1tbx+4sQJXlxclL696OzsFLFfCkFoBLV2wxpGR0f59OnT7Ovry25ubq9t6KqqKjk5vJaiQE9NTbGzs7P06TMqEot+lZqamnR3tMy+BeGXhi0tLWmbeHt72+DXFTxtRQTIguL6sWPHeHJyUvpwv5CQEOlXX1VL7FuQJVZWVjggIEBODNlOoN8/np6e2qroOXfunPRnZ2cbfJZ4a4KwUki9YlLYZ2NjY9JfXl4ufX5+fry8vGyIR1YU/uTkZIPPEm9FEMQkJSXJCcFQH/SUlpZaFBQRESH9V69eNfgsceCCNjY2DCsDS0lJUYeJAmhWkH6FTMWb40AFvXz5ki9dumQQk5aWpg7T0Gc/7KG5uTmDX7+HbKlFByZoc3OT4+LiDGJyc3PVYZJ79+7JcUjfapZDKRD+O3fuGGItcWCCsHHFBNAtFBUVqUMMIEE4OjrKmEePHkkfVktfh2w5KR+IoNraWvlwGDY0Cuf4+Dg/efJEGooufn2AxBEaGipj0PYIKisr5XU0u+r+ssS+BaHhDAwMNAhCt6z/9YUFBwdrr6agoqLC4Ef7c/36da3jENcsvbam2LegxsbG1yZuztAe6QXhbzWJ6A37yJY+DuxbEOqJv78/nzx5ck/Dl6OtrS1D/Pr6Oufl5bG3t7cUcvz4ca2O4YxlK6YE2dRt45XD4W1tbc2iYYx6HtKDlcAJF8nhTQ52AlOConDBlvOQPWHqPPQpLjQ0NKhjjwRYiF1BkUKQHxGt7lVH7JXdzLlFRMFCEPgtPDxcHXsk2D32/0lETnpBeagjg4OD6ni75vHjx+JzwPd6McCDiP7GN66jxOXLlyFmydz/j/CNmAsLC9U4uwQN7G4y+FYVoudnDLL3BKETU68KUMHGqsFgLOfQ0JB6r0NlZGSEExMThZhfiMhVFWAOfLyfRaKIiYnhsrIybm1t1ao6KvO7Mnz7xlEC3yNw9tr9DDaPJKZO2Bp8iaiAiH4nog3Rdx2Soc6MEFEREb2vTvRNQNFCJUZ7gZ7pXRmeh9bsQ3VCpvgPtygNGFjAxJ8AAAAASUVORK5CYII=",
  21: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAASASURBVGhD5ZlHSC5XFMf/dsEQolgCsUNcJoRYQI01iQ0XogRXtizzLBAwii6MKMaChUCQkIULRdGV2EusoE8lBDGKCiI+lPdiiRAb1hPO4Awz931NQ/xG8oOzcObcO+c/d75zzr0C/2M8AHwE4DMAnwOIfwLj5/DzPgbgKQb0GFwAfAPgVwCHAMiK9heAKQAFAN4VA7WETACbPFlISAiVlJRQe3s79fX10djYGI2Ojv7nxs/p7++njo4OKisro/DwcFncFoAcMWBT1PDAiIgIGhoaIj0xMTFBsbGxsrAfAdiIwYv8wM75+fl0e3srzqcbSktLZVE/iQLUfCWLeQ6Ul5fLor4WhTDvAHgVGhoqjtM1cXFxLOhPAK6ioBesdnJyUhyja5aWlsjOzo5FfScKmg0ODhb9LWJ/f5/m5uZoZGSE5ufn6ejoSHQxyezsLHV3d0sJ6ObmRrxtlvsk8bs6QXDR+pu/yYewu7tLOTk55ObmpqkZnp6eVFRURKenp+KQt9je3iYHBwdpnKurq0VjROrq6nj8FQA/WdAnPGFXV5foa5S9vT0KCAjQCBEtOjqazs/PxaEKl5eXFBMTo/h7e3vT2dmZ6GaWwcFBeQ7uKiRi+MLw8LDoa5S8vDwlEBsbG2mlqqurKTExUSOqtrZWHCrBKxMfH6/x9fHxeZSgqakpeY4vZUHcM0mV2RIODg7I0dFRCaSqqkpzPyUlRbnHXYaaw8NDamxsJHd3d42YfyNofHxcnuMLWRA3glK7YQlra2sUFBREXl5e5OLiQltbW5r7ra2tSpD8WaoLdGFhoUaEn5+ftMJWFXR3dyfZ8fEx7ezsvJWZVAWPxMyZm5srXXdycqKWlhbq6emx/gqZ4uTkhHx9fZUgOdupKS4upszMTFpZWZH+1rUgXqm0tDQlQP6dbWxsaHw4s6np7OzUpyAWk5GRoQTHxvXBHLoUxG9dvTJs2dnZoptBdCfo4uKCkpKSNGL4h28puhJ0dXVFycnJGjEFBQWim0l0JSgrK0sJhmtJZWWl6GIW3Qhqa2vTrExUVJTUzmxubtL6+rpiXHS5ZhlDF4K44fT399cIcnZ2JltbW801tsDAQOnTNIYuBKmLoTnj9siUID5Jkn09PDysI6impkZq9fmNmjM+Obq+vhanUOjt7VXmCgsLM7ndMIYhQQ/qtvmhvBHjt2nK2MdcgCyW/SzxNYYhQbF84SH7IT1haD/0KV/gff1zhBfiXlC0LOh9AKePqSN6oLm5mcVcAwiUBTEvIyMjRd9nwf22/w8AdmpB33IdWVhYEP11zerqqnwc8L1aDPMegDd8xvWcSE1NZTHHxv5/xGfEVFFRIY7TJfX19XIyKBSFqPmZnfSeIFRiukQBIvzD+oWdeTkXFxfFuazK8vIypaeny2I6ATiLAozBh/evOVEkJCRQQ0MDDQwM0PT0tFSZn8pmZmakM++mpiZp72Vvb89C9jmJiQFbgheAcgC/Abi8fyvWMq4zywAqAXwgBvoYuGhxJeb2gnumpzJ+HrdmH4oBGeIf3LOdnMyQCFgAAAAASUVORK5CYII=",
  22: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUnSURBVGhD5ZlZLD9XFMePNUIjJSlNilpiiYimKR5QShdbJASNJ1vfqC0eSkioELUkJESk6QMJqfBoX2r7W4pUQ0oQSwhaVEnt62nO5DeTmfv7/Rh/20/6Sc6DuefMPd/fvXPPmQHwP+Y9AHAFgE8B4AsA+PwZjOah+T4CADM2obfBCACSAOAXAPgbAPAF7R8AGACAFAAwZhOVQzQALNLN3N3dMTMzE+vr67GlpQV7enqwu7v7yY3maW1txYaGBszOzkZPT09e3DIAxLEJ30YRBXp5eWFHRwdqEn19fejn58cLqwQALTZ5lh/IOTk5Ga+vr9n7aQxZWVm8qGpWgJiveTGvgZycHF7UN6wQ4h0AWPfw8GDjNBp/f38StA0AJqygb0ltf38/G6PRTE5Ooo6ODon6jhX0xs3NjfWXxc7ODo6OjmJXVxeOjY3h3t4e66KStbU1HBgY4E60qakpPDs7Y11koTgkfhcfEFS0/qU9eR82NjYwLi4OTU1NJTXDzMwM09LS8OjoiA3hGBkZwYCAADQ0NJTEOTo6YmVlJet+JyUlJRR/AQAf8oI+phs2NjayvmrZ3NxEGxsbSUKs+fr64snJiSSOVkNPT0/JV2wZGRmSmLtob2/nY6mr4PiMLnR2drK+aklISBAS0NLS4laqsLAQAwMDJckVFxcLMbSl7O3thTETExMu+dzcXHR2dpbEDQ8PS+a7Ddq2iriveEHUM3GVWQ67u7uor68vTF5QUCAZDwkJEcaoy+ChIi3+EcQHED13lpaWwnh6erowdhe9vb183Je8IGoEue0gh7m5OXRwcEBzc3M0MjLC5eVlyXhNTY2QGG1LvkBTC2NnZ8c9c66urpIYIioqSoiLjo5mh9XyYEE3Nzec7e/vcyfV1dWVZFxU8FB8clLM5eUlbm9v49bWliSG8Pb2FuJSU1PZYbU8WNBtHB4eopWVlZAYnXZyoONeV1dXiKMmWC5PJohWKjw8XEiKnrOFhQXWTYn19XXJYUHbkVZSLk8iiMRERkYKSZFRfbgLEuPk5CSJu88JRzy6oPPzc8nKkMXGxrJuSiwtLXGHiziuurqadbuTRxV0enqKQUFBkqTi4+NZNyVoK1pbW0viqqqqWDdZPJqgi4sLDA4OliSVkpLCuimxsrIiEaOtrY11dXWsm2weTVBMTIyQFBXK/Px81kWJg4MDdHFxEeKMjY25pvYhPIqg2tpaycr4+Pjg6uoqLi4u4vz8vGBUdKn+8FCLJI6jjoCaXNqC4jhVdUodDxZEDSe7/w0MDLitI75GZmtry21NYmZmRmmcOm5aXfb6s3YKzc3NSgmoM2qPeEGJiYlK4+osLCyMnVYtDxZUVFSEFhYWXDN5l9GXI741oqNdThz5JCUlsdOqRZWge3XbtOXo5e34+PhWIx/x+xD9LTfuPm+vqgT50YX7vA9pEqrehz6hC01NTazvq4AWQiHIlxf0PgAcyakjmkhFRQWJuQQAW14Q8Su9j7xGFK/9fwCAjlhQBtWR8fFx1l+jmZ2d5T8HfC8WQ7wLAH/RN67XRGhoKInZV/f/I/pGjHl5eWycRlJaWsofBqmsEDE/kpOmHxAiMY2sABZ6sH4iZ1rOiYkJ9l4vyvT0NEZERPBifgYAA1aAOujj/Z90UNBn27KyMmxra8PBwUGuMj+XDQ0Ncd/zysvLuXcvxYeUHTrE2ITlYA4AOQDwGwCcK36VlzKqM9MAkA8AH7CJvg1UtKgSU3tBPdNzGc1HrZk9m5Aq/gOsKCqaske6tgAAAABJRU5ErkJggg==",
  23: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAVgSURBVGhD5ZlvLF53FMePP0FL2CRjMgzJmtSLLctaL9qO6WyKSDSVRSNR7OVWpF7oH038jbUIyZJFlkm8IFMSb1Ba2lKh8y9IpkqjLalslElWrfhTZ/ne+N3c+6vnee5Tqo/sk5w393fO756ve+855/cg+h/zARF9SkRfElEEEX29C4b74H6fEZGXnNCb4EpEPxDRLSKaJyJ+h/YPEXUQURoRucuJGiGBiCaw2eHDh/n8+fNcXV3NjY2N3NbWxjdv3nzrhvs0NTVxTU0NX7p0iY8cOSLETRJRspywOYoQePToUW5paWFb4vbt2xweHi6E/UxEdnLyMj/B+ezZs/zq1St5P5vhwoULQtQvsgAt3wkxe4Hs7Gwh6ntZCHAjoumQkBA5zqY5fvw4BM0S0fuyoB+h9s6dO3KMTdPf388ODg4QlSUL6jp06JDsb4i5uTnu6enhGzdu8L1793hhYUF22ZKHDx8qHzn+iE+ePJGXDbNZJIa0BQJN61+8k9bw9OlTTk5OZk9PT13P8PLy4oyMDF5aWpJDFAYGBjgyMpKdnZ3VmP3793N8fDw/evRIdrfI1atXsccqEX0sBH2OTWtra2Vfk8zMzHBgYKBOiGxhYWH88uVLXdzg4CC7ubm95ivMz8+PHz9+rIuxxPXr10U8pgqFr3ChtbVV9jVJamqqmoSdnZ3ypAoLC/nEiRO6BK9cuaLGoA2gt2mTz83N5czMTN63b596PS4uTncvS3R0dIjYb4UgzExKZzbCs2fP2MnJSU2goKBAtx4TE6OuYcoQDA8P68SOjo6qa6Wlpep1d3d3np+fV9cs0d7eLmK/EYIwCCrjhhHu37/PBw4cYG9vb3Z1deXJyUndekVFhZocXkvRoGdnZ5XxKSsriy9fvqyLQWEQMdgTvkbZtqCNjQ3FFhcXeWpqitfX13XrmobHRivnmTNn1JjQ0FBlf6NsW5A5nj9/zv7+/mpyqHamWF5e5vz8fO1splTIoaEh2dUsb00QntTJkyfV5PCdjY+Py24qY2Njqi8MrzD6mbW8FUEQgz6iTRD9wRyY5NF/RBl3dHTk4OBgbmhokF3NsuOCVlZWdE8Ghm/CEqhk09PTSoO+ePGiGos20NfXJ7ubZEcF4TuIiorSiUlJSZHdDKE5wHFiYqK8bJIdE7S6usrR0dE6MWlpabLblqytrcmXlAIi9jFaHcGOCUpKSlITwGuSl5cnu+i4du2a0nQPHjzICQkJ8jKfPn1a3e/YsWPyskl2RFBVVZXuyaB3YAabmJjgBw8eqIamK3pKZWWl6o+Rv7u7W90PU4OHh4e6fu7cOc3dzLNtQRg4AwICdIJcXFzY3t5edw0WFBSkvJoAPUo70EJAenq6Mjn4+Pio1zHX4WhhlG0Lqq+vfy1xU4beIgSBrq4us9M2Sjf2t4ZtCyoqKmJfX19lWrZkmK7lAoBJIDY2VpnZhBA84YiICL57967O1whbCbJq2sYrh8PbixcvzBp85POQFsyBEIDxXx5wrWErQeG4YM15yJbY6jz0BS7U1dXJvnsCPIhNQWFC0IdEtGSpj9gq5eXlELNGREFCEPjDmmZmS2we+/8kIgetoEz0kd7eXtnfpkFD3vw5IFcrBrxHRH/jsLWXQPknokVT/z/Cb8Sck5Mjx9kkxcXFohiky0K0/AonWy8QGjG1sgAZfFi/wRmP05qD1m4wMjLCp06dEmJ+JyIXWYAp8OP9XygU+Nm2pKSEm5ububOzU+nMu2WYJHBULysrU85emPWIaA5FTE7YCN5ElE1Eg0S0Iuaud2ToMyNElEdEH8mJvgloWujEGC8wM+2W4X4YzT6RE9qK/wChCEMDsSh+ggAAAABJRU5ErkJggg==",
  24: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAT+SURBVGhD5ZlZSD9VFMePCy4YamIa5A764IMRqYiapi1u+CBK9OTWg4hpgmCKouaCK6gEIdGD4JIo+OK+5BZorimkqCBqqJXmlvt64gy/GWbu/7f51/yN9IHzMvecmfP9zZ1z7r0/gP8xbwGABwB8AAAfA8BHz2D0HHreuwBgwyb0OpgBQAoA/AQAfwMA6tAOAGAEANIAwJxNVBs+B4BVupmXlxdmZWVhY2MjdnR04MDAAPb39//nRs/p7OzEpqYmzMnJQV9fX17cGgDEswmro5QC/fz8sKenB+XE0NAQBgUF8cK+BQA9NnmWMnJOTU3Fu7s79n6yITs7mxf1HStAzGe8mJdAbm4uL+oLVgjxBgD87u3tzcbJmuDgYBL0FwC8yQr6ktQODw+zMbJmenoaDQwMSNTXrKCfPT09WX+t2N3dxfHxcezr68OJiQnc399nXTQyOTmJbW1t2N7ejkdHR+ywWhRF4ldxgaCm9Q/NyYewtbWF8fHxaGVlJekZNjY2mJ6ejqenp2yIUhYXF9HY2FiIn52dZV3UUlFRQXHXAODIC3qPbtTS0sL6qmR7exudnZ0lQlgLDAzE8/NzNlQCVVIfHx9J3Pz8POumlu7ubj6WVhUcH9KF3t5e1lcliYmJQgJ6enrcmyopKcHQ0FBJcuXl5WyohMLCwld+iIcKGhkZ4WM/5QXRmonrzNqwt7eHRkZGQgLFxcWS8YiICGGMVhmqmJmZQUNDw0cLGhwc5GM/4QXRQpBbbmjD0tISurm5oa2tLZqZmeHa2ppkvK6uTkiOpqWyBn11dYUeHh6Cn1jYswu6v7/n7PDwEDc3N/H29lYyLmp4qKpyZmZmCj4pKSno6OioO0HqODk5QQcHByE5qnYsY2NjwjgVjoODA7S2tpafIHpTUVFRQmL0na2srEh8qJS7urpKxqkSmpuby0sQiYmJiRGSIqP+wJKcnCyMl5WVcdd2dnbQwsJCPoLoAxe/GbK4uDjWTdwvuErIQ/GWlpbC2MbGhiROE08q6OLiAsPCwiRiEhISWDc8Pj5GJycnwYc2jA0NDVhfX4+1tbVoamoqjBUVFXHXtRX2ZIKur68xPDxcIiYtLY1146DklPUcddbc3MzeRilPJig2NlZ4OK0WqOurggSxCWuyZxVEU0L88ICAAFxfX8fV1VVcXl4WjJou9SxaQdNUysvLk1h+fj5mZGSgiYmJcK+kpCRujBat2vBoQVRmxd8DGSWkr6//yq/s4uLCTU11XF5e6rYo0L6FTVyV0fJIkyAq2+Ltw9zcHOuilkcLKi0tRTs7O7S3t9dodHJ0c3PD3kICbRDd3d2FGG2nGo8yQQ9abdOUo45/dnam1shH036IoG+M9ydTtphVhzJBQXThIfshOaFsP/Q+XWhtbWV9XwT0IhSCAnlBbwPAqbo+ImdqampIzA0AuPCCiF/8/f1Z3xeBYtv/GwAYiAVlUB+h46SXBFVExXHAN2IxhCUA/ElnXC+JyMhIEnOo6v8jOiPGgoICNk6WVFZW8sXgK1aImO/JSe4FQiSmhRXAQh/WD+RMr3Nqaoq9l05ZWFjA6OhoXsyPAGDCClAFHd7/QYUiJCQEq6qqsKurC0dHR7nO/FxGByr0p1t1dTW391LsqXapiLEJa4MtAOQCwCwAXCl+FV0Z9ZkFACgEgHfYRF8HalrUiWl5QWum5zJ6Hi3NXNmElPEv+WExSGG5PIAAAAAASUVORK5CYII=",
  25: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAVkSURBVGhD5ZlZSF1XFIaXUzSxVOsMVWvU+hJoKU3EeeqgJvggSlEQh/jgQ6s1VKiiD1ai1gEVCkWLguJQUd9inGfRVIMUoc5DsCi2DhXrrElWWYd7Duds770e5yv9YIHcvfbe6797WGtfAf7HmAPARwDgCQCfA8BnV2A0D833MQBYsAGdBUMA+BoAOgFgDQDwGu0fAOgBgAQAeJcNVA5hADBNgz148ACTk5OxqqoKnz17hu3t7djW1nbpRvM0NjZidXU1pqamopubGy9uDgCi2YDVkU0d3d3dsbm5GTWJrq4u9PX15YX9BABabPAsP5JzfHw8vnnzhh1PY0hJSeFF/cwKEPMVL+YmkJaWxouKZYUQ7wDAn87Ozmw/jcbPz48E/Q0A77GCviG13d3dbB+N5uXLl6ijo0OivmcF9d+/f5/1l8XKygoODg5ia2srvnjxAtfX11kXCUdHR7i6uopra2tKjcbb399nu6lEcUn8Lr4gKGn9S3vyNCwuLmJ0dDSamJhIcoaFhQUmJibi9vY224WjrKwMTU1N0czMTKlRW01NDdtNJbm5uTTvIQB8wAv6hAKpra1lfVWytLSEd+/elQhhzdvbG3d3d9muGBsbe8yXtZKSErabSpqamvh+VFVw+NAHLS0trK9KHj9+LEyupaXFrVRmZiYGBARIAsvJyWG7oqurq9B+69YtvHPnDt6+fVswfX19rKioYLuppKenhx/vS14Q1UxcZpYD7X8KhA/q6dOnkvZHjx4JbVRliNnc3ERzc3NBTH9/Py4sLEhsfn4et7a2JP3U0dHRwc/3BS+ICkGu3JDD+Pg4Ojk5oaWlJRoaGuLc3Jykvbi4WBBE21KcoEdGRoQ2R0dHSb+zcm5Bb9++5WxjY4P7Rl+/fi1pFyU8ZG/OyspKoY1yCO0K2qpZWVmy52c5tyB10FaxtbUVgqbbTkxSUpLQRmeF/5s3T09PnJ2dlfQ5iUsTRCsVHBwsBEdnZGpqSuLj4+NzTARrDg4OXC6Sy6UIIjGhoaGSwCg/iDk4OMB79+4J7RERETgwMMCdyYyMDNTW1hba2JVVx4ULokDFK0MWFRXFunHnjvLX0NCQ0rni4uKE/jY2NkpzmDIuVNDe3h4GBgZKxMTExLBusqDHHD8GrdbMzAzropQLE3R4eIgPHz6UiElISGDdZNPZ2SkZa2JignVRyoUJioyMFCanaoHOgTpGR0fxyZMnGB4ejiEhIdwXIoaqA348Y2NjrlCVw4UIKi8vl3ybXl5e+OrVK5yensbJyUnBKOnS2SHoaSLuU1paKoxHPvTk59uo2pDLuQXRYbWzs5MEZ2BgILmleLO3txdWgioGFxcXoU1PT49bMSpEFQ81weg5IpdzC6qvrz8WuCqj8ki8tcbGxtDKyuqYn9gKCgok853EuQVlZ2ejtbU1d7WeZLSN6EEnhrZhWFgYGhkZCSLo1UmFbENDg8RXDsoEnarapi1Hj7ednR21Rj7qcsny8jK3tehssRXFaVAmyJc+OM17SJNQ9h76lD6oq6tjfW8EtBAKQd68ICsA2D4pj2gqRUVFJOYIAOx5QcRvHh4erO+NQPHs/wMAdMSCvqM8QoXjTYLSgOLngB/EYghjAPiLfuO6SQQFBZGYDVX/P6LfiDE9PZ3tp5Hk5eXxl8G3rBAxv5CTpl8QIjG1rAAWOlil5EzLOTw8zI51rVDFTpW6QsyvAGDAClAF/Xi/TBeFv78/5ufn4/Pnz7G3t5fLzFdlfX193D/dCgsLubeXrq4uCVmhS4wNWA6WAJAGACMAcMDXXddklGdGASADAN5nAz0LlLQoE1N5QTXTVRnNR6XZh2xAyvgPQ6IgVqXDrFEAAAAASUVORK5CYII=",
  26: "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",
  27: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUBSURBVGhD5ZlHSHRXFMePBWsIsUQjsUvElSFGXajRaBIrLkQJ2VizERKNkEVsCyOKFVSiQUIWLpSILu0ltk80KhKEWLGgWBJL7L2ccB6+x3t3nOeMYxnJD85m7jnvnv/cO/ec+wbgf8y7AOAOAJ8AwOcA8NkzGM1D830IAFZsQg/BFAC+AYDfAWAbAPAF7V8A6AOAVAB4m01UFb4CgDl6mJeXF6anp2NtbS02NTVhV1cXdnZ2PrnRPM3NzVhXV4dZWVno4+PDi1sAgAQ2YTkKKNDX1xfb2tpQm+jp6cHAwEBe2E8AoMMmz1JIzikpKXh9fc0+T2vIyMjgRf3MChDzJS/mNZCdnc2L+poVQrwFACve3t5snFYTFBREgv4BADNW0Lektre3l43RasbGxlBPT49E/cAKeuPp6cn6q8Tm5iYODQ1hR0cHDg8P487ODuvCcXJywvlub2/L2tbWFu7t7bHhSrk9JP4UHxBUtA5oT6rD6uoqJiQkoLm5uaRmWFlZYVpaGh4dHUn8Kyoq0MLCAi0tLWWNfCIiIiSxchQXF9O8FwDgwAv6iBKpr69nfZWytraGTk5OEiGsBQQEcKvCk5ubq+CjzNzd3SXzydHa2srHUVfB8Sl90N7ezvoqJSkpSZhcR0eHW6n8/HwMDQ2VJFZUVCTElJaWoqGhIRobG0vMxMQEzczMuOfwcYmJiZL55Ojr6+PjgnlB1DNxlVkVaI8bGBgIk+fl5UnGabvwY9Rl8Ozv7+Pi4iIuLy9LbGNjA6uqqlBXV5eL8fDwUNiucnR3d/PzfcELokaQazdUYWpqCl1dXdHa2hpNTU1xYWFBMl5dXS0Iom15X4E+ODhABwcHzp9WbHp6mnWRRWNBNzc3nO3u7nLf8NXVlWRcVPBQlZMzNjZW8C8sLGSH70VjQXIcHh6ivb29kCCddnIMDg4Kvi4uLnh+fs663MuTCaKVioqKEhKk39ns7CzrJiE4OFjwp636EJ5EEImJiYkRkiOj+iDH+Pi44Eu/oePjY9ZFJR5dEG0T8cqQxcfHs24KJCcnC/6ZmZnssMo8qqDT01MMCwuTiFGlhtBq2NjYCDEjIyOsi8o8mqCLiwsMDw+XiElNTWXd7oT6Pj7Gzc1N4aRUh0cTFBcXJyRFVZ5aG1WhK706KyrHowiqqamRrIy/vz8uLS3h3NwczszMCEZFl2oWi+gajZWVleywWmgsiBpOR0dHiSAjIyOhdRGbs7MztzXFnJ2dCZ0BWX9/v2RcXTQW1NjYqJC4MqP2iBW0srIiEU+tlCZoLKigoABtbW3Rzs7uXqM3R5eXl5L4+fl5rpugceoJ19fXJePqcpcgtbpt2nLUDdPRK2fkI74P8dCJRmPKxtXlLkGB9IE69yFt4q770Mf0QUNDA+v7KqCFuBUUwAt6DwCO1Kkj2kR5eTmJuQQAZ14Q8Yefnx/r+yq4vfb/BQB6YkHf01GqSU/1EkxOTvKvA34UiyHeAYC/qYK/JiIjI0nMrrL/j+gdMebk5LBxWklJSQl/GHzHChHzCzlp+wEhElPPCmChH9av5EzLOTo6yj7rRZmYmMDo6GhezG8AYMQKUAa9vN+ggyIkJIR7SdjS0sI1klSZn8sGBga4P93Kysq4u5e+vj4J2aRDjE1YFawBIBsAxgHg/PZbeSmjOjMBALkA8D6b6EOgokWVmNoL6pmey2g+as0+YBO6i/8Awl6Er2vJ3gMAAAAASUVORK5CYII=",
  28: "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",
  29: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAVqSURBVGhD5ZltSJZnFMePGiU4hm/TydSlmR9EJmPpB3U53ZxlLxDpEBTT5gdh0ylCK/RDiuLUoGA1Yu6DH4xJoSCaljnfRjYNGcJ8BZGGOV/mlKWJmp3xv/G6uO9LH5/HNHtkP7hA7nOu+z7/+7qec851S/Q/5h0i+oCIPiaiz4jo010YeA6eF0REbmpAr4IDEX1FRL8Q0d9ExG9w/ENEbUSUSURvq4FaQgIRDeNmwcHBfPHiRa6srOS6ujp+8OABNzU1vfaB59TX1/OtW7c4NzeXQ0NDhbgRIkpRA96MYkwMCwvjxsZGtiZaWlo4MjJSCPueiGzU4FW+g3NGRgavrq6q97MaLl26JET9oArQ84UQsxfIy8sTor5UhYC3iOjPkJAQdZ5VExUVBUGTROSkCvoaaltbW9U5Vs3jx4/Zzs4Oor5VBf165MgR1d8ipqamuLOzk+/fv8+PHj3imZkZ1WVD+vr6tEz28OFDnp6eVs0Ws5YkftcnCBStf7Ent8LY2BinpKSws7OzoWa4ublxVlYWz8/Pq1M0IAIp2NbWVs5xdXXl7Oxsfv78uepultLSUtxjmYjeF4I+xE2rqqpUX5M8ffqUfXx8DELUERERsS7AmpoatrGxWeern7OwsGCYY46GhgYxH12Fxie4cO/ePdXXJOfPn5dBIECsVFFRER87dswQYElJiZyDrenk5CRtgYGBfPnyZY6PjzfMycnJMTzLHG1tbWLu50IQeiZtK1gC9vv+/ftlAIWFhQb7iRMnpA1dhuDGjRvy+sGDB3lubk7a0tLSpM3BwYEnJyelzRzNzc1ibrQQhEZQazcsob+/n/39/dnd3V17+MjIiMF+8+ZNGRy2pSjQSUlJ8np6erphztDQEO/bt0/at7L9ty3o5cuX2pidneUnT57wixcvDHZdwWN95jx16pS8riagZ8+esYeHh7Sjd7SUbQvaDATm7e0tA0O2EyQmJppcIWw/fbZMSEgw2DfjtQnCSp05c0YGhd8ZtpLg6tWr0ubn58crKyvSdvv2bWnDOHnypLSZ47UIgpi4uDhDUKgPeiYmJtjR0VHajx8/zrW1tXz9+nVD9sOIjo42zN2MHRe0tLRkWBmMc+fOqW4a1dXVJusQEoz4+/Tp0+pUk+yooMXFRe1N6wNLTU1V3QzgOThrHThwQPPHbw4pHSLEPZARLWXHBC0vL3NsbKxBTGZmpupmktHRUR4YGNASCQgKCpL3yc/PV91NsmOCkpOTZQDYRgUFBaqLAQSOJNHe3r6uqx8fH2d7e3t5v7t37xrsm7EjgioqKgwrc/ToUe2NDw8P8+DgoBwouqhZQPdgreVHDRNcuHBB2lCwTTW2G7FtQWg40broBeHt6rtnMXx9fbWtCVCIXVxcpC0gIEBbVX0HgYH0vhW2LejOnTvrAjc18LaFIFBeXr7ORz9QfMWKWsq2BRUXF7Onpyd7eXmZHchm+gIK0OupR49Dhw7xlStXDH6WspGgLXXb2HLY4zi3bDbgo56HBLB1dXVpp9yenh6TfpawkaBIXNjKecia2Og89BEuoJ/ai2Ah1gRFCEHvEtG8uTpirVy7dg1iVojIVwgCv4WHh6u+e4K1Y/8fRGSnF5SDOoIf6l4Cn8LWPgfk68UARyKawDeuvcTaKXjW1P+P8I1Y+xKzFygrKxPJ4BtViJ4f4WTtCUInpkoVoIIf1k9wxnJ2d3er93qj9Pb28tmzZ4WYn4nIXhVgCny8/wuJIiYmRmtL0NKj9Udl3q3R0dGh/dMNDSvOXmufuqaQxNSALcGdiPKIqIeIlvS91xsYqDO9RFRARO+pgb4KKFqoxGgv0DPt1sDz0JodVgPaiP8A7UUCIKWOVwEAAAAASUVORK5CYII=",
  30: "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",
  31: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAASwSURBVGhD5ZlHSHVXEMf/flZQREGNWBAD2VmIFdQQSxJFEARBFFxE3AhfjIhgrKAiJNg1EDSIulBsCxUr9rKxEKKoIIEoZmMlig37hLl89/Le8b3PEuK7j/xgFt4zc+78vcc5c47A/xhHAD4AvgDwFYCoNzB+D7/PF4CTmNBrsAbwHsAkgCMAZED7G8AMgO8B2IqJPockAH/wZIGBgZSbm0ttbW00MDBA4+PjNDY29p8bv2dwcJDa29upoKCAQkJCZHF/AvhWTPhj/MiBoaGhNDIyQmpiamqKIiIiZGE/AzARkxf5iZ0zMjLo/v5enE815OXlyaJ+EQVokiiLMQYKCwtlUWmiEMYGwF9BQUFinKqJjIxkQfsA7EVB37Ha6elpMUbVLC8vk6mpKYv6QRQ0HxAQIPo/i93dXZqfn5eq0vr6ujj8JBzb3d0tFaC7uztx+Ek+FInfNQsEb1qnvCZfwv7+PqWmppK9vb2yX5iYmFBwcLBUdp/D9vY2mZubS7E8z/n5uejyJOXl5Rx/A8BDFvQ5T9jZ2Sn66uX4+Ji8vb0VIaKxsP7+fjFMi+vrawoPD1di3Nzc6OLiQnR7kuHhYXkO7iokwvnB6Oio6KuX/Px8JRErKyvKysqi0tJScnd3V547OzvTycmJGCrBXyYqKkrrl8CxrxE0MzMjz/GNLIh7Julv4DlcXV2Rk5OTkkhTU5Mytri4qCwhNnHpHR0dUXV1NTk4OGiJ+TeCJiYm5Dm+lgVxI/jo5fq4ubmh3t5eKisro7S0NDo9PVXGzs7OyMbGRklyaGhIKzYzM1NLhIeHh7Q8DSroYzQ0NCjJsrC9vT2tcS4iPGZpaUl1dXXU09Nj+C+ki9bWVkpMTJT3BDIzM6Pm5mbRjXJycigpKYnW1takn1Up6OHhgVxcXJTEWExtba3oJsGVTZOOjg71CTo4OCBXV1eys7NTkmOBJSUlousjVCno8vKStra2pI6Bzy7W1tZKkvX19aK7FqoUJKLR2pOnp6dU5vWhOkFcwkX6+vqUJLma7ezsiC4KqhC0srJC8fHx5OfnR76+vo+SaGxsVJK0tbWVej59qELQ5uamkgRbVVWVMsbLi4XKY3wfwVVQH6oQxCQnJ2uJSklJoaKiIvL399d63tXVJYZqoRpBvIy8vLy0khctOztbDHsE3yTJ/o6OjoYTxHCjmZ6e/qjR9PHxoZaWFtFdJ3zE4GMDfx0+R/EW8FJ0CXpRty1yeHhICwsLNDk5KbU0L7ktur29lQ51bK8Rw+gSFMEPXnIeUhO6zkP+/IDP9cYIf4gPgr6UBTkDOOcTpzHCjTCAWwCfyoKYhbCwMNHXKIiJiWFB6wBMNQVlv3v3TjpCGxMbGxtkYWHBgko0xTB2APb4jsuYiIuLYzHH+v5/xHfEVFxcLMapkoqKCrkYZIpCNPmVndReIDTEdIoCRPgPq4md+XMuLS2JcxmU1dVVSkhIkMV0ALASBeiDL+93uVBER0dTZWWldCU1Ozsr7cxvZXNzc9Kdd01NDcXGxkr3FQAOuIiJCT+HTwAUAvgNwLVmv2YA431mFUApAFcx0dfAmxbvxNxecM/0Vsbv49bsMzEhXfwDi760pgSwqTcAAAAASUVORK5CYII=",
  32: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAViSURBVGhD5ZlpSB1XFMePGhdURENqxYVqgwvBhapRMJZWuyQEBEUUhXxQ/CJaK+KHuoWoCC1qiKiIFnGBSI1foiSauMYl4kasK5iSBhRxt4pLsIqecobMZeb65uW9aPRJf3C+3HvOzPnP3HfumfsA/sd8BgCeAPA1AHwPAN+dgdF96H5eAGDNJ/QxmAFAIgB0AsAaAOA52j8A0A0APwOABZ+oJkQBwF90sevXr2NaWho+fPgQnzx5gu3t7djW1vbJje7z9OlTrKurw8zMTAwICBDF/Q0AMXzC6viVAm/cuIHPnj1DXaKrqwuDgoJEYSUAoMcnz/MbOSclJeHh4SF/PZ0hPT1dFFXGC5ASKYq5CGRlZYmi4nghhDkAzPn5+fFxOk1wcDAJWgYAK17QT6T2xYsXfIxOMzIyggYGBiTqF15Qn6+vL++vEYuLi9jX1ydUpampKX5akdnZWezu7hYq2ujoKO7t7fEuGvG+SPwpLRC0aW3RmtSG5eVljI2NRSsrK7Zf6Onpob+/v5CkEv39/Xjz5k00NTWV7jPo6uqKJSUlvPsHyc/Pp/h9APhCFPQVXbC+vp73VWRjYwM9PDxkCUmNhDU1NfFhglBDQ8Nj/lJLTU3lw9TS0tIixlJXIfAtDTx//pz3VSQjI4MlYGJigikpKZibm4sODg5s3MbGBjc3N1kMLSlnZ2c2T2+Wkr937x5eu3ZNJurly5ey+6mDlu37uB9FQdQzCb8BTaDErK2t2c0rKyvZ3NDQkOwNSJcebdLiOL1BaQFaX1+XPQx6QJrS0dEhxv0gCqJGUO26l7K/v4+PHz/GvLw8jIuLw62tLTa3vb2N5ubmLLHm5mY2Ry3M1atX8fLly+jp6cnGRSIiIlhcVFQUP63IiQWpo7y8nCVFwpaWltjc0dERHhwcCMVkYWFBFkcEBgay2OTkZH5akU8iqKamBiMjI8U9AS9duoRVVVW8myIDAwNCjCiImmBNOXVB9ORtbW1ZMpRYUVER76bI3NycrFjQcqQ3qSmnLmhlZQXt7OzQ0tKSJUUCc3JyeNdjkBg3NzcWR6ZNhSNOXdC7d+/w7du3QsdAP3wzMzOWXHFxMe/OePPmDbq4uMjElJWV8W4f5NQF8Uhae3RyclLZ0rx+/RodHR1lYkpLS3k3jThVQVTCeRobG1mSxsbGQs8mhd6mVIy+vj7W1tbKfLThxILGxsYwNDQUvb290cvLC3d3d2XzFRUVLFkLCwuhTItQ5+Du7i6bb21tlcVry4kFzczMsITI7t+/z+ZoeZFQcY7OI6gKisTExMhiqSOYn58XliBdVzRV+5QSJxZEREdHyxK7c+cO3r17F318fGTjjx49YjETExOyOTLquKkN4sfPvFOgZSRdOqqM75oTEhKO+SgZLWlNORVBxNraGsbHx+OVK1dkydDGWF1dzbtjWFgY2tvbC02oOiOfxMREPlwRVYK06rZ5VldXcXBwEDs7O3FyclLxtGhnZ0cwKiLqjHxUlXolVAkKogFtvod0CVXfQz400NDQwPteCOhFvBf0jSjIBgB26IvzIkKNMAAcAMCXoiBikL5HLiK3bt0iQVMAYCAVlEotCH1CXySmp6fRyMiIBOVIxRCWALBEZ1wXiZCQEBKzofT/EZ0RY3Z2Nh+nkxQUFIjFIJkXIuV3ctL1AiERU88L4KEfViU50+scHh7mr3WujI+PY3h4uCjmDwAw4QUoQYf3i1Qo6Ni2sLBQOJLq6ekRduazst7eXuE878GDB3j79m3xIGWFihifsCZ8DgBZAPAKAP6V9mvnYLTPjANALgDY8Yl+DLRp0U5M7QX1TGdldD9qzZz5hFTxHzobQaFd+2kJAAAAAElFTkSuQmCC",
  33: "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",
  34: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUjSURBVGhD5ZlrKG5pFMeX20HkUsaIhClf5NK4h8kwFyJ1SmmUD0c+IGMmKeN+r5FLLlMTk1AO41IHueZ++eI+hGJqpPngHnIQBmtau7Of9vsc3vO+h+HV/Gp9eZ619l7//e69nvU8L8D/mE8AwBEAvgCArwHgq0cwug/dzwkATPmEPgY9AIgDgCEA2AcAfEI7AIBRAPgBAAz4RBXhOwD4ky7m5uaGycnJ+Pr1a+zs7MSBgQHs7+//z43u09XVhQ0NDZiWloZeXl6iuL8A4BWfsDx+pkBvb2/s7e1FVWJ4eBj9/PxEYb8AgBqfPE8BOcfHx+P19TV/PZUhJSVFFPUrL0BKmCjmOZCeni6KiuKFEPoA8Le7uzsfp9L4+/uToB0AMOYFfU9qR0ZG+BiVZmZmBjU0NEjUT7ygCVdXV95fIba2tnBiYkKoSsvLy/y0QkxNTWFrayu+efMGj46O+Gm5vCsSf0gLBC1ax/ROKsPOzg5GRkaisbExWy/U1NTQw8NDKLuKsrKygtra2uwac3NzvItcCgsLKe4SAKxEQZ/ThZqamnjfOzk8PEQHBweWBG8krKOjgw97D6qknp6eMrELCwu8m1x6enrEWOoqBL6kgb6+Pt73TlJTU1kCOjo6mJCQgLm5uWhpacnGzczMPvj6UAz/MJQVNDo6KsZ+Kwqinkn4BhTh/PwcTU1NWQLV1dVsjr4FLS0tNifv1ZudnUVNTc17CxocHBRjvxEFUSMo9+ZSLi8vsa2tDfPz8zEqKgqPj4/Z3Nu3b1FfX58l193dLRMrcnFxgY6OjsxPKuzRBcmjsrKSJUbCtre3eReBpKQk5hcXF4dWVlaqJaiurg7DwsLENUF44jU1NbybwPj4OEve19cXDw4O0MTERHUE3dzcoLm5OUuIxJSVlfFuAicnJ2hrayv4vXjxAtfW1vDs7AwNDAxUR9Du7i5aWFigkZERS4oE5uTk8K4YGxvLfAoKCoSxzc1NNDQ0VB1B9ITX19eFjoH2Lnp6eiy5iooK5idZLzA4OJiNU4GQPoyNjQ02pwgPLohH0tqjjY0NXl1d4enpKVpbW7Nx2jDW19cL3155eTnq6uqyuby8PGFcUWEPKohKOE97eztLjlqavb09oU26bc2RZ42Njfylb+Xegugdf/nyJTo7O6OTk5Pw9KVUVVWxpOhjp3WKyjef8Ifs0QStrq7K3LikpITNURdBQsU5Oo8gSBS9SpmZmTKWlZWFiYmJQvskxkRHRwtz1LQqwr0FEeHh4TKiIiIiMCMjA11cXGTGm5ub+dD3oIfw5EWBvgl7e3uZ5HmjJ68IVLal24f5+XneRS4PIojY39/HmJgYmVWejHq02tpa3v1OaB2zs7MTOnUyRV81kdsEKdVt81AVm5ycxKGhIVxaWlL6tIg6DSos1EWQKRt/myA/GlBmP6RK3LYfcqGBlpYW3vdZQD/EO0G+oiAzADih3eNzhBphAPgHAD4TBRGTPj4+vO+zIDAwkAQtA4CGVFCiurq6sIV+TlBFpG0IAORIxRBGALBNZ1zPiZCQEBJzeNf/R3RGjNnZ2XycSlJUVCQWgx95IVJ+IydVLxASMU28AB76sKrJmX7O6elp/lpPyuLiIoaGhopifgcAHV7AXdDh/RYVioCAACwuLhaOpMbGxoSV+bGMDlToT7fS0lIMCgoS91S7VMT4hBXhUwBIB4A5ALiQ9mtPYLTOLAJALgBY8Il+DLRo0UpM7QX1TI9ldD9qzWz5hG7jX9YMSN/tKoQtAAAAAElFTkSuQmCC",
  35: "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",
  36: "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",
  37: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAU0SURBVGhD5ZlbLKVXFMeX2zCIEKouQaZS8TBImSEZmpa2yIhEIpFKRlLxImm1EQ+dcUkQScUlLjUNjRgPpHhBBjPut3nAREpc06iEB/fUMMhQrGZ9mW/nnOWc4xzOcKS/ZL18e62z19/+vrXX3gD+x3wEAL4A8DkAfA0AX12B0Tw0nx8AOPKELoIVAHwPAL0AsAUAeI32DwAMAMCPAGDDE9WGbwHgL/qx+/fv4+PHj7Gurg6fP3+O3d3d2NXV9cGN5mlra8P6+nrMyMjABw8eyOL+BoDveMKa+IUCg4OD8cWLF2hI9PX1YWhoqCzsVwAw4slz8sk5JSUFT05O+O8ZDE+ePJFF/cYFKBIni7kJZGZmyqKSuBDCGgCWAwMDeZxBExYWRoLWAcCOC/qB1Pb39/MYg+b169doYmJCon7mgobv3bvH/bVidXUVh4eHpao0PT3NhwUHBwe4sbGBW1tbGm1zcxPfvHnDw9Xyvkj8qVggaNPapXdSF9bX1zExMRHt7OzEfmFkZIRBQUFS2eWUlZWhvb09Ojg4aDTyiYqK4uFqKSgooLmPAMBDFvQZJdPQ0MB91bK9vY0+Pj5CCDcS1traqhSTm5t7xk+d+fr6KsVqoqOjQ46jrkLiS3rw8uVL7quW9PR0MbmFhQWmpqZKCbu5uYnnTk5OSq9OUVERmpub4+3bt5XM0tJSWmX6I8ixtPLaMjAwIMeFy4KoZ5K+AW149+4dOjo6ismrq6vF2OjoKJqZmYkxxVdvZ2cHFxcXcWlpScnoG3z69CkaGxtLMf7+/ri3tyfizqOnp0ee7xtZEDWCKt97VRwdHWFzczPm5eVhUlIS7u7uirG3b9+itbW1ENTe3q4UqwqK9/DwkPxpxebm5riLRi4tSBOVlZVCDAlbW1vjLmdISEgQMfn5+Xz4XD6IoNraWoyLi5P3BDQ1NcWamhrudoZXr14JMZ6ennh4eMhdzkXvgk5PT9HFxUUkRmJKS0u5m0rCw8NFHK3uRdC7INosXV1d0dbWViRHAnNycrirEuPj48KfvqH9/X3uohV6F0QdAFUvqlZ0drGyshKJlpeXc3dBcnKy8KNt4KLoXRBHobXHO3fuSGWeQ6vh7Ows/KjcXxS9CqISzmlpaRGJ0kZKew2ns7NT+Hh7e+Px8TF30ZpLC5qYmMCYmBhpA/Tz8zvz7ldVVYlkbWxspJ6PQ0d62UeXrkAVlxY0Pz8vkiErLi4WY/R6kVB5jO4jqApyFI7RWFFRwYd14tKCiPj4eCVRjx49wqysLAwICFB63tjYyEMl0XJnQDY4OMhddEIvgug1unv3rlLy3NLS0niYxPLysujbyGZnZ7mLTuhFEEGHMSq9dIZRFEKt/7Nnz7i7YGFhAd3d3aXO3MvLC1dWVriLTqgSpFO3zaET5sjICPb29uLU1NS5t0VU0aibJqM97LKoEhRKD3Q5DxkSqs5DAfSgqamJ+94IaCHeC/pCFuQEAHt04ryJUCMMAP8CwCeyIGIkJCSE+94IIiMjSdA0AJgoCkqjUnqZnuo6mJmZwVu3bpGgHEUxhC0ArNEOfpOIjo4mMdvq/n9Ed8SYnZ3N4wySwsJCuRj8xIUo8js5GXqBUBDTwAVw6MOqJmdazrGxMf5b18rk5CTGxsbKYv4AAAsuQB10eb9KhSIiIkK6JKQrKWokaWe+KhsaGpL+6VZSUoIPHz6U7isAYIOKGE9YGz4GgEwAGAeAQ8V+7RqM9plJAMgFAFee6EWgTYt2YmovqGe6KqP5qDX7lCekiv8AcWmbuW63uuIAAAAASUVORK5CYII=",
  38: "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",
  39: "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",
  40: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUgSURBVGhD5ZlZSCVXEIYr6oiouAyYCFEiivogREJcQI3GcRcEN2JmQA1EGMQlD4pRRmeML+ooKoQJGkV8UBR9ct9nHPNgRglByDwYUTG4xAWjQXGJWqGaew7dJ96ee71ul3xQL6fqnK7/nu461X0B/sc4AMDHAPAZAIQDQNgNGF2HrucNAO+LCV0GKwDIAoBxANgGALxF2wGACQDIBQAbMVFd+BIAfqfFfH19sbCwEFtbW7G3txdHR0dxZGTk2o2u09fXh21tbfjkyRMMCAhg4hYA4CsxYTXKaWJgYCAODg7iXeLly5cYGhrKhH0PAO+JyYtUUHBOTg6enZ2J690ZioqKmKgfRAFyvmBijIHi4mIm6mtRCGENAH/4+fmJ8+40Dx48IEEbAGAvCsomta9evRLn3GlmZmbQ1NSURH0rCvrJx8dHjNeb5eVl7Orqkmx+fl50K1hfX8fJyUkcHx/HhYUF0a0zmiLxq7xA0KH1N92ThnB8fIze3t787KioqBBDJPb29vDx48doZ2fHYy0sLDAhIUH6QfTl+fPntMYJAHzEBH1Ci3Z0dIixelFQUMATJKupqRFDpMoZFhamiJObm5sbbm1tidNUGRgYYPOpq5D4nAaGhobEWJ2hW0dM7iJBTU1NihjaKSrB1tbWfCwrK0ucpsrExASbG8kEUc8kncyXYX9/H93d3d8p6Pz8HOk5Zf6MjAzua2ho4OO2tra4vb2tmKvG2NgYmxvBBFEjKLUblyEzM5Mnc+/ePa2ClpaW0MzMjPvlFXVzc1OxSz09PYq5alypIGqNWBLx8fEYExOjVdDw8LBC+OLiIvfR7nl6enJ/ZWWlYq4aVyZoZ2cHnZ2dpcVsbGxwbW0NExMTtQqSPz/29vbSrsjx9/fn/uzsbIVPjSsTlJqayhOgZInw8HCtgmpra7nP0dERd3d3Ff6QkBDuT0tLU/jUuBJBdHCyi9NtxoiIiNAqqLy8XFVQcHAw96ekpCh8ahgsaHV1FR0cHKRFXFxc8ODggPtiY2N5UlS55GgOQK2C5DuUnp6u8KlhsCD5rUaFoL29HVtaWiTz8vLivocPH0pjb968kebV19dzHz1DGxsbinXlz5A+Z5HBguQX1sXYedPd3c3HqHyLVc7Dw4P7q6qqZFdUx2BB8n5NF2OC5ubm0MTEhI+/fv2ar0m7JT+H9HlTNlhQc3MzlpSU4NOnT7k9e/YMS0tL0dXVlScVGRkp+ajXIk5PTxW3JLU9jBcvXvDx+/fv/+f5UsNgQWqoFQWirq6O+8lo9/Lz89HKyoqP5ebmitNUuVZB8kpFVU3k5ORE0U2IRs+RPn0cca2CHj16hE5OTpI1NjaKbomjoyPMy8vjpZ/M0tISk5OTcWVlRQx/JxcJMqjblnN4eCh132S0G2rQTkxNTUnF4TIvdoyLBIXSgCHvQ7fJRe9Dn9JAZ2enGGsU0EZoBIUwQY4AsF9WVibGGgWayvkPALgyQcTPQUFBYqxREB0dTYJ+AwBTuaA8OsVZ32UsvH37Fs3NzUnQd3IxhB0A/EnfuIyJuLg4EvOXtv+P6Bux1MIYA9TAaorBN6IQOT9S0F0vEDIxHaIAEXqwmiiYtnN6elpc61aZnZ3FpKQkJqYdACxEAdqgj/frVCiioqKwuroa+/v7pVOdTuabMvqASa8S9D2CGl/NZ7BNKmJiwrrwAQAUA8AvAHDM+q5bMjpnZgGgDAA+FBO9DHRo0UlM7QX1TDdldD1qzdzFhC7iX9MdFEfiAiiUAAAAAElFTkSuQmCC",
  41: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAQ2SURBVGhD5ZlZSD11FMe/7kEhCVpBiiL4JCSRC6jhvoIgCqGCkOSLlPkQmOKaL4mKC0FgivigKPokrqi5vqQSIeZLEmKgLYhBuODGifPD33Dv8Xrv/V/JO5c+8Htw5sxvzmdm7vmdGYH/MSEA3gPwIYBMABkvMPg8fL4YAG/JhFzhdQCfAvgewCkAcuM4A7AG4HMAgTJRZygB8AtPFhcXR3V1dTQyMkLT09O0tLREi4uL//ng88zMzNDo6Cg1NDRQYmKilvsVwMcyYXt8zQcmJSXR/Pw8mYmVlRVKS0vTYt8A8JLJS9o5uLq6mu7v7+V8pqG+vl5LfSsFLPlIy3gCjY2NWuoTKcK8AeC3+Ph4eZypSU9PZ6E/AQRJoc/YdnV1VR5janZ2dsjHx4elvpRCm7GxsTL+lTk6OqLJyUk1Dg4O5G6bbG5u0sTEhCpAd3d3crdDHorET5YFghetf/iZfA7X19cUExNjrB3t7e0y5BGHh4fk5+en4oOCguj8/FyGOKSjo4OPvwEQroXe5wnHx8dl7CtRW1tryPDo7u6WIVbwBUhNTTXiQ0ND6eLiQoY5ZG5uTs/BXYUilTcsLCzIWKfZ2NiwknEkxHcmIyPDKj4sLMwlobW1NT1HthbinkmtzK7Aj0lUVJRTQqenp2p7cHDwo3hXhZaXl/UcWVqIG0HVbrhCVVWVkZT+PTwlVFNTYyURHh5OXl5e5hHiyqSTKywspLy8PLtCFRUVal9AQAD19fWpaqjj3S50dnamkuBjAwMD6eTkhIqKiuwKceEoKSmhvb099bephMrLy41kBgcH1bbMzEy7QlzZLBkbGzOHkOWV5cdMk5WVZVdIYgqh4+NjCgkJUZNERERYJZGfn28k2N/fb3WcLUwhZPmocSHgpIaHh9WIjo429pWWlqptW1tbcgoDUwglJCQYSTgzKisr5RQGphCy7NecGaYXGhoaoqamJmpubjZGS0sLtba2UmRkpJFgdna22se91lOYQsgeHlkU7JGSkmIkyG29I/hLko7nymk6obKyMvUawGNgYEDufsTU1JSK5bvDxeby8lKGOMSW0LO6bUuurq5U983j5uZG7n7E7e2tEe+KDGNLKI03POd9yJ3Yeh/6gDfwe70nwjfiQShFC70D4LytrU3GegS9vb0scwsgUgsxPyQnJ8tYjyA3N5eFfgbgYyn0hbe3t92+y4zs7++Tv78/C31lKcO8CeAP/sblSRQUFLDM30/9/4i/EasWxhPo7OzUxaBGiljyHQeZvUBYyIxLAQn/sAY5mG/n9va2nMut7O7uUnFxsZYZA/CaFHgK/nj/OxeKnJwc6urqotnZWVpfX1cr80sN/oDJX5Z6enpU4+vr68sif3ERkwk7w9sAGgH8COD64aq4a/A6swugDcC7MlFX4EWLV2JuL7hneqnB5+PWLEomZIt/AfiopMuywOIYAAAAAElFTkSuQmCC",
  42: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAATsSURBVGhD5Zl7KL5nGMcvP4eWLY2yrUYTOYRozaGwjDmXcshCKTX/YCYpI8f5Z0IoUkaiiPjT+TCnkaFZNMWYImzWmBY5xbWuO8/Tc9+/9329r9ePV/vU9c99Xdf9XN/3ud339dwA/sdYA4AHAHwKACEA8PkTGD2HnucJAO+JBT2EtwEgEwB+AIC/AQCf0Y4BYAoAvgYAC7FQbUgEgN9oMm9vb8zPz8eOjg7s6+vDsbExHB0dfeNGz+nv78fOzk4sLCxEPz8/SdzvAJAqFqyJ7yjR398fh4aG0JCYmJjAoKAgSVg9ABiJxYtUUHBWVhbe3NyI8xkMBQUFkqhGUYCSLyQxL4GioiJJ1JeiEOIdANj18fER8wya4OBgEnQIAJaioK9I7eTkpJhj0CwtLaGxsTGJ+kYU9KOXl5cYrzM7OzvY29vLbHNzU3RzUOzU1BTb0ZaXl/Hi4kIM0Yq7TeIX5QZBh9a/tCb14fLyEj09PeWzo6KiQgxhzM3NYXh4OJqbmyvPGXR2dsb6+nox/F4qKysp/woAPpIEfUwTdnd3i7E6kZeXxxVYU1MjhrC3YWpqysWJlpubK6ZpZHBwUMqlroLxGQ0MDw+LsVozMzPzWmGiIFpSjo6Ost/S0pIVX1paiq6urlzu7Owsl6sJWrZ3eWGSIOqZ2Mn8EE5PT7lC1QmiQ1ryGRkZcRvQ0dER2trayv6cnBwuVxPj4+NSXqgkiBpBthweQnp6ulyIcjmJgqiFcXBwQCsrK/Tw8OB8REJCgpybmJgoutXyqIKUv3pMTAxGRkaqFXR7e4vX19d4eHiIBwcHnI8ICAiQc7Ozs0W3Wh5N0PHxsbxMLCwsWJFxcXFqBWlifn4eTUxM5FxqgrXl0QSlpKTIBbS0tLCxkJAQnQXt7u5yf4O0HOlNasujCKKDUyqAlplEaGioToJIjIuLi5xDpssOR+gtaH9/H62trdkkdnZ2eHZ2JvuioqLkwpqamrg8ka2tLXRycuLENDY2imH3orcg5VKjjaCrqwvb2tqYubm5yb6kpCQ2trCwIE6BGxsb7MdQimloaBDDtEJvQb6+vlwh91laWhqXv729zYl59eoVtre3czG6oLcgZb+mjSkFnZycoLu7u+yj3XFkZISbX1f0FtTa2orFxcVYUlIiG7UvZWVlaG9vLxcbFhbGfNRrSaSmpnJiqSPY29tjS3B9fV02VeeUOvQWpAlNm8Lq6ionhow6bmqDxPFn6xREAgMD5aKorVeSkZHxWuHqjDYbbXmjgpKTk9HGxoZZc3Mz54uNjWXj1F1oMorJzMzkcjWhSpBe3baS8/Nz1n2TXV1dcT5pnM4tTUYxuny9qhIURAP6fA89J6q+hz6hgZ6eHjH2RUAv4k5QoCToAwA4LS8vF2NfBHV1dSTmGgDsJUHET/Q98hKJiIggQb8CgLFSUC61IKr6LkNmbW0NzczMSNC3SjHEuwDwJ91xvSSio6NJzD/q/n9Ed8SshXkJVFVVSZtBtihEyfcUZOgbhEJMtyhAhP6wWiiYXufi4qI417OysrKC8fHxkpguAHhLFKAOurz/gzYKuratrq7GgYEBnJ6eZifzUxldYNLNUm1tLWt87y5S/qJNTCxYG94HgCIA+BkALu9+lecyOmdWAKAcAD4UC30IdGjRSUztBfVMT2X0PGrNHMWCVPEfpwUxxo77HwwAAAAASUVORK5CYII=",
  43: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUdSURBVGhD5ZlXSHRXEMfHXtHkQY1EiQo+qBAJsYAajcYuwofKhwUxEF8kMYqCvfsQsUMgYCz4oNjge7GixooSGypEhIgVNTGIgaiIik6Yy3cPd0/cdde1rOQH83Jm5tz5u3fPzFkB/sdYAMCnAPAFAAQCwFfPYPQcep4rAFjyBT0EEwD4FgB+AYBjAMAXtBMAmACA7wHAjC9UGWIB4HfazN3dHXNycrCtrQ17e3txZGQEh4eHn9zoOX19fdje3o75+fno5eUlitsEgK/5ghXxAyV6e3vj4OAgahJjY2Po7+8vCvsRALT44nkqKDg1NRVvbm74/TSG3NxcUdRPvAApb0Uxr4GCggJR1De8EMIUAPY8PDz4PI0mICCABB0BwIe8oO9I7fj4OJ+j0SwsLKCOjg6JyuYFTbu5ufHxKrO7u4s9PT2CbWxs8G4ZyE9fcvoj7uzs8G6leX9ILEsPCGpa/9A7qQ6Xl5fo6urKekdFRQUfIrC4uIghISFoYGDAYo2NjTEmJga3trb48HuprKykPa4A4BNR0Ge0aWdnJx+rEllZWaxAstraWj4El5aW0NTUVCZOara2tri9vc2nKWRgYEDMp6lC4EtaGBoa4mOVZmpq6j/F8YKoDVBvkxZfWlqKmZmZaGRkxNbfvHkjk3cfExMTYm6wKIhmJqEzP4SzszN0dHS8V9DKyoqMf21tjflqamrYupmZGR4fH8vkKmJ0dFTMDRIF0SAojBsPISUlhRWjp6cnV9DR0ZEwPmVnZ2NhYaGMjw4GMc/ExESIVZZHFUSjkVgIvSphYWFyBSkiKSmJ5fn6+uLt7S0fIpdHE3RyciJ8DyiXXpPDw0OMiopSWtDFxQWWl5dLZzO0tLTE5eVlPlQhjyYoMTGRFdLU1CSsBQYGKi1ofX2dxZJZWVnh7OwsH3YvjyKIGqdYCL1mIkFBQUoLoteV+o94jOvq6qKzszO+e/eOD1WI2oIODg7QwsJC2MTOzg7Pz8+ZLzw8nAlqaGiQyeOhk2xvbw/39/cxLy+P5WlpaeH8/DwfLhe1BUlfNToIOjo6sLW1VTAXFxfmi4uLE9bm5ub4Le5EcoHDhIQE3i0XtQV5enqyBytjycnJ/BZ4fX3NL2F6ejrLUWWuVFuQdF5TxkRBXV1dGBERgU5OThgbG8tvK3yiYo6Pjw/vlovaglpaWoTGWFRUxKy4uBhLSkrQwcGBFRUcHCz4aNYimpubmY9G/pmZGbYnTQ3m5ubMn5GRIXmiYtQWpAhFh8Lp6Sna29szPwlIS0sTJgdra2u2TnPdfVcPKU8qyM/PjxVGYz3P9PS0wmmbjm5qCarwpILi4+PRxsZGsMbGRt4tQJNAZGSkMLOJQgwNDYWmTFO7qtwlSK1pWwqNMzR9k11dXfFuGeh2SwJo/N/c3OTdSnOXIH9aUOc+9JLcdR/6nBa6u7v52FcBfRDvBfmJgj4CgLOysjI+9lVQX19PYq4BwEEURPyqSjPTJEJDQ0nQbwCgIxWUqa2trfTcpSlQQ9bX1ydBpVIxxAcA8Cddtl4TdPwDwN/y/n9EvxELI8xroKqqSjwM0nghUn6mIE0/ICRiOnkBPPTFaqJg+jhVuWg9B6urqxgdHS2K6QAAQ16APOjH+z/ooKCfbaurq7G/vx8nJyeFzvxcRpMEXdXr6uqEwZdmPQD4iw4xvmBlsAKAAgBYAoBLce56IaM+swoAZQDwMV/oQ6CmRZ2YxguamZ7L6Hk0mjnyBd3Fv7z9SjIAt7ETAAAAAElFTkSuQmCC",
  44: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAASjSURBVGhD5ZlZSH1VFMa/nKVQA60gxQF8UVAiJ9QwzVkEUYgUhKAHEdMeBFMcUyFRcSAITBHBEQVfHFFz6iXHFPLBJNBAKwmTcA7dsTbuw7nb6/VcL3/vvfSD9bL3Wvusz3v22utsgf8xHgCCAHwAIB7ARy9g9Bx6XjCAt+SEnsPrAAoAfA/gLwDMjHYCYBFAEQAXOVEtfALgF1osNDSUlZaWsr6+PjY2NsZmZ2fZzMzMKzd6zvj4OOvv72fl5eUsMjJSiPsVwKdywob4mgKjoqLY1NQUsyTm5+dZbGysEPYNgNfk5GUayLmwsJDd3t7K61kMZWVlQtS3sgA1Hwsx1kBFRYUQ9ZkshHgDwG9hYWFynEUTFxdHgv4E8KYs6HNSu7CwIMdYNGtra8zW1pZEfSkL+iEkJET2N5qDgwM2MjLCbW9vT55+lJWVFR4zOjrKTk9P5WmD3BeJn9QFgg6tf+idNIXr62sWHBysnB0NDQ2yi152dnaYo6OjErexsSG7GKSxsZHibgB4C0Hv0UJDQ0Oyr1GUlJQoSZG1tLTILg+gShoREaETt7W1JbsZZHJyUsRSV8H5kAamp6dlX80sLy/rJKVVUG1t7YM4YwUtLi6K2EQhiHomfjI/h7OzM+bv7/8gsacEra+vMzs7uwdxxgqam5sTsQlCEDWCvN14Dvn5+Uoy9vb2mgTRfgsKClJ81cLMKohaI5FIRkYGS0lJ0SRIvd8KCgqYt7e3+QWdnJwwLy8vvpiLiws7OjpimZmZTwpS77eYmBi+jru7u/kF5ebmKkl0dXXxsfj4eIOC1PvNwcGB7e7usouLC/4HMasgOgRFAvSaCRISEgwKUu83cU7RL+vq6mo+QYeHh8zDw4Mv4uPjw87Pz5W51NRUJbGOjg6dONV5wdLS0pRxKhBubm7K3P7+vk7cU5gsSP2qUSEYHBxkPT093AIDA5W57OxsPkYn/9XVFfP19VXm6IOxt7eXz7e3tzNnZ2dlrq6ujo9rFWayoPDwcOXhWqyoqIjvHXVJ12IDAwPyo/VisiB1v6bFhCB5/Cl7MUHd3d2ssrKSVVVVKVZdXc1qamqYn5+fklBiYiKfow7k5uaG1dfX68SIuOLiYubk5KTE5eXl8TlqWrVgsiBDGCoKj0H7y6xFwRB0UIrEqK3XApVt9efD5uam7GKQVyooJyeHeXp6cuvs7JSn9XJ8fMwCAgJ410Gm9VUT6BNkUret5vLykhcAMto3Wri7u+NnmYgz9rZJn6BYGjDle8ic6Pseep8GhoeHZV+rgH6Ie0ExQtA7AM7o69EaaWtrIzH/AvATgogfo6OjZV+rIDk5mQT9DMBWLajYxsaGXydZE1QR6TMEwFdqMYQbgD/ojsuaSE9PJzF/P/b/I7oj5i2MNdDU1CSKwReyEDXfkZOlFwiVmCFZgAxtrC5ypp9zdXVVXsusbG9vs6ysLCFmEICTLOAx6PL+dyoUSUlJrLm5mU1MTLClpSV+Mr+U0YUK3Sy1trbyxvf+uuuYipicsBbeBlABYAPA9f1fxVxG58w2gFoA78qJPgc6tOgkpvaCeqaXMnoetWb+ckL6+A9cGzmCkbi3IQAAAABJRU5ErkJggg==",
  45: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUgSURBVGhD5ZlpSF1XEMenrtUWtaitUKWi8ZNQKXU34lI1LgghQlBRKCiItLZKxSpqtIKpqJhAoWDVEHFF880V9wWlKlKEqlgXsKC2RixWTVzQKXPxXK6T915efBqf9AcDjzsz58yfe+85c+4D+B9jCwCfAoA/AIQAwBdvwWgems8NAD7kBV2E9wDgKwDoB4AtAMBrtG0AGAKAbwDAgheqDbEA8AcN5uHhgdnZ2VhfX49tbW3Y29uLPT09V240T3t7OzY0NGBubi76+voKccsA8CUvWBM/UqKfnx92dXWhPjEwMIBBQUFC2E8A8A4vnlNCwWlpaXhycsLH0xtycnKEqJ+5ACX3hZibQF5enhCVxIUQ7wPAn56enjxPrwkODiZBfwPAB1zQ16R2cHCQ5+g1U1NTaGhoSKK+54JG3d3defwbs7q6iq2trZItLi5yt8Tx8TE+f/4ct7a2VNrm5iYeHBzwNLWcLRK/KRcI2rT+pWdSFw4PD9HNzU3eO0pKSniIRE1NDVpbW6ONjY1KI19jYyNPU0tpaSnNdwQAnwhBn1EBzc3NPPaNyMrKksWQVVRU8BCJpKSkc3GqrLKykqeppbOzU+RRVyERSBe6u7t5rNaMjIy8UpQ6QT4+PnKMiYkJmpubo5mZmWympqZYW1vL09QyNDQkxgsTgqhnknbmi7C3t4cuLi5aCdrZ2UFbW1tZzOjoqPTeKW1lZQV3d3d5qlr6+vrEnKFCEDWCUrtxEVJTU2URxsbGGgVNT0/L/lu3bnH3hbhUQdQaiQLv3r2LERERGgXV1dXJftpD6KkoLi7Ghw8fXmh+4tIEbW9vo4ODgzSYhYUFrq+v47179zQKyszMlP30rojfwvz9/XFpaYmnaeTSBCUmJsqFVFdXS9dCQkI0CgoMDHxFBDdnZ2dpL9KWSxFEG6cogB4zQWhoqFpBtE+5urrK/oSEBBwbG8O5uTksKipCAwMD2Zeenn4uVxM6C1pbW5NXKkdHR9zf35d9kZGRclF8Lzk9PZVyJyYmVM6VkpIi59Kj/OLFCx6iEp0FKR81Wgiamprw6dOnkinvQFxcnHSNBGgDHeZELt0tda0TR2dBXl5e8sTaWHJyMh9CJf39/efy5ufneYhKdBak7Ne0MSFoZmYGMzIypDsXExODR0dH58al7kDkWFlZSY2qNugs6MmTJ5ifn48PHjyQraCgAAsLC9HJyUkuKiwsTPJRr0XQ0UQpVKyMBL1fdOQXvqioKMWMmtFZkCY0LQp0pPf29pb91FXQHaO4s4OabOPj4+dyNXGlggICAuSiqK3nzM7Oop2d3bniufHl/nVcqaD4+Hi0t7eXrKqqirsllpeXMTY2Fi0tLWURdOqkz2XPnj3j4a9FlSCdum0lL1++lLpvMv7SczY2NqRHi96thYUF7tYaVYKC6IIu56HrRNV56HO60NLSwmNvBHQjzgQFCEF2ALBH/dRN5PHjxyTmGACchCDi19u3b/PYG0F4eDgJ+h0ADJWCvqP+Sdu+S1+gbYCO8gDwg1IMYQUAf9E3rptEdHQ0iflH3f9H9I1YamFuAmVlZWIx+JYLUfILBen7AqEQ08wFcOjFqqZgup2Tk5N8rGuFOnbq1M/ENAHAu1yAOujj/QYtFHfu3MHy8nLs6OjA4eFhaWd+W0YfMOnL0qNHj6TG18jIiIRs0iLGC9aGjwAgDwCmAeBQ9F3XZLTPzABAEQB8zAu9CLRp0U5M7QX1TG/LaD5qzVx4Qar4D1+XJ4X0xQKqAAAAAElFTkSuQmCC",
  46: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAU1SURBVGhD5ZlrLGVXFMeXZ0obIaFtUqNImIQYqXqFETX1lkkmiHhE0gQfaKcNEozHUF9qHkGUihIR7/DR+/1oSBFKEEnpB01GtUGjIYYOq1knzs45e9zrXtfrpr9kJTd77bXP+t9z7tprnwvwP8YCAO4BgA8A+APA59dgdB26njMAvM8ndBHeBYAvAWAIALYAAG/QdgBgFAC+BgATPlFViAaAX2kxNzc3zMrKwsbGRuzo6MCBgQHs7++/cqPrdHZ2YlNTE+bk5KCXl5co7jcA+IJPWBnfUaC3tzf29PTgbWJ4eBj9/PxEYd8DgA6fPE8RTX78+DEeHx/z690anjx5Ior6gRcgJUoUow3k5uaKohJ4IcR7APC7u7s7H3erefDgAQn6EwDMeEFfkdqRkRE+5lYzMzODenp6JCqTF/STq6srP19t1tfXsb29XbDV1VXe/RbLy8s4ODiIo6OjuLm5ybtV4rRI/CItELRp/UPPpCYcHh6is7Mz2zuKior4KQwq/x4eHqirq8vmm5qaYmpqqrCOOjx//pzijwDgY1HQJ7Rga2srP1ctMjIyWHJkxcXF/BSByspK2Tze4uLi+BCldHd3i7HUVQh8RgO9vb38XJUZHx9/K7GzBM3NzaGOjg6b4+Pjg8+ePcPIyEhZbF9fHx+qEHpcT+MCRUHUMwk780XY29tDOzs7lQRFRUUxv6enJ75584b57t+/z3zp6emyOGXQb/A0LkAURI2g0G5chOTkZJaIgYGBQkE7OztoZmbG/A0NDTL/2tqa8KWurKzg1taWzKeMSxVErZGY4KNHjzAkJEShoMnJSeajYvDq1Ss8OTnBpaUlXFxclM1Vh0sTRN/4nTt3hMVMTExwY2MDw8PDFQqi5lb0WVpaYnNzM7q4uLAxGxsbfPHihSxGFS5NUHx8PEumpqZGGPP391coqKKigvmMjIzYZ95SUlJkcedxKYJo4xQToMdMJCAgQKGg0/1CZmlpacIRITMzUzZO+5SqaCyInn0LCwthEWtra9zf32e+0NBQllRVVZUsrqysTJZ0QkKCzJ+UlMR8QUFBMp8yNBYkfdSoELS0tGBdXZ1gjo6OzBcTEyOMTU1NCXH0WSqINkQp0gJDv82DgwOZXxEaC6KWRZrYeZaYmCjEUeMrHR8aGpKtOzExwXzm5ua4vb0t8ytCY0HSfk0VEwVR82lsbMzGy8vLZet2dXUxn5WVFb5+/VrmV4TGgmprazEvLw+fPn3KLD8/HwsKCtDW1pYlFRgYKPikjxaNif67d+/i7u4u80VHRzPfw4cP2fh5aCxIGcqKAkGdgOgnc3JywuzsbAwLC5ONq9PLXakgX19flhSV6bOguyZNnjd1+jjiSgXFxsYKXQBZdXU172bU19ejg4ODTIi9vb2w+arLWYI06ralUKml7pvs6OiId8sg//z8vFDtZmdnVS4CPGcJ8qMBTc5DN8lZ56FPaaCtrY2fqxXQjTgV5CsK+hAA9goLC/m5WkFpaSmJ+RcAbEVBxM90YtRGgoODSdASAOhJBaXToUvsu7QFeg1maGhIgr6ViiFMAWCT3nFpE9RRAMDfiv4/onfEQgujDdDp9rQYfMMLkfIjTbrtBUIippUXwEM/rBqaTLdzenqaX+tGWVhYwIiICFFMCwC8wwtQBL28/4MKBZ0eX758KbT3Y2Njws58XUYvMOngV1JSIjS++vr6JOQvKmJ8wqrwAQDkAsAsABxK+64bMNpnFgCgEAA+4hO9CLRp0U5M7QX1TNdldD1qzez4hM7iP8nPCfFwSDy5AAAAAElFTkSuQmCC",
  47: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAATASURBVGhD5ZltKH5nGMAv72NLE7Zp3iJSimZeCsuYdylRa5Sa5gNttg8r8858mZCXxmIkH4j46P3l720+DGmpocYUhQ0Z8j661nXnnM6593h2HgePp/3q+nLu677P9Xuex31f5wD4H2MPAD4A8AEARALAR88QdB+6ny8AvMUX9BBeB4DPAeAVABwAAOoxDgFgCgC+BABrvlAlfAIAv9FiAQEBmJeXhx0dHdjX14djY2M4Ojr65EH36e/vx87OTiwsLMTg4GBB7ncA+JQvWBvf0cSQkBAcGhrCl8TExASGh4cLYt8DgBFfPE8FJefk5ODt7S2/3oshPz9fkPqBF5DysSBjCBQVFQlSn/EixBsAsBUYGMjPe9FERESQ0J8AYMMLfUG2k5OT/JwXzcLCApqYmJDUN7zQT/7+/ny+zmxubmJvby+LtbU12dj5+Tnu7e3hwcGB1tjf38ejoyPZXG3cbRK/SDcIOrRO6DephqurK/T19RXPjoqKCtl4fX092traop2dndagnISEBNlcbVRWVtL9rgHARRB6jwro7u7mc3UiNzdXlKGoqamRjZeXl8vGtYWPj49srjYGBweFedRVMD6kC8PDw3yuYmZmZv5VFC9UXV2NFhYWaGlpKQsrKyu0sbFBIyMjcW5GRoZsrjampqaEedGCEPVM7GR+CKenp+jh4fGfQsfHx7ixscH+zqSxu7uLjY2NaGxszOb5+fmxNZUyPj4u3DNKEKJGkLUbDyE7O1uUMDMzu1foPk5OTtDFxYXNoW9sdXWVT9HKowpRayQIJCUlYVxcnM5C6enp4hx+I1HCowkdHh6ik5MTW8za2hp3dnYwOTlZJ6HZ2Vkx393dne2UuvJoQtJPtrW1lV2LjIzUSSg6OlrMb2pq4ocV8ShCdHAKhdDPTCAqKkqx0OLiophLf0NnZ2d8iiJUC21vb6O9vT1bxNXVVVZIfHy8WGRzc7NsHk9WVpaYW1BQwA8rRrWQ9KdGG0FXVxe2t7ez8Pb2FsdSU1PZtbm5OX4J9iE4ODiIuZpylKJaKCgoSCxESWRmZvJL4MjIiDju5eWFNzc3fIpiVAtJ+zUloUmIHumFcV26Ak2oFmpra8Pi4mIsKSkRo7S0FMvKytDNzU0slHYwGqNei0fyGI0NDQ38sE6oFtKGkk3h8vJS7Awopqen+RSdeFKhsLAwsVBq6zWxtbUl9m0UKysrfIpOPKlQWloaOjo6smhpaeGHGevr6+js7My6DE9PT9ZhqEGTkKpuW8rFxQXrlCmur6/5YQbtaEIOPcWqRZNQOF1Q8zykTzQ9D71PF3p6evhcg4C+iDuhMEHoHQA4pUdkQ6Suro5k/gYAN0GI+Dk0NJTPNQhiY2NJ6FcAMJEKfU1bqZqeSh8sLy+jubk5CX0rlSHeBIA/6AQ3JBITE0nmr/v+f0TviFkLYwhUVVUJm8FXvIiUHynppW8QEpluXoCH/rBaKZm+zvn5eX4tvbK0tIQpKSmCTBcAvMYL3Ae9vN+ljSImJoa9JBwYGGCNJJ3MzxX0ApPeLNXW1rLG19TUlET2aBPjC1bC2wBQBACLAHB196noK+icWQKAcgB4ly/0IdChRScxtRfUMz1X0P2oNfPgC9LEP95Ti95WUm1EAAAAAElFTkSuQmCC",
  48: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAU/SURBVGhD5ZlZSB5XFMePcV9Qq9iWNFIR9EVwa1RwNy7RgCIKpQpCoaJItX1oYhV3fagmwSQUCi5oEEVRfHLFfQFxqyLqi9oH61It1aIook1yyrk4l5lbnXz6afykPzgv95xz7/kzM/eemQH4H2MHAK4AEAAAYQAQ+gGM1qH13ADgY7Ggy2AOAN8CQD8A/AUAeIO2CwBDAPAdAFiKhWrCVwCwRJN5eXlhVlYW1tfXY1tbG/b29mJPT8+1G63T3t6ODQ0NmJOTg76+vpK43wDga7FgNX6iRD8/P+zq6kJdYmBgAENCQiRhPwOAnli8SCkFZ2Rk4Nu3b8X5dIbs7GxJ1C+iADlfSmJuA7m5uZKob0QhhAUA/O7t7S3m6TQPHjwgQdsA8JEoKJ3UDg4Oijk6zdTUFOrr65OoH0VBo/fv3xfjL8zq6iq2tLQwW15eFt0KlpaWsL+/n+1mCwsL+O7dOzFEI043iVn5BkGH1j7dk9pwfHyMbm5u/OwoLS0VQxijo6MYGBiIhoaGPFZPTw89PDywtbVVDH8vT58+pTlOAOBzSZAHTdrU1CTGXojMzExeIFl5ebkYghMTE2hiYqKIE625uVlMU6Wzs1PKpa6CEUwD3d3dYqzGjIyM/KewswQFBwdz/927d/HFixdYWVmJrq6ufNze3h4PDg7E1HMZGhqSciMkQdQzsXv5MtDiTk5O7xW0vb2NFhYWzEe3WF9fH/etra2htbU1zx0fH1fkqkHznOaFS4KoEWTtxmVIS0vjhcifC1HQ5uYmmpqaMp+xsTHu7Owo/M7Ozjx3bGxM4VPjSgVRayQVERsbi1FRUecKevPmDXp6enJ/bW0t95EA6dmysbHB3d1dRa4aVyaIFqX7nXItLS3ZFYiLiztXEEG3tZmZGfMbGRlhUlISpqamMhFSXnV1tZimypUJomLEIsLCwlQFEbRt29ra8jjJDAwMLiyGuBJBdHBKhdBtJhEeHq4qiB5+Ek3Fi4Joo3B3d79wt6K1oI2NDbSzs2OTODg44OHhIfc9evSIF1hRUaHIo+49ICCA+0NDQ9m5NDMzgykpKXzc3NwcV1ZWFLlqaC1IfqvRRtDY2IivX79m5uLiwn0JCQlsjIomZOcFWllZ4dbWlmLeoKAg7n/8+LHCp4bWgnx8fPjCmlhycjLLKysr42PU+oiUlJRwP109TdFakLxf08QkQcXFxXyMnhWRJ0+ecD81nJqitaCamhrMy8vD/Px8bgUFBVhYWIiOjo68qIiICOajXouQbyS0KcjX29/fV3Qb6enpshXV0VqQGmqbwt7eHt67d4/76Tmiq/bq1St2xaRxsunpaUWuGtcqSP5gU1svQh85aBeTFy9aUVGRmKbKtQpKTExkV4GsqqpKdDNmZ2cxJiaGN6qSUcddV1cnhr+XswRp1W3LOTo6Yt032cnJiehWsL6+zl47aN35+flLf2U6S1AIDWjzPnSTnPU+9AUNXPRNUVegC3EqKEgS9CkAHNCOcxt5+fIlifkHABwlQcS4v7+/GHsriIyMJEELAKAvF/TDnTt3eN91W1hcXGTvVQBQJBdDWAPA1kVaDl0gOjqaxPx93v8j+kbMWpjbwLNnz6TN4HtRiJxKCtL1DUImpkkUIEIPVjUF0+WcnJwU57pR5ubmMD4+XhLTCAAmooDzoI/3f9BG8fDhQ3z+/Dl2dHTg8PAwO5k/lFEnQV+W6IMkNb6nr+5/0iYmFqwJnwBALgD8CgDH8r7rBozOmTkAKAaAz8RCLwMdWnQSU3tBPdOHMlqPWjMnsaCz+BfWsQNz+hZRPwAAAABJRU5ErkJggg==",
  49: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUhSURBVGhD5Zl7KLdnGMcvZ2XJcaZhjv+gac2hsIw5h0SWY63mD9psSRk5+2dCqLEYyR9ERMn5MKelhrTUJNOSFXNYaE6h17WuJ8/d89yv38+Pn2P71FVvz3Vd931939/zXvf13C/A/xhzAPgQAD4BgEAA+OwJjPah/dwA4F2+oPtgAABfAcDPAPAPAOAz2j4ATAHANwBgyBeqCvEA8Act5uHhgTk5Odja2op9fX04NjaGo6Ojj260T39/P7a1tWFeXh56e3uL4v4EgC/4gpXxPSX6+Pjg0NAQviQmJibQ399fFPYDAGjwxfOUUXBGRga+efOGX+/FkJubK4r6kRcg5XNRzGsgPz9fFPUlL4R4BwD+8vT05PNeNAEBASRoBwCMeUFfk9rJyUk+50WzsLCAWlpaJOo7XtAv7u7ufPyd2djYwK6uLsHW1tZ4t4zl5WWhk83OzuLe3h7vVpnrJvGbtEHQofUvvZPqcH5+jm5ubuzsKCsr40MESAS1YE1NTRZrZmaGmZmZeHp6yoffSnl5Oa1xAQAfiII+okU7Ojr42DuRnZ3NCiSrqqriQ7Cnpwc1NDRkcVLz8/PDk5MTPk0pg4ODYj5NFQKf0oPh4WE+VmVmZmbeKo4XtLu7i8bGxszv6uqKxcXFGBcXJ8vLysqS5d3G1NSUmBssCqKZSXgV7sPx8TE6OTndKqiuro75bG1t8fDwkPlSU1OZz8DAAHd2dmS5yhgfHxdzg0RBNAgK48Z9SE9PZ8Xo6OgoFJScnMx8aWlpMt/q6ipqa2sz/11e/wcVRKORWER0dDSGhYUpFBQZGcl8fAM6OjpCS0tL5qfZUVUeTND+/j5aW1sLixkaGuLW1hbGxMQoFJSUlMR8/C9Er5+JiQnzx8fHy/zKeDBBKSkprICmpibhWWBgoEJB1dXVzOfo6IiXl5fM19nZyXxkERERslxlPIggOjjFzek1EwkKClIoaHt7G42MjGR5vb29WFtbK+t+ZLSOqqgtaHNzE83NzYVFqFtJz43w8HBWVENDgyyP6O7uVngOUXcT/xwVFcWnKkRtQdJXjRpBe3s7trS0CObi4sJ8CQkJwrO5uTlZPu1D31p6enpCnI2NjdDSSYSYSx1RVdQW5OXlJfubvc3ojLmJ9fV1XFlZETocIR2bSkpK+HCFqC1IurEqJgqiwum8mZ6efmuqpw6pr6/PcgYGBmR+ZagtqLm5GQsKCrCwsJBZUVGRMMbY29uzooKDgwUfzVqEZGNh5KfJXEQ6B1pYWAjTh6qoLUgZyprCwcEBmpqaMr+zszOWlpbKJggyau934VEF0bQsFkZjPU9jY6OseN7o8L26uuLTlPKoghITE9HKykowKv4m6uvr0c7OTibEwcEBKysr+VCVuEmQWtO2lLOzM+H9J7u4uODdDPJTOx8ZGcHFxcV7fdiJ3CTInx6o8z30nNz0PfQxPaB56jVCP8S1ID9R0HsAcEwd5zVSU1NDYi4BwF4URPzq6+vLx74KQkNDSdDvAKAlFZRFtzD83PXSoaswXV1dElQiFUMYAcA23XG9Jq6/gg8U/f8R3RELI8xroKKiQmwG3/JCpPxEQS+9QUjEdPACeOgfVhMF0885Pz/Pr/WsLC0tYWxsrCimHQD0eQGKoMv7v6lRhISECGMJjfQ0+tPJ/FRGF5h0s0QDKw2+11ddu9TE+IJVwQIA8gFgEQDOpbPXMxidM0sAUAoA7/OF3gc6tOgkpvGCZqanMtqPRjMnvqCb+A8JSQlPGVRePwAAAABJRU5ErkJggg==",
  50: "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",
  51: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAASiSURBVGhD5ZlZSDZlFMf/7qkRiVphiqB0W+Tnghu5VIrglZh640J49WmC4YYK6k2i4kLQReaVu16JK+7LTW7kkghKit5kLqm553LiDN8MM4+vr0uk89IPDvLOc55nzp9n5pwzj8D/GGcAHwMIBvA5gPBnML4P3+8TAO+JAT0FewCvAQwB2AVAL2h/AhgF8A2Ad8RAH0I8gBVezMfHh3Jzc6mhoYE6OztpYGCA+vv7/3Pj+3R1dVFjYyPl5+dTQECALO43AMliwMb4jicGBgZSb28v6Ynh4WEKDQ2VhX0PwEwMXqSUndPT0+n6+lpcTzfk5eXJon4QBaj5ShZjChQUFMiivhaFMG8D2PT19RXn6ZqwsDAW9AcAB1FQGqsdGRkR5+ia6elpsrCwYFE5oqAJb29v0f9ednd3aWdnR/or2vb2Nh0dHYlTbjExMUFtbW1SArq6uhKH7+VNkvhFnSC4aP3Fz+RjWFhYIGdnZ3JycjJojo6OlJaWJk7TsL6+TlZWVtK74ODgQMfHx6LLvZSVlfH8vwG4y4I+5QVbWlpEX6PU19erC59Bi4uLE6cpXFxcUEhIiOLr6upKJycnotu99PT0yGtwVyERwhf6+vpEX6NkZWUpwfBzbGdnR7a2torZ2NhQamqqOE2CdyY8PFwj3s3N7UmCRkdH5TW+lAVxzyRV5scQERGhBFNTU0Obm5u0sbGh2NramvQuqeHflZWV0iOpFvNvBA0ODsprfCEL4kZQajceyvn5Obm7uyvBcPAPISMjQyOC1zAzM3t5QSsrK2RpaSktxImBn+OqqioqLi6m1tbWO7NbSkqKNIcfR97V9vZ2fexQR0eHEoicpdTm4eEhiRTJzs6m+Ph4WlxclH7rRlBhYeEtEaLxDnKNUcOZTU1zc7M+BCUnJyuB+Pv7S23+8vKy1OqrX3gvLy+jxVI3gjhbzc/PS4/V6empZqypqUmzU7Ozs5pxNboRZIytrS2yt7dXAuUCfBcmIWhvb09qe+RA6+rqRBcFXQjigPkDKykpiSIjI2lpaUkzvrq6qqT0+9bVhaD9/X1NwImJiZrxzMxMZYwTxOHhoWZcjS4EMTk5OUogbLGxsVRbWyv1burr3AkbQzeCuM0PDg7WBC9aQkIC3dzciFM18EmS7M8dx4sJYvjmXPldXFw0Qjw9Pam0tFR0Nwh3HPzZwLvj5+d3qwQ8BEOCntRtyxwcHNDMzAwNDQ3R3NwcnZ2diS53cnl5Ke0221PEMIYEhfKFx34P6QVD30Ov+AJ/15sivBFvBH0mC/oAwHFJSYnoaxJUV1ezmEsAHrIg5uegoCDR1yTg4g7gVwAWakHfmpub0+TkpOiva7hLsba2ZkHFajHMuwC2+IzLlIiOjmYx+3f9/4jPiKmoqEicp0vKy8vlZJAhClHzIzvpPUGoxLSIAkT4xfqJnXk7p6amxLVeFP6ojImJkcU0A3hLFHAXfHj/OycKPn+rqKig7u5uGhsbkyrzc9n4+Lh05s0nSlFRUXKXv81JTAz4IbwPoADALIALdb/2AsZ1Zh5ACYAPxUCfAhctrsTcXnDP9FzG9+PW7CMxIEP8A45qkdV8YtnmAAAAAElFTkSuQmCC",
  52: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAViSURBVGhD5Zl5LGV3FMcPZumMRmrQNmpCTDAR0VSHZCxp6UIkEiLK/DMjafxjqSC2IJZ/KkwyhPijOn9MLMV/wox9JzoMRSrWSoiotai1tZzm3Pjd3Pvz3vPeeMOTfpITyT3n3Hu+9/3u+Z17AfyPMQMARwDwAICvAeCrCzC6Dl3vUwD4kC/obTAEgAgAaAGAVQDAS7S/AKAdAH4AACO+UHUIAYBJOpmzszMmJSVhaWkp1tTUYFNTEzY2Nr5zo+vU1tZiWVkZpqSkoKurKxP3BwCE8gWr4kdKdHNzw7q6OtQlWltb0dPTkwkrAAA9vniebAqOiorCo6Mj/nw6Q3JyMhNVxAuQ8h0TcxVITU1lor7nhRDvA8Cci4sLn6fTeHl5kaAlADDmBUWS2ra2Nj5Hp+nv70cDAwMSlcgL6nrw4AEffyarq6u4srIi/OVteXkZt7a2+BSR2dlZbG9vFzra4OAg7u/v8yFqcdIkfpM2CNq0/qY1qQkjIyNoZmaGpqamCs3ExAQjIyP5NOzp6UFvb2+8ffu2bK+xs7PDgoICPvxMcnJyKP9fALBkgj6jE1ZUVPCxKikpKZEVpMiCg4NlOfRrXL9+/VSc1OLi4mQ5Z/Hq1SuWS1OFwJd0oL6+no9VSXx8vFgErWO647du3RLt5s2bGBYWJsbTkrKxsRFzjI2NheLT09PR3t5eJqq7u1t2LVXQsj3J+5YJoplJ2Jk1gZYNKyA/Px/n5uaE54LZzMyM8CwxaJNm8Xp6erIGtLa2hnfv3hX9MTExou8smpubWd43TBANgsJyUBe625aWlmIBVPxZ0Ahz7949vHPnDjo6OvJuDAoKEs8XEhLCu5WiFUGTk5N47do14UTUGGgdP3v2DDMzM7GyslJhdzs+PsaDgwNcWlrChYUF3o3u7u6ioOjoaN6tFK0Iqq6uFi+u6CG3trYWRKpLb2+veIPIaAhWF60ISktLOyWCNyqwq6uLTz0FPXvSZkHLkX5JddGKoNDQULGAhw8fCmP+2NiY8JzQHsR8Tk5OeHh4yKeLkJj79+/LboQmHY7QiiDqXsPDw8Ky2t3dlfnKy8tlBQ4MDMj8jOnpabS1tZXFFhUV8WFnohVBqlhcXERDQ0OxSNqAeSYmJtDKykomprCwkA9Ti3cuiPYUGntYoc+fP5f5qcVLxejr6+OLFy9kMZpwbkFUML1gPXnyBH18fHB0dFTmn5qaknUs6Xk3NjbQwcFB9BkZGWFDQ4MsX1POLWh9fV1W8OPHj2X+2NhY0UcNYnNzU/RJmwkZTQTz8/PCEhwfHxdN0T6ljHMLIhITE2WF0S5fXFwszG7S4zQJM2g6l/rIaP6jMYg/fuGTwvb2Nnp4eJwqRGqPHj0SpgNGeHj4qRhl5u/vL7ueKrQiiNjZ2cGEhAQ0NzeXFUPzWnZ2Nh+OAQEBaGFhIQyhqoxiIiIi+HSlKBL0VtM2gx70N2/eYEtLCw4NDeHe3h4fIkC/KhndCFVGMZq8vSoS5EkHNH0f0hUUvQ99Tgeqqqr42CsB/RAngr5ggj4GgO2srCw+9kqQl5dHYg4AwJoJIn6l95GrCG3uAPA7ABhIBcXRCPL69Ws+XqehKeXGjRskKFMqhvgAABbpG9dVws/Pj8SsK/v/EX0jxoyMDD5PJ8nNzWXNIJoXIuUnCtL1BiERU8EL4KEH62cKpp+zr6+PP9elQi+VgYGBTMwvAPAeL0AZ9PH+T2oU9P3t6dOn+PLlS+zo6BB25ouyzs5O4XsefVHy9fVlU/4yNTG+YHX4CABSAWAAAP6RzmuXYLTPDANAFgB8whf6NtCmRTsxjRc0M12U0fVoNLPhC1LEfzzHHtB+37M8AAAAAElFTkSuQmCC",
  53: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAV2SURBVGhD5ZltSJZnFMePaeYbulHqdIho7EN92FgzQUs3m1tSBEE6jSCF0afZFCVNU0j9MPGFisE+rPlJa1oQRFriW5oY0xJfmAiGlSJMTWdNKzfTM/43XTfPffa8mU/2yH5wkVzXOdd9/vd9Peec+47of4w/EX1MRDFEFE9EX67DwHVwvU+IKEAG9CZ4E9F3RNRKRDNExO9w/ElE7UT0PRH5ykDtIYWIRrDZ7t27+fTp01xTU8M3btzg5uZmbmpqeusD16mvr+dLly7xmTNnODo6WokbJaI0GbA1foDjnj17+NatW+xMtLW1cVxcnBL2IxG5yOAlpTA+efIkLy8vy/2chry8PCXqJynAlG+UmI1AQUGBEvWtFAJ8iGg8MjJS+jk1+/btg6ApInpfCkqH2tu3b0sfp+bevXvs6uoKUblSUGdERIS0t8nMzAw/efJE+1eO6elpnp+fly46Dx480H7kuImPHz+Wy3bzOkn0mSYIFK2/cCZXw+DgIPv7+/O2bdvMjq1bt3J6erp04/v37/P+/ft5y5Ytep3x8vLixMREfvjwoTS3SVlZGfb4h4hClaBPsWltba20tUp1dbUekKWRnJxs8Ont7WUfH5//2KkREhLCjx49MvjY4ubNm8ofXYXGF5hobGyUtlY5deqUHgjOMe6yp6enPvAETpw4odujDKC2mQZfVFTE2dnZmr2aP3z4sOE6tmhvb1e+XytB6Jm0yrwacGxUEBcuXODx8XEeGxvTB44PfkuK/v5+w9MYGhrS1yorK/V5X19fg58tWlpalO9XShAaQa3dsJfFxUUODQ3Vg7Dn7E9NTWntU25uLhcWFhrWkBjUXt7e3pqtvThE0MjICLu5uWkbITHgHJ87d047QnV1dVazmzlSU1N1QbGxsbyysiJNLOIQQdevX9cD2Lx5s/63GuHh4ZpIa7x8+ZJLSkpMezMOCAjgvr4+aWoVhwjCkZEi5MAT7OzslK46w8PDBvvAwEC+e/euNLOJQwSlpaXpgURFRWltPgJEq48apNZ27drFr169ku4a6OSRGVUaxw3YuXMnX7t2TZpaxSGCkIUGBga0Y/XixQvD2uXLlw13HrXHHNgDmXFiYoLz8/N1excXF+7p6ZHmFnGIIGtMTk5qmUoFiAJsDyYvcHzs2DG5bJG3Lmh2dlZre1RwVVVV0oSXlpbkFGdmZuo+q+kr1ywIAeMFC6k2ISHBUCABmk6V0k33RTo/ePAg79ixg1NSUgw+4OjRo7rP3r175bJF1ixobm7OEPDx48cN61lZWfoaEsSzZ8+0eTwpNY9WqaurS/fBTfHz89PXsYe9rFkQQLVXF8dISkriixcvar2b6Tw6YQWKbVhYmL4GARkZGdpeQUFB+jz6Ojxle3GIoIWFBY6JiTEELweOkKz4qEvWum08+atXrxp8bOEQQeD58+eck5PDwcHBhqC2b9/OpaWl0lwHncChQ4cMmdDDw4Pj4+P5zp070twm5gS9UbetePr0qfbS1traqnXUaGnsAR05BKD9Hx0dlct2Y05QHCZW+z7kLJh7H/oME1euXJG2GwI8iNeCPleCPiCiheLiYmm7ITh//jzELBFRuBIEfltNMXMmUNyJ6HcicjUVlL1p0ybu7u6W9k4NCrK7uzsEFZmKAe8R0SRetjYSSP9ENGfp/4/wjZjPnj0r/ZyS8vJylQwypBBTfoaRsycIEzG1UoAEP6xfYIzHuZoXrfUAL5VHjhxRYn4lIg8pwBL4eP8HEgW+v1VUVHBDQwN3dHRolXm9BjoJvKrji9KBAwdUlz+NJCYDtodAIiogol4i+lv1Xe9ooM4MEFExEX0oA30TULRQidFeoGdar4HroTX7SAZkjn8BUr83PMMfhvYAAAAASUVORK5CYII=",
  54: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUPSURBVGhD5ZlZLHVXFMdXPzMNGrSNRgjx4kFTRYxROhDiSVS9IGlEBG1CTDHzUEGCNOlD9UFiJpGIMebpocYiFUEqeFA1p+YWq1kn39k5Z7uue/k+9970l6yQc/Y6Z/33OWettfcF+B9jAwCuAOAPAF8AwOcvYHQfut/HAPA+H9BTMAOAJAAYAoADAEAN2hEAjALAdwBgzgeqCt8AwBpdzMPDA7OysrC+vh47OztxYGAA+/v737rRfbq6urChoQFzcnLQx8dHFPcHAMTxASvjB3L09fXF3t5e1CaGh4cxMDBQFPYjALzDB89TSoNTUlLw9vaWv57WkJ2dLYr6iRcg5WtRjC6Qm5srivqWF0K8CwDbnp6evJ9WExQURIL+AoD3eEHJpHZkZIT30WpmZmZQT0+PRGXygibc3d358Y9ycHCA+/v7wl/e9vb28PT0lHdRyNTUFLa1tWF7ezuenJzwp5XyOkn8Jk0QVLT+pndSHZaWltDGxgatra0VmpWVFSYnJ/Nu91heXkYjIyNWc+bm5vghSikrKyO/fwDAXhT0CV2oubmZH6uUuro6FsRDFhUVxbvJoEzq5eUl81lYWOCHKaWnp0f0pa5C4DM60NfXx49VSnp6OguC3mNTU1M0MTFhRrMeHx/Pu8koLi6+NwnqChodHRV9vxIFUc8kVGZ1CA4OZkFUV1fj9vY2bm1tMdvY2BC+pYeYnZ1FfX39ZwsaHBwUfb8UBVEjKLQbqnJ1dYX29vYsCApeHa6vr9HV1ZX5S4VpRNDa2hoLghIDvceVlZVYVFSELS0tj2a3jIwMJiApKUk2ORoR1NHRwQIwMDBg/4vm6OgoiFTE+Pg4GxcQEIBHR0dCVtSooLy8vHsieKMnODExIfM7OztDZ2dn4byhoSGurq7ixcUFmpuba1ZQXFwcC8Db21to81dWVoRWXzrbbm5ueHNzw/wSExPZudLSUuHYzs4OWlhYaFYQZa/FxUXhtaIZltLY2MiCIxMLpaReYFhYGBtPCcLS0pKd29zclFztcd6IIGXs7u6imZkZC5AKNj0lBwcHdowWjFSYa2trhZRPdUs8V1JSIhxXVdhbF3R4eCi0PWKATU1NeH5+rrDmKDN60qrwbEEUMC2wYmNjMSQkROjFpKyvr8uCn5ycFBpOPuDH7MUEHR8fywKOiYmRnU9NTWXnKEHQN0ZPiF6l/Px8mRUUFGBaWhoaGxszn4SEBOEcP1EP8WxBRGZmpmw2IyMjsaamRujdpMepE34M6jo0nhSonvj7+8uC5y06Ohrv7u5413tQ2pYuH+bn5/khSnkjggh6jaiFsbW1lQlxcnJiNUYVaCHo4uKCdnZ2gqn6qokoEvSkbluEPnjqnIeGhoSieHl5yQ9RCj1Fmhx66mTq7jYpEhRIB9RdD2kLitZDn9KB1tZWfqxOQA/itaAAUdCHAHBGq0ddpKqqisT8CwCOoiDiVz8/P36sTkDFHQB+BwA9qaC0V69eCdtJugRlRFqGAECRVAxhCQC7tMelS4SHh5OY44d+P6I9YiwsLOT9tJLy8nIxGXzPC5HyMw3S9gQhEdPMC+ChD+sXGkyPc3p6mr+WRqFFZUREhCimCQCMeQEPQZv3f1KioP23iooK7O7uxrGxMaEyv5TRhgr96EY7SqGhoWKXv0dJjA9YFT4AgFwAmAOAa2m/pgGjOrMIAMUA8BEf6FOgokWVmNoL6pleyuh+1Jo58wEp4j8TOCZ6FzZIqwAAAABJRU5ErkJggg==",
  55: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAVqSURBVGhD5Zl9SFZ3FMfPUnPqWI5yE4cI2v4SNtZM0pSpe1GS/oqWQZQg4h+zzZqlooIKOqlQYzBo6R/iy1T8J3oTTbNEWS+yjFmoKSiEzmzm0srXM74Xf5f7nOelR33KR/aBg3J/59x7vvf+fr9z7n2I/sf4ENGnRBRJRF8T0VdvwXAdXO8zIvpQJrQavIjoeyJqJaIJIuJ1tH+IqJ2IfiCi92Wi9pBARP042c6dOzkzM5Orq6v54sWL3NLSws3NzW/ccJ1Lly5xTU0NZ2dnc3h4uBI3SESJMmFb/IzA3bt389WrV9mZaGtr4+joaCXsFyJ6RyYvKYbz0aNHeXFxUZ7PacjKylKifpUCjHynxGwEcnJylKgkKQS8R0QjoaGhMs6piYmJgaC/iegDKSgVaq9fvy5jnJo7d+6wi4sLRGVIQR0hISHS/7VMTEzwkydPtL/SxsfH+fnz5zKE5+fnrcaouFevXskwqyxvEn8aNwgUrX8xJ1fC/fv32cfHh7dt22bRtm7dyqmpqTKMKyoqtDHpb4yrra2VYVY5deoUBM0RUYAS9DmmW11dnfS1SVVVlSx+ZnbgwAEZxklJSWZ+0s6dOyfDrHLlyhUVh65CIwoHmpqapK9NTpw4oSeAeezp6ckeHh66ubu7c3JysgzjsLAwPW7z5s0W4yorK2WYVdrb29X5vlWC0DNplXklxMbG6omdPXuWR0ZGeHh4WLehoSFtTRiZmprSpqkS09HRYRKj4iytPWtcu3ZN5fGNEoRGUGs37AWLNiAgQBeEJOyhu7tbj9m+fbscXhUOEdTf38+urq7aiXDHMY9LS0s5Pz+f6+vrrd5h47pDDcGsKCws5KKiohVd34hDBF24cEFPzM3NTf9fWWBgoCZSkp6ervtgrci4yMhIfvTokQyziUME5ebmmiUjDU8Qa8RIVFSUmZ+0oKAgrRbZi0MEJSYm6glg10Kb//DhQ63VRy1RYzt27OCFhQUtZnZ2loODg/WxQ4cOcWdnJz948IALCgp406ZN+lhaWpq8pFUcIgi7V09PjzatXrx4YTKGomi849gIwNLSEj9+/Jhv3bpl8VopKSl6jL+/v9l5reEQQbYYGxtjLy8vPTlsBPaAp6xi8LQGBgaki0XeuKCnT59q7YtKDq2OPbS2tpo8WUxhe1izICSMF6wjR45wXFwc9/b2mozjzqot3XheTNFjx47xwYMHed++fTw3N2cSh+5AxXh7e5sVZWusWdDk5KRJwocPHzYZP378uD6GDQLdAcCrifEJlJeX6zFYX3jlV2Px8fGGM9pmzYJARkaGSXL79+/n8+fPa72b8Tg6YQVe6Xft2qWPoX7hiaERXX5R062rq8vkerZwiKDp6WmtCBqTkIaphTtvBNPT19fXzNdoJSUlJjGvwyGCwMzMDJ88eZL9/PxMEkJhLC4ulu46g4ODnJCQwFu2bNFj0K3jc1ljY6N0fy2WBK2q21Y8e/aM7969q+1S9+7d45cvX0oXi4yOjmpTC2urr69PDtuNJUHROLDS9yFnwdL70Bc40NDQIH03BHgQy4K+VIJ8iWga/dRGpKysDGLmiShQCQJ/RERESN8NAYo7Ef1FRC5GQT+hf0LjuJFAGcCrPBHlG8UAbyIawzeujcTevXshZtLa70f4Rsx5eXkyzik5ffq02gx+lEKM/AYnZ98gDGLqpAAJFlY5nPE4b9++Lc+1rqBjR6e+LOZ3InpXCrAGPt6PYqPA97czZ87w5cuX+caNG1plflt28+ZN7Uc3fFHas2eP6vLHsYnJhO3hIyLKIaJuIppVfdc6GepMDxEVENHHMtHVgKKFSoz2Aj3T2zJcD63ZJzIhS/wH9UoUj/9aV3wAAAAASUVORK5CYII=",
  56: "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",
  57: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAURSURBVGhD5ZlZLH9HFMePvWiaCtrSICEVT5oqElta2iLEk6h6sKWRSEqbaOwkeKkgQUojVU+W4k3s+/pQW4oUCSrhAbXUvrWY5tz87+Te4+f6/fj9+Uk/yQmZOTNzvu7cM2cugP8x1gDgCgB+APA5AHz2BIbr4HofAsA7NKCHYA4A3wBAPwDsAgB7RvsbAIYA4FsAeIsGqg5fAcASTubh4cEyMjJYXV0da21tZb29vaynp+e1G67T1tbG6uvrWXZ2NvP29hbF/QkAcTRgJX7AgT4+Pqyzs5PpEgMDA8zf318U9iMA6NHgKYXonJyczK6vr+l8OkNmZqYo6icqQMqXopiXQE5OjijqayoEeRMA1j09Pek4nSYgIAAF/QUAFlRQEqodHBykY3SayclJZmBggKLSqaBRd3d36n8vu7u7bGdnR/hJbXt7mx0fH3Pfs7MzoY36UcP5Dg4OZOso8SpJ/C5NEHhoHeGe1IS5uTlmbW3NrKysVJqlpSVLSkri/uXl5UIb9aOGPqGhobK1lCgqKkJB/wCAgyjoI9xujY2N1FeR2tpaevjdssjISO5fUFBwq/8uc3V1la2lREdHhzgOqwqBT7Ghq6uL+iqSmprKA8B9bGZmxkxNTbmZmJiwhIQE7l9SUiK0SX3QcJyFhQXT09Pj88XHx8vWUmJoaEgcFygKwppJOJk1ISgoiAeA22l9fZ2tra1xW11dFd4JkcPDQ6FN6oO2ubnJKisrmb6+vjCXm5sbOzk5ka2lRF9fnxjHF6IgLASFckNdLi4umIODAxeEgT6Uo6MjPhc+scXFReqiiFYELS0tMUNDQ2EiTAy4j0tLS1l+fj5ramqSZbf7iI6O5n+YwsJC2n0vWhHU0tLCgzAyMuK/i+bo6CiIvI+xsTE+xsnJiV1eXlKXe9GKoNzc3FsiqOETHB0dpUNlBAYGcv+qqirarRZaERQXF8cD8fLyEsp83PtY6uNZIvbhC351dUWHC0xPT3M/fIdOT0+pi1poRRBmr9nZWWFbYQUgpaGhQfakMHBVJCYmcp+srCzarTZaEaTE1tYWMzc358HiAUzBp2FjY8N9xsfHqYvavHZBe3t7QvkiBltTU0NdWHd3N+93cXG5c1uqw6MFYcB4wYqNjWXBwcFsfn5e1r+8vMxT+l3z4pVe7NekKlDFowXt7+/LAo6JiZH1p6Sk8D5MEFghUCTXaFZRUUG7NeLRgpD09HQeEFpERASrrq4WajdpO1bCFFplDA8PUxeN0IogrLX8/PxkwVOLiopiNzc3dKhQ84l1G9rCwgJ10QitCEIwU6WlpTFbW1uZEDzxlUqYlZUVZm9vz+zs7JizszPb2NigLhqhStCDqm0RvF1OTU2x/v5+NjMzw87Pz6mLDMxo+ITR6Bn2EFQJ8scGTe9DuoKq+9DH2NDc3Ex9XwT4IF4J+kQU9B4AnOAV+SVSVlaGYv4FAEdREPKbr68v9X0R4OEOAH8AgIFU0PeYSh9TUz0HWKUYGxujoHypGORtANjCE/wlERYWhmL27/r/EX4jZnl5eXScTlJcXCwmg++oECk/o5OuJwiJmEYqgIIv1i/ojI9zYmKCzvWs4KUyPDxcFPMrALxBBdwFfrzfxESB39/wI2F7e7tQSOLJ/FQ2MjIi/NMNvyiFhISIVf42JjEasDq8CwA5ADANAJfSeu0ZDM+ZWQAoAID3aaAPAQ8tPImxvMCa6akM18PS7AMakCr+A3QVeOgyVfjMAAAAAElFTkSuQmCC",
  58: "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",
  59: "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",
  60: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAWXSURBVGhD5ZlbTFVHFIYXIEqAVFBoCX0gwRRJSNo0FYiXQoG2gJEHxGifsKEaSUAwgEEuDxQeuCYQG24Nb0JKfFO5KAICDVpFQjCVSAEVYkJBUuRiYqmymn+Hmew9wOHgoXBIv2ReZs3aM/+ZmbXW3ofof4w7EX1KRF8S0ddEFLYJDfNgvs+I6EN1Qe+DExElEFEbEU0REW9h+4uIOogoiYg+UBdqDt8R0R94mL+/P1+6dIlra2v5xo0bfPv2bW5pafnPG+ZpaGjguro6zsrK4kOHDglxI0T0vbpgU+TD8fDhw9zc3MzWRHt7O4eEhAhhPxGRjbp4lQIMPn/+PL979059ntWQkZEhRFWoAvScFGK2A9nZ2ULUD6oQ4ExEYwEBAaqfVRMaGgpBE0TkqgpKhNo7d+6oPlZNT08P29nZQVS6KujXAwcOqOPN5vnz59zZ2alFpqGhIdW8IuPj49zV1cVtbW08MjKims1mKUj06QMEktYszuR6GRgY4OPHj7OTk5PMGbt27eLo6GgeGxtTh2vMzMzwuXPn2MXFRfo4ODhoPqOjo+rwNSkqKsIzFojISwj6HA+tr69Xx5rk/v37hkWpzdfXl6enpw0+iJxhYWHLxoq2b98+fvnypcFnLZqamoQ/qgqNr9Bx8+ZNdeyqzM7Osre3t1yIl5cX5+XlcXJyMu/cuVP2IyHqqampMQjATiEEOzs7y76EhASDz1p0dHQI32+FINRM2vk3l8rKSrmAPXv2GO5AZmamtB08eFD2Ly4uMu6psJ05c0baqqurZf/u3bt5ampK2taitbVV+H4jBKEQ1MoNc9Efm7i4OIMNu4dj8OjRI37x4oXsf/bsGe/YsUP66SPq5OSkYZeuX78ubWthsaD5+Xn29PSUk6PWAsPDw9zf38+vXr1SXTRu3bolfezt7fnp06fSht3bv3+/tBcWFhp8TWGxIIRlW1tb7SHIAeXl5RwTE6MtEn179+7l+Pj4ZQFBf39cXV21XdETGBgo7YmJiQabKSwW1NfXJyeGICFEbajScfwEpaWl0ubh4bFsJ4ODg6U9NjbWYDOFxYK6u7uXLf7YsWN87do1rqioMITyCxcuSL/8/HyTgoKCgqT91KlTBpspLBb08OFDgxg/Pz/tDgiuXLkibYiA4ugtJcBVBel36PTp0wabKSwW9OTJE7axsZGTp6SkGOxIjPqIde/ePa2/qqpK9uEOTUxMGPz0d2g9uchiQfjF3dzc5ORq8pybmzNEQUQ3gCMp+hC+1Sjn4+Mj7cXFxbonmsZiQQAJU0x+4sQJgw3RS1/boUQCg4ODMjqioZgVYLf0u7qeN+UNEVRQUCAnx0J6e3ul7fLly9KGnUQxCt6+favdN2FD2SNA6Bf9uHfq/TLFhgjCL6o/dsg96enpfPbsWfF+orWLFy8a/MrKyqQNDeVPWlqaYUeTkpIMPmuxIYIAjgVeFfQL1LcjR47w69evDT4LCwscGRm5bKxouEfrqePAhgkCd+/e1V6D9RW2u7u7ln8QHFbizZs3nJqaqo0TPo6Ojtpd1Nd+5rKSoHVX2yqo41DGI0Sb+wtjHMYjOLzPi51gJUEh6FjP+5A1sdL70BfouHr1qjp2W4CNWBIULAR5ENF8bm6uOnZbsBQ5/yEibyEI/IaotB2JiIiAoN+JyE4vKBVZXGT17cLjx49FhP1RLwa4ENGf+Ma1nYiKioKY6dX+P8I3Ys7JyVH9rBIUsEvBIFkVoudnDLL2AKETU68KUMHFqsFgbOeDBw/UZ20p+AiD7xdLYn4hIgdVwGrg4/04AkV4eDiXlJRwY2OjltWRmTer4ds3akZ8jzh69Kj4DDaJIKYu2Bw+IqJsIuolor9F3bVFDXmmn4hyiehjdaHvA5IWMjHKC9RMm9UwH0qzT9QFrcS/fvXjYwjQzUAAAAAASUVORK5CYII=",
  61: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAATOSURBVGhD5ZlZSG5VFMeXc6CEY4UpghAIQpfIAYfIoVJEHxwIH9MQLpQphrMPpg+JilOoEL2pKPokzkNeB9TUK6KmSCqKCOVAjlcwhxX/g+dwzr6ffmqlR/rBfvCsvfdZ/7P3t9baW6L/MU5E9D4RfUREnxBR6AM0vAfve0ZEb4kO3QdrIvqKiH4mol0i4kdsfxLRIBF9Q0Rvio7ehngi+g2TeXt7c1ZWFtfX13NbWxv39fVxb2/vf97wnvb2dm5oaODc3Fz29/eXxa0S0ReiwzfxPQYGBARwV1cX64mBgQEODg6Whf1ARCai8yJF6JycnMwXFxfifLohOztbFlUjClDzuSzmKZCXlyeL+lIUAmyIaMPHx0ccp2tCQkIgaIuI7ERBX0PtixcvxDG6Zmpqis3MzCAqUxQ04uXlJfa/Nevr6zw0NCRFpuXlZdF8IyMjI9zc3CwFoPPzc9FslKsgMaMOEEhah9iTd2VxcZFjYmLY2tpayRlWVlYcHR3NGxsbYvfXWFtbYwsLC2mcnZ0dHx8fi12MUlxcjPF/EZGbLOgDTNjU1CT2vZGJiQm2tbVVhIjNw8OD9/b2xGEKp6enHBQUpPR3cXHhV69eid2M0tnZKc+BqkIiCA+6u7vFvtdyeHjI7u7uijNubm5cWFjIKSkpbGlpqTxHQjQEViY0NFTzAVxdXe8laHBwUJ7jM1kQaiZp/9+W2tpaxRF7e3teXV1VbDk5OYrNz89PM253d5fLysrY0dFRI+afCOrv75fn+FQWhEJQKjdui/rrJiYmamxYPWyDubk53tzc1NiwgmoRWFkTE5PHFYQfrrOzs+IUai2wsrLCs7OzvL+/Lw5RSEhIkMYgcFRWVnJLS8vjrxDCsqmpqTQJckB1dTXHxsYq0crBwYGfP39uMCBkZGRwfHw8z8/PS3/rQtDMzIziBATJQsSGKh3bTw0im5rGxsbHFzQ6Ovqa85GRkdza2so1NTWaUJ6amioO16ALQS9fvtSI8fT05MvLS8VeV1en2BABDW09GV0IWlpaUiITWlpamsa+s7PDNjY2in18fFxjV6MLQfji6jwiJs+joyNNFOzp6dHY1ehCEEDClB2Ji4vT2La3tzW1HUqk69CNoKKiIsURbK/p6WnFVlVVpdiwkgcHB5qxanQjaGtrS7PtkHsyMzM5KSlJPp9ILT09XRyqQTeCAM4vyPiyQ2ILDAw06iBukuT+Tk5ORvsb4l8TBMbGxqRjsLrChmPIPwgOxkDuwrEBq+Pr68snJydiF6MYEnTnalsEdRzKeIRoVNS35ezsTKoN0e4jBhgSFIwHdzkP6QlD56EP8QDn+qcIFuJK0MeyoHeI6LigoEDs+ySoqKiAmDMicpcFgV8QlZ4i4eHhEPQrEZmpBX2LM85NWV2PLCwsyBH2O7UYYEtEf+CO6ykRFRUFMXvX/f8Id8Scn58vjtMlJSUlcjBIEYWo+RGd9B4gVGKaRAEi+GH9hM5YzsnJSXGuRwWXMLi/uBLTSERviAKuA5f3vyNQhIWFcWlpKXd0dEj318jMD9WGh4elmrG8vJwjIiLY3NwcQrYRxESHb8PbRJRHRNNEdCrXao/UkGdmiaiAiN4VHb0PSFrIxCgvUDM9VMP7UJq9JzpkiL8BpIBz9qSAuMQAAAAASUVORK5CYII=",
  62: "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",
  63: "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",
  64: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAU8SURBVGhD5ZlrKK5ZFMeXOxHmYGYyyUkpUaZpkIPJYGbo4INL03ycTELGEOG4hOHDCOUyoab5oFxGlJJr5ziuOcb90hCDkmTG/S5MrGk92U/Ps/F6HWe8r+ZX+8tee+1n/59nv2utvV+A/zEWAOAIAJ8BwBcA4PMAjZ5Dz/sYAN7nF/Q2GAJANAC8BoBNAEAVtm0A6AaAHwDAmF+oMnwDAH/SZM7OzvjixQusqqrCpqYmfPXqFb58+fI/b/Sc5uZmrK6uxrS0NHRzc2PiFgHgW37BiviJHN3d3bGtrQ3Vic7OTvTy8mLCfgYADX7xPLk0OCYmBs/Pz/n51IaUlBQmqowXIOVrJuYxkJ6ezkR9xwshjABg2cXFhfdTa7y9vUnQGgC8xwv6ntR2dXXxPmrN8PAwamlpkahkXlCfk5MTP15plpaWsKenR4hM8/PzvPlWBgcHsb6+HhsaGnB3d5c3K+QySIxLAwQlrX3ak3dlZmYGg4OD0dDQUMwZenp6GBQUhMvLy/zwa5menhZ8mP/o6Cg/RCF5eXnkdwYA1kzQJzRRbW0tP1Yh9FZNTU3FhfDNzs4Od3Z2eDcZFEldXV1lfhMTE/wwhbS2tjJfqioEPqeO9vZ2fuyN7O/vo42NjbgIa2trzMnJwdjYWNTV1RX7KSEqIjs7+8qLuKug7u5u5vsVE0Q1k7D/laW8vFxcwJMnT3BxcVG0paamirZnz57J/KSMjIygtrb2vQV1dHQw3y+ZICoEhXJDWXx8fMQFhIWFyWz09WgbTE1N4crKiszGOD09RUdHR3EOqbAHF3R4eIiWlpbiAqjWIhYWFnByclKpKJWUlCT6R0dHC1tWZYIoLGtqagqTUA4oLS3FkJAQ1NHREfrMzMwwMjLyxoDQ29srLt7T0xO3t7fR3NxcdYLGx8fFh5MgJoRvVKXT9pNCX9fW1lawU/CYm5vD4+NjNDY2Vp2g/v7+K4sPCAjAxsZGLCsrk4XyuLg4mW9UVJRoy83NFfpWV1fRxMREdYIoOknFODg44MXFhWivrKwUbRQB2daT5Av09/cXx1OAkL4Eqjruwr0Fzc7OooaGhriA+Ph4mX1jYwONjIxEO72As7MzfPr0qdhHB0YSXlFRgcXFxWhgYCDaKJ9Rv7LC7i2I3rj0R8wnz4ODA1kU7Ovrw729vWtzjqJWU1Mjm/cm7i2IoITJHhwaGiqzra+vy2o7ykdbW1tXFnxbe1BB9INmD6btJS0oS0pKRBt9SYpiR0dHwlbKyMiQtczMTExISEB9fX3RJyIiQrBR0aoM70TQ2tqabNtR7klOTsbw8HB2PhFaYmIi73qFk5MT1QYFBl2gSMt+vnl4eAhf5jYobEvnGRsb44co5J0JIt68eSMcg6UVtoWFhZB/KDgoA/3m7O3t0crKSmjKbjXGdYLuXG3zUB1HZfzAwABubm7yZoVQDqMvSVUEtbveNl0nyIs67nIeUieuOw99Sh11dXX82EcBfYhLQZ5M0IcAcEinx8dIUVERifkHAGyYIOJ3ikqPET8/PxL0BwBoSQUl0BmHLj4eExQRLyPsj1IxhCkA/E13XI+JwMBAErNz0/9HdEeMWVlZvJ9akp+fz4JBLC9Eyi80SN0DhERMLS+Ah35Yv9Jg+pxDQ0P8XCqFLmHo/uJSzG8AoM8LuAm6vP+LAoWvry8WFBRgS0uLcH9NmfmhGl2oUM1YWFiIz58/Z2eqdQpi/IKV4QMASAeAUQA4ZbWaihrlmUkAyAaAj/iFvg2UtCgTU3lBNdNDNXoelWa2/IKu419pHAh9/9RaUgAAAABJRU5ErkJggg==",
  65: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAWcSURBVGhD5ZltSJZnFMeP2stKWc6XTRohaANB2BjTyEqmuS2x+mDKUPDDSERhc1orTPNDM3JiodHIcNgHUZtEX6R8z3yjnG+IMtNWRoniNJlpFr5QZ/xvvC7u+/J59LHH5SP7wQG5z3Xu6/o/13Wfc+5bov8x7kT0KREFEtFXRBTyDgzzYL7PiOhDdUFvgyMRfU9EdUQ0TkS8hvYPETUQ0Y9E9L66UEuIIqK/cDN/f38+deoUFxcX882bN7m2tpZramr+c8M8t27d4pKSEj59+jTv2bNHiBsgou/UBS/FLwjcu3cvV1ZWsi1x584dDg4OFsJ+JSI7dfEqWRicmJjIr1+/Vu9nM6SmpgpReaoAPd8KMeuB9PR0ISpWFQKciGhw165dapxNs3//fggaJaIPVEE/QG19fb0aY9O0t7ezg4MDRKWogpr9/PzU8Rbz5MkTbmxs1DLTw4cPVbeB+fl5fvbsGY+Pj5u0sbExnpmZUcPMspAkuvQJAkVrCmdypdy/f5+PHDnCjo6OsmZs3ryZw8PDeXBwUB2ucfXqVXZ1dWU3NzeTBt+1a9fUMLNkZ2dj3jki8hSCPsdCSktL1bFL0trays7OzlKIaj4+PjwxMaGGcWxs7KKxquXn56thZqmoqBBx6Co0gnChqqpKHWuWqakp9vLykgvw9PTks2fPclJSEm/atEleR0FUCQgIkH6M3bp1K2/ZskUadriwsFANM0tDQ4O43zdCEHom7fxbypUrV+SiXFxceGBgQPrS0tKkD4vXMzk5ye7u7lJMc3MzP3361GCPHz/mFy9eGOKW4vbt22K+r4UgNIJau2EpISEhctFHjx41+LB7OAY9PT08NDRk8HV2dsq4nTt3Gnxvi9WCpqenefv27XJh6LXAo0ePuLu7m58/f66GSIqKimQcaghOxblz5zgzM9Pi+VWsFoS0bG9vr90ENeDy5cscERHBGzdu1K4hSyUkJJhMCCdOnJCC8KyIv4UFBgZqP8xKsFpQV1eXXAAECSGqoUvH8dMTFBS0aJxq3t7eWi2yFKsF3b17d9EiDh06xGVlZZyXl2dI5cnJyTJudnaWfX19pS8mJka7F2pZRkaG3HU1bjmsFtTR0WEQg0W+efNG+vXPCTKgOHoYMzw8rNUvU3PFx8fLuB07dvCrV6/UISaxWlB/fz/b2dnJyY8fP27wo61xcnKS/paWFoPfHHiZEzHYreVaKIHVgvCLo0URk6vFEzVEnwWrq6sNfnPU1dXJGFhfX586xCRWCwL6ah8ZGWnw4YHW93Y4YgAp/dixYxwdHa1lxbm5OUMcugMRg+cQjaolrIqgrKwsOTmOFwqm4NKlS9KHnUR3APBqIq7DCgoKZAyeL7zyC9/BgwelbzlWRdDo6Kjh2KH2pKSkcFxcnHg/0ezkyZMyBq/0u3fvlj6ke+wYGtGFFzVp9+7dM8y3FKsiCOADiqniKGzfvn388uVLQ0xvby97eHgsGqu3nJwcQ8xyrJoggF8Sv66+w0bziTpirsFEIxsVFcXbtm2TMdhVFOIbN26ow5fFlKAVd9sqaFfQxiNFW/owj4yMaD8Inq0HDx6obosxJSgYF1byPmRLmHof+gIXrl+/ro5dF2AjFgR9KQR5ENE0+qn1yMWLFyFmnoi8hCDwB7LSeiQ0NBSC/iQiB72gn9A/iaq+XkAZWMiwP+vFAGci+hvfuNYThw8fhpgJc/8/wjdiPnPmjBpnk5w/f14kgyRViJ7fMMjWE4ROTKkqQAUPVgEGYzvb2trUe60p6NjRqS+I+Z2I3lMFmAMf70eQKA4cOMAXLlzg8vJy7fs1KvO7sqamJq1nzM3N5bCwMN6wYQOEjCGJqQu2hI+IKJ2IOoloVvRda2SoM91ElEFEH6sLfRtQtFCJ0V6gZ3pXhvnQmn2iLsgU/wILb/ahtV0LEgAAAABJRU5ErkJggg==",
  66: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAW0SURBVGhD5ZlZTJVHFMcPIlKFNKDQNhQkkrgEY5taN7a0QIvE5QEkhMSXBkJCbC0qMbgQQ/GhICQQqpg0fTBuJbwR3FlEGrQqCCguqZIYoikgAiJEocBp/l/uTL4ZuNeLl8Il/SXzMmfON/P/7sw5Z75L9D/Gl4g+I6IIIvqGiKKnoWEezPc5EX2kL+h98CCi74momoi6iYhnsPUQUS0R/UhEH+oLtYckIvoLD1u7di3v27ePT58+zRUVFVxZWclXrlz5zxvmOXfuHJ85c4YPHjzIoaGhQlwbEX2nL9gWP8MxLCyML168yM5ETU0NR0ZGCmG/EJGLvnidXAzeuXMnj46O6s9zGvbv3y9ElegCzCQKMbOBrKwsISpFFwI8iah93bp1up9TExUVBUGdROStC/oBaq9evar7ODW3b99mV1dXiMrUBf2xZs0afbzdPH36lK9du2ZEpsePH+tmq9y/f5+rqqq4traWOzo6dLNdWIJEkzlAIGn1Y09OlgcPHnB8fDx7eHjInOHu7s5xcXHc3t6uD5cg/K9fv57nzJkj/by8vHj37t08NDSkD7fJkSNH4D9MRIFC0Bd4YGlpqT7WJjdv3jQWYUp+SluxYgX39vbqbnz8+PFxY81t+/btuotNLly4IHxRVRh8jY5Lly7pY63S39/PQUFBchGBgYF8+PBhTk9P53nz5sl+JEQzd+7cYRcXF2mPiIjgvLw8TkhIUERdvnxZ8bMFtqvFL0YIQs1k7H97Mb/lhQsXcltbm7QdOHBA2kJCQhS/xMREaduwYQOPjIxIW3h4uLRlZGQofrbAGbT4fSsEoRA0yg17iY6OlpMnJycrNvx62AZ3797lZ8+eyf6enh729vaWfqdOnVL8njx5YrzUhw8fcnd3t2KzhcOCBgYG2M/PTy4MtRbAglpaWrivr093Mbh+/br0QTB4/vw5j42NcWtrK9+7d08fbjcOC0JYFtEJOeDYsWO8bds2dnNzM/oWLVrEaWlp4wICilshyN/fn8+ePcurV6+WfUuWLOH8/HzFxx4cFtTU1CQXAUFCiN5QpWP7CSBc2ObPnz9uvGg7duxQ5nsXDguqr68ft4gtW7ZweXk5l5SUKKF8165d0s+SL5S2Z88e44qQmZmp9CNP2YvDghoaGpTJV65caZwFAQ67sCECiq1XXFys+KWkpJieypyamiptGzduVGy2cFjQo0ePlFyCt2zmxYsX7OnpKe03btww+k+cOKEIQiQ0g7uXsAUEBPCbN28UuzUcFoQ37uPjIyfXk+fr16+VKCiSJApfs6Dq6mrFz7yV8fyXL18qdms4LAggYYrJkeXNdHV1KbUdSiSA4nPBggWy/+jRo4rf+fPnpW3x4sX89u1bxW6NKRGUm5srJ8f2amxslDbzWcGbfvXqlbTFxMRI2/LlyxVbUlKStG3dulX2v4spEdTZ2alsO+QeRCocbMv9xGh79+5V/FAJCBvaqlWrjFJp8+bNSv9karkpEQRwiHFVMC/E3FCbDQ4O6m586NChcWPNbTJ1HJgyQQDlDK7B5grb19fXyD8IDtY4efIkBwcHK0KWLVtmJN/JMpGgSVfbOqjjUMYjRNtbWA4PD3Nzc7MR7XAG7Q0COhMJikTHZO5DzsRE96Ev0VFWVqaPnRXgh7AI+koI+oSIBnJycvSxs4KioiKI+YeIgoQg8Cei0mwkNjYWglqJyNUsKAN3HJHVZwv4DGaJsD+ZxQAvIurAN67ZBCoKIuq19v8RvhFzdna27ueU4HZrCQbpuhAzv2KQswcIk5hSXYAODtZvGIyf89atW/qzZhR8hMH3C4uY34noA12ANfDx/m8ECtweCwoKjPIe36+Rmaer1dXVGTVjYWEhb9q0iefOnQshXQhi+oLt4WMiyiKiRiIaMtddM9CQZ1qIKIeIPtUX+j4gaSETo7xAzTRdDfOhNFuqL2gi/gV1p9kNn+wI9QAAAABJRU5ErkJggg==",
  67: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAVGSURBVGhD5ZlbSJxHFMePmmijUjRqWy0qKJWAYKlNAhpLa2xVvDx4ofQlUC2C0FrFEkyiD9Y8VFRItN4ofUtCxTcx8ZIYb1XTREUM1Ui9YERivVAToyG26in/j8yw32R3s+tas9IfnJfvnPlm/jvjOWc+if7H+BBRGBF9RESfElHMPhjmwXzvE9Fb6oJ2gxsRfU1Et4lohYj4NdpfRNRNRN8S0ZvqQi3hCyL6Ay87ceIEnzt3jq9evcrNzc1869Ytvnnz5n9umOf69et87do1Liws5MjISCFumoi+VBdsjh8w8NSpU9za2sr2RGdnJ0dHRwthPxKRg7p4lVIE5+Tk8Pb2tvo+u+H8+fNCVK0qwJDPhZiDQFFRkRD1lSoEuBPR3MmTJ9Vxds3p06chaJGIPFVB30BtV1eXOsauGRwcZCcnJ4gqUAX9evz4cTXeYmZnZ7mnp0fLTJOTk6pb49mzZ7y0tMQrKytmbXl5mR8/fqwON8mLJDFimCBQtNZwJq1lfHycU1NT2c3NTdYMFxcXTklJ4bm5OV1sZWUle3l5sbe3t1lDTGJiom6sOcrKyjDv30QUKAR9gIU0NDSosWa5e/cue3h4SCGqHTt2jFdXV2V8SUnJSzGmLCwsTDeXOVpaWsQ4dBUan+BBW1ubGmuStbU1DgoKkgsIDAzkixcvcm5uLjs7O8vnKIiCiooKbfeOHDmiM1dXV/b09GQHBwc5LiMjQzefObq7u8W4WCEIPZN2/i2lrq5OTn706FGenp6WvgsXLkhfRESEfP7kyROemZnhhw8f6mxhYYFramrY0dFRGxMeHs7r6+ty3Kvo6OgQ830mBKER1NoNS4mJiZGLzszM1PmwezgG9+/f5/n5eZ3PGIjHDuNd2LEHDx6oIWaxWRB+PT8/PykIvRaYmpri0dFRqzIUOHPmjHxXaWmp6n4lNgtCWhbHAzUAxyUtLY0PHz6sPUOWys7O1iUEU/T19UkxwcHBvLm5qYa8EpsFjYyMyEVAkBCiGrp0HCdzxMbGyvj6+nrVbRE2C+rv739p8UlJSdzU1MS1tbW6VJ6Xl6cOlwwPD8s4/A1tbGyoIRZhs6ChoSGdmNDQUN7Z2ZH+K1euSB8yoKmjh2Mp4pAZd4vNgiYmJnQ1Iz8/X+dH6+Lu7i79d+7c0fkBdsPX11fGoEjvFpsF4RdHiyIWY1g8wdOnT3VZsL29XecHeCb86Ci2trbUEIuxWRBAwRQLSk9P1/nQgBr2dsZ+fVzphd+arsAYeyII9UIsCMcLf+CCqqoq6cNOokNQMbhGc3V1teq2ij0RtLi4qDt2qD0FBQWclZUl7ieanT17Vh3Kz58/l50BDNcOW9gTQQAfUNBsioWpFhUVZTQV41ohCjMMVxBb2DNBYGBgQLsGG3bYPj4+Wv1BcjAGWqSAgAD29/fnkJAQfvTokRpiFcYEWd1tq2CRaOORonHrNAcyGvpBGG6xtmJMUDQeWHMfsieM3Yc+xIPGxkY19kCAjXgh6GMh6B0iWscV+SBy+fJliPmHiIKEIPAbstJBJD4+HoJ+JyInQ0HfIZUaq+r2zNjYmMiw3xuKAR5E9Ccq+EEiOTkZYlZN/f8I34i5uLhYHWeXlJeXi2SQqwox5CcE2XuCMBDToApQwR/WzwjGdt67d09912sFH2Hw/eKFmF+I6A1VgCnw8X4BiSIuLk77SHjjxg2tkURl3i/r7e3VesZLly5xQkICHzp0CEKWkMTUBVvC20RURETDRLQperXXZKgzo0RUQkTvqgvdDShaqMRoL9Az7ZdhPrRm76kLMsa/iitbCXJcoc4AAAAASUVORK5CYII=",
  68: "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",
  69: "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",
  70: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUbSURBVGhD5ZlrKL9nGMevOeV8ClOjZRpJbdYcCmPYnEoSsVesprxw2AtqhDIvmFPUshjyhibvnM/Hf3JKSyhWTqOc5xAyw72up9999zy33+/xYz/8tE9db57ruu/n+rp/93Vf9wPgf4wtAHwCAF8AwFcAEPIChu/B930KAHZ8Qk/BBABSAWAIAA4BgLyi/QUAowCQAQDmfKLq8A0A/IGTeXl5kezsbNLU1EQ6OjrIwMAA6e/vf3bD93R2dpLm5maSm5tLfH19qbhVAPiWT1iOYhzo5+dHenp6iDYxPDxMgoKCqLCfAeA9PnmenzA4PT2d3N7e8vNpDTk5OVTUL7wAMfFUzFsgLy+PivqOF4KYAsCf3t7e/DitJjg4GAXtAYAVLygN1Y6MjPBjtJrZ2Vmiq6uLon7gBb3z9PTk45VyfX1NDg4OyOHhoaxhzNHRkcq9uLOzQ8bHx8nQ0BBZXV3l3WqjKBK/iwsEHlpn+JtUh6mpKWJjY6OWOTk5kd3dXcn409NTkpKSQiwtLdk5Y2hoSGJiYsjm5qYkVh1KS0txjmsA+JAK+gwnbWlp4WOVMjo6yhJRx7a2tthYXK2QkJB7MdScnZ2FlX0M3d3ddDx2FQJf4oPe3l4+VikTExPEyMhIqVlZWREDAwOWoJubGzk7O2Nj6+vrJQJwpbAEm5qasmepqamS9z2E6A8cSgVhzySczOpwdXVFNjY2BMOfCDVcCRSLonA+MzMzsri4yMbd3d0R3Kc08eTkZOarra1lzy0sLIQ9qC6Dg4N07NdUEDaCQrvxX4mKimKJNTY2Snzr6+tET0+P+cUVdX9/X7JK7e3tkrFyPJughoYGllBoaCjvJn19fcyvr69P1tbWmA9Xz9XVlflLSkokY+V4FkHHx8fEzs5OmBjPhYWFBT5Esn/wZ4mrIsbHx4f509LSJD45nkVQUVERSyYuLo53C1RWVrIYe3t7cnJyIvEHBgYyf2JiosQnh8YFXV5eEkdHR7Y6c3NzfIhAcXGxrKCAgADmT0hIkPjk0LggPL9oIv7+/ryboTgAVQoSr1BSUpLEJ4fGBUVHR7NEKioqeDejpqaGxeEe2tvbk/jFe+gxZ5FGBeGpbm5uLkyIlWt5eZkPYbS1tbGEsXzzVc7FxYX5y8rKJGPl0Kigrq4ulgSWXUxMFSsrK0RHR4fFj42NMR+ulvgcesxNWaOC8vPzWRLx8fG8W8LNzQ1xd3dn8dj2UKqrq9lza2vre/tLDo0KwhJNE8HS/RBVVVUsHg3bn6ysLGJiYsKeZWRk8MNk0aggDw8PlkhdXR3vvgfepSIiIiSixIb76DF9HKIxQbhf8HLl4OAgmLq/e2xuMzMzia2tLRNibGwsrPb29jYf/iDKBD2q2xZzcXFBzs/PBVN1M1UFrsTk5KRQHJ5ysaMoExSED9S9D2kbyu5Dn+OD1tZWPvZNgAuhEBRIBdkDwHlhYSEf+yZQVM5/AOAjKgiZkuvDtJnw8HAUtAgAumJBmXiKT09P8/FazdLSEv2G8aNYDGIJALtYht8Siiv/sar/H+E3YlJQUMCP00qwgVUUg+95IWJ+xSBtLxAiMS28AB7cWPUYjMs5MzPDz/WqzM/Pk9jYWCrmNwAw5AWoAj/e72ChCAsLI+Xl5cJVAU91PJlfyvDbN7ZU+D0iMjKSfgbbxyLGJ6wO7wNAHgDMAcDftO96JcNzZh4ACgHgAz7Rp4CHFp7E2F5gz/RShu/D1uxjPiFl/As55GXa27ofAAAAAABJRU5ErkJggg==",
  71: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAQySURBVGhD5ZlLSEVVFIZ/30kpCWpBihiEOCkjH6BGPkpFEAeKOEwaljkITFHBHFj4lqBBJk4URUfiEzWfCD6QEBUkCDUHWWqCqPhesTaew3FxvPd6I93SB2ty9trrrP/se9Ze+1zgf0wIgHcBfAjgYwDpT2B8H77fewBCZULu8CqAzwH8DOAAAD2j/Q1gGsCXAAJloq5QCOBXDhYXF0dlZWXU2dlJAwMDND4+TmNjY/+58X0GBwepq6uLKioqKDEx0RD3G4BPZcKO+JYnJiUl0cjICOnE5OQkpaamGsK+B+Ahk5d8x87FxcV0c3Mj42lDeXm5IeoHKcBKgSHmJVBZWWmI+kwKYV4D8Ht8fLycpzVpaWks6E8AQVLQF6x2ampKztGa5eVl8vLyYlFfS0FzsbGx0t+Wy8tL2t/fp4ODA4fGPoeHh07fxbm5Oert7VUF6Pr6Wg475a5I/GItELxpHfNv0hUWFhYoODjYJYuMjKS9vT0ZwmRra4t8fHzUuxAUFEQnJyfSxSl1dXU8/xJAhCHofQ7Y09MjfW2Znp62bnhObXd3V4ZQXFxcUEpKiukXFhZGp6en0s0pw8PDRgzuKhQpfGF0dFT62jI/P0/+/v62xk/Z19fXTDI6OpqOj49lCLUy6enp94SHh4e7JcjygDMMQdwzqZ3ZFc7Pz2l7e1vZzs6OabwSLJZFcbyAgABaX1+/N5ffraamJvVzlCvprqCJiQkjxieGIG4EVbvxb8nJyTET7OjokMNUUlJyT0RERAR5eHjoKai9vd1MNCMjQw4rioqK1Lifnx+1trZSX1+fnit0dHREoaGhKjDvC2tra9JFUVpaSoWFhea4toJqa2vNxPLz8+WwCVc2K93d3foJOjs7U8kYq7OysiJdHkRLQbx/GUklJyfLYYdoKSg3N9dMqrGxUQ47RDtB3K8FBgaqgNzCbG5uSheHaCdoaGjITCgqKopub2+li0O0E1RVVWUmVFBQIIedop0gLtFGQly6H4t2gmJiYsyE2tra5LBT+EuSMT8kJOR5BfH7wocrbvvZ3PlC1N/fr+by6iQkJKg97bHYCXpUt22FnygfyticnUztuLq6Mue7I4axE5TKF1w9D+mG3XnoA77A5/qXCC/EnaCPDEFvAjipqamRvi+ClpYWFnMF4G1DELPw2D5MF7KysljQOgAvq6CvPD09aXFxUfprzcbGhvEN4xurGOZ1AHtchl8Sd0f+o4f+P+JvxFRdXS3naUl9fb1RDEqkECs/spPuBcIipkcKkPCL9RM783IuLS3JWM/K6uoq5eXlGWK6AbwiBTwEf7z/gwtFZmYmNTQ0qKPCzMyM2pmfymZnZ1VL1dzcTNnZ2eTt7c1C/uIiJhN2hTcAVAJYAXBx91Sey3ifWQVQA+Atmag78KbFOzG3F9wzPZXx/bg1e0cmZMc/X2/2Xj+FDtsAAAAASUVORK5CYII=",
  72: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAATnSURBVGhD5ZlbLHVXEMfHNaUlJapNSkrFNaKa4gFNS1tEIohLPEr6RlXEQwlCPWjjEhLioSoiIRWv7pe6fcQtNG4JjYpblKJEECVMMzvWztnjnGP7enCkv2RezprZe/577TVr1j4A/2PeAQA/APgUAL4EgC+ewOg+dL+PAMCRJ/Q6vAkA6QDwKwAcAAA+o/0NAEMA8C0A2PJE1ZACAL/TxQIDAzEnJwebmpqwra0N+/r6sLe399GN7tPe3o7Nzc2Yl5eHwcHBQtwfAJDKE9bHDxQYEhKCXV1daEwMDAxgWFiYEFYNACY8ec6P5JyRkYHX19f8ekZDbm6uEFXLBWiSLMS8BPLz84Wor7kQ4i0A2AwKCuJxRk14eDgJ2gMAOy7oG1I7ODjIY4ya6elpNDMzI1HfcUGvAgICuL9WLi8vcX9/Hw8ODvQa+RweHupcixsbGzg0NCRVtNnZWby4uOAuqrgtEr9pFgjatE7onVTDxMQEOjg4qDJXV1fc3d1VxI+NjWFkZCRaW1sr9hpPT0+srq5W+KqhtLSU4i8B4AMh6GO6YEtLC/fVCj1Vtunpta2tLTmWZsPCwuKOj6ZlZ2cr7ncfnZ2dIpa6ConP6Yfu7m7uqxV6wlZWVlrNzs4OLS0t5eS8vb3x5OREiqNXyt3dXR4jX0q+sLAQfXx8FKJGR0f5bXWi8YAjhCDqmaSdWQ2U2Pr6umS0DoTRTJBYSpSuZ2Njg4uLi3IcbdIiYRMTE0UBorXm7Owsj2dlZclj99Hf3y/ivhKCqBGUXof/SkxMjJxUQ0ODYoxaGDc3N7S3t0c/Pz/FGJGUlCTHpqSk8GGdPJqg+vp6OaGIiAg+jDc3N3h1dYV7e3u4s7PDhzE0NFSOz8zM5MM6eRRBR0dH6OjoKF2Y9oWFhQXuopfx8XE0NzeXBVETrJZHEVRSUiInk5iYyIf1srm5qSgW9DrSTKrF4ILOz8/lBU2zMzMzw110QmK8vLxkMWQPqXCEwQXR/iWSoXWgltXVVfTw8FCIqa2t5W73YnBBsbGxckIVFRV8WCsrKyvo4uKiEFNTU8PdVGFQQdSv2draShekDmB5eZm73GFtbU0hxtTUFBsbG7mbagwqqKOjQ06MejEqzfo4Pj5GX19fOYYeRk9PD3d7EAYVVFBQICeXnJzMh++Qmpoq+5NRR7C9vS29gjS7wrTtU7owqCAq0SI5Kt36mJ+fV4gho46b2iD++7N1Cv7+/nISdXV1fFhBWlrancR1WVxcHA/XicEE0Xqhw5WTk5Nk930hio+Pl/xoz9Jn5JOens7DdaJN0IO6bU3Ozs7w9PRUMl0nU4Hwoxh9Rj4POb1qExRGP6g9Dxkb2s5Dn9APra2t3PdFQBNxK+gzIeg9ADgtLi7mvi+CqqoqEnMFAB8KQcTEQ/owYyIqKooELQKAmaagbGpBJicnub9Rs7S0JL5hfK8phngbAHapDL8kbo/8R7r+P6JvxFhUVMTjjJKysjJRDDK5EE1+IidjLxAaYlq4AA4trJ/JmaZzamqKX+tZmZubw4SEBCHmFwB4gwvQBX28/5MKBX22LS8vl44Kw8PD0s78VDYyMiK1VJWVlRgdHS0+pPxFRYwnrIZ3ASAfAGYA4J/bp/JcRvvMHAAUA8D7PNHXgTYt2ompvaCe6amM7ketmTtPSBv/Ag3Mg1m71gREAAAAAElFTkSuQmCC",
  73: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUJSURBVGhD5ZlZLH13EMfHvlOJUilpNGmEh1ZTPKBaqggREiLEg3/Sx1YJD/bEkmhjCUmTPlTFC6kl8WINag2xBkFIxS4pRSX2EqaZX/xOzv259zr08r/STzIv5zfzO/O955yZOecC/I95FwA+BoDPASAEAL56AaPz0Pk+AQBHMaGnYAUA3wLA7wBwAAD4Fu1vABgAgO8BwFZMVAkJAPAHbebj44NZWVlYV1eHra2t2NPTg93d3c9udJ62tjasr6/H3Nxc9PPz4+JWAeCNmLA2fqBAf39/7OzsRH2ir68Pg4KCuLCfAMBATF7kR3JOSUnBm5sbcT+9ITs7m4v6WRQgJ56LeQ3k5eVxUd+IQghrANjy9fUV4/Sa4OBgErQHAPaioO9IbX9/vxij10xOTqKRkRGJyhQFDXt7e4v+arm6usL9/X08ODjQauRzeHio8VlcWVlhDzn9iBsbG+KyYu6KxIy8QFDTOqZ7UgljY2Po4OCgyNzc3HB3d1clfmpqCsPCwtDMzEzqM5aWlhgXF4dra2sqvkooLS2lPa4A4AMu6FPatKGhQfRVy8DAgJSIEtve3pZip6en0dra+p4PN1dXV1xfX1c530N0dHTweJoqGF/Sga6uLtFXLSMjI2hhYaHW7O3t0dTUVErQw8MDj4+PWRzdetTb5MkXFhZiRkYGi+XHY2JixFNqRfYDh3JBNDOxzqyEy8tLds+TbW5uSkZXgsSSKNrPxsYGFxYWpLjZ2VmVq7G4uCitVVRUSMdtbW3ZM6iU3t5eHvs1F0SDIBs3/itRUVFSYrW1tSpre3t7bHzKzMzE/Px8lTUqDDzOysqK+Srl2QTV1NRISYWGhorLWklOTpZiAwMD8fb2VnTRyLMIOjo6QkdHR7Yx9YX5+XnR5R4XFxdYXFwsn83YHjMzM6KrVp5FUElJiZQUlV8lLC0tSTFkTk5OODo6Kro9iM4FnZ+fs4rFrw6VZiXQJE/9h5dxY2Nj9PT0xJaWFtFVKzoXRP2L/8oBAQHiskaokm1tbeHOzg7m5ORIexgYGODExITorhGdC4qOjpaSofL7VGQvcJiUlCQua0Sngmheo75B8SYmJri8vCy6qOX6+lo8hGlpaZIgpXMloVNB7e3tUhLu7u5ay21jYyNGRkay6SEhIUFcxsTERGmvx9y6OhVEDZInER8fLy6rIO9TVDxoquDQ1GBnZyetp6enq8RqQ6eCqETzJKh0a+Pk5IRN39yfBKSmprLJwdnZWTpOcx29WihFp4K8vLykRKqrq8XlewwPD2udtql0Nzc3i2Fa0Zkgel6oy7u4uDBT+oWIJgGa+Whm40LMzc0xJCQEh4aGRPcHUSfoUdO2nLOzMzw9PWWm6c1UEzShkwAa/1dXV8VlxagTFEQHlL4P6Rvq3oc+owNNTU2i76uALsSdoC+4oPcA4LSoqEj0fRVUVVWRmGsA+JALIsYe08z0ifDwcBK0AABGckEZhoaGOD4+LvrrNdSQ775hFMrFEO8AwC6V4dfE3Sv/kab/j+gbMRYUFIhxeklZWRkvBqmiEDm/kJO+FwiZmAZRgAg9WL+SM13Ox7xovQRzc3MYGxvLxfwGAOaiAE3Qx/s/qVDQZ9vy8nL2qjA4OMg680sZTRI0UlVWVmJERASb9QDgLypiYsJKcAKAPACYBoB/+Nz1loz6zBwAFAHA+2KiT4GaFnViGi9oZnopo/PRaPaRmJA6/gUjxJvFR5u2CQAAAABJRU5ErkJggg==",
  74: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAStSURBVGhD5ZltKH9nGMevedbmqcxWI6wkXvhbQ8IytiGSF6S9XO2FZP57oYw8jjJ5CK32YibKw0R54zHM01KeRyhWK6bMPBPC4lrX6Xfujmvn/9vBb5xf+9T15tzXfZ/re875Xdd13z+A/zFvAkAAAHwAAB8DwEdPYHQfut8LAHDjAT2E1wEgAwB+AoB9AMBntEMAGAOAlwDgyAPVwqcA8CstFhwcjDk5OdjS0oLd3d04NDSEg4OD/7nRfXp6erC1tRXz8vIwLCxMFvcbAHzGAzbGNzQxPDwc+/v7UU+MjIxgVFSULOxbAHiNB88pJ+fMzEy8ubnh6+mG3NxcWdR3XICSVFmMOZCfny+L+pwLId4AgN9DQkL4PF0THR1Ngv4EABcu6AtSOzo6yufomtnZWbS0tCRRX3FBPwcFBXF/Va6vr3Fvbw/39/eNGvkcHBxo+i1OT09jZ2cndnV14fHxMR82iiFJ/KJMEFS0Tumb1MLU1BS6urpqMm9vb9zZ2eFL3GF1dRVtbW1FzZmfn+cuRqmoqKB51wDgKQt6jxZqb2/nvqqMjY2Jm2uxra0tvoSA3l5oaOgd/8XFRe5mlL6+PnkudRUSH9KFgYEB7qvK5OQk2tvbq5qLiwva2NiI4Pz8/PD09JQvISgpKfnHA7ivIMUDjpEFUc8kVWYtXF5e4sbGhmSbm5vC6E2QWBJF6zk4OODKygqfLpibm0MrK6tHCxoeHpbnfiILokZQajceS2JiogissbGRDwuurq4wICBA+CqF6UZQQ0ODCComJoYP3yE7O1v4ZmRkoKenp74EHR0doZubm7Qw1YXl5WXuIpiYmBDBR0ZG4uHhoZQRdSWorKxMBJSSksKHBWdnZ+jj4yP5UfJYX1/Hi4sLdHR01I8gCsjDw0O8HWN1JD09XQReXl4uXdve3kYnJyf9CKL6JQcTERHBhwWKeoEJCQniOiUIZ2dnMUbZ8z6YXFBSUpIIprq6mg9LnJycoJeXl/CjDWNzczM2NTVhXV2dVMPksdLSUum6VmEmFUT9mvz9W1tb49raGneRoODUao4xa2tr48uoYlJBvb29IgBfX1+8vb3lLhIkiAf8b/YsggoKCkQAqampfFhAHTR9SoWFhXesqKgIs7Ky0M7OTqyTlpYmjVHTqgWTCqIULQdCqfshUCulm6QQGBgoAqmvr+fDmqC0rdw+LCwscBejmEwQ/V5oc+Xu7i7ZQ0+Idnd30d/fX6plZFo/NRk1QffqtpWcn59L1Z9My85UDXowj1lHTVAUXdC6H9Ibavuh9+lCR0cH9zUL6EUYBEXKgt4GgDPaPZojtbW1JOYvAHhXFkRMGevD9ExcXBwJWgEAS6WgLAsLC+k4yZygjGg4w/haKYZwBoAdSsPmhGHLf/Sq/4/ojBiLi4v5PF1SWVkpJ4MvuRAl35OT3hOEQkw7F8ChH9YP5Eyvc2Zmhq/1rCwtLWFycrIs5kcAsOMCXgUd3v9BiSI2NharqqqkrcL4+LhUmZ/K6ECFWqqamhqMj4+X91S7lMR4wFp4CwDyAWAeAK4MT+W5jOrMEgCUAMA7PNCHQEWLKjG1F9QzPZXR/ag18+EBqfE3fB2KCnMFaCgAAAAASUVORK5CYII=",
  75: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUTSURBVGhD5Zl5SD1VFMdPLj/TcsMlISVMQ0QoIxVXUisVRfxDEQX/CPzDP8pSElJUMEELdwgCUxFxScT/XHFfUFyQEPdc0EQ0l8QVF/TEGd4d3lyf85vsqU/6wAF555w75ztz595zR4D/MTYA8CEABADA5wDw2RMYXYeu9xEA2PIFPYS3AOArAOgFgH0AwGe0vwFgAAC+AQAzvlAlxAHAHzSYp6cnpqenY11dHba0tGB3dzd2dXU9utF1Wltbsb6+HjMzM9HX15eJWwWAL/mC5fiREv38/LCjowN1ib6+PgwKCmLCfgaAN/jieX6i4OTkZLy5ueHH0xkyMjKYqF94AerEMjEvgaysLCYqkRdCvA0Af3p5efF5Ok1wcDAJ+gsALHlBX5Pa/v5+PkenmZycRH19fRL1PS9o2MPDg4/XyNXVFe7t7eH+/r6sUczBwcGdd/H6+lo2f3d3Fy8uLiQ5cqgWid/VFwjatI5pTiphbGwMra2tFZmjoyPu7OxI8quqqtDKyupOLDPyNTQ0SHLkKCgoIEFXAPAeE/QxTbfGxkY+ViMDAwP8pidrm5ubkvzExMQ7MbyVl5dLcuRob29nedRVCATSD52dnXysRkZGRtDY2FijWVpa4qtXr8TCXF1d8fj4WJLv4+Mj+inWxMREMoaRkRHW1NRIcuRQu8EhTBD1TMLOrASa3+vr64JtbGyIRk+CxJIoGs/U1BRnZ2cluUdHR2hjYyOKGR4eloxBtra2hicnJ5I8OXp6epigL5ggagSFduO/EhkZKd796upq3o1TU1Oi39nZmXc/iEcTRC87KzYkJIR3C9TW1ooxtIfQrMjLy8P8/PwHX/9RBB0eHqKtra0wMO0LMzMzfIhAWlqaKIjeFfY3s4CAAFxZWeHTZHkUQXSHWVExMTG8WyQwMPCOCN6cnJyEvUgpWhd0fn6ODg4OwqD0dOg90cTl5SW6ubmJhSckJAiLyPz8PObm5qKenp7oS0lJ4dPvReuCaP9ihfj7+/NukdvbW9za2sLx8XGN10pKShLHoRtEN0oJWhcUFRUlFlJcXMy7FUOHOTYOPa3l5WU+RCNaFUQ9mZmZmTCgoaEhLi4u8iGK6e3tFQWRLSws8CEa0aqgtrY2sQAXFxdhWt3H9PQ0pqamYnx8PEZHRwtNrjrUHbCxLCwshEZVCVoVlJ2dLRYRGxvLuyXQ0UT9CVRWVoo+uhF05Ge+iIgISa4cWhVESzQrgpZuOegY4e3tLcbTFKUnRo2o6qAm2ujoKJ9+L1oV5O7uLhZRUVHBu+8wNzeHdnZ2kuJ5Kykp4dNk0ZogmiZ0uLK3txdM6Rei1dVVjIuLQ3Nzc1EE7V/0uay5uZkPfy2aBP2rbluds7MzPD09FYw/mb6O7e1tYWrRu7W0tMS7FaNJUBD9oPQ8pGtoOg99Qj80NTXxsS8CehAqQZ8yQXYAcEr91EukrKyMxFwDwPtMEDEm14fpMmFhYSRoFgD01QV9R/0TNY4vCdoGVN8wflAXQ1gAwA4twy8J1ZH/8L7/H9E3YszJyeHzdJLCwkK2GHzLC1HnVwrS9QVCTUwjL4CHXqxKCqbHOTExwY/1rFDHTp26SsxvAPAmL+A+6OP9Ni0UoaGhWFRUJBwVBgcHhZ35qWxoaEhoqUpLSzE8PBwNDAxIyC4tYnzBSngHALIAYAoALlnf9UxG+8w0AOQCwLt8oQ+BNi3aiam9oJ7pqYyuR63ZB3xBmvgHxk95GHXl32sAAAAASUVORK5CYII=",
  76: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUTSURBVGhD5ZlbSC1lFMeXV9JSFMyC1DBIUbDMvOGF0krliPigiOBLIL5YFuqDdzEfLG8opgmZKN6S83q83/WgeEFTVJQw0ESy9CiK99QVa9jzMfM1ezuetrqlHyw4zFrfN+u/58xa6xsB/se8DgDvAUAwAHwKAJ/cg9F96H7vA4A9n9DL8CoAfAEAgwCwCwD4gLYHACMA8BUAWPOJqiEeAH6lzXx8fDAzMxNbWlrw2bNn2N/fj319fXdudJ+Ojg5sbW3FnJwcDAgIEMX9BgCf8wnr4ltaGBgYiN3d3WhIDA0NYUhIiCjsewAw4pPn+Y6CU1JS8Orqit/PYMjKyhJF/cALkBIninkM5ObmiqISeSHEawDwu6+vL7/OoAkNDSVBfwKALS/oS1I7PDzMrzFoZmZm0MTEhERl8IKee3t78/GKXFxc4M7ODu7u7uo0innx4oXOd3F5eRkHBgZwZGQEt7e3ebcqNEXiF2mBoKZ1SP8n1TA5OYl2dnaqzNnZWTFRKv9+fn5obGzMeo2NjQ2mpqbi+fk5H66TkpISWn8BAG+Lgj6gDdvb2/lYRejXVGh8Wm1zc1O2vra29l8xUktISJDF30RXV5e4lqYKgY/pQk9PDx+ryPj4OFpYWCiara0tmpubs+Tc3Nzw8PCQrZ2bm0MjIyPmDw4OxuLiYoyNjZWJ6u3tld1TF5IfOEwURDOT0JnVcHZ2huvr64JtbGwwoydBYkkU7WdlZYVLS0uytXFxcSxpf39/vLy8ZL6goCDmS09Pl63TBb2DmnWfiYJoEBTGjf9KVFQUS6qhoUHm29vbY2LJmpubZf61tTXhR11ZWRGKilruTFB9fT1LNiwsjHfjxMQE81Mx2Nrawuvra+EpLi4u8uGquRNB+/v7aG9vL2xMfUEpQRpuRUEODg7Y1taGXl5e7BpVxNLSUn7ZjdyJoKKiIpYYveBK1NTUsBgqIOK/eUtOTuaX6kTvgk5OTtDR0VHYlJ7O7OwsHyKg6RcyS0tLE44IGRkZsuvUp9Sid0HUv8REqFJpo6qqSpZ0YmKizJ+UlMR84eHhMp8u9C4oOjqaJVJeXs67GY2NjTJB1BCl0NlL9NETPz09lfm1oVdBNK9ZW1sLG5qZmeHq6iofwqDBVypocHBQ5qceJvpobKI5UA16FdTZ2cmScHV1FcqwNmims7S0ZPHV1dUyv3QvJycnoYGrQa+C8vLyWBI0BdwE9SfpD3BwcMB88fHxzEcNWi16FSSdwah03wRNAmI8mYeHB2ZnZ2NkZKTs+m1mOb0K8vT0ZEnU1dXxbkXy8/NlyfN2mzmO0Jsgel/ocEVdn+w2X4iamprQ3d1dJsTFxUVovrdFSdCtpm0px8fHeHR0JJiuk6kSdPKdn58Xqh01Y7VFgEdJUAhdUHseMjSUzkMf0oWnT5/ysY8CehAaQR+Jgt4EgKPCwkI+9lFQWVlJYv4GgHdEQcSkrjnMkImIiCBBSwBgIhWUToeuqakpPt6goc9gmm8Y30jFEDYAsE1l+DGhOfLva/v7EX0jxoKCAn6dQUKnW00x+JoXIuVHCjL0AiER084L4KEX6ycKpsc5PT3N7/WgLCwsYExMjCjmZwB4hRegDfp4/wcVCjo9lpWVCeP96Oio0Jnvy8bGxoSRqqKiAp88eYKmpqYk5C8qYnzCangDAHIBYBYAzqVz1wMY9ZkFACgEgLf4RF8GalrUiWm8oJnpvozuR6PZu3xCSvwDMJZbhJvwXPsAAAAASUVORK5CYII=",
  77: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAS9SURBVGhD5ZlZLHVXFMf/5tKSCtVKDVGp4UVV8YCmpS1CxAORvkgqfZG02oc+1JioBxVDkNJIVbyQilfzUNNHYog0YkyjYkhQQw1BUOxmnTgn91vO9Z37ubjSX7Je7l5r7/U/+5y9194X+B/zBoAAAB8C+BTAJw9gNA6N9x4AF57Qy/AqgK8A/A5gB4B4RPsHwACAbwA48ES18DmAP6mzkJAQkZmZKRoaGkRLS4vo6ekR3d3d9240Tmtrq2hsbBQ5OTkiLCxMFvcXgC94wrfxIwWGh4eLjo4OYUr09fWJyMhIWdhPAMx48pwics7IyBCXl5e8P5MhKytLFvUzF6BLiizmKZCbmyuL+pILIV4DsBoaGsrjTJqoqCgS9DcARy7oa1Lb39/PY0yaiYkJYWFhQaK+54KeBQcHc39Vzs/Pxfb2ttjZ2bnVyGd3d1f5Fk9OTsTW1tYNP24Ut7+/z4fVy/Ui8YfuAkGb1iG9k1oYHR0Vzs7OmszLy0tsbm5KcZWVlcLJyemGDzfyiY+P58Pqpbi4mASdA/CUBb1Pr1tTUxP3VWVgYIBverfa2tqaFFdQUHCjTZ8FBATwYfXS3t4ux1FVIfEx/dDZ2cl9VRkZGRG2traq5ujoKKytrZXE/P39xeHhoRRXWloqbGxsbsTY2dlJcWZmZkpcWloaH1YvOg84WhZENZO0M2vh9PRULC8vS7aysqIYzQSJpeSoP3t7ezEzM6PEHRwciKWlpediyDY2NkR1dbUwNzeX4oKCgsTR0dFzY95Gb2+vLOgzWRAVglK5cVcSEhKUp1xfX8+bVaEZ9PT0lGJoxubn57nLrdyboLq6OkVMdHQ0b9ZLamqqEldUVMSbX8i9CNrb2xMuLi5Sx7QvTE9PcxdVhoeHFTHe3t7i7OyMu7yQexFUWFioJJacnMyb9UIzKcfV1NTwZk0YXRBtmO7u7srsTE5OchdVyE8WQ9/Q8fExd9GE0QXR/iUnFhERwZv1kp6ersRlZ2fzZs0YXVBiYqKSWFlZGW9WhWbD1dVViRsbG+MumjGqIKq7HBwcpA6trKzEwsICd1Glq6tLEePn5ycuLi64i2aMKqitrU1JzNfXV1xdXXEXVehIL8cZUhWoYVRBeXl5SmIpKSm8WS86x2hRVVXFmw3CqIJoiZYTo6VbC1Q6yZUB2eDgIHcxCKMKCgwMVBKrra3lzaqsrq4qdRvZ3NwcdzEIowmi74VeHTc3N8m03hAtLi4KDw8Pae/y8fER6+vr3MUg1AQZVG3rQssvVcZkWm+JaEWTY2hTvitqgiLpB63nIVND7Tz0Af3Q3NzMfZ8ENBHXgj6SBb0F4IiOyE+RiooKEvMvgHdkQcSoIXWYKREbG0uCZgBY6Ar6jpbSu9RUj8Hs7Kx8h/GDrhjidQCbtAw/Ja6P/Hv6/j+iO2KRn5/P40ySkpISeTH4lgvR5RdyMvUFQkdMExfAoQ/rV3Km6RwfH+d9PSpTU1MiKSlJFvMbgFe4AH3Q5f0GLRQxMTHSJSEdFaiQpJ35oWxoaEgqqcrLy0VcXJywtLQkIVu0iPGEtfAmgFwAkwDOrp/KYxntM1MACgC8zRN9GWjTop2YyguqmR7KaDwqzd7lCanxH0Ua3XHMRx2mAAAAAElFTkSuQmCC",
  78: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUsSURBVGhD5ZlZLHVXFMeXeZ5VGyWNL0G8KIoH81REggcifZE0acJDaR9KSpAYEq0pSJM+GCIRUiGejDEPIWYRJCKaUKLUFIKoYjfrxN45d7v3fOd+vbjSX7Jezlpr7/W/5+y91zkX4H/MRwDgCQDBABAFAJEvYDgPzvc5ADjwBX0IZgDwLQAMA8AxAJBXtFMAGAOA7wDAki9UDl8BwCYO5ufnR3JyckhLSwvp6uoig4ODZGBg4NkN5+nu7iatra0kLy+PBAQEUHG/A8DXfMFS/ISJgYGBpK+vj2gTIyMjJDw8nAr7BQB0+OJ5fsbgzMxMcn9/z4+nNeTm5lJRv/ICxKRQMW+B/Px8KuobXghiDgB/+Pv783laTUREBAo6BAAbXlAGqh0dHeVztJr5+Xmip6eHon7kBU36+vry8Uq5vb0lR0dH5Pj4WNIw5uTkROVa3NzcJMPDw8Jutra2Rh4eHvgQWTxuEsviDQIPrQt8JuUwMzND7O3tZZmLiws5ODhQyJ+cnCQhISHEwMCAnTM6OjrE29ubdHZ2KsTKoby8HMe4BYDPqCBvHLStrY2PVcrY2BgrRI7t7u6y3NnZWWJsbPwkRmzt7e0K872P3t5emotdhUAYXujv7+djlTI1NUVMTEyUmo2NDTE0NGTFeXh4kIuLC5YbFhbGfI6OjqS6uprU1dURT09Pdt3Z2ZlcXl4qzCmF6AeOpoKwZxKeZTnc3NyQ7e1twXZ2dpjhnUCxKArHs7CwENYG5fDwkJibmws+fMSGhoaYD3Otra2ZKHys5YLjPOZ9SQVhIyi0G/+V+Ph4VlRTU5OCb39/X7iL6DMyMhI2DDFubm4sd3p6WsEnxbMJamxsZAVFR0fzbnJ3d0d8fHyUCkYBdG3Z2tqS09NThVwpnkXQ2dkZcXBwEAbGc2F1dZUPEcDH2tTUVIjDtZaamkrS09MFEVRoQ0MDnybJswgqLS1lBSUnJ/NuBXDbtrOzY/HU9PX11RaDaFzQ9fW1sDPRu7O4uMiHMHDxR0VFCcXzgnCj8PLyUrtb0bggPL9oUUFBQbybgR1DcHAwi42MjBTOpaWlJZKWlsaum5mZka2tLT5dJRoXlJiYyIqpqqri3QzxgWxlZfWkgwgNDWX+rKwsBZ8UGhWE/ZqlpaUwILYyGxsbfAijrKyMFYytD09JSQnz492Ti0YF9fT0sCLc3d0lG8zi4mIWi2uFJzs7m/mx4ZSLRgUVFBSwIlJSUni3Ah0dHSwWNwXxfNgeubq6Mn9GRoZCrhQaFYRbNC0Ct24pzs/PiZOTE4vHdYR3rba2Vrhj9DrawsICn64SjQoSF1JfX8+7n4AfOXAXExfPW1FREZ8micYE4XrBZx1/dTS5X4iWl5dJQkICa1SpYcfd3NzMh78XZYLU6rbFXF1dCa0+mqo3U1Xs7e2RiYkJYV5sldTNpygTFI4X5L4PaRvK3oe+wAvqvilqC3gjHgWFUkGfAMAl7jhvkZqaGhTzDwC8o4KQGak+TJuJjY1FQWsAoCcW9IOurq7QLL4l1tfX6TeMIrEYxBoADtRpObSBx1f+M1X/H+E3YlJYWMjnaSUVFRV0M/ieFyKmDoO0fYMQiWnjBfDgwmrAYLydc3Nz/FivysrKCklKSqJifgMAY16AKvDj/Z+4UcTExJDKykrhVWF8fFw4mV/KsJPAlgo/SMbFxdFX979wE+MLlsPHAJAPAIsA8Le473oFw3NmBQCKAeBTvtAPAQ8tPImxvcCe6aUM58PWzJUvSBn/Aj14VQYM3zRCAAAAAElFTkSuQmCC",
  79: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUbSURBVGhD5ZlrKL9nGMcvZzJyGtMw/7FEGmt4gWVsjjm8cEjxYjUv1GZLygiFF7YcosbSTN7QpCg5H+a0lENaQlgtTJnjyCkj7nXdue+e5/b7/Tx/++Gnfeqqf7/ruu/7+v6f57nu674B/I95EwDeB4CPAOBTAPjkCQzXwfV8AMBeTOghmAPAFwDwCwDsAwB5RvsbAMYA4CsAsBQTVUIqAPyOk/n7+5O8vDzS3NxMurq6yNDQEBkcHHx0w3W6u7tJS0sLKSgoIIGBgUzcHwDwmZiwJr7FgUFBQaSvr4/oEiMjIyQ0NJQJ+x4A9MTkRb7D4KysLHJ9fS3OpzPk5+czUT+IAqSkMDEvgcLCQibqc1EI8gYA/BkQECCO02nCwsJQ0A4AWIuCvkS1o6Oj4hidZnZ2lhgYGKCob0RBv/r5+YnxKrm8vCR7e3tkf39fo2HMwcGB2m9xaWmJVrLJyUka+1Bui8Rv0gKBm9YxvpNKmJqaInZ2dors1atXZHt7WzYeRWAJ1tfX5/sMxmZnZ5Pz83NZrBLKy8txjksAeIcJ+gAnbW1tFWNVMjY2xhNRYpubm3xsR0cH0dPTuxPDLCQkhJydncnWu4/e3l42HrsKysf4Q39/vxirEnxFzMzMVJq1tTUxNjbmCXp6epLj42M6bnd3l/qZz9vbmxQXF5Pk5GSZqJycHHFJjUj+gyOYIOyZ6KughIuLC7K+vk5tY2ODGz4JFMuStrCwIIuLi3xcXV0dT9rV1ZUcHR1xX0ZGBveZm5uTnZ0d7ruP4eFhNjacCcJGkLYb/5W4uDieWFNTk8yXnp7OfZmZmTLf6uoqMTQ05H6lrz/yaIIaGxt5QhEREaJbJlYsQCcnJ8TR0ZH7sXdUyqMIOjw8JPb29nRi3BcWFhbEEJKWlqb2CeHrZ2Njw/2pqakyvyYeRVBZWRlPJikpSXRTqqureYy7uzu5urrivra2Nu5Di42NlY3VhNYF4d7h7OzMn87c3JwYQsH9yMrKiicdHR1NOjs7SW1traz6oYWHh4vD1aJ1QfgBs0SCg4NFt4z29na1+xBWN/bv+Ph4cahatC4oISGBJ1JVVSW674Dr4FnLxMSEjnFxcaElHUWwebAiKkWrgrAHs7S0pBMaGRmRlZUVMUQta2trZHl5mVY4xMfHhwsqKSkRw9WiVUE9PT08CQ8PD3JzcyOGcDBx3G/Gx8fvdPVbW1vE1NSUz4XzKkWrgoqKingSKSkpoluGZGFaPLC7YOTm5nKfg4MDOT09lY3VhFYFYYlmiWDp1gTuVba2tjzey8uLlJaWyjoINCzvr4NWBfn6+vJEGhoaRPcdMEaavGi4+Wp6bVWhNUG4MB6unJycqCm9Iaqvr6fnJKkQNzc3UllZKYYqQpWg1+q2peDZBd93NHUnU1Vg/PT0NBkYGKAb8UMOdgxVgkLxB6XnIV1D1XnoQ/wB+6mXCD6IW0EhTNBbAHCKFeclUlNTg2KuAOBdJgiZuq8P01WioqJQ0CIAGEgF5eAtDH6oLwm8Cru9wyiRikGsAGAby/BL4vYUfKju70d4R0xvYl4CFRUVrBh8LQqR8iMG6XqBkIhpFQWI4If1Ewbj45yZmRHnelbm5+dJYmIiE/MzAJiKAtSBl/d/YaGIjIykbQm29Nj64878VDYxMUFbKmxYY2Ji2FXXLhYxMWElOABAIQDMAcA/0t7rGQz3mXkAKAWAt8VEHwJuWrgTY3uBPdNTGa6Hrdl7YkKq+BdwAVriUMH/5QAAAABJRU5ErkJggg==",
  80: "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",
  81: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAASySURBVGhD5ZlZSH9VEMfHPUglcSlQUQJBffin5IJL5FIpog8qpo+pL4qZYOCCaz5Y7hsERi+CouiTuC+5IqSSIaYPQQiitrgkbuA6MZf/uZw7+nOL9P7oA/PgPTPnztd775w55wfwP8YRAF4BwAcA8BEARD6D0X3ofu8BgBNP6Cm8CQBZAPADAOwCAL6g7QPAFAB8AQC2PNGHkAIAv9Jk/v7+WFBQgO3t7djX14djY2M4Ojr6nxvdp7+/Hzs6OrCoqAiDg4OFuN8A4DOe8F18TYEhISE4NDSEemJiYgLDw8OFsBYAMOHJc74h5+zsbLy6uuLz6YbCwkIh6lsuQOZTIcYYKC4uFqLSuRDCGgA2AgICeJyuiYiIIEF/AoAdF/Q5qZ2cnOQxumZxcRHNzMxIVD4XNOvn58f9H8TW1hbOzs7iyMiIcoOjoyPucicU293drRSgy8tLPnwvr4vEz3KBoEXrkN7Jx0BCUlJS0NbWVrNmuLm5YWVlJXe/lfX1dbSwsFDi7Ozs8Pj4mLvcS3V1NcWfA4CbEORLE3Z1dXFfg+zv76Onp6dGCLecnBwepuHs7AzDwsJUfxcXFzw5OeFu9zI4OCjmoK5CIYwuDA8Pc1+DlJWVqYlYWlpifn4+trW1YUJCgkbU0tISD1WgJxMZGanxdXV1fZKgqakpMccnQhD1TMrK/FCkVVtZxWUCAwPVsdraWs3Y7u4u1tfXo4ODg0bMvxE0Pj4u5vhYCKJGUGk3HkpQUJCaCLVDMmlpaepYVVWVZoxeQ1kEfW8mJiYvLyg3N1dNigQIDg4O0MPDQx2jCiaTmpqqXLeyssKmpibs6elRfV9U0ObmpibxmJgYzMrKQi8vL/UaJc/Jy8tTKuPKyoryt24EEXt7e+jj46MmJFt6ejp3V6DKJtPZ2akPQefn55iRkYHW1tY3xJA5Oztjc3MzD7uBbgSVlpaqibi7u+PAwACura1hXV2dUsbFGCV8F7oQRK8areoikd7eXs241Norr+T19bVmXEYXgubm5tQkbGxslLVFRlro0N7eXukqDKELQbRrFElQ+aWKJ0PbZjFOT5ILltGFoI2NDbWhJKM2SCYpKUkd8/X11f8rRyQnJ6uJkFHFa21txfj4eM31lpYWHqpBN4K2t7fR29tbkzy32NhYpbzfBZ0kCX9HR8eXE0Ts7OxgZmYmOjk5aYTQf7qkpAQvLi54yA2oQtK2gWKoqT09PeUu93KboEd32zJUxufn59Ud6+HhIXcxCImmTR3ZU8QQtwkKpwuP2Q/pidv2Q+/TBdrXGyP0IF4L+lAIegcAjisqKrivUdDY2EhiLgDgXSGI+DE0NJT7GgXR0dEk6BcAMJMFfWlqaqp83MbE6uqqaIa/ksUQbwHAH3TGZUzExcWRmL8N/X5EZ8RYXl7O43RJTU2NKAY5XIjMd+Sk9wIhieniAjj0YX1PzvQ4FxYW+FwvyvLyMiYmJgoxnQDwBhdgCDq8/50KRVRUlHK+RrvS6elpZWV+LpuZmVHOvBsaGpSDGHNzcxLyFxUxnvBDeBsAigHgJwA4k/u1FzBaZ5YBoAIAnHmiT4EWLVqJqb2gnum5jO5HrZkHT+g2/gGcv268sUR/UwAAAABJRU5ErkJggg==",
  82: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAVLSURBVGhD5ZlpLG1XFMeX6WkUKVJtgpCKZ4goqSGGpqUDEZL3nigfDV8eqoIPiLE+aE0hiETTLxJCfBPz8Myij1TjpYRGJQStGiqGRAmrWSf2yTnbvdfl3nJv+kvWl73XOnv979ln7XXOBfgf8y4AeAHAxwDwOQB89gBG69B6HwKALZ/QfXgbANIA4BUA7AEAPqIdAMAYAHwDAJZ8ouoQDwC/0cX8/PwwNzcXW1pasKurC4eGhnBwcPA/N1qnu7sbW1tbMT8/H4OCgpi43wEggU9YFd9RYHBwMPb19aEuMTIygqGhoUxYPQAY8MnzfE/O6enpeHl5yV9PZ8jLy2OiGnkBUr5iYvSBgoICJiqZF0KYA8CGv78/H6fThIWFkaAdALDiBX1NakdHR/kYnWZubg6NjIxIVA4vaNLX15f3V4utrS2cnJzEgYEBYYHj42PeRSHr6+s4NjYmVLT5+Xk8OzvjXdTiukj8Ii0QdGgd0Z68CyQkPj4eLS0tZWeGo6MjlpWV8e4i09PTGB4ejmZmZrI4V1dXrK+v591vpaKiguLPAcCRCfKhC7a3t/O+Sjk4OEA3NzdZQrxlZGTwYcLdMDExueErtezsbD5MJb29vSyWugqBT2mgv7+f91VKcXGxmMCTJ08wJycHm5ub8cWLF7LkaCsxaEu5uLiIc1ZWVkLydC0PDw9Z3NTUlGw9VdC2vY77kgminkk4mdVFcmoLp7iUgIAAca6qqkocp0OajRsYGMgK0P7+Pjo4OIjzmZmZ4txtDA8Ps7gvmCBqBIXtoC6BgYHi4tQOSUlKShLnysvLxXFqYZydndHa2hq9vLxkMURsbKwYR8+mumhFUFZWlrg4CWAcHh7KthVVP8bV1RVeXFzgzs4Obm9vi+OMkJAQMU7R86cMrQja3NyUJR4ZGYlpaWno7u4ujiUmJvJhSpmZmUFjY2Mxlr/rqtCKIIL2vbe3t5iE1JKTk3l3pWxsbMh+HNqOdCfVRSuCzs/P8eXLl2hubn5DDJmdnR3W1dXxYTcgMXz5v0uFI7QiqKioSEzAyckJe3p6cGlpCaurq4Uyzuba2tr4UJHV1VV8+vSpTExjYyPvdisaC6KtRmcIS6Kzs1M2L2nthS1JxYBnZWVF+CGkYhoaGng3tdBYELUuLAkLCwvc29uTzUsOOrSxsRG6Cilra2syMYaGhsKhfF80FkRvjSwZU1NToeJJoddmNk93UiqYyrqnp6c4T30gNbWaoLEgepCl/Ri1LlKkB6SPj49syyUkJIhzZNQR0A9CW3B5eVk0ReeUMjQWRMTFxckSo4rX1NSEz58/l41Lu+c3b97I5sio46Y2iB9/8E6BfkG+oeQtKipKKO+M1NTUGz7K7NmzZ7L1VKEVQcTu7i6mpKSgra2tLBlqMgsLC28cjnT37O3thXlVRj7UdaiLIkF37ralUBl//fq1+MZ6dHTEuwicnJwIdnp6qtLI5y5vr4oEhdLAXd6HdAlF70Mf0UBHRwfvqxfQjbgW9AkT9D4AnJSWlvK+ekFtbS2JuQCAD5gg4id6H9FHIiIiSNCvAGAkFZRNLQg93PrE4uIia4a/lYoh3gGAP+kblz4RHR1NYv5W9v8RfSPGkpISPk4nqaysZMUggxci5Qdy0vUCIRHTzgvgoQfrR3Km2zk7O8tf61FZWFjAmJgYJqYNAN7iBSiDPt7/QYWCPtvS9zV6Kx0fHxdO5oeyiYkJ4XteTU2N8CHm+kPKX1TE+ITV4T0AKACAnwHgH2m/9ghG58wCAJQCgB2f6H2gQ4tOYmovqGd6KKP1qDVz4RNSxL9LHPuoFek5awAAAABJRU5ErkJggg==",
  83: "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",
  84: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUwSURBVGhD5ZlrLGVXFMdXMZgUQ6g2ESFNJPigI2UwNEofJh4fEOVjjQ+IqoRkEAYlUY8RjKSJph+IZ0gk4v2oZyap8ShJSTRpIsJoMSoY8QirWTvOzjnbvdc1Wvfe9JesL3uvdfb6n3vO2mufC/A/5j0A8ACATwDgcwD47BaM1qH1PgIAezGht+FdAEgBgJ8BYBsAUIe2AwBjAPAtAFiJiWpDHAD8Thfz9vbGrKwsbGxsxK6uLhwaGsLBwcH/3Gid7u5ubGpqwpycHHz48KEk7g8A+FpMWBPfU6C/vz/29fWhPjEyMoJBQUGSsBoAeEdMXqSEnFNTU/Hs7Ey8nt6QnZ0tifpBFCDnK0mMIZCbmyuJShCFEBYAsPrgwQMxTq8JDg4mQX8BgI0o6BtSOzo6KsboNdPT02hsbEyiMkVBk15eXqK/Vqyvr+Pk5CQODAywBfb390WXK5mamsL29nbs6OjA3d1dcVojF0XiV3mBoE1rj57J60BC4uLi0MrKSrFnODk5YXFxseiulsXFRTQzM+Pxs7OzootGysrKKO4EAJwkQZ50odbWVtFXLTs7O+jq6qoQIlpaWpoYdgmqpL6+voq4+fl50U0jvb29Uix1FYxPaaC/v1/0VUt+fj5PwNTUFDMzM7G+vh6joqIUyc3NzYmhCgoLCy/diOsKGhsbk2K/lARRz8R2Zm2R7dpsF5fj4+PD5549e6aYkzMzM4MmJiY3FjQ8PCzFfiEJokaQtRva4ufnxxOgdkjO48eP+VxpaaliTuL4+Bg9PDy4n1yYTgSlp6fzBEiABFUoFxcXPkfVTxVPnjzhPikpKayQ6FTQ2tqaIvHQ0FCWmJubGx+Lj48XwxgTExPcJzAwkBUYOzs73QoiXr9+jffv3+eJyC0hIUF0ZxwcHPAbQcVkeXkZDw8PFaVfJ4JOTk4wKSkJLSwsLokhc3BwwOfPn4thmJyczH1KSkrY2KtXr/DevXu6FZSXl8cTcHZ2xp6eHlxaWsKKigp256W5lpYWHiPbLzAsLIyPU4GwtrbmcysrK3xOG24siB41GxsbnkBnZ6diXtbas0eSoHaIhEvjdGBsaGjAuro6rK6uxrt37/K5oqIiNq6tsBsLevHiBV/c0tISt7e3FfOyjQ5tbW3ZO7KxsaFyz9Fkzc3Niuuq48aC6NQoLUo9GFU8OXRslubpl3zz5g0TJCZ8ld2aoNXVVbxz5w5fmNogOTExMXzO09OTjdEjR48SvXtyo9iMjAw0NzfnMYmJiWyOmlZtuLEgIjY2VnE3qeLV1tZiZGSkYrympkYMvcTR0ZFuiwJBpdbd3V2RvGjh4eGsvF8FXUt+fLiqoRX5VwQRW1tbbF+xt7dXCHF0dMSnT5/i6empGKKSzc1NdnMojkzbR01ClaBrd9tyqIzTiVM6se7t7YkuGjk/P2eFg7oIsut+bVIlKIgGrnMe0idUnYc+poG2tjbR1yCgH+JCUKAk6AMAOKDToyFSVVVFYk4B4ENJEPFLQECA6GsQPHr0iAT9BgDGckEZRkZG7OU2JKgiXjTD38nFENYA8Cd94zIkIiIiSMzf6v4/om/EWFBQIMbpJeXl5VIxSBOFyPmRnPS9QMjEtIoCROjF+omc6ed8+fKleC2dsrCwgNHR0ZKYFgAwFwWogz7eb1ChCAkJYd/X6FQ6Pj7OdubbMvqgQn+6VVZWsg8xF2eqTSpiYsLa8D4A5ALALAAcy/s1HRjtMwsAUAgADmKibwNtWrQTU3tBPdNtGa1HrZmLmJAq/gFAbQLaVXM8XwAAAABJRU5ErkJggg==",
  85: "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",
  86: "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",
  87: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAVHSURBVGhD5ZlbSB1nEMdHjbFJjUS0tiJRaVDUB6tpVKIprbZVUfOgweqLUPUlYq2QPhijJqkP1njBeCtJ6YugKL6J16iNd6iKFqEaKDageGmNsYk34nXKLPmW3ck5J2vOqR7pD+Zlv5n95n++3fnm2wPwP+Y9APAFgE8A4AsA+PwQjOah+T4CACee0NvwLgCkA8AvALAMAHiEtgIAvQDwLQDY8US1kAgAf9DNAgIC8MaNG1hbW4vNzc3Y1dWFnZ2d/7nRPC0tLVhXV4c5OTkYHBwsxP0JAF/zhA3xAwWGhIRge3s7mhOPHj3C0NBQIawSACx48pxCcs7IyMC9vT1+P7MhOztbiPqRC1DylRBzHMjNzRWiUrkQwhYAZgMDA3mcWRMWFkaC/gYAey7oG1Lb09PDY8ya0dFRtLKyIlFZXNDAxYsXub8m5ufncWBgAB8+fChNsLa2xl0kNjc3cWlpCZeXlw3a06dP8fnz5zxcL6+KxG/KAkGb1io9kweBhCQmJqKdnZ1qz3Bzc8OCggLujuXl5ejg4ICOjo4GjXyio6N5uF6Kiopo3m0AcBOC/CmRhoYG7quXlZUV9PLyUgnhlpmZqYrJz89/zUef+fr6qmIN0dbWJuKoq5D4jC50dHRwX73cvn1bnvzkyZOYlZWFNTU1GBcXp0psfHxcjikpKUEbGxs8deqUyk6fPo329vZoYWEhxyUnJ6vmM0Rvb6+ICxeCqGeSdmatKHZtaRdXEhQUJI+RCMGLFy/wyZMnODMzo7LFxUWsrq5GS0tLKebChQu4vr6uuqchuru7xXxfCkHUCErthlYuXbokJ03tkJKUlBR57O7du6oxXayurkrvHfnTij1+/Ji7GMQkgq5fvy4nTQIEVJ08PDzkMap+byIpKUn2Lyws5MNvxCSC5ubmVIlHRUVheno6ent7y9e0vAeDg4Oy//nz53Fra4u7vBGTCCKePXuGfn5+ckJKS01N5e46CQ8Pl2Pu37/PhzVhEkHb29t47do1tLW1fU0MmYuLC1ZUVPAwFWNjY7I/vUMbGxvcRRMmEXTr1i05GXd3d2xtbcWpqSksLS2VyrgYq6+v56Ey9IMIv5s3b/JhzRgtiB412jdEMk1NTapxRWsvPZL7+/uqcYJWw9nZWfYbHh7mLpoxWtDQ0JCcyJkzZ6T+S4lio5PaGOoqONT3CR/qOHZ3d7mLZowWRKdGkQzt/FTxlNCxWYzTSnLBBB3phY+WamgIowXNzs6itbW1nBC1QUri4+PlMX9/f52PnOIYjVVVVXz4QBgtiEhISJATIqMX/MGDBxgbG6u6XllZyUPx5cuXcmdA1tfXx10OhEkELSwsoI+Pjyp5bjExMVJ559AKi76NjKqjMZhEEEEHsbS0NHRyclIJOXfuHObl5eHOzg4PkZienkZXV1fJz9PTU/pxjEGXoAN320qojFPZFSdWajYNQRWNumkyOsUaiy5BoXThIOchc0LXeehjutDY2Mh9jwW0EK8EfSoEfQAA63REPo7cu3ePxOwAwIdCEPHr5cuXue+xIDIykgT9DgBWSkHfUSk1pqc6CiYnJ0Uz/L1SDHEWAP6iHfw4ceXKFRLzj77/j+gbMd65c4fHmSXFxcWiGGRyIUp+IidzLxAKMQ1cAIderJ/JmZZzZGSE3+tImZiYwKtXrwox9QDwDhegD/p4v0iFIiIiQvq+RqdSaiRpZz4s6+/vl/50Kysrkz7EnDhxgoQsURHjCWvhfQDIBYAxANhS9mtHYLTPTABAPgC48ETfBtq0aCem9oJ6psMymo9aMw+ekC7+BYJqVc/sZKnHAAAAAElFTkSuQmCC",
  88: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAWwSURBVGhD5ZlbLJ1ZFMeXS5VSdcmYSbUhkxD1YKqDpupaZtoIEsTwIpmSaBuMpKOhqdaliRktUSSTEIloNIT0oXGp+zUENRUpHqTz4jZM1ZR6GbRrsr7YO9+3cZzDGY7ML9kve+317fU/395rr/0dgP8xXwCACwB4A0AgAAQcQKN5aL5vAMBGDGgvmAJAAgC0A8AiAOAhtiUA6AKAnwDAXAxUHaIBYJIe5u7ujmlpaVhZWYl1dXXY2tqKLS0t/3mjeerr6/HZs2d479499PT0ZOL+AIAfxYBV8Qs5Xr58GV++fIm6REdHB/r7+zNhxQCgJwYv8isNTkpKwk+fPonP0xnu3r3LRP0mCpDzAxNzFEhPT2ei4kQhhBkATHl4eIh+Os2VK1dI0AIAWIqCEkltZ2en6KPTvHr1Cg0MDEhUqiio183NTRyvFrOzs9jb24vNzc3SBB8/fhSHbMvk5CS2t7dL2WxsbAw/f/4sDlGLzSQxIk8QdGit0JrUBBISHR2N5ubmijPDzs4Oc3JyxOEcEu/j44PHjh3jPnp6eujq6orPnz8Xh+/Ko0eP6BlrAGDHBLnSQ6urq8WxO7K0tIROTk4KIWJLTk4W3XBwcBCNjY23jJW3mpoa0U0ljY2NzJeqCgk/6mhqahLH7khGRgYPwMjICFNTU7GiogLDw8MVwb1+/Vrh5+fnx22nT5/GgoICLC0tRRcXF95/9uxZXF1dVfipoquri/l+zwRRzSStZXWRndrSKS7n4sWL3JaXl8f7FxYW0MzMTOqnJdbW1sZt09PTaGFhwf0GBga4bTfoOZt+3zFBVAhK5Ya6XLp0iU9O5ZCc2NhYbsvNzeX9c3NzaGJiIvUfP34c379/r/BzdHTkfv39/QqbKrQi6Pbt23xyEsD48OEDOjg4cBslAMbGxgZeuHCB28rLy7mNBLC9ZWVlJe1RddGKoJmZGUXgQUFBmJCQgOfOneN9169fF92kZX3ixAnJTnsvJiYGb9y4IYlgfmVlZaKbSrQiiKAlc/78eR6IvMXFxYnDOfTWrK2tt/gYGhpqLIbQiqC1tTW8efMm3+Ris7W1xaKiItFN2vyBgYFS8KIPJQr6gTStVrQi6MGDBzwQe3t7bGhowImJCczPz5eWErNVVVVxH6revb29uS0gIEA6lyi1x8fH835TU1N8+/atYj5V7FsQLTVLS0sewIsXLxR2WWkv/eKspJGdF3jq1Cmcn59X+Pn6+nJ7SkqKwqaKfQvq6+vjE588eRIXFxcVdnngtFdYxqIUzvqp9BF5+PAht9PbU5d9C6JbI5uYzhPKeHLo2szs9CaZ4OzsbN5Pb07kzp073E4Fp7rsW9DU1JSisKQySE5kZCS3UcHJllxtbS3vp6Qgn29lZUVxDCQmJsqeqJp9CyKioqL45NQo45WUlGBYWJiiv7i4mPssLy/jmTNnuI32Eb21wsLCLel/eHhYMZ8qtCKIyhhnZ2dFEGILDg6W0rscWq6UxcSx8paVlaXw2Q2tCCLevXuHt27dQhsbG0VAVC3fv38f19fXRReJkZERDA0N3XKGUcX99OlTcfiubCdI42pbDqVxOk/YjZX2gzpQMunp6ZHmffPmzZ6/Mm0nyJ86NLkP6RLb3Ye+pQ5Nb4q6Ar2ITUG+TNBXALBKGeco8uTJExKzDgBfM0HEgJeXlzj2SHDt2jUSNAYABnJBP+vr60ub+ygxPj7OiuEsuRjCAgDmNSk5dIGQkBAS8/dO/x/RN2LMzMwU/XSSx48fs2SQLAqRU0qDdD1ByMRUiwJEaGOV0WB6nUNDQ+KzDpXR0VGMiIhgYqoAwFgUsBP08f5PShRXr16Vvq/RrbS7u1s6mQ+qUSVBf7rRB0n6ELN5df+LkpgYsDp8CQDpAPA7APwjr7sOodE5MwoA2QBgKwa6F+jQopOYyguqmQ6q0XxUmjmIAW3Hv3rIzVVxlaOqAAAAAElFTkSuQmCC",
  89: "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",
  90: "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",
  91: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAS0SURBVGhD5ZlJSHRHEMfLPWiIcUvihiDoTSW4gRriliiCIC7BYyQnTYygYFzBeElQcSGQgwa8KG4n9zWuBz+VGJeoEBGXSzSKwbiAa4V62M2b+saZ+Qxx3pAfNDKvq/vVf3q6qroF+B/jAQDBAPARACQCQMILNHoPvS8EAN7jDj0HJwD4EgB+BoATAEAztlMAmAaArwHgHe6oKWQDwO80WXh4OJaUlGBbWxv29/fj+Pg4jo2N/eeN3jMwMIDt7e1YXl6OUVFRQtwOAHzOHTbEdzQwOjoah4eHUUtMTk5iXFycEPYDAFhx5znfk3F+fj7e39/z+TRDaWmpEPUjF6DmMyHGEqioqBCivuBCiLcB4CAiIoKP0zTx8fEk6AgAXLigr0jt1NQUH6NplpaW0MbGhkR9wwXNhYWFcXuT2N/fV74Iant7e7zbKHNzc9jd3a0EoLu7O95tlMcg8as6QFDS+pt+k2/Czs4OZmVloZOTk8wXjo6OmJ6ervSZwu7uLtrZ2SljXVxc8OLigpsYpaamhsbfAICfEPQhTdjZ2cltn2R7exu9vLykEN6ob3Nzkw/T4fr6GmNjY+UYHx8fvLy85GZGGRoaEnNQVaEQSw9GRka47ZMkJCRIR9zc3LCsrAwLCgrQ3t5ePqcAc3t7y4cq0Mqo56Dm6+v7LEHT09Nijk+FIKqZlMxsCisrK9IJW1tbnJ2dlX2tra06TvIv6eTkBOvr69Hd3V3H7t8ImpiYEHN8IgRRIaiUG6bQ0tIinQgKCuLdGBgYKPtzc3N1+mgV1SL8/PzQysrKvIIeN6HSaA9wUlNTZX9MTIxOX05OjvLcwcEBm5qasKenR9qaTVBzc7N0Qt8KqfeGv7+/zj4qLi7G7OxsXF9fVz5rQhAlM+EEhVzaU4KDgwN0dnaW/Z6ennh2dib7KbKp6ejoML+gh4cHnXBLe4ZK/K6uLgwODpbPqdHmp0DwFJoQRBjKQ+pE6+3tjefn53y4RDOCCKoGMjMz5U+MMn1hYSFWVlZKJwMCAgyWM5oSJDg6OsKtrS3lL6EOzYmJidxcB00Iurm5UVZnYWFBOZJzQkNDpZMU1QyhCUG0ydWZXl0D9vX1yURJjUQbQhOCiOTkZOmIq6urUssVFRXpBIS0tDQ+7DU0I2h5eVnJ9sIZ3kJCQvDw8JAPew26SRJjPDw8zCeIoAMdXXGphdBPMS8vD09PT7m5Xnp7e5VjA61OZGQkXl1dcROj6BP0RtU2Z21tDUdHR3F+fh6Pj495t0GoLKJDHbXniCH0CYqjB7zUtxT0nYdC6QGd6y0RWohHQR8LQR8AwEV1dTW3tQgaGxtJzC0A+AtBxCt+drEUHlPIbwBgoxZUZG1tbTQRao2NjQ1xj/GtWgzxLgAc0h2XJfF4Ov7rqf8f0R0xVlVV8XGapLa2VgSDAi5ETTMZaT1AqMR0cgEc2lg/kTEt5+LiIp/LrKyurmJGRoYQ0wEAb3EBT0GX939QoEhKSsK6ujocHBzEmZkZJTO/VKM7P7rzbmhowJSUFOUeEAD+pCDGHTaF9wGgAgB+AYBrdc1mhkZ5ZhUAqgHAmzv6HChpUSam8oJqppdq9D4qzQK4Q/r4B82YdJtvxLRyAAAAAElFTkSuQmCC",
  92: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAVmSURBVGhD5ZlvSJV3FMePWTly2Ez3pzQmihmRyjB7oY1puhmBkFrDl8JepXNigitTdL7ZsMDA8IUNIkhmvaw0/7QyZ7iStZSlNv9QEk7nn+UyaJqe8X24vx/Pc7r3dm+36ZV94CA855znOd/n/v6c5yfR/5h3iSiGiD4molQiSlkGw3PwvFgiek8W9Dr4E1EeEf1IRFNExCtoM0TUTkRfEVGALNQVsonod9wsPj6ejx49yufPn+fLly9zW1sbt7a2/ueG51y5coXr6+v5+PHjnJCQoMQNE1GOLNgZ3yIxMTGRr169yt7E9evXOTk5WQmrISIfWbzkOwTn5+fz4uKivJ/XcOzYMSWqVgow87kSsxooLS1Vor6QQsDbRDS6e/dumefV7N27F4ImiChQCvoSam/cuCFzvJru7m729fWFqK+loJ927dol413i0aNHxouAPXz4ULodgrz29nZjRbt79y4/f/5chriEbZH41bxAYNP6G2PSHYaHh/nQoUPs7++v94sNGzZwZmam4XPErVu3OC0tzYg17zVRUVFcU1Mjw19JVVUV8ueJ6EMl6CPcsKGhQcY6ZHBwkLds2WIpyGzw9fX1yTTj11i3bt1L8WYrKiqSaU5pampSuegqDJJwobm5WcY6JCUlRRcQFBTEJSUlXFBQwOvXr9fXscAsLCzoHAypyMhI7Q8MDDSKLy8v5x07dlhEdXZ2Wp7nDAxbW95nShB6JmNndoV79+7pB69du5Y7Ojq07+zZs5bCzC8Jm7S67uPjY1mApqeneevWrdpfWFiofa/i2rVrKu9TJQiNoDEcXOHMmTP6wdHR0dLN27Zt0/7Dhw/r62hhIiIieNOmTRwTE2PJAZiPKi87O1u6HeKxINskNCwpKUm6OT09Xfv37Nmjry8tLRlDcGJigsfGxiw5ALEqD8PXVTwWVFdX5/QXMs+v8PBwyzxyRFdXlzF8VR6aYFfxWBA2M/VgrFiYU4rR0VHeuHGj9m/evJlnZ2ct+RLkmBcLDEdXXoLCY0EYOhhqqgDMGcyPCxcuGMWo67Dg4GCempqSt9BAzPbt2y057qxwwGNBwNk+ZN5oQ0JC+OnTpzLdYGhoyLKAwGpra2XYK3kjggC6gYMHD+ohhn3lyJEjXFZWpgvEUHrx4oVM5QcPHnBYWJhFzOnTp2WYS7wxQQqsWv39/cZfgBVKFZmamirDeWRkxCJmzZo1fO7cORnmMh4Lmp+fN36d27dv212N4uLidLHFxcUW35MnT3jnzp3aHxAQwC0tLZYYd/FYECY5JrsqytwDXrp0yegClA+izeTk5GgfDB3B48ePjSE4MDCgzd4+5QiPBYF9+/bporDzo5dDX2ZeEA4cOGDJ6e3ttYiBoeM2vwBly9opAHy/+Pn5vVSIstjYWB4fH7fk5ObmvhTnyOTLcMYbEQTQXOKIy1wIhiIKn5mZkeGckZHBoaGhRhPqzBCTl5cn0x1iT5Bb3bYEQwkTG+3L5OSkdGvm5uYMe/bsmVNDjDtfr/YEJeOCO99D3oS976E4XLh48aKMXRXgh7AJ+kQJ+oCI5iorK2XsquDUqVMQs0BE4UoQ+Nn87bKasG0hvxGRr1lQEVoQuRF6O/fv31fnGN+YxYB3iGgcZ1yrCdvX8V+O/n+EM2KuqKiQeV7JiRMn1GJQIIWYqUOQty8QJjENUoAEE+t7BOPnvHPnjrzXitLT08NZWVlKzA9E9JYU4Agc3v+BhQLHtidPnuTGxka+efOmsTMvl+HMD+d51dXVvH//fnWQ8icWMVmwK7xPRKVE9AsR/WPu2VbAsM/0EFElEYXIQl8HbFrYidFeoGdaLsPz0JpFyoLs8S979QGWKkPEQgAAAABJRU5ErkJggg==",
  93: "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",
  94: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUmSURBVGhD5ZlXSHRHFMePvQV7iooogj4IKsGCqMGaKIoglpDHQB5EjREUjF2jDxEVVAJ50ICCJRbwwd5iy4OxRiUqRgT1wdgwGgsW/E44l73D3XHdb/fTuCv5wbAwZ+bO+c+9c865dwH+x7wPAJ4A8AkARABA+As0WofW8wKAD3iH3gUzAEgDgF8A4BgAUIPtBAAmAOAbADDnHVWFLwDgT7qYr68v5uTkYHNzM/b09ODIyAgODw//543W6e3txZaWFszPz8eAgABR3BYAfMk7rIzvaWJgYCAODAygNjE2NoahoaGisB8AQId3nqecBqenp+P9/T1/Pa0hNzdXFPUjL0DK56KY10BBQYEo6iteCPEeAOz6+fnx87SasLAwEnQAAFa8oK9J7fj4OD9Hq5mbm0M9PT0S9S0v6FcfHx9+vErs7OwIG0Fte3ubN6vEzMwMdnZ2YldXF56envJmpciCxO/SAEFJ6x96JtVha2sLk5KS0MzMjOULU1NTjI+PF2yqsrq6ikZGRuwaCwsL/BClVFRU0LxbAHASBX1MF2pra+PHPsrm5iba29szJ/hGtrW1NX7aAyiS+vv7y81dWlrihymlv79fnEtVhUAIdQwODvJjHyU8PJw5YGNjg3l5eZiRkYGGhoasnwLM3d0dP1WO0tLSB5uhrqCJiQlx7meiIKqZhMysCrSguLi+vj5OTU0xW0NDg5xzyjZpfn5emP9UQaOjo+LcT0VBVAgK5YYq1NfXs8U9PDx4M7q5uTF7SkoKbxa4ublBT09PuY3RmCDZIRRaSEgIb8bY2FhmDwoK4s0C2dnZbExaWho6OTlpTlBdXR1bXNEdkp4vFxeXB+eIHlHRHhwcjCcnJ2hra6s5QZTMxMUNDAzkHNjd3UULCwtmt7Ozw7OzM2a/uLhAV1dXwUYBZGNjA6+urtDc3Fxzgt68eSM8aqIDdGaoxG9vb5c7F9Ro54+Pj9lcOlOirby8XOjb29uT24QXF0Qoy0PSROvg4CDcFUKSLzAmJoZdiwKEpaUls6lbbTyLIIKqgcTERLa7VlZWmJmZiYWFhcw5erzojp6fn6OzszPrpxfGpqYmbGxsxNraWjQxMWG2srIyoV9VYc8mSOTg4ADX19eFX4ISrOhcRESE0EdnS1HOUdZaW1u5lRTzZEG3t7fC3aGCkl7Jeby9vZlTFJ4J2m3e4be1FxNEh1waZqU1YHd3N+ro6DAbiSaogqZHqaioSK4VFxdjVlYWGhsbsznJycmCjYpWVXiyICIqKoo5YG1tLdRy5Jg0IMTFxfHTFHJ9fa35oLC4uChX8vPNy8sL9/f3+WkKobAtvRZdWx2eRRBBL3T0iUsqhB7F1NRUIfuryuHhIbq7u6Ojo6PQVH3URBQJUqva5llZWcGhoSGcnp7Go6Mj3vxWKKxfXl4K+Yqaul+bFAkKpQ5lpb42o+h9yJs6Ojo6+LGvAroRMkHBoqCPAOCC3h5fIzU1NSTmDgBcREHEb4+9u2g7shTyBwDoSQVl6erqskT4WqCIKPuO8Z1UDGEJAPv0jes1IXs7/vux/4/oGzGWlJTw87SSyspKMRhk8EKk1NEgbQ8QEjFtvAAeOlg/0WC6nbOzs/y1NMry8jImJCSIYn4GAGNewGPQx/u/KFBERkZiVVUV9vX14eTkpJCZX6rRBxX60626uhqjo6PFd6pDCmK8w6rwIQAUAMACANxIazYNNMozywBQCgAOvKPvAiUtysRUXlDN9FKN1qPSzJV3SBH/AjZHCFzA3RWiAAAAAElFTkSuQmCC",
  95: "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",
  96: "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",
  97: "iVBORw0KGgoAAAANSUhEUgAAADQAAAAoCAYAAACxbNkLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAU8SURBVGhD5ZlbSJxXEMfHe6OlVk0vRlRQFF9UGqMPaqlGq6II4qX0JVDpS7S1QgSbeAHrS0UDUZoU0YIvCVUf1cRrjVohjcE0hmqgxhB98I71GrxPmYPn8O24u+5mra70B4PwnZk98/++75yZ8wnwP+YDAAgBgE8BIB4A4k7AaB6aLxQAPuQJvQ0uAPANAPwGAIsAgKdoSwDQBwDfAcB7PFFT+BIA/qYfCw8Px+vXr+Pdu3extbUVu7u7saur6z83mqetrQ3v3buHxcXFGBkZKcVNAMBXPGFj/EiBUVFR2N7ejtZEb28vxsbGSmE/AYANT55TQc55eXm4t7fHf89quHHjhhT1Mxeg5Qsp5ixQUlIiRX3NhRDvAsBUREQEj7NqLl++TILmAMCNC/qW1D58+JDHWDVPnjxBOzs7EvU9F/T7pUuXuL9JTE5OihtB9vr1az6sePPmDc7Pz+Pi4qJRW1hYwOXlZR5ukINN4k/tBkFFa5XeSXOYmJjArKwsdHFxUfXC2dkZ09PTxRinpqYGPTw88Pz580aNfFJSUni4QSorK2nubQDwlYI+oWQaGxu5r0HGx8fxwoULSgg3GhsbG9OJKS8vP+RnyEJCQnRijfHgwQMZR12FIIYudHR0cF+DxMXFqcnpjhYVFWF+fj46Ojqq67TB7OzsqJibN2+ik5MTnjt3Tsfoqbq5uaGNjY2Kzc7O1pnPGH19fTIuQQqinklUZlN49uyZmtje3h4HBgbUWENDg86d1t6klZUVfPXqlVhzWpuZmcE7d+6gra2tiLl48SKur6+ruKPo6emR830uBVEjKNoNU6ivr1cJBwcH82EMDAxU4zk5OXz4EKurq+jr6yv86Ym9ePGCuxjFYkEHi1BYTEwMH8bU1FQ1Hh0dzYcPceXKFeVfUVHBh4/EYkF1dXVGn5B2ffn5+emsI87g4KDy9ff3x62tLe5yJBYLomImk3BwcBBrSjI1NYWurq5q3NPTU6wdQyQkJCjf2tpaPmwSFgva398Xr5pMhNYMtfhNTU1iu5XXyaiuULHUx/DwsPKjNbSxscFdTMJiQYSxOqQttF5eXri2tsbDBVevXlV+tO2/LcciiKBuIDMzU71iVEuuXbuGpaWlKtGAgADc3d3loeJp0Oso/R4/fsxdTObYBEnm5ubEVkt/CSqwMtH4+HjuLujs7FQ+QUFBekWbisWCtre3xdOhu0pHck5YWJhKtrCwkA8L6EgvfczpCvRhsSBa5LTYZULaHrClpUWnhTH0KmmO0Xj79m0+bBYWCyKSkpJUQu7u7mJRFxQU6GwIaWlpPEywubmpOgOy/v5+7mIWxyLo6dOnotGUSXELDQ3F2dlZHiagWiX7NjLelZvLsQgi6EBHn7i0QuhVzM3NxaWlJe6uePnyJfr4+KC3t7eoYdPT09zFLPQJMqvb5jx//lzsWo8ePRKnzaOgHY26aTI6xVqKPkGxdMGc85A1oe88FEYXmpubue+ZgB7EgaDPpKCPAWCdjshnkerqahKzAwB+UhDxhylnF2vkoIT8BQB2WkEFtJUaKoTWyujoqPyO8YNWDPE+AMxSBT9LHJyO/zH0/yP6RoxlZWU8ziqpqqqSm0E+F6KljpysfYPQiGnkAji0sH4hZ3qcQ0ND/LdOlZGREczIyJBifgWAd7gAQ9DH+xnaKBITE8VHwvv374tGkirzSRl986N/ut26dQuTk5PFd0AAmKdNjCdsCh8BQAkADAPAlrZnOwWjOjMCAOUA4MUTfRuoaFElpvaCeqaTMpqPWrMAnpA+/gWzQ1uumxik+gAAAABJRU5ErkJggg==",
  98: "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",
  99: "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"
};


/**************************************************************
 * seq(1~99)에 맞는 뱃지 이미지 정보(Blob + 표시 크기)를 돌려줍니다.
 * 0~9는 정사각 뱃지, 10~99는 가로로 넓은 뱃지를 씁니다. 100 이상은
 * 준비된 뱃지가 없어서 null을 돌려주고(그 사진에는 번호를 생략),
 * 이런 경우는 실제로는 거의 없습니다(원인별 상세 시트 행 수 기준).
 **************************************************************/
function recoveryBadgeInfoForSeq_(seq) {
  if (seq >= 0 && seq <= 9) {
    const blob = Utilities.newBlob(
      Utilities.base64Decode(RECOVERY_BADGE_BASE64_BY_DIGIT[seq]), "image/png", "badge_" + seq + ".png"
    );
    return { blob: blob, width: RECOVERY_BADGE_SIZE_PX, height: RECOVERY_BADGE_SIZE_PX };
  }

  const wideBase64 = RECOVERY_BADGE_BASE64_BY_NUMBER[seq];
  if (seq >= 10 && seq <= 99 && wideBase64) {
    const blob = Utilities.newBlob(Utilities.base64Decode(wideBase64), "image/png", "badge_" + seq + ".png");
    return { blob: blob, width: RECOVERY_BADGE_WIDE_WIDTH_PX, height: RECOVERY_BADGE_SIZE_PX };
  }

  return null;
}


/**************************************************************
 * 사진 하나의 오른쪽 아래 모서리에 seq(순번) 숫자 뱃지 이미지를
 * 겹쳐 올립니다.
 *
 * 가로/세로 위치 모두, 우리가 직접 계산해둔 값이 아니라 방금 넣은
 * 사진 자신이 실제로 어느 열/행/오프셋/크기로 자리 잡았는지(photo.
 * getAnchorCell(), getAnchorCellXOffset()/YOffset(), getWidth(),
 * getHeight())를 다시 읽어와서 그 값 기준으로 딱 붙입니다 — 사진마다
 * 가로세로 비율이 달라 실제 적용된 크기가 우리가 미리 계산한 값과
 * 조금씩 다를 수 있어서, 사진 자신의 실제 위치/크기를 기준 삼아야
 * 사진마다 번호 위치가 들쑥날쑥하지 않고 항상 정확히 그 사진의
 * 오른쪽 아래 모서리에 맞습니다(호출 전에 SpreadsheetApp.flush()로
 * 먼저 반영을 확정해둬야 함 — insertCauseSheetImages_ 참고).
 **************************************************************/
function overlaySeqBadge_(targetSheet, seq, photo) {
  const badgeInfo = recoveryBadgeInfoForSeq_(seq);
  if (!badgeInfo) return;

  const anchorCell = photo.getAnchorCell();
  const anchorCol = anchorCell.getColumn();
  const anchorRow = anchorCell.getRow();
  const photoOffsetX = photo.getAnchorCellXOffset();
  const photoOffsetY = photo.getAnchorCellYOffset();
  const photoWidth = photo.getWidth();
  const photoHeight = photo.getHeight();

  const badgeOffsetXWithinAnchorCol = photoOffsetX + photoWidth - badgeInfo.width;
  const badgeOffsetYWithinAnchorRow = Math.max(
    photoOffsetY, photoOffsetY + photoHeight - RECOVERY_BADGE_MARGIN_Y_PX - badgeInfo.height
  );

  const xAnchor = pixelXToColumnOffset_(targetSheet, anchorCol, badgeOffsetXWithinAnchorCol);
  const yAnchor = pixelYToRowOffset_(targetSheet, anchorRow, badgeOffsetYWithinAnchorRow);

  const badge = targetSheet.insertImage(badgeInfo.blob, xAnchor.column, yAnchor.row, xAnchor.offsetX, yAnchor.offsetY);
  badge.setWidth(badgeInfo.width).setHeight(badgeInfo.height);
}


/**************************************************************
 * 원인별 상세 시트의 lastDataRow(마지막 데이터 행)에서 두 줄 아래부터
 * 사진을 가로 6장씩 붙입니다.
 *
 * - 세로: CAUSE_SHEET_IMAGE_TARGET_HEIGHT(약 5.5cm)에 맞추는 것을
 *   기본으로 하되, 그 비율대로 계산한 가로폭이 칸(슬롯) 폭을 넘으면
 *   칸 폭에 맞춰 다시 줄입니다(원본 비율은 항상 유지) — 세로가 너무
 *   커져서 다음 줄과 겹치는 일을 막기 위함입니다.
 * - 가로: 1~CAUSE_SHEET_IMAGE_AREA_LAST_COLUMN(X)열의 실제 너비 합을
 *   6등분(간격 CAUSE_SHEET_IMAGE_GAP_PX 포함)해서 칸 폭을 정합니다 —
 *   그래서 6번째 사진도 X열 끝을 넘지 않고, 시트의 실제 열 너비와
 *   무관하게 6장이 항상 고르게 간격을 둡니다.
 * - 칸 위치는 사진이 있는 것끼리 압축해서 채우지 않고, 각 건의 실제
 *   순번(seq) 기준으로 "(seq-1)/6이 몇 번째 줄, (seq-1)%6이 몇 번째
 *   칸"을 그대로 씁니다 — 예를 들어 9번 행에만 사진이 있으면 1번
 *   칸이 아니라 두 번째 줄의 세 번째 칸에 옵니다. 사진이 전혀 없는
 *   줄은 아예 건너뛰어 공간을 낭비하지 않습니다.
 * - 번호는 셀 텍스트가 아니라, 사진의 오른쪽 아래 모서리에 숫자
 *   뱃지 이미지를 겹쳐 올리는 방식으로 표시합니다(overlaySeqBadge_).
 *   한 줄의 사진을 전부 넣은 "다음"에 뱃지를 올려서, 칸 폭을 넘치는
 *   뱃지(특히 두 자리 번호)가 옆 사진에 가려지는 일이 없게 합니다.
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
    const placedPhotos = [];

    // 1단계: 이 줄의 사진을 전부 먼저 넣습니다.
    block.forEach(function(item) {
      try {
        const fileId = driveFileIdFromRecoveryLinkUrl_(item.url);
        if (!fileId) throw new Error("드라이브 링크에서 파일 ID를 찾을 수 없습니다.");

        const pixelX = item.slotIndex * (slotWidth + CAUSE_SHEET_IMAGE_GAP_PX);
        const anchor = pixelXToColumnOffset_(targetSheet, 1, pixelX);

        const blob = DriveApp.getFileById(fileId).getBlob();
        const image = targetSheet.insertImage(blob, anchor.column, blockStartRow, anchor.offsetX, 4);

        const nativeWidth = image.getWidth() || slotWidth;
        const nativeHeight = image.getHeight() || CAUSE_SHEET_IMAGE_TARGET_HEIGHT;

        let finalWidth = nativeWidth * (CAUSE_SHEET_IMAGE_TARGET_HEIGHT / nativeHeight);
        let finalHeight = CAUSE_SHEET_IMAGE_TARGET_HEIGHT;
        if (finalWidth > slotWidth) {
          finalHeight = nativeHeight * (slotWidth / nativeWidth);
          finalWidth = slotWidth;
        }

        image.setWidth(finalWidth).setHeight(finalHeight);

        if (finalHeight > maxHeight) maxHeight = finalHeight;
        placedPhotos.push({ item: item, image: image });
        insertedCount++;
      } catch (error) {
        failedCount++;
      }
    });

    // setWidth/setHeight 직후에는 이미지의 실제 위치/크기가 아직
    // 확정 반영되지 않은 상태로 읽힐 수 있어(한 사진은 맞고 다른
    // 사진은 안 맞는 식으로 들쑥날쑥해짐), 뱃지를 올리기 전에 먼저
    // 전부 반영시켜 둡니다.
    SpreadsheetApp.flush();

    // 2단계: 번호 뱃지는 이 줄의 사진을 전부 넣은 뒤에 올립니다 — 먼저
    // 넣은 사진 위에 뱃지가 가려질 일이 없게 순서를 분리했습니다.
    placedPhotos.forEach(function(p) {
      overlaySeqBadge_(targetSheet, p.item.seq, p.image);
    });

    if (maxHeight === 0) maxHeight = CAUSE_SHEET_IMAGE_TARGET_HEIGHT; // 전부 실패했을 때 대비
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
