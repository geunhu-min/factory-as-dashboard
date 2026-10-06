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
const RECOVERY_BADGE_WIDE_WIDTH_PX = 64; // 두 자리 뱃지 가로 크기(세로는 동일)
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
  10: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAASPSURBVGhD7ZpbSGxVGMf/qahQSIFlkpoIvYkR3vASpuUFQVTU8M2gBx/Kk5LgFTV9KFRQSHyQRAQlEUHB+yVvIeSNFI4oQiKJqKGmKYimfvFtXJu915k5yXbmtJ3TDxbo+q+1Z3//WfPtb60Z4H/uzZsAggB8COATAB+bsPF98f29D+AtOQAjvArgCwA/ATgEQI+oHQOYBvAEgIcc2H3IBrDJFwsNDaXi4mLq6Oig/v5+Gh8fp7GxMdM1vq+BgQHq7OyksrIyioyMFGb8BuAzOcDn8S1PjIqKouHhYXrMTE5OUmxsrDDiewCvyMHKfMeD8/Ly6ObmRr7eo6WkpESY0CwHrOVTEbwjUl5eLkz4XA6ceQ3A72FhYfI8hyIuLo4NOADwhmzAl+zO1NSUPMehWFxcJGdnZzahSDbg55CQEHm8Q3KXFH/VJkQuGv7iz4i92N/fp42NDdrd3ZWlZzg8PKTt7W06OzuTJZtQW1vLBlwBeFcY8AEv/66uLnmsTTg/Pyd/f38lAaWnp8uyChvEuqenJ7m7u5OPj4+SkE9PT+WhD2JoaEgkQ64aFT7ijpGREXmsTcjJyREvSKmpqbKssLW1RV5eXuo4bYuJiaGrqyt5imGmp6fFtROEAVxDK5WULTk4OKDs7GxdMNZWQFZWljomMzOTWlpaKCgoSO1ramqSpxhmYmJCXDdeGMAbCaWctAWXl5fU2NhIfn5+z7yblgzgvMBLnnVfX1+6vr5W+peWlkTGpuDgYLq9vZWnGsLuBuzt7akBu7q6UkZGBrm4uFg1gPcXYjyvBAEHHBAQoPSzQfdJoPfB7gbs7Owo71xERITy3F1eXn7uCqivr1f1goICncb7EaHNzc3pNKPY3QDO2jMzM+r/mhe0aICmTqeKigqdFh8fr2q807MFdjdA5t8MyM/PV/XKykqdpjWgp6dHpxnFdAYUFhZaNSAhIUHV+vr6dJpRTGcABy300tJSncY1gNBs9Zg2nQGtra2qnpubq9MCAwOVficnJ1pfX9dpRjGdAfPz86oeHh6u9h8fH5OHh4fSzzXFxcWFbp5RTGcAF07iec/vdHd3N52cnFBRUZHVlfEQ/lMD0tLSZFmhra1NHcPN29tb/dvNzY02NzflKYZ54QaMjo6qwSQnJ8uySk1NjRKs1gjeIPX29spDH8QLN+Do6EgxgXebKysrsqyDE11zczNVV1dTe3u7co5gaywZYJfdoFmxZEAsd9jrPMBsWDoPCOYOzr4vA/xG3xkQIwx4G8A5f+5eBvisAsDfAAKEAcwv0dHR8liHJCkpiQ14CsBZa8DXXIRwVebIrK2tKYc0AL7RBs+8DmCfz8wdmZSUFA7+T2u/H+DvzKiqqkqe5xDU1dWJ5PeVHLiWFh7kaAlRE3yXHLAMJ4YfeDAvl4WFBflaj4rV1VXlUPYu+B8BuMsBW4O/LN3jxJiYmKgcXA4ODipnfVxJmbXNzs4qP+poaGhQ9h13p9F/cJKXA7wPXgDKASwDuLxz8bE0fs6vAqgG8I4cmBG4aODKictHrqHN2vj+uLR/Tw7AEv8AlxyHEiin8fwAAAAASUVORK5CYII=",
  11: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAOUSURBVGhD7ZpLSBtRGIVPjEqhpbSLPiRIoNBtS6m6UEMb+1AEV0IpuIlkmZouuqhiFjabFhUVA12ULpXEx0J8v6rRbqpSSiSKVuqimzZaWoh10Ra55R+8YfKbBDEJZoZ8cBeZe+bOPSeTO//MBMhybC4BuAHAAuA+gHsZ2GheNL+bAC5zAyfhLAAHgHcAfgAQGmo/AfgBOAGc58aOw2MAn2mw4uJi0djYKHp6esTIyIiYmZkR09PTGddoXqOjo6K3t1c0NzeL0tJSGcYXADZuMBEvaceysjIxMTEhtMzc3JywWq0yCA8AAzfLeUXihoYGcXBwwMfTLE1NTTKE19ywmkfSvB5xuVwyBDs3TpwD8LWkpITvpysqKioogBCAizyAJ5TO/Pw830dXrKysCKPRSCE85wG8Lyoq4npdcrgoflIviFQ0hOk3kk42NjaUtre3x7uOsLm5qWjD4TDvSprW1lYK4C8AswzgFp3+Pp+Pa1OG1+sVBoNBWYSolkjEwMBARDs4OMi7k2Z8fFwuhlQ1KtylDZOTk1ybEra2tkRBQYE8qBgeHuaSCNvb28JkMkW0FEaq8fv9cvyHMgCqoZVKKtVQhVZYWBgxlOgMoOObzeYobTrOgNnZWTn+AxkA3Ugok00V9E3abLYoM/ECIK3dbj+i03QA3d3dERN0dbFYLHED8Hg8cbWaDaCjo0Pk5eUpVeX+/r5wOp1xAyBtbm6ucDgcR7SaDWB9fV0Eg8HIZ/XPgQdA2tXV1ZhazQbAqa+vjxsAR63NBpANIBsA706abADZAE4hgESXQY5am457gVMJoK6uLmJqaGiId0eh1vb19fHupDmVAKgoorvNqakpsbu7y7ujUGtDoRDvTppYAaTtbjATiRWAlTak63lAphHrecBt2tDf38+1uoS+6MMA7sgArgL47Xa7uVaXdHV1kfl/AK7JAIgP5eXlXKtLqqqqKIAgAKM6gGc5OTliaWmJ63XF2tqayM/PpwBeqM0TFwB8p2fmeqampobM/4r3/wF6ZyZaWlr4frqgra1NLn5PuXE1b0iktwVRZd7HDXNoYXhLYjpdlpeX+ViaIhAIiNraWmneC+AMNxwPeln6jRbGyspK0d7eLsbGxsTCwoJSSWVqW1xcVP7U0dnZKaqrq5WHrQB2aJHnBo/DFQAuAB8B/DlMUSuNrvMBAG4AJm7sJFDRQJUTlY9UQ2dqo/lRaX+dG4jFf/vd9+NaP6WDAAAAAElFTkSuQmCC",
  12: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAARMSURBVGhD7ZpbSDVVFMf/XtCgzFKsIDAJ9LGMVLxiWqYIoiLk5UGDHtMCe0hR8PaQd4XAhwyf1NQnwfslb72kIimkSBhCImqoYfhioSvWcPYwZznnw8Zzvuacrx+shzNr7ZlZ/7PPmrX3HOB/HkwEgLcApAH4AMD7NjS+L76/twG8IhOwwvMAPgXwPYBzAORFdglgBcBnAF6UiT2EEgC/8Mni4+OppqaGBgcHaWJighYWFmh+ft52xvc1OTlJQ0NDVFdXR8nJyUqMXwF8LBN8El/xwJSUFJqZmSFvZmlpiTIyMpQQXwPwk8lKWjm4qqqKbm9v5fm8ltraWiVCn0zYyEcqeV+kvr5eifCJTJx5AcBvCQkJcpxPkZmZyQKcAXhZClDJ6iwvL8sxPsXm5iYFBASwCF9KAX6Ii4uT8T6Joyj+ZCyI3DT8yb8RT3F6ekr7+/t0fHwsXfc4Ozujw8NDOj8/ly630N7ezgL8BeANJcA7PP1HRkZkrFu4vr6mqKgorQAVFhZKt87Y2Bilp6dTWFgYBQcHU0REBGVlZdHs7KwMfRTT09OqGHLXqPEeH3D3hRQVFRXqgpSfny/dGm1tbXqMmQ0PD8shlllZWVHn/VAJwD201km5E57KJSUlTomYzYCDgwPy9/fX/H5+flRZWUkDAwNUXFysjwsPD6fLy0s51BKLi4vqvFlKAF5IaO2kO7i5uaHe3l6KjIy8902aCdDS0qL7y8rKnHyJiYm6b25uzslnFY8LcHJyot90UFAQFRUVUWBgoEsB+LrV1dVUWlqqTU8j5eXl+rnGx8edfFbxuABHR0faszYpKUl77m5tbT1xBriCW/GYmBh97N7engyxhMcFuLq6otXVVf2z4YL/SgBe0alxaWlpdHd3J0Ms4XEBJFYEaG1t1cdwgdzY2JAhlrG9AMaiyNbX1ydDHoWtBTCs2DTr6uqSIY/GtgI0Nzc7Jd/f3y9D3IItBeDtNhUTEhKitauewnYCcMcYGhqqx0RHR1N3dzc1NTVRY2OjZg0NDd7zGJQYBSgoKJDue1PflfFmpzt46gJwC6uSyM3NdfLxsz02NvZesmY2OjrqNNYqT12Ai4sLTQRebW5vbzv5WABumtivtrfNjP38U3EHZgJ4ZDVoV8wEyOADntoPsBtm+wHv8gHekXkW4C/aIUC6EuA1ANdcjZ8FeK8CwN8A3lQCMD+mpqbKWJ8kJyeHBfgZQIBRgC941bW+vi7jfYrd3V1tkwZAkzF55iUAp7xn7svk5eVx8n+4+v8AvzPTWk9fpKOjQxW/z2XiRr7hIF8riIbkR2TCEi4M33IwTxd37sT8F+zs7Gibso7kvwPwnEzYFfyy9IQLY3Z2NnV2dtLU1JTWtnInZVdbW1vT/tTR09OjrTscu9G/c5GXCT6EVwHUA9gCcONQ0VuMn/M7AJoBvC4TswI3Ddw5cfvIPbRdje+PW/tomYAZ/wC/IZx+UFPKhwAAAABJRU5ErkJggg==",
  13: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAR9SURBVGhD7ZpbSHRVFMf/qYhSij5UQqISBD5IEangDVMrRfCGEL6IQY9lPgReUNBELFRQCBI00QdFEZ+83/IOpiIpJEiYl0TUzNRQRFNXrI3nMLO++cTGM19nhn6wH2bWWmfO+p89a6+9Z4D/eTSvAngbQByADwAkmXDwffH9vQPgNZmAPbwM4DMAPwD4AwA50fgTwBSALwD4ysQeQw6AX/hiERERVFxcTO3t7dTX10djY2M0OjpqusH31d/fTx0dHVRaWkrR0dGaGL8C+EQm+BBfc2BMTAwNDQ2RMzMxMUEJCQmaEN8CeEkmK/mGnfPz8+n29lZez2kpKSnRRPhOJmzJx1ryrkhZWZkmwqcyceYVAL9FRkbKOJciMTGRBTgE4C8F+JzVmZyclDEuxdLSErm7u7MIRVKA2fDwcOnvktwXxZ8sCyI3DX/xd8RRHBwc0Pr6Ou3t7UnTMxweHtLOzg5dXFxIkyHU1NSwANcAgjUB3uXp39XVJX0N4fz8nEJCQlQBysrKkmadnp4eiouLI39/f/L29qbg4GAqKCig09NT6fokBgcHtWLIXaPifX5jeHhY+hpCXl6e9oGUkZEhzYrGxkbdRw4uzGdnZzLEbqamprRrf6QJwD206qSMhKdyTk6OVTK2ZsD+/r564mz38vKiyspKamlpodDQUD2uvLxchtnN+Pi4dt0PNQF4I6HaSSO4urqihoYGCgoKeuZp2hKA22sPDw9l55ZbY2ZmRo/jr4ZROFwAfqLajXt6elJ2draeoC0BLi8vaXt7W90YzxqNubk5/TppaWlWMU/B4QLs7u6qtTYqKkqtu8vLyw/OAMnJyQnNzs6qTZgWx5sco3C4AFywpqen9dcWH/goAXJzc3V/Pz8/6uzslC5PwuECSP6NAHd3dxQWFqb7+/r6UmFhoaGrgKkFuLm5UdN/YWFBJa7FxcfHq+JqBKYWQJKUlKTH9vb2SrNdmE4AnvbcLW5ubtL19bWVraioSI+trq62stmL6QTgJicwMFAtmd3d3Va29PR0PbapqcnKZi+mE6CtrU23c/e3srKi+v/m5ma9f+BldWNjQ4baxX8qQGZmpjSrac+dnubDSQcEBOiveXB7bBQvXICRkRE9kdTUVGlWHB0dWU13bfj4+FBVVZV0fxIvXIDj42MlAu82eXo/xPz8vNpH8BNvbW2lra0t6fJkbAngkN2gWbElQAK/4ajzALNh6zzgPX5DLkGuCj/oewHiNQECAJwbWWnNDNcYAH8DeFMTgPkxNjZW+rokKSkpLMDPANwtBfjSzc1NbUJcmbW1NdVxAvjKMnnGD8ABn5m7MnyyBODkef8f4N/MqKKiQsa5BLW1tVrxK5CJW9LETq5WEC2S75IJS7gwfM/OPF0WFxfltZyK1dVVdSh7n3wnAC+Z8PPgH0v3uTAmJydTXV0dDQwMqLM+7qTMOvgonf/UUV9fr/Yd97vJ37nIywQfw+sAygAsA7i6V9FZBq/zqwAqAbwhE7MHbhq4c+L2kXtosw6+P27t35IJ2OIfapWyF9OjfqYAAAAASUVORK5CYII=",
  14: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAARDSURBVGhD7ZpbSHRVFMf/XgnNT/OSgmCSqG9meAEvUdr3pQiiIoRvBj2a9dCDij6YCIUKKmFC9CRKIvrk/fJ5ffCGpJIgaiCJeEvFULykrFgbz2FmNVPDMH6dOfWD/TB7rbXPWX/27L32ngH+x2HCACQAeA/AcwAfGrDxe/H7vQPgTZmAM/gDKAPwEsBvAMiN2hmAaQCfA3gmE3OEEgBbPFhKSgpVVlZSZ2cn9ff30/j4OI2NjRmu8XsNDAxQV1cXVVdXU3p6uibGLwA+kQn+HV9zYEZGBg0PD5M7Mzk5SVlZWZoQ3wLwkMlKvmHn8vJyenh4kOO5LVVVVZoI38mELflYS96M1NTUaCJ8KhNnXgfwa2pqqowzFdnZ2SzAEYA3pACfsTpTU1MyxlQsLy+Tl5cXi1AhBZhLTk6W/qbkcVH8yXJB5KLhd/6OPBWHh4e0ublJ+/v70mSXi4sLFbO9vU339/fS7DQNDQ0swB2AtzQB3uXp393dLX1dwuXlJUVHR6sFqKioSJptcnd3p+oPjgkJCaGTkxPp4jRDQ0PaYshVo+ID7hgZGZG+LqG0tFR7IBUUFEizTSoqKvSYoKAglwowPT2tjf2RJgDX0KqSciVHR0dUUlKiJ+LoDJibmyMPDw89xtUzYGJiQhv7hSYAHyRUOekKbm9vqaWlhaKioqySd0SAq6sriouLs4pxOwEODg70l/f19aXi4mLy9vZ2SICysjLlFxAQQGFhYe4pwN7entpr09LS1L67srLi0AzgNUjza29v14oW9xOAt6+ZmRn9s8UD7Qpwfn6uf2VycnJUX1JSknsKIHFEAG2n4Km/u7ur+hISEv4bAvT19Smbj48P9fb26v3aDAgNDXXpydRQApydnVFERISyhYeHU1tbG7W2tqoWGRmp+v39/am+vp46OjqsYp3FUAJsbW3ptn9qvCu4AkMJsLOz85dE7bWYmBirWGcxlAA3Nze0uLhICwsLeuPP3GJjY1VMYGCg2iLX19etYp3lXxWgsLBQmu2SmJioYoKDg+n6+lqaneaVCzA6OqoLkJeXJ812iY+PVzF+fn50fHwszU7zygU4PT1VIvA0Xl1dlWa7zM/Pqxi+peLjsauwJcCTnAaNii0Bsrjjqe4DjIat+4Ak7ujp6ZG+psTi0PW+JkAEgMu6ujrpa0r4rgLAHwDe1gRgFjIzM6WvKcnNzWUBfgbgZSnAl56enqoAMTMbGxvqkgbAV5bJM0EADvnO3Mzk5+dz8uf2/j/Av5lRbW2tjDMFjY2N2uL3hUzcku/ZyWwLokXy3TJhCS8MP7AzT5elpSU5lluxtramLmUfk/8RwGsyYXvwj6UHvDDyHV1TUxMNDg6quz6upIzaZmdn1Z86mpub1bnj8Tb6mBd5maAjhAOoAbAC4PZRRXdpvM+vAagDECkTcwYuGrhy4vKRa2ijNn4/Lu1jZQK2+BNZVKLqQNq2OQAAAABJRU5ErkJggg==",
  15: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAASFSURBVGhD7ZpZSHVVFMf/jgRG2EMlhlMiig9FfM6apQ1OiIqY4otBb5om9JCiD05YqKAYCEq+KYnoi/OU0+dDTqTwKZLkQypqZGE4oKIr1uaew737u/rJ6V4799AP9oNnrXXuWf+z79pr7yvwPw/mNQBvA3gPwEcAPtTh4Ofi53sHwOtyAlrwAFAE4EcAfwAgBxp/ApgFUALgFTmxh5AH4Be+WXh4OJWVlVFXVxcNDg7S5OQkTUxM6G7wcw0NDVF3dzdVVFRQTEyMIsavAD6TE7yPbzgwNjaWRkdHyZGZnp6mhIQERYjvADjJycp8y87FxcV0c3Mj389hKS8vV0RokxM251MleSNSWVmpiPC5nDjzMoDfIiIi5DhDkZiYyAIcAXhVFuALVmdmZkaOMRTLy8vk4uLCInwtC/A0LCxM9jckpqL4s3lB5Kbhb/6O2IvDw0Pa2tqi/f192fToNDQ0sABXAPwUAd7l6d/T0yP72oTT01Py9/cXBSgrK0s2q3Dx5VkYGRlpMbguxcfH08HBgRyiiZGREaUYctco+IAvjI2Nyb42oaCgQPlAysjIkM2Cy8tL8vLyUv2sjZ2dHTlME7Ozs8o9P1EE4B5adFK25OjoiPLy8iySuGsGbG9vk6urq/Dx8PAgPz8/8vHxUUdQUBDt7e3JYZqYmppSnudjRQDeSIh20hbw22xpaSFfX9/n3uJdAgwMDKg+9fX1dHV1Je5zfn4uvkJnZ2d0e3srh2nC7gLwd1VJxt3dnbKzs9W3e5cAdXV1wu7k5ESLi4vimr06UbsLsLu7K9ba6Ohose6urq6+cAbk5OQIu5ubG6WlpVFISAgFBARQamqqzfcjdhfg5OSE5ubm1L/NPtCqAPymQ0NDVR9ro62tTQ7TjN0FkHmRAMqMYTt/VUpKSqizs9OigDo7O9Pm5qYcqgndCXBxcSGWpo6ODlpYWLCw5ebmqrF8JmELdCfAffT396ux6enpslkTuhXg+vr6uco/PDysxqakpFjYtKI7Afr6+igqKooCAwOptrbWwmbq28UoKiqysGlFdwKY2z09PcUKwk3P2toaeXt7qza5PmjlPxUgMzNTNguUPoAHN0PBwcGiiVKuFRYWyiGaeXQBxsfH1US4sbEGt7v5+fmqnzJ4WSwtLbVZG8w8ugDHx8dCBN5t8rS+j5WVFWptbaXq6mpqb2+njY0N2eVfY00Au+wG9Yo1ARL4gr3OA/SGtfOAJ3yht7dX9jUk/KJNAryvCOAF4LSmpkb2NSR8VgHgGsBbigDMT3FxcbKvIUlOTmYBngFwMRfgK95xKYcRRoVXFVN/UW2ePOMJ4JDPzI0Mb6gA/HXX/w/wb2ZUVVUlxxmCxsZGpfh9KSduTgc7Ga0gmiXfIycsw4Xhe3bm6bK0tCTfy6FYX18Xh7Km5H8A8JKc8F3wj6UHXBiTkpKoqalJ7M15p8adlF7H/Py8OERtbm4W+w7TafTvXOTlBB/CGwAqAawCuDSp6CiD1/l1ADUA3pQT0wI3Ddw5cfvIPbReBz8ft/ZBcgLW+AeeS5cLNikAZQAAAABJRU5ErkJggg==",
  16: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAASRSURBVGhD7ZpbSHRVFMf/XpBQiXooxUAl6UUwjVTwkqVliiAqgvgiCj2IloUmpOiDilB4QUXoIQIFFS+v3i95CzQVTUFB0l7yHmkqopiXFesw5zCzvpkvGc98Had+sME5a+0zZ/1nz9prrxH4nwfzGoC3AbwH4CMAHxpw8HPx84UCeF0GYA9eAD4F8AOAPwDQExonAKYBfA7gZRnYQ8gG8AvfLCIigsrKyqizs5P6+/tpfHycxsbGDDf4uQYGBqirq4sqKiooOjpaFeNXAHkywOfxNU+MiYmh4eFhespMTk5SfHy8KkQrABcZrOQbdi4qKqK7uzt5vydLeXm5KsK3MmBzstTgnZHKykpVhE9k4Iw3gN8iIyPlPKciISGBBTgC8KoU4DNWZ2pqSs5xKpaWlsjNzY1F+EoK8GN4eLj0d0pMSfFn84TIRcM5f0ccxeHhIW1ubtLe3p40PQMn393dXdrZ2ZEmXairq2MB/gIQoArwDi//np4e6asLFxcXFBgYqCSgjIwMada4v7+nxsZGCgkJIW9vb/Ly8qKwsDBqa2uTro9iaGhITYZcNSp8wBdGRkakry7k5uaqb0hpaWnSrHBzc0Pp6emanxwtLS1yit1MT0+r9/1YFYBraKWS0pOjoyPKzs62CMTWCjAtS2UEBwdTe3s71dTUqAmLXFxcaH9/X06zi4mJCfW9ElUB+CChlJN6cH19Tc3NzeTv7//MJ2lNAP6K+Pn5KXZPT0/a3t7WbAUFBcp9QkNDaWFhwWKevThcgIODAy1gDw8PyszMJHd3d5sCmC1JZZ+W3N7eykuPwuECcPbmpRsVFaXsu8vLy89dAa2trZqdS9aVlRUqLCyknJwcxXZ+fi6nPAqHC3B2dkYzMzPaa7M3tCoAn9xUO2d99W91BAUF0fr6upxmNw4XQPJPAhQXF1sE7OPjQyUlJZSamqpd48R4eXkpp9qF4QQoLS3V7Lz3m3/a5ruIXnWK4QSorq7W7HFxcRa2wcFBzabXKdVwAnR0dGj2xMREC9vs7Kxmy8vLs7DZi+EE2NjYUAodtvv6+tLV1ZVm6+7u1uZyW04PDCcAY9a/o/z8fDo5OaGtrS2LXUGvo/q/KgDX+9aYm5vTyl4evBNwVai+tnWGsIcXLsDo6KgWSEpKijRrcMc5ICBA81VHVlaWrsXQCxfg+PhYEYFPm6urq9JswenpKfX29lJtbS01NTXR/Py8dHk01gRwyGnQqFgTIJ4vOKofYDSs9QPe5Qt9fX3S1ynhD9okwPuqAL4ALrgB8V+AexUAbgC8qQrA/BQbGyt9nZLk5GQWYB2Am7kAX7q6uurWdTEqXHVykwZAtXnwzCsADrln7syYjth/2vr/Af7NjKqqquQ8p6C+vl5Nfl/IwM35jp2cLSGaBd8jA5ZwYvienXm5LC4uyns9KdbW1pSmrCn4bgAvyYBtwT+WHnBiTEpKooaGBqU5wb0+rqSMOrh/wP/UwWU0nztM3ejfOcnLAB+CD4BKAMsArk0qPpXB+/wagBoAb8jA7IGLBq6cuHzkGtqog5+PS/u3ZADW+BvWkH+My7w9vwAAAABJRU5ErkJggg==",
  17: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAPxSURBVGhD7ZpLSFRRGMf/qYhQpi16QGAShCuL8P2I0h6KICqCCC4MWoiUtWiRogtzUz5AJWgRISJq4tb3Ix1tk4qkkCCjuUjEMbMw3EwhJ76D5zJ9jjKO99qdWz84iPN95879/++d737nzAD/8ZjTAC4DuAbgFoCbJhx0XnR+VwCc4QK84TiA+wDeAvgKQPjQ+AbABuAhgJNcmCfkA7DTwWJiYkRpaalobW0VXV1dYmhoSAwODppu0Hl1d3eLtrY2UV5eLhITE5UZnwDc5QL34xlNTEpKEn19fcKXGRkZESkpKcqIFwCOcbGc55RcUlIitre3+fF8lrKyMmXCSy7YlTwl3opUVFQoE+5x4cQJAJ9jY2P5PEuRmppKBqwBOMUNeEDujI6O8jmWYmpqSvj7+5MJT7gB76Kjo3m+Jdkpih9cCyI1DT/oM2IUDodDzM/Pi5WVFR46cmpqasiAnwAuKAOu0u3f0dHBc3Vha2tLhIeHywKUk5PDw6K6ulpERUWJuLi4PQfF29vb+VSv6O3tVcWQukbJDXqhv7+f5+pCYWGhekORlZXFw6KgoECL7zcqKyv5VK+w2WzqmHeUAdRDy05KT9bW1kR+fv4fItzdASQsIiJi14iMjFQFS/6lpkYPhoeH1fncVgbQQkK2k3rgdDpFQ0ODCAsL23UV3RmwFy0tLZoBjY2NPOw1hhuwurqqCQ4MDBS5ubkiICDgQAZQsQwNDZVzMjIyePhQGG7A8vKyvHIJCQnyuTs9PX3gOyAvL0/mBwUFiYWFBR4+FIYbsLm5KcbGxrT/Xd7QIwMmJia0/KKiIh4+NIYbwDmoAfSRUVffbrfz8KExtQFLS0uybniS6y2mNqCurk7LpY0NIzC1AWrzIjg4WPYRRmBaA9bX10VISIjMi4+P52HdMK0B9ORQecXFxTysG3/VgOzsbB7WaGpq0vLq6+t5WDeO3ICBgQFN2H5dXVVVlZbX3NzMw7px5AZsbGxIE2i1OTMzw8Mai4uLModyjSqAhDsDDFkNmhV3BqTQC0btB5gNd/sBUfRCZ2cnz7UkdKF3DLiuDDgHYIuK0L8A7VUA+AXgojKAeJ+cnMxzLUl6ejoZ8BGAv6sBj/38/ORS1MrMzc2pxdZTV/FEKAAH9eJWJjMzk8R/3+v3A/SdmW67r2ajtrZWFb9HXLgrryjJagXRRXwHF8yhwvCakul2mZyc5MfyKWZnZ7UdJgBvAARxwXtBX5auUmFMS0uTGxU9PT1yxUadlFnH+Pi4/FEHLaZo3bGzG/2FijwX6AlnAVQAmAbg3HHRVwY952cBVAE4z4V5AzUN1DlR+0g9tFkHnR+19pe4AHf8Br7x5w92yx0hAAAAAElFTkSuQmCC",
  18: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAASjSURBVGhD7ZpZSHVVFMf/DmhQ5gCVEJoo6fdkRCo4ZGmZ4izCh+KDafggZYk9pDhgIhYqDiQpEb44JIpPzkNOoaQiDihKEEIiTpkYgprojrXxHO5d3/VDTvd+HW/9YD941lrnnvXf+66z9r4C/3NvXgLgD+BtAO8DeE+Hg56Lnu8NAC/zBLTwPICPAfwI4HcA4gGNPwBMAfgUwIs8sfuQBuAXullgYKAoLCwU7e3toq+vT4yNjYnR0VHdDXqu/v5+0dHRIYqLi0VISIgixq8APuQJPo2vKDA0NFQMDQ2Jh8zExISIiIhQhPgGgA1PlvM1Oefl5Ynr62t+vwdLUVGRIsK3PGFDHivJWyMlJSWKCB/xxIkXAPwWFBTE46yKyMhIEuAAgCsX4BNSZ3JyksdYFYuLi8LOzo5E+IIL8FNAQAD3t0pui+KyYUGkpuFP+o5Yiv39fbG1tSV2d3e56QkODg7E9va2ODk54SazUF1dTQL8BeA1RYA3afl3dXVxX7NwdnYmvLy8ZAFKSUnhZpXOzk4RHBwsXF1dhYODg3B3dxdJSUlieXmZu/4jBgcHlWJIXaPkXbowPDzMfc1CZmam8oEyIVM0NTWpPnw4OTmZVYSpqSnl3h8oAlAPLTspc0JLOS0tzSgZUyvg+PhYuLi4SDvNfEVFhew4MzIy1Lj4+Hgeppnx8XHlvlGKALSRkO2kObi8vBQNDQ3C09Pzidk0JcDs7KxqT09PN7J5e3vL6x4eHuLq6srIphWLC7C3t6cmRDOampoq7O3t7xRgbW1N9c/JyTGyPXr0SF739fU1W2dqcQF2dnbku5YKGr13l5aWnroCaGbDw8Ol3dnZWRbjzc1NUVpaqsZVVlbyMM1YXIDT01MxPT2t/m3wgSYFII6OjkRiYqLqpwxaOdTHmxOLC8C5jwCHh4fSxgWglVRQUCDOz895iGZ0JwD1Cf7+/uqMl5WVid7eXpGVlaXGJScni5ubGx6qCd0J0NbWptrz8/ONbFFRUaptbm7OyKYV3QlA22/Fzg9gqqqqVFtLS4uRTSu6E4BmXbE3Nzcb2XJzc1Vba2urkU0ruhOgp6dHtfv4+Mi+gL7vMzMzws3NTbWtr6/zUE38qwJQMeNQ52hwgCkcHR2Fn5+fsLW1Va9lZ2fzMM08cwFGRkbURGJjY7lZQq/BuLg41U8ZNjY2MvmLiwseoplnLgBtdkgE2m2urKxwsxFU6evq6uT5XWNjo+wizY0pASyyG9QrpgSIoAuWOg/QG6bOA96iC93d3dzXKqGJvhXgHUUAdwBndBDxX4DOKgBcAfBWBCB+DgsL475WSUxMDAmwDsDOUIDP6b07Pz/P/a2KjY0NeUgD4EvD5AkXAPt0Zm7NJCQkUPInd/3/AP1mJsrLy3mcVVBTU6MUv8944oZ8R07WVhANku/iCXOoMHxPzrRcFhYW+L0eFKurq/JQ9jb5HwA8xxO+C/qxdI8KY3R0tKitrRUDAwPyrI86Kb0O2kHSmUJ9fb3cd9yeRh9SkecJ3odXAJQAWAJweaviQxn0nl8FUAHgVZ6YFqhpoM6J2kfqofU66PmotX+dJ2CKvwFuB3rQ9VpCAQAAAABJRU5ErkJggg==",
  19: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAASVSURBVGhD7ZpLSHVVFMf/+SIwtEQtCEwFJw6MSAUfadpDEURFSEHQsIGDsgYN8gWaDgoVFIIGGTpQUZyJ70e+B6mIigqS+CDxiaaGgpa6Yh3uPlzXd/2S8937dbxfP9hwz157nXvW/5yz9tr7XuB/HowfgDAA7wH4EMAHJmx8XXx9bwPwlwEYwRPA5wB+AXAEgB5R+wPAGIAvAXjJwB5CNoDf+GQRERFUVFRELS0t1NXVRUNDQzQ4OGi6xtfV3d1Nra2tVFpaStHR0UqMdQCfygCfxnfsGBMTQ319ffSYGRkZoYSEBCXEDwBeksFKvufBhYWFdHNzI8/3aCkuLlYi/CgDtuYTFbwzUlZWpkT4TAbOvALg98jISOnnVCQmJrIABwBekwJ8weqMjo5KH6didnaWXF1dWYRvpACT4eHhcrxTYkmK89YJkYuGP/kdcRT7+/u0urpKOzs70vQEu7u7tL29TdfX19JkF6qrq1mAvwC8pQR4hx//9vZ2OdYunJ+fU2BgoJaAMjIypFmnqalJqzm8vLzI09OTQkNDqba2Vg57Znp7e1Uy5KpR433u6O/vl2PtQl5envpCSktLk2aNkpISfYxsOTk5dHt7K10MMzY2ps79sRKAa2itkrInBwcHlJ2dfScYW0/A1NSUbvfx8dHuOj+m3t7een9DQ4N0M8zw8LA670dKAF5IaOWkPbi6uqL6+noKCAh44m7aEqCgoEC3c7mtaGtr0/vDwsLslhMcLsDe3p5+4R4eHpSZmUlubm73ChAfH6/ZeHra3NzU+y8uLsjf31+zubi40Nra2h0/ozhcAM7gHExUVJQ2787NzT31CYiLi9MF2Nra0vv5jgcHB+u+9lqXOFyAs7MzGh8f14+tvtCmALm5ubqdM7SCheTZQNl4lWcPHC6A5N8E4NlH2Xnqm5ycpOXlZUpOTtb7uTU2NkpXQ5hOACY/P/9OsKpxDlGfm5ubpZshTCkAU1VVpb3z7u7uWvKrrKyk9PR03bezs1O6GMK0Aig2Njbo9PRU+2y1q0NLS0tyqCFMJwC/71zo8Fpkfn5e7+f1g0qCQUFBdHl5ecfPKKYTgIsfZY+NjaXDw0M6OjqirKwsvZ/3+OzFfyoAv9MSXizxHVZjfH19yc/PTz8OCQmhk5MT6WaY5y7AwMCAHkxKSoo0a8zMzOgrRuvGRdL6+roc/kw8dwGOj481EXi+X1hYkGYdvstc7JSXl2uLIXsvzhS2BHDIatCs2BIggTsctR9gNmztB7zLHR0dHXKsU2JVescrAd4AcM6V14sA71UA+BtAsBKA+ZXn4BcByyJrGYCrtQBf86bD9PS0HO9UrKysqAXWt9bBM68C2Oc9c2cmNTWVgz+57/8D/JsZVVRUSD+noKamRiW/r2Tg1vzEg5wtIVoF3y4DlnBi+JkH8+PCJepjZnFxUduUtQTfBuBlGfB98I+le5wYk5KStL36np4eba+PKymztomJCW3ztK6uTlt3WHajDznJywAfwusAygDMAbiyqPhYGs/ziwAqAbwpAzMCFw1cOXH5yDW0WRtfH5f2ITIAW/wD2lF/IGxgsUUAAAAASUVORK5CYII=",
  20: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUaSURBVGhD7Zp7KL5nGMev/ZwzjGzGHErtH8mac1gOm3OKMOSf1f7wh9loyjGMP7YcoiZ/aJKiSUI5hp+zmsOSmhyHlJgwh0kM17ruPE/Pc//e95fegx/vb5+6S9d13c/7XN/3fq77up8XwP88mPcBwA0APgOALwDg8yc46L7o/j4BgA/4BFTBFADSAeAlABwCAD6jcQwAYwDwLQCY84k9hGQAWKOLeXl5YW5uLjY3N2N3dzcODQ3h4ODgkxt0Xz09PdjS0oIFBQXo5+cniPEnAHzFJ/g6fqSJ/v7+2N/fj8+ZkZERDA4OFoT4GQDe4ZPl+YmCMzIy8Pb2lr/esyUvL08QoY5PWMqXQvK6SGFhoSDC13zixLsAsOPt7c3P0ylCQkJIgL8AwJIX4BtSZ3R0lJ+jU8zNzaGenh6JkMMLMOnp6cnH6yT3RXFBWhCpaTijZ0Qb7O7u4vb2Np6dnfEuhRweHrL48/Nz3qURysvLSYBrAHASBPiUln9raysfqxY1NTXo7u6OZmZmaGxsjHZ2dpiSkoJLS0t8KGNlZQXj4uLQ2tqaxdvb27OCfHp6yoeqRV9fn1AMqWtkBJFhYGCAj1WJu7s7lqiC7owNS0tLXFhYkM3Z3NxEGxubV2JpBAYG4vX1tSxeHcbGxoRrhwkCUA/NOilN0NHRId68hYUFlpaWYn19Pfr6+or2oKAg2ZzExETRl5CQwOLd3NxEW21trSxeHYaHh4XrhgoC0EGCtZOaIDIyUuGNHx0dseVNdhMTE9zb22N2qhG05Mnu4OCANzc3zD4/Py9UbPTw8GArSxNoXYDGxkZMT0/HpKQk3NnZkflcXFzYhxsaGuLW1haz0flCEIxWggAl7OzszOwkEAmlCbQugDI2NjbYN0+f5eTkhJeXl8xeWVkpCpCVlSWbQ+cRwTc9PS3zqcobEyA6OlpMJjs7W7RL+nQsKiqSzQkNDRV9dNLTBG9EgNTUVDER2t5orxfIzMwUfcXFxbJ5UgHa29tlPlV5VAGooCUnJ4tJ0LM8NTUli6HVoEyAsLAw0dfV1SXzqcqjCUB7d2xsrJiAqampwl6DkhZi8vPzZT7qAQSfprbpRxGAKjjt58LNW1lZ4eTkJB/GaGhoEOPS0tJkPldXV2Z/8eIFLi8vy3yq8igCSJ9r2sqot1fGzMyMGOvj4yPaj4+P0dzcnNkdHR3FXUNdtC7A+Pi4mBANOoNXVVVhSUmJOKg7pASJq6srcb+nb7qtrQ1PTk4wJydH6cpQB60LIN3uXjdWV1fFOdQ8SX22trbi30ZGRri2tib7DHXQqgAHBwdoYGDwSrKKxvr6umxuWVkZS1YaQwekzs5OWZy6aFUAOvNTpZe+rlY0qKJfXFzw01mhq6urY49IU1MT7u/v8yFqo0gAjZ4GnzqKBAgmg6I9WhdR9D7AgwxUfd8G6Iu+FyBQEOBDAPiHnru3AXpVBwD/AoCzIADxW0BAAB+rk0RERJAAfwCAnlSA76kJoa5Ml6EXsvQyBgB+kCZPvAcA+/TOXJeJiYmh5P9W9v8D9JsZa1V1kYqKCqH4fccnLqWegnStIEqSb+UT5qHC8AsF03KZnZ3lr/WsWFxcxPj4eCH5XwHAmE9YGfRj6R4VxvDwcPbisre3l530qJN6qmNiYoL9U0d1dTVGRUWhvr4+JX5ARZ5P8CHYAEAhAPwOAFf3Kj6XQfv8IgCUAsBHfGKqQE0DdU7UPlIP/VQH3R+19h/zCSjiPyHxLQ+roYPpAAAAAElFTkSuQmCC",
  21: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAARYSURBVGhD7ZptKH53GMe/8xQZ5mEPTIlayYuteUpYwx48pRQNebPaS7OpqRHKvNnyEJr2QitvaFK88Rjm+cXQkppiiBRGWInEcK3vyTndfm7/7vAv52yfut78zvU79/l+z7mvc/1+9w38j8O8DuBdAB8A+BjAR88weF28vvcAvKEKeAieAIoA/ArgAICYKI4ATAD4CoC3KswR8gH8yZPFxMRIWVmZtLe3S29vr4yMjMjw8PCzC15XX1+fdHR0SEVFhcTHx+tmrAP4XBX4Ir7nxISEBBkcHBQzMzY2JsnJyboRPwJ4RRWr8gOTi4uL5erqSj2faSkvL9dN+EkVbMtnungrUllZqZvwhSqcvApgKzY2Vp1nKVJSUmjAHgBf1YAv6c74+Lg6x1LMz8+Ls7MzTfhWNWA6OjpazbckN0VxwbYgsmk45nfkZbC9vS2bm5tyfHysHroX5i4vL8v6+rpcX1+rhx9FbW0tDbgAEKIb8D4f/87OTjX3UTQ1NUlkZKR4eXmJu7u7BAUFSUFBgSwtLampd0hPT9cKVlhYmFxcXKiHH8XAwIBeDNk1aiRxYGhoSM19ELxjFKp0Zkb4+vrKwsKCOs2gubnZyA0JCXlyAyYmJvTzf6obwB5a66Segp6eHkOAj4+P1NTUSGtrq8TFxRnjSUlJ6jS5vLyUqqqqW2a9jCdgdHRUP/8nugFcSGjt5FOgP76MlpYWY/zw8FACAgK0cQ8PD9nd3TWOdXd3S1RU1C3xpjWgra1NioqKJC8vT7a2tm4di4iI0D7czc1NNjY27owzsrOzDaNMacB9rK2taXeen8Xv9tnZmTbOmhEeHq6JZSE+OTkRf39/6xmQmZlp3OXS0lJjnN/9yclJ4zW5t7cn3t7e1jKgsLDQEB8cHCwHBwdqigFrA4unJQzg3c3PzzfEsx+YmZlR025hGQN44SxounhPT0+Heg1LGMDClpuba4j38/OT6elpNc0uljCgpKTEEE8RXAs4iukNYEXXxTO4Bm9oaJDq6moj2B0eHR2pUzVMb4Dt6+5FsbKyok7VoAFcQDEnNDTUXAbs7++Lq6vrHbH2YnV1VZ2usbOzo3WKzAkMDDSXAWxmWOltt6vtBRdep6en6nSN8/NzbTeX52HhfOr9AHsGPOlq8Lljz4BkDjjyjrYC9vYDojjQ1dWl5loS3ugbAz7UDXgLwAlfTf8FuFUH4B8AYboB5LfExEQ115KkpaXRgD8AONsa8I2Tk5PMzs6q+ZaCG7I3r9jvbMWT1wD8xT1zK5OVlUXxf9/3/wH+Zqa1qlakrq5OL35fq8JtaWWS1QqijfhOVbAKC8PPTObjMjc3p57LVCwuLkpOTo4u/hcA7qrg++CPpbssjKmpqVJfXy/9/f3aSo+d1HONqakp7U8djY2NkpGRIS4uLhS+zyKvCnSENwFUAvgdwPmNi2YJvucXAdQAeFsV9hDYNLBzYvvIHvq5Bq+Prf07qgB7/As+z55bDNnESwAAAABJRU5ErkJggg==",
  22: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAATTSURBVGhD7ZpZSD1VHMe/7oqZuVWYIAj64EOZqYgLLmVuCIKSG0jQo1piQoqC20PuCokPGvqipj6I4L7k/pBKiJCCYSiCmKIGEkim/uI3OMPc470q13td7r8P/F7O+c7M+X1n5szvnHuB/3kwbgA+BBAO4DMAn77A4HHx+D4C8K6YgD7YA8gB8DOAYwD0iuIUwByArwG8LSb2ENIB/M4nCwwMpKKiIurq6qKhoSGampqiycnJFxc8ruHhYeru7qaSkhIKCQmRzfgDwJdignfxPR8YGhpKY2Nj9JqZmZmhqKgo2YgfAJiJyYpUszgvL4+urq7E871aiouLZRNaxYTVfCEnb4qUlpbKJnwlJs68BWAvKChIPM6kiI6OZgMOATiJBuSyO7Ozs+IxJsXq6ipZWFiwCd+JBiwGBASIepPkZlJcU0+IXDSc8TtiDPb392l3d5fOzs7ELq0cHh7Szs4OHR8fi10Goba2lg24AOApG/AxP/69vb2i9lE0NzeTv78/OTg4kK2tLbm7u1NGRgZtbGyIUon+/n6KiIggZ2dnsrGxITc3N4qJiaHx8XFR+ihGR0flyZCrRolIbjDUha6vr6VEtVRnUjg5OdHa2prGMTU1Nbd06ujp6dHQP4a5uTn5vJ/LBnANLVVShmBgYEAZuKOjI1VWVlJbWxsFBwcr7ZGRkYp+e3ubzM3NpXYzMzPKzc2ljo4OSktLU/QuLi50enqqcR19mZ6els8bIxvACwmpnDQE8fHxysBbWlqU9pOTE3J1dZXa7ezs6ODgQGqvqqpS9JmZmaozkYZpExMTGn36YnQDOjs7KScnR7qDe3t7Gn2+vr7Sxa2traWJjuHrFhQUSK8NP55qsrOzFQMGBwc1+vTF6Abogh91vvN8LU9PTzo/PxclGnAp7uPjoxiwubkpSvTi2QxITExUkiksLBS7b8ErOlkfHh4uTa6G4FkMyMrKUpLx8PC49xtfXV2t6HmCXFlZESV686QGXF5eUnp6upIM1wNLS0uiTAP1pMjR2toqSh7FkxlwcXFBycnJSiL29vb31hqqFZsUDQ0NouTRPIkB/L6mpqYqiXB1t7i4KMo04HpBnXx7e7soMQhPYkB+fr6SiJeXl7QWuAvebpP1XDpzuWosjG7A/Py8xp3kNTg/yuXl5Urw3ebCiOHFD1eMst7b25saGxupoqJC0ZeVlb2ez6D6c3dXbG1tSXrx0dcVvNlpCIxqwNHREVlZWd0avLbgwojx8/O71act+vr6xMvphVEN4DU/z/TyVrWu4IWXXAnyK8N1vqhRB/fzq2IItBlg0NXgS0ebAVHccN832lTQth/wCTfwjsybAN/oGwMiZAPeB/A3z8ZvArxVB+BfAF6yAcwvYWFhotYkiYuLYwN+A2ChNuBbXnUtLy+LepOCN2R5MwZAhTp55h0Af/KeuSmTlJTEyf+l6/8D/JuZVHqaInV1dfLk942YuJo2FpnahKhKvldMWIQnhh9ZzI+LIXdinoP19XVKSUmRk/8JgK2YsC74x9IDnhhjY2Opvr6eRkZGpLKVK6mXGgsLC9KfOpqamighIYEsLS058SOe5MUEH8J7AEoB/ArgnxsXX0vwd34dQCWAD8TE9IGLBq6cuHzkGvqlBo+PS3tvMQFt/Afpm0KHSQMVxQAAAABJRU5ErkJggg==",
  23: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUGSURBVGhD7Zp7KP5XHMc/c4vMNdtkSqmVpK0Zkktu29wiRUNSq/1pRk0uUUyy5RI1TTHxB5HkD/cw9zK0UPOHMbeSuV8iMXzW5+Scnuf8Hr/8novf8zztVZ9/zvmc85zP+3uez/mc7/MA/M+zeQ8APgaAYAD4HAAi9NBoXbS+TwDgfTkAdbAGgAwA+A0AjgAADchOAGACAL4DAFs5sOeQAgB/0WS+vr6Yn5+Pra2t2NvbiyMjIzg8PKx3Ruvq6+vDtrY2LCwsxICAAC7G3wDwtRzg6/iRBgYGBuLg4CAaMmNjYxgWFsaF+BkA3pGDlfmJnDMzM/H+/l6ez2ApKCjgIvwiB6zIVzx4Y6SoqIiL8I0cOPEuAOz4+fnJ44yK8PBwEmAfABxkAb4ldcbHx+UxRsXCwgKampqSCHmyANM+Pj6yv1HymBQXFRMiFQ0X9B3RBbu7u7i1tYUXFxdyl0r29/dxe3sbr66u5C6tUFFRQQLcAoAbF+BT2v4dHR2yr0bU1tait7c32tjYoKWlJbq4uGBqaiqurKzIroyuri4MDg5GBwcHtLKyQjc3N8zKysKzszPZVSMGBgZ4MqSqkRFKDUNDQ7KvWjw8PLBAVVRnzCjAxcVFpTH19fWv+HGjxHx+fq7krwkTExN87i+5AFRDs0pKG3R3d4vF29nZYWlpKTY0NKC/v79oDw0NFf57e3vsiVM77RTyb2pqQg8PD+FfXFys9BmaMDo6yuf9ggtAFwlWTmqD6OhosfC6ujrRfnx8jE5OTqydAqbACSqvzczMWDuV3JypqSkxD301tIXOBWhubsaMjAxMTk7GnZ0dpT5PT0/24RYWFri5ucnarq+vWZKkhVEC5MzMzAgB4uLiFGbRDJ0L8BTr6+tiq1OCo8BVcXp6itPT0+wSxgWgS462eGsCxMbGioBycnLkbkF6errws7e3x/b2dtlFI96KAGlpaSIoV1dXPDo6kl0YdIJ4eXkJX1tbW8zNzdXqKfCiAtzd3WFKSooIiLI8fbefgvxp+8/NzbHA+biQkBC8ubmR3dXixQS4vb3FhIQEEYS1tfUb1xoRERFifE9Pj9ytFi8iAG3lpKQksXhHR0f2ZFVBvpeXl7ixscFEUyQvL0/MUV5ertSnLi8iQHZ2tli4u7s7O+aegoocygt0NHZ2dir1xcfHi3momNIGOhdgcnJSLJqM7uDV1dVYUlIijKq9k5MT5t/S0iJ8qfpbWlpi9X9jY6MokOgKS8eoNtC5AIrH3etsdXWV+dO2p0qPt1PQzs7OSr4kmLbQqQAHBwdobm7+SrCqbG1tTYw7PDxU2u7c6CZZVlam9BmaolMB6M5PmV7xdbUqo4uXqvv+7Owsu0bTE6eSmpfL2kSVAFq9Deo7qgQIo4Y3PaMNFVXvAz6jBvkIMlboQT8KEMIFcAaAS21mWn2GcgwA/AsA7lwA4vegoCDZ1yiJiooiAf4EAFNFAb43MTFhlxBjhl7IUsUJAD8oBk/YA8A/9M7cmKE3SwBw+tT/B+g3M1aqGiOVlZU8+WXJgSvSQE7GlhAVgu+QA5ahxPArOdN2mZ+fl+cyKJaXlzExMZEH3w4AlnLAT0E/lu5RYoyMjMSqqirs7+9nNz2qpPTV6FU6/amjpqYGY2Ji+G3ygJK8HOBz+AAAigDgDwC4eVTRUIzO+WUAKAWAD+XA1IGKBqqcqHykGlpfjdZHpf1HcgCq+A/1W1gU8U4cgQAAAABJRU5ErkJggg==",
  24: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAATXSURBVGhD7ZprKK1ZGMf/rrmMcR9CFKF8MMYtYRrMxTmkFA06H0zNR2NGjUJ8MFIzucRpGKUpJRoJXzjnCOO4lVuTaE45mEjJjGs0Ygye6Vmd923vdTDaZ5P9zvzq+bDXetZ61/Pf613rWWtv4H9ujDuAUADvA/gIwIf30HhcPL53AbwjB2AI9gDyAPwMYAcAmZDtARgB8CWAt+XAbkI2gCXuLCoqioqLi6mtrY16e3tpcHCQBgYG7p3xuPr6+qi9vZ1KS0spNjZWEeM3AJ/JAV7Ht9wwLi6Onj17RqbM8PAwJSYmKkJ8D8BMDlbmO3bOz8+n8/NzuT+TpaSkRBHhBzlgXT5VgtciZWVligify4EzbwFYj46OlttpiqSkJBbgDwDOsgBfsDrPnz+X22iK2dlZsrCwYBGKZAHGIyMjZX9N8mpRnNNdEDlpOOR35DbY2NigtbU1Ojw8lKuu5eDggBYXF2l5eZnOzs7kaoOpqqpiAU4B+CkCvMfTv6OjQ/Z9I+rr6yk8PJwcHBzIxsaGvLy8KCcnh168eCG7vsbp6anIP3hcrq6utL29LbsYzNOnT5XFkLNGQQIX9Pf3y74GcXFxIQKVMjPVnJ2daW5uTm6mR1FRkerv5ORkVAFGRkaUvj9RBOAcWmRSxqCnp0cdvKOjI1VUVFBzczPFxMSo5QkJCXIzlfHxcTIzM1N9jT0DhoaGlL4/VgTgg4RIJ43Bw4cP1cE3NDSo5bu7u+Tm5ibKbW1taXNzU68dc3R0REFBQWp7kxSgpaWF8vLyKCsri9bX1/XqQkJCxMOtra1pdXVVr47hdlzP64a7u7tpCnAVKysr4pvnZ/n5+dHx8bFePa9Byrfe1NSkJC3aESA1NVUNsLCwUK9uf3+ffH19RV1ycrIoi4iI0I4Ajx49UoP38fGhnZ0dvfrc3FxRx1OfcwYmNDTU9AXgBCY7O1sNnvOBiYkJPZ/u7m5RZ2VlRV1dXWq5MgN44TTmyfTOBOBkJj09XQ3e3t7+tVxjb2+PPD09Rb2Hhwc1NjbS48ePhXl7e6vtKisrqbW1Va+todyJAJwMZWZmqsG7uLiI/V1maWlJ9fk3413BGNyJAAUFBerA/f391fdahncGOdCrLCAgQG5uELcuwOjoqN7AeTurra2l8vJy1Tg75Ol/cnJC09PTNDU1pRp/ZgsMDBTtOZvkV2dhYUF+lEHcugC629119vLlS7mpHmFhYcKPXx85Z3gTblWAra0tsZrLwV5mfMy9juDgYOFnZ2cn+jUWtyoAn/l5uupeV19mfPDivP86JicnRV98S8U7irG4TACjngbvO5cJkMgF8h6tVS67D4jggs7OTtlXk+gcuj5QBPAE8CdvTf8F+KoOwN8A/BUBmKn4+HjZV5M8ePCABfgVgIWuAF+bm5uLBETL8IUsX8YA+EY3eMYJwO98Z65l0tLSOPj9q/4/wL+ZiVRVi1RXVyuL31dy4Lo0s5PWFkSd4DvkgGV4YfiRnXm6zMzMyH2ZFPPz85SRkaEE/xMAGzngq+AfSzd5YeQ7upqaGnry5Ik46XEmdV9tbGxM/Kmjrq6OUlJSyNLSkgPf4kVeDvAmeAAoA/ALgL9eqWgqxvv8PIAKAN5yYIbASQNnTpw+cg59X43Hx6l9oBzAZfwD5BpI5woHrN4AAAAASUVORK5CYII=",
  25: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUNSURBVGhD7ZpZLP1HFMePNUTVEm1FiSBCPLSp2oIqutgikaC2lyZ9o1RSSQkPtmhjCVGRkIoXUhE82IPaH2ppkJREKWKJVtBELLGe5kzML/c37v3H/7rXcttPcl5mzvx+c7537pkzcy/A/9ybtwDgPQD4CAA+BYBPnqHRvGh+7wPA22IA6mAGAGkA8AsA7AMAviA7BIBRAMgAgDfFwO5DIgD8QQ/z9vbG7OxsbGpqwq6uLhwcHMSBgYFnZzSv7u5ubG5uxtzcXPT39+di/AkAX4oBvorvaWBAQAD29fXhS2Z4eBhDQkK4ED8CgJ4YrMgP5Jyeno7X19fi814sOTk5XIRaMWBFvuDB6yJ5eXlchK/EwIk3AGDTx8dHHKdThIaGkgB/A4CVKMDXpM7IyIg4RqeYmZlBAwMDEuE7UYAJLy8v0V8nuU2Kc4oJkYqGI/qOaIOdnR3c2NjAo6MjsetJKC0tJQEuAMCRC/ABLf+WlhbR90FUVVWhp6cnmpubo4mJCdrZ2WFSUhIuLi6KrgxKvrQKfX19ZUZ5KSgoCHd3d8UhatHb28uTIVWNjGBq6O/vF33V4ubmhgWqpDpjZmVlhXNzc7Ix5+fnaGtre8dX0dbW1mRj1GV0dJQ/83MuANXQrJLSBB0dHdKkLSwssLCwEOvr69HPz09qDw4Olo1ZWVlBQ0ND1mdmZoaOjo7o4OAgmaurK25vb8vGqMvQ0BCfx2dcADpIsHJSE0REREiB1tTUSO0HBwdoY2PD2k1NTWVLurOzUxpTUlKCFxcXbFWcnp7i8fExnpycsJWlCbQuQGNjI6alpWFCQgJubm7K+jw8PNjLjY2NcX19XWovLi5m7Xp6ejg1NcXatFWJal0AVayurrJPnt5FS/zs7Ezqi4+PZ+1GRkYYFRWF7u7u6OTkhJGRkRo/jzyZABQYX+ZZWVlSO33SfGWostraWtmzHsKTCJCSkiIFY29vj/v7+1Lf1tYWr85YIszIyMCGhgZMTEyUxujr6+PS0pLsmeryqAJcXV3JAqF6YHJyUuZDXwXamminEPsoj/CxdCehCR5NAMrkMTExUgC0vb1urdHe3i6Nj46OFrvV4lEEoC0rLi5Omry1tTVOTEyIbjIuLy/vZP6enh7pGbS9aoJHESAzM1OauLOzMzsLqKKtrY0VSS4uLlhUVCTru63bmdHWqgm0LsDY2Jg0aTI6g1dUVGB+fr5kVB0eHh4yf4UJoaWlJRtPK2h+fp6dH3ifmB/UResCKG53r7Ll5WVpDK8DyKgYcnNzY8USb0tNTZW94yFoVYC9vT1WzIjBKjOq/zlU7iYnJ9/xoW2Rvk6aKoMJrQpAZ37K9IrX1cqMDl5U34vMzs5idXU1FhQUYF1dncqj80NQJoBGT4PPHWUChFDD6+7RLxVl9wEfUkNra6voq5PQB30rwMdcAFsAOKat6b8AXdUBwCUAOHMBiF8DAwNFX50kPDycBPgdAAwUBfiWTlz8MkJXoV3ltr4oUAyesASAv+jOXJehAxUA/KPq/wP0mxkrVXWRsrIynvy+EQNXpJ6cdC0hKgTfIgYsQonhJ3Km5TI9PS0+60WxsLCAsbGxPPifAcBEDFgV9GPpLiXGsLAwLC8vZ2dzOqlRJfVcbXx8nF2iVlZWsgvV298a9ijJiwHeh3cAIA8AfgOA81sVX4rRPr8AAIUA8K4YmDpQ0UCVE5WPVEM/V6P5UWnvKgagjH8BKSA9CNJJGyMAAAAASUVORK5CYII=",
  26: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUoSURBVGhD7Zp7KL5nGMcvp1+OM7KRiaL9o4wZEmbYzClR5JCU2h/CbJga8QdSWw4h2h9aUcjhD/84hjnH0Ejxh7HIcRMmOTSna11P73P3PPfv9Uve9/V7vdunruK6r/t57+v7PM99X/f9vgD/82jeA4CPAOBTAPgCAD7XQqNx0fjcAeB9PoGnYAYAWQDwCwAcAQC+IDsBgHEA+AYA3uETewxJAPA7Xczb2xsLCgqwtbUVe3p6cHh4GIeGhrTOaFy9vb3Y1taGRUVF6OfnJ4rxBwCk8Qm+iR+oo7+/Pw4MDOBLZnR0FIODg0Uh6gFAj0+W50cKzs7Oxru7O/56L5bCwkJRhJ/4hKUkiMnrIsXFxaIIX/GJE+YAsO3j48P30ylCQkJIgL8AwIoX4GtSZ2xsjO+jUywsLKCBgQGJ8D0vwJSXlxcfr5MoJsUl6YRIRcMZvSOaYG9vD7e2tvDs7IxvUgpNvru7u7izs8M3qYWKigoS4BoAnEQBPqbHv6Ojg49VidraWvT09EQLCws0NjZGe3t7TE5OxtXVVT5U4P7+Hqurq9HNzQ3Nzc3RzMwMPTw8sKmpiQ9Vif7+fnEypKpRIIgcg4ODfOyToEQoUSXVmWBWVla4tLQk63Nzc4OxsbGvxYpWV1cni1eF8fFx8bpfigJQDS1UUuqgu7ubDdzS0hLLysqwsbERfX19mT8oKEjWR/FYCubq6orNzc1CP8WEhXp6eri/vy/r81RGRkbEzwoVBaCNhFBOqoOIiAiWTENDA/MfHx+jjY2N4DcxMcGDgwPBf35+Lrwe5Dc1NcWNjQ3WJyMjAx0dHdHd3R3n5uaYXxU0LgC9s1lZWZiYmIjb29uyNrq79FmvXr3Czc1NwSd5JIV1muf29pZ3qYTGBXgIurN05+mznJyc8OrqSvDX19czAahkXVxcxMzMTExNTRXaHrt6PJa3JkBUVBRLND8/n/lp5yb6adYX/xbNxcUFV1ZWZNdShbciQEpKCkvIwcEBj46OWFtubq4sYVtbW8zLy8Po6Gjmo1fn8vJSds2n8qwC0PublJTEEqF6YHp6WhZDT4PYTmu/9G5L+6qrTnk2Aa6vr2VrOyWnrNYoLS1lMYGBgbK2vr4+1qauXeqzCEDFUHx8PBu8tbU1Tk1N8WECLS0tLC40NFTWNjk5ydrS0tJkbU/lWQTIyclhA3d2dhb2Ag9BpTEVOhRrZ2fHVgeivb2dXYeO5dSBxgWYmJhggyajtZ1q/JKSEmZU5Z2cnLA+kvM7TE9PF9rW19dlq4K6tuoaF0C63L3J1tbWWJ+ZmRlW9pLRSkBVofh/TEyM7DNUQaMCHB4eopGR0WvJKjO6w1LoxJkKJD4uISFBrcWQRgWggdJMLz2uVma08bq4uOC74+npKXZ2dmJ5eTnW1NTg7OwsH6IyygRQ625Q21EmQDA5lK3Ruoiy84BPyNHV1cXH6iR0oxUCfCYKYAcA57Q0/RegozoAuAEAZ1EA4teAgAA+VicJDw8nAVYAwEAqwHf6+vpqO3XRVqjqpMMYACiVJk+8CwB/0pm5LqPYYv/90O8H6DszoVTVRSorK8XJ71s+cSmNFKRrE6Ik+Q4+YR6aGH6mYHpc5ufn+Wu9KJaXlzEuLk5Mvh0AjPmEH4K+LD2giTEsLAyrqqqEwwna6VElpa1G5wf0ow4qoyMjI9HQ0JASP6RJnk/wMdgCQDEA/AYA/yhUfClG6/wyAJQBwAd8Yk+BigaqnKh8pBpaW43GR6X9h3wCyvgXYWUlifKyhwgAAAAASUVORK5CYII=",
  27: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAASbSURBVGhD7ZpZSH1VFMa/nFDMHLABEwQhBMF/5IwaqQ1OCIrigC9BDyJmCQkp+mC+lBMqSQ8SKqImPvjiPOT8kEqIkA8OqQhijoEkkqk71sGzOWd771+59xy9nvrBQjhr7XvP993rumvve4H/eTCvA3gB4H0AHwH40AaD7ovu710Ab4gCLMEVQCGAnwEcA2DPKE4BTAP4AsBrorCHkANgnR4sLCyMlZaWss7OTtbf38/Gx8fZ2NiYzQXd18DAAOvq6mLl5eUsKipKNuN3AJ+KAl/Gt7QwOjqaDQ8Ps+fM5OQki4uLk434HsAroliR76i4qKiIXV9fi4/3bCkrK5NN+EEUrCRLFm9EKioqZBM+E4UTrwLYDQ8PF9cZivj4eDLgAICnaMDn5M7U1JS4xlAsLS0xe3t7MuFr0YC50NBQsd6Q3DbFZWVDpKHhjP5H9GBvb4/t7Oyws7MzMfUk1NTUkAGXAPxkA96jt39PT49YaxWNjY0sODiYubm5MWdnZ+bj48Nyc3PZ6uqqqq66upqFhISwiIgIs0H57u5u1TpLGRoakpshTY0SsXRhZGRErLWIm5sbSaiJ6UwKT09Ptry8zOvz8vLu1JiKyspK1fNYyvT0tPyYn8gG0AwtTVJa0NfXx2/a3d2dVVVVsZaWFhYZGcmvx8bG8noSFhAQcCeCgoLkhiX9paFGCyYmJuT7+Fg2gDYS0jipBUlJSVxoc3Mzv35ycsK8vb2l6y4uLmx/f1+1TqSjo4Mb0NTUJKYtRncD2traWGFhIcvOzma7u7uqXGBgoPTkTk5ObHt7W5VTQo3Tw8NDqk1OThbTVqG7AebY3NyUXnl6Lj8/P3ZxcSGWcLKysqQ6ap4bGxti2iqezICUlBT+r1FSUiKmOQsLC7wuPz9fTFvNkxig7PS+vr7s+PhYLOFkZGTwV399fV1MW82jGnB1dcVycnK4eBI1Pz8vlnG2trak/kC16enpYloTHs2Ay8tLlpaWxsW7urreO2vU1dXxejrY0INHMYCGoczMTC7Gy8uLzc3NiWV3kA8vaHo8ODgQ05rwKAYUFxdz8f7+/tJe4D6Ojo6kwYnW0NCkF7obMDMzw8VT0B68vr5emvjkoOnw9PTU7LqCggJVTkt0N0D5cfeyWFtbU61rbW3luYaGBlVOS3Q14PDwkDk6Ot4RayrEAYfeFXKuvb1dldMSXQ2gPT91euVxtamgjdf5+blqLU2KtHZ0dFS3BkiYMkDT3aCtY8qAOLpw32e0UTB1HhBCF3p7e8VaQ0Iv9K0BH8gGvAXgL2pC/wXoqA7APwD8ZQOIX2JiYsRaQ5KYmEgG/AbAXmnAV3Z2dtJW1MjQgeztZusbpXjCA8AfNIsbmdTUVBL/p7nfD9B3ZpqdvtoatbW1cvP7UhSupIWKjNYQFeJ7RMEi1Bh+pGJ6uywuLoqP9axYWVnhJ0wAfgLgLAo2B31Zuk+NMSEhQTqoGBwclHZsNEnZaszOzko/6qDNFJ0mOzg4kPBDavKiwIfwJoAKAL8C+PvWxecS9Dm/AqAKwNuiMEugoYEmJxofaYa21aD7o9H+HVGAKf4FScaNDAWdRPoAAAAASUVORK5CYII=",
  28: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAU3SURBVGhD7ZpbLGZXFMeX2yCqLtFW1ERiUiMeNFXErYpWZ4zbCCnipTQeRLWikhKXqIg2bmNSKZFmPGAqhBfMCOo6JkVECBGlIRJRxiUR4m41a8c5OWf7TOS7GL72l6yXvdY5Z6//3vZee38A/ufSvAMAzgDwCQB8DgCfXUOjflH/PgSAd/kElMEEAJIB4A8AWAcAvEG2CQB9APAtALzNJ3YZYgDgL3qZm5sbZmRkYF1dHba2tmJXVxd2dnZeO6N+tbW1YX19PWZlZaGXl5cgxt8A8BWf4Ov4iR709vbG58+f402mp6cH/f39BSF+AQAdPlmenyk4JSUFT05O+PfdWDIzMwURfuUTlvKlkLw2kp2dLYjwNZ848RYALLm7u/PPaRUBAQEkwCoAWPACfEPq9Pb28s9oFaOjo6inp0ci/MALMOjq6srHayVni+K4dEGkomGb/kY0wfLyMi4uLuL29jbvUsjq6iouLCzg1tYW71ILRUVFJMAhANgJAnxE07+hoYGPVYny8nJ0cXFBU1NTNDIyQhsbG4yNjcXp6Wk+lPH06VP09PRECwsLvHXrFlpbW2N4eDiOj4/zoSrx7NkzYTGkqpHhRw0dHR18rFKcnp6yRBVUZ8woQT6pioqKc3GCkYB8vCr09fUJ7/5CEIBqaFZJqYOWlhax82ZmZpifn4/V1dXo4eEhtvv5+YnxGxsbaG5uztpp5CmeKs64uDgxPiQkRPYNVeju7hbeGygIQAcJVk6qg6CgILHjNLIClKiVlRVrNzY2xpWVFdY+NDQkxtPMkWJvb8/ab9++jUdHRzKfsmhcgJqaGkxOTsbo6GhcWlqS+ZycnMSRpoWOmJycFAVITEyUxTs6OrJ2BwcHtVWmGhfgIubn59nI07fs7Oxwb2+PtdPI+vr6snb6k6HFeGZmBnNyckRhCgoK+NcpzRsTIDg4WEwoPT1d5nv16hWGhYWJfsH09fVZHa9O3ogA0gXN1tYW19fXZf61tTWMiIg4JwBVbWlpaeJsUQdXKsDx8THGxMSICVE98OLFC1nMzs4OOjs7iyOem5uLzc3NGB8fLz738OFDtr2qgysT4PDwkHVcSMLExERhrVFbWyvGpKamynyBgYGi7+XLlzKfslyJADRaUVFRYuctLS1xcHCQD2PQ8VuI4y9gCgsLRV9VVZXMpyxXIgCNpNBx2svpLHAR0tjKykqZLykpSfQ9efJE5lMWjQvQ398vdpqMzuClpaWYl5cnGlV7m5ubLL6pqUmMvXPnDqsLaAYNDAywmSP4pqam+E8phcYFkG53r7PZ2VkWf3BwIL3ARENDQ7x79y7q6uqKbQkJCfxnlEajAtB2ZmBgcC5ZRTY3Nyd7TpFwOjo6LPn9/X3Zd1RBowLQmZ9Weul1tSKjg9fu7i7/OFvpy8rK2P3d48ePcWxsjA9RGUUCqPU0eN1RJIA/NSjao7URRfcBH1NDY2MjH6uV0ECfCfCpIIA1AOzQ1vRfgK7qAOAIAOwFAYg/fXx8+Fit5P79+yTAFADoSQX4nvbd4eFhPl6roAtZuowBgB+lyRPmAPAP3ZlrM6GhoZT81kX/P0C/mbFSVRspLi4WFr/v+MSlVFOQti2IkuQb+IR5aGH4jYJpuoyMjPDvulFMTExgZGSkkPzvAGDEJ3wR9GPpCi2M9+7dw5KSEmxvb2cnPaqkrqvRCZLuFB49eoQPHjxgN0wAsEaLPJ/gZXgPALIBYAwADs5UvClG+/wEAOQDwPt8YspARQNVTlQ+Ug19XY36R6X9B3wCivgX+M0gzdgJpKQAAAAASUVORK5CYII=",
  29: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUZSURBVGhD7ZpZSF5HFMdP3VBsTA1qixVcwBcfLE2jiKZuXdwIBJSqCFLsgw/WVqhQN4jVh5ZEUaj0QUvyoKL4IIhmIVr3h6gUkeiDdUlQXCpuiIK1mlP+gzNcJ58in99n9Et/MKDnnrl3zv/OPXNmlOh/To0nEQUT0adE9DkRfXYBG8aF8X1ERF56AObgSkQ5RPQHEa0SEV+itk5EvUT0HRG56YGdhjQi+gs3CwkJ4YKCAm5oaOD29nbu7Ozkp0+fXriGcXV0dHBjYyMXFxdzeHi4FGOGiL7WAzyJn9ExIiKCHz9+zJeZ7u5ujomJkUL8SkTv6MHq/ALn3NxcPjg40O93aSksLJQi/KYHbOQrGbwtUlJSIkX4Rg8cvEtEc6GhoXo/myI2NhYC/E1E7roA30Kdnp4evY9NMTIywvb29hDhR12AgRs3buj+NslhUhw1JkQUDVv4RqzBwsICv3z5kre2tvRLJllcXOT5+Xne39/XL1mEu3fvQoA9IvKVAnyM6d/c3Kz7nonq6mq+fv06X7lyhZ2dndnb25vT09N5YmJCdxU8ePBA1Bxubm7s6urKQUFBXFFRobudmUePHslkiKpREA3DkydPdF+zePXqlQjURHUmmru7O4+Ojh7pU1RU9JqfbBkZGeKelqK3t1fe+0spAGpoUUlZgtbWVjX4q1evcllZGdfW1nJYWJiyR0dHK//BwUFlv3btmnjrmKboK+11dXVHnnEWurq65H2/kAJgIyHKSUuQkJCgBl5TU6Psa2tr7OHhIewuLi68tLQk7NnZ2cof5bakqalJ2YODgy2WE6wuAL7lnJwcTk1N5bm5uSPX8F3jWU5OTiIxgqioKGHD8vTixQvlu7Ozw15eXuKanZ0dT01NGe5kPlYX4Dimp6fFm8ezfH19eXd3V9gjIyOVAFIUgDceEBCgZoGl9iVvTICkpCQVTH5+vrJnZmYqOzK0BEshVgN5Dbs8S/BGBEAml4H4+Pjw6uqquobVR17DJzIwMMDj4+McHx+v7Gj3798/ck9zOVcBMI3T0tJUEKgHkPV1srKyjgQrG3KF/Lm+vl7vZhbnJsDe3h7fvn1bBYDpfFKtUV5eLr55R0dHkfywfBr7t7W16V3M4lwEQOGSkpKiBo/1HVP7NMzOzvLm5qb42XCqw8+fP9ddzeJcBMjLy1MDx1s1ZncdfO8odLAXMVaIy8vLKgn6+/urVeOsWF2Avr4+FTwa9uCVlZVcWlqqGqb3+vq68EfxI31v3rzJKysrIkmijpB2nPFZCqsLYFzuTmqTk5PCf3t7W7xhaUe16OnpqX4PDAzkjY0N/TFmY1UB8PaQxPRgTTVjZTc8PMx+fn6v+aBImpmZOfKMs2JVAbDnR6Y3Hlebath4odQ1greMYufOnTtiM2SpzZmOKQEsuhu86JgSIAaGk9ZoW8LUecAnMLS0tOi+Nomh9I6SAnxARNtYmt4GcFRHRP8SUYAUADzDGvw2cLjJGicie6MAP+DQYWhoSPe3KXAge7jB+skYPHiPiJZxZm7L3Lp1C8FvHPf/A/ibmShVbZF79+7J5Pe9HriRWjjZWkI0BN+sB6yDxPA7nDFdUKJeZsbGxjg5OVkG30REznrAx4E/li4hMcbFxYmz+ocPH4qdHiqpi9r6+/vF4WlVVRUnJiayg4MDAl9BktcDPA3vE1EJEf1JRP8cqnhZGtb5MSIqI6IP9cDMAUUDKieUj6ihL2rD+FDaB+oBmOI/9cUlI/8ZjucAAAAASUVORK5CYII=",
  30: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUsSURBVGhD7Zp7KLdnGMcvh1AObTHGsJC/nJpjYRkb3l4pwtqf08Qfmw1TjmEObTlEJn9okno1vSnkGF6nUXMYqYmE5BBbMmbl7FrXk/vuee6XN/0O4rd96v7DdV/38zzX97nv67nu+wfgf+7NOwDgCQAfAsAnAPDxI2z0XPR8XgBgLQagCqYA8CUAvAKAfQDAJ9QOAGAUAL4GAAsxsPvwGQCs0MX8/PwwOzsbX7x4gV1dXTg4OIgDAwOPrtFzdXd3Y0tLC+bl5WFgYCATYw0APhcDfBPf08CgoCDs6+vDp8zw8DCGhoYyIX4EAD0xWJEfyDk1NRWvrq7E6z1ZcnJymAj1YsByPmXB6yL5+flMhC/EwAkzANj09/cXx+kUYWFhJMAfAPC2KMBXpM7IyIg4RqeYmZlBAwMDEiFLFOAXX19f0V8nuUmK8/KESEXD37RGtMHOzg5ubW3dO6nu7+/jxsYGHh8fi10aoby8nAQ4B4D3mQAf0PRvbW0VfdWipqYGvby80NzcHE1NTdHNzQ0rKytFN87y8jLGxsailZUVmpiYoL29vZSQj46ORFe16O3tZcmQqkaJj8jQ398v+qpMUlISu8lrLTExUXTH9fV1tLGxec2XWkhICJ6fn4tDVGZ0dJRdO4IJQDW0VElpAkqk7OFtbW2xvr5emg2Wlpbc3tPToxiTkJDA++Lj47GhoQE9PT25ra6uTuGvDkNDQ+y64UwA2khI5aQmKCoqQj09PekmnZ2d3F5RUcEDysjI4HbKETTlye7g4ICXl5eSfXZ2lmVs9PHxwevraz5GHbQuAK1ZWs+0pE5PT7m9qqqKC1BaWsrttL9gdpoJDArY2dlZspNAJJQm0LoAIru7u9jR0SEtB7qPmZkZrq2t8X5KjEyA9PR0xVjaj7C+yclJRZ+qPKgA9Onz8PDgQbi4uOD8/LzCR1anY0FBgaIvPDyc99FOTxM8qAA0bQ0NDXkQTk5OWFtbq/BJS0vj/YWFhYo+uQBtbW2KPlV5UAEODw+lz87ExATGxMTwYORTPTMz804BIiIieB8tI03woALIoYRoZ2cn3VxfX1+q9ggKmgWZm5urGEM1AOvT1Gda6wJQ9j44OOABypEdTvCAGhsbuS0lJUXh7+7uzgVbWlpS9KmK1gWgQoaqOip/V1dXFX2UA1iwc3Nzkm1qaorbAgICuC+JaGFhIdkdHR3x5OREdiXV0boAycnJPKCoqCjc3NyUNjjyZOfq6sprhLOzM/69pzf98uVLKXdkZWVxf3FmqIPWBdje3kZra2v+8PQWaYPD/qYmftKampoU/axmoGZsbIwrKysKf3XQugAETW9vb29FUNRoKtMbvo2SkhIpWLk/LaX29nbRVS0eRADi4uJCKodp/11WViZ9x2lqvwlKdLR5Ki4uxubmZtzb2xNd1OY2ATS6G3zs3CZAKBk0eR7wmLntPMCHDHetTV2DXvSNACFMgHcB4B9ad/8F6HAGAC4AwJkJQPwaHBws+uokz549IwF+BwADuQDfUhFCVZkus7i4iEZGRiTAd/LgibcAYI9qdV0mOjqagv/rrv8foN/MpDM9XUR2JvmNGLicBnLStYQoC75VDFiEEsNP5EzTZXp6WrzWk2JhYQHj4uJY8D8DgIkY8F3Qj6W7lBgjIyOlg0s6yx8bG5MqqcfaxsfHpX/qqK6uxufPn7PjuD8pyYsB3gcbAMgHgN8A4OxGxafS6Du/AADFAPCeGJgqUNFAlROVj1RDP9ZGz0elvasYwG38C6fUQmNDULSpAAAAAElFTkSuQmCC",
  31: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAR1SURBVGhD7ZpZSG1VGMf/DkjgVGiaBoqKIIpJTg9qlDY4IQhK3Mck8eVmqT2kqGCiFA44BD5Ej0oiCoojas4PqaQICXJREMcMMRUfnL/4Fq7N2atj1/RcOHvTD9bL2t/ae///Z+1vfXvtA/zPg3kTwDsA3gPwEYAP7bDxffH9RQHwUQU8BlcAzwH8AuAQABmoHQGYAvAlAA9V2EN4BuAFnywuLo5KS0upvb2d+vv7aWxsjEZHR+2u8X0NDAxQR0cHlZeXU0JCgjRjA8BnqsB/4zsemJiYSMPDw2RkJiYmKDk5WRrxAwAHVazK9xxcWFhINzc36vkMS1lZmTShTRVsyadSvBmpqKiQJnyuCmfcAGzFx8er40xFSkoKG3AA4A3VgC/YncnJSXWMqVhcXCQnJyc24RvVgNnY2Fg13pTcJcVly4TIRcMpPyOvgt3dXdre3v5PSfX09JTW1tZoY2ODbm9v1cNPoq6ujg24BBAoDXiXp39nZ6ca+ySam5spKiqK3N3dydXVlSIiIqihoUENs0p6erpIWMHBwXR5eakefhJDQ0MyGXLVKPiAO0ZGRtTYR5Ofny8v8o+Wl5enhutoaWnRYgMDA21uwNTUlDz/J9IArqFFJWULOJFKAX5+ftTW1iZmg5eXl9Y/ODioDqPr62uqrKzUmfUqZsD4+Lg8/8fSAH6REOWkLaiqqiIHBwdxkb6+Pq2/vr5eE1ZSUqIb09PTQzExMTrxhjXg5OREJDB+pM7Pz7X+xsZGTVhNTY1uTHh4uHYsOzubvL29jWuAyv7+PvX29orHga/j5uYmsruEs3xYWJgQy4n47OxMe1wMbwAvfZGRkdqvGxISQsvLy7oYfvanp6fF0sccHByQh4eHOQzgGsDZ2VkzICgoiFpbW9UwHTxjPD09zWHA8fGxWHbm5ubEsy2NKC4uVkM1TGWAJZwQ/f39xcUdHR1pc3NTDREY3gBOakdHR1YFWmxO3Ft3GN6A3Nxc8vX1FeXv+vq67hjnAGnA0tKS7pjE8AYUFBRoIjMzM2lra4sODw+pqKhI6w8NDdXVCJYY3oCdnR3y8fHRxPKSJgsb2Xjz8j7YAH6B4jieMYYzgOHpHR0drRPNLSAggLq6utRwHXt7e+Ti4iLiuXgypAHM1dWVKIf5/bu2tpa6u7vFsvgyLi4uxG4uj52dnbX5foA1A2z6NmjvWDMgmTtsuR9gz1jbD4jhjpc9m2aBf+g7A96XBrwF4Ky6ulqNNSW8OQPgCkCwNID5NSkpSY01JWlpaWzA7wCcLA34muvz+fl5Nd5UrK6uyiX2W0vxzOsA/uBa3cxkZWWx+L/u+/8AfzMTe3pmxGJP8itVuCU/cpDZEqKF+E5VsAonhp84mKfLwsKCei5DsbKyQjk5OVL8zwBeUwXfB38s3efEmJqaKr7o8F4+791xJWWvbWZmRvypo6mpiTIyMuR23J+c5FWBD8EXQAWA3wBc3LlolMbr/AqAagBvq8IeAxcNXDlx+cg1tL02vj8u7UNVAdb4G8Sys6/n0FlsAAAAAElFTkSuQmCC",
  32: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAATuSURBVGhD7ZpbLHxXFMaXS2ji1rqWJhoVXkSpIuISpVVCJBJS5UEqFS+lRR9KSNxTd6KJh2o8oYgH4n6pWz0UKZGUSEOIS1QjStMHrbCatWPvnNkzGpn/GR3T/pL9MHuvfc75vjlnnbX3DMD/PBoXAHgTAKIA4D0AeNcIG10XXV8AALjKAvTBBgA+AYDvAOAcAPAZtQsAWACATwHAXhb2GD4EgJ/pYCEhIVhcXIzd3d04MjKCMzMzOD09bXSNrmt0dBR7enqwtLQUw8PDuRl7APCRLPCf+JImRkRE4MTEBD5n5ubmMCYmhhvxFQCYyWJl6ig4Pz8fb29v5eM9W0pKSrgJHbJgJR9w8aZIWVkZN+FjWThhCwCHoaGh8jyTIjY2lgw4A4BXZAPyyJ35+Xl5jkmxtraGFhYWZMIXsgHfBwcHy/EmyX1S3FAmRCoafqdnxBCcnJzg0dHRo5Pq2dkZ7u/v4/n5uTykCg0NDWTAXwDwOjfgLbr9+/r65NgXoq2tDQMCAtDOzg5tbGzQz88Pm5qa5DDBwMAARkdHo6OjI1pbW6OLiwvGxcXh5OSkHPpCjI+P82RIVSPjHepQ80Q5OTn8JFotOztbDsf6+nqtOGXr7e2Vp+jNwsICP+773ACqoVklpQaUSPmFu7u7Y0dHB7sbnJycRP/Y2JiI393dRXNzc9ZvZmaGeXl52NXVhenp6SKe5l5cXGicR19mZ2f5ceO4AbSQYOWkGlRUVDAhdMzh4WHR39jYKAQVFRWJ/urqatGfmZkp+omwsDAxNjU1pTGmLwY34OrqCnd2dtgjdX19Lfqbm5uFmJqaGtFP5yVDMjIy2O2pJCsrS8wZGhrSGNMXgxsgc3p6yi6eHgc6j62tLe7t7clhWtBbw9fXVxiwvb0th+jFkxpAIvz9/YUIb29v3NjYkMN0Qis6Pi8qKgrv7u7kEL14UgOoBrC0tBRCvLy8sL29XQ7Toq6uTsyhBLm6uiqH6M2TGnB5ecme6+XlZUxJSRGiCgsL5VCBMilSo7eImjypAUooIXp4eIhv9eDgQA5RrthYo8SpNgY3gJ5VemfrEqjYnNCqO6qqqjTEd3Z2aoyrhcENSEtLQzc3N1b+UpGjhHIAF7i+vi76abuN91PpTOWqoTC4Abm5uUJMUlISHh4esoVNQUGB6Pfx8RE1Ai1+HBwcNMZaWlqwsrKSFVXUysvLn89r8Pj4GF1dXYUge3t7dHZ2Fp+p0eYlR771H2q02akGBjeAoNs7KChIS4Snpydb9XEoXwQGBmrF6Wr9/f0a59CXJzGAuLm5YeUwrb9ra2txcHCQvRaVkAGLi4uszufb27oajdOjoga6DFB1NWjs6DIghjrU3A8wZnTtB7xNHcpn05ShL/regGhuwKsA8Adl4/8CtDkDADcA8AY3gPghMjJSjjVJEhISyICfAMBCacDnVJ+vrKzI8SbF1tYWWllZkQGVSvHEywDwC9XqpkxycjKJ/+2h/w/Qb2as9DRFFHuSn8nClXxNQaaWEBXi+2TBMpQYvqFgul3U3In5N9jc3MTU1FQu/lsAeEkW/BD0Y+kpJcb4+Hj2iw7t5VPZSpWUsbalpSX2p47W1lZMTEzk23G/UpKXBT4GNwAoA4AfAeDPexefS6P3/CYAVAHAa7IwfaCigSonKh+phjbWRtdHpb2PLEAXfwPP2VfPaNy+SgAAAABJRU5ErkJggg==",
  33: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUxSURBVGhD7ZprKK1ZGMcfh0QuEeM2ojNSCqMx+OCSwcw4ESma5osajXyZMS5TLlGMMJNLZIpyRnxwIvngThj3MofGpUadpoPjljHkMpG7Z3pW1mrvxZmjbW9tu/nV+rDX+3/fdz3/vd7nfdbaG+B/7sx7APAhAAQBwKcAEKaFjcZF4/MCABs5AFUwAYBvAOBXANgBAHxEbRcARgDgOwAwlwO7C18CwJ90MV9fX8zKysLGxkbs7OzEgYEB7O/v17pG4+rq6sIXL15gTk4O+vv7czMWAeArOcD/4kc6MSAgAHt7e/ExMzQ0hCEhIdyInwFATw5W5icSJycn4+XlpXy9R0t2djY3oVoOWJEvePC6SG5uLjfhazlwwhQAVv38/OTzdIrQ0FAyYAsALGUDviV3hoeH5XN0iunpadTX1ycTMmUDxn18fGS9TnKdFGcVEyIVDf/QM6IJNjY2cG1t7c5JdWtrC1dWVvDo6Eg+pBZKSkrIgDMAcOYGfETTv7m5Wdbei8rKSvTy8kIzMzM0MTFBd3d3LCsrk2WC1tZWDAoKQktLSzQ2NkZnZ2dMSUnB/f19WXovenp6eDKkqpHxCXX09fXJWpVJTEzkN7nREhISZDnW1NTc0PFGifng4EA+RWVGRkb4tT/nBlANzSopdUCJlA/e3t4eq6ur2WywsrIS/d3d3UK/ubnJvnHqNzIywoKCAqyrq0M3Nzehz8vLU7rHfRgcHOTX/YwbQAsJVk6qg/z8fNTT02M3aW9vF/2lpaUioPT0dNFP5bWBgQHrp5KbMzY2JvT0aKgLjRtA0/XVq1fskTo5ORH95eXlIqDCwkLRf3x8jG/evGEDowTImZiYEPqoqCjRf180boAMTfG2tjb2ONB9TE1NcXFxUZYJ9vb2cHx8nC3CuAG0yFEXD2oAvfo8PT1FIC4uLjg7OyvLlIiPjxd6CwsLbGpqkiX34kENoBqAP9/Unj59ilVVVbJMcHV1hR4eHkJvbm6OGRkZan0LPKgB9A6n1w49zzExMSKwtLQ0Wcq4uLhg0//ly5cscK4PDg7G09NTWa4SD2qAIpQQHRwc2M2fPHnCEt+7CAsLEyZ0dHTIh1VC4wbQNN7d3b01QIXNCVF3kP7w8BCXlpbw7OxMSZ+ZmSn0xcXFSsdUReMGxMXFoa2tLSt/X79+rXSMcgAPaGZmhvVRkePo6IiGhobY0tKipI+Ojhb62tpapWOqonEDkpKSxKAjIyNxdXUVd3Z2MDU1VfS7urqKGqGhoUH0U/U3NzfHcsfz589FAqUlrGymqmjcgPX1dbSxsRFBUSa3trYWn6kpvtdp2lOlx49R0HZ2dkp6Ko/VhcYNIGh6e3t7KwVBzcnJ6cY0J7a3t5WmO2+0klSsGtXBgxhAnJ+fs3KY1t9FRUVsufuupe3k5CRbONE3Xl9fj8vLy7Lk3txmgFpXg9rObQaEUIc69wO0mdv2Az6mjtueTV2EvuhrA4K5AXYAcKjOTKvNUI4BgHMA+IAbQPwWGBgoa3WSZ8+ekQF/AIC+ogHfU31OixBdZmFhgVWcAPCDYvCEBQD8RbW6LkM7SwCw97b/D9BvZmxPTxdR2JNMkQNXpJZEupYQFYJvlgOWocTwC4lpukxNTcnXelTMz89jbGwsD74JAIzkgN8G/Vi6SYkxPDyc/aJDe/mjo6OsktLWRlvp9KeOiooKjIiI4KvJvynJywHeBVsAyAWA3wHg9NrFx9LoPT8PAAUA8L4cmCpQ0UCVE5WPVENra6PxUWnvKgdwG/8Ce01taPfIyoYAAAAASUVORK5CYII=",
  34: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAATqSURBVGhD7ZpJSDVHEMfLleAe9wUMRtSDuMTtoIZEsyiKICghNyMRL8ZEzUFFD0aUBBdcggohJ1EioqC4osYdV6IoESQqiGtccAkRjVuFar4exvb5fX4v7328eeQHdXg91TNd/+mprul5AP/zbBwAwB8A3geAjwHgIx00GheNLwAAHMUA1MEcADIA4FcAOAYAVJCdAMAoAHwNAFZiYM/hcwD4g04WGhqKeXl52NTUhF1dXTg4OIgDAwM6ZzSu7u5ubG5uxoKCAgwPD+dibADAF2KAL+N76hgREYF9fX2oZIaHhzEqKooL8SMAGIjBivxAzpmZmXh3dyeeT7Hk5+dzEerFgOV8xoPXRwoLC7kIX4qBExYAsBUWFib20yuio6NJgAMAeFsU4CtSZ2RkROyjV8zPz6ORkRGJkCsKMBESEiL66yUvkuKiPCFS0fAXPSPaYHd3F7e3t187qZ6fn+Pq6iqura3h7e2teFhtysrKSIBrAHiHC/AeTf+WlhbR9z9RXV2NAQEBaGlpiebm5ujr64sVFRWim0qur69Z/UHjsrOzw6OjI9FFbXp7e3kypKqR8SE19Pf3i75qk5aWxi/yyFJTU0X3R+Tm5kr+NjY2GhVgdHSUn/tTLgDV0KyS0gSUSPngXVxcsL6+ns0GupO8vaenR+wmMTExgQYGBpKvpmfA0NAQP/cnXAB6kWDlpCYoKiqSAujs7JTay8vLpaBycnIe9OFcXFygt7e35KdIAXjyokfq6upKaq+srJSCKikpedCHk5GRwY5T3nBwcFCmACL7+/vY0dHBHge6joWFBW5sbIhuTDAuUENDAy9alC0ALX1+fn5SYJ6enri4uCi64enpKbq7uzOfmJgY1hYcHKx8AagGMDY2lgTw8PDA2tpa0Q1TUlLYcZr6m5ubrM3f31/5ApydnbFlZ3JyEhMTEyUhsrOzJZ/29nbWZmJigm1tbVI7nwH29vavXUS9jDcqgBxKiK6uruzihoaGuLe3x7K+s7Mza3NycsK6ujqsqalh5ubmxtqpiKKk2djYKJ5SLbQuwP39PZ6cnEhTWY5scwKnpqbw8PBQ+v0qo1VBE2hdgOTkZHY36c6tr68/OEY5gAdES+XBwcGjQJ8ySqCaQOsCpKenS4OOj4/Hra0tPD4+xqysLKndy8uL1ftks7OzODMzIxn9JiMf8rW2tmZL5PLysngptdC6ADs7O+jo6CgFa2VlxRKZ/G7S5uWrCAwMZL62trZ4eXkpHlYbrQtALCwsYFBQ0IOgyWitb21tFd1V4uPjw/qYmZmxXKEp3ogAxM3NDZu69P5dWlrKljhaFp/L9PQ0608vV/SoaApVAmj0bVDXUSVAFDVocj9Al1G1HxBMDc99NpWO7KXrAy6AMwD8XVxcLPrqJbQ5AwA3APAuF4CYiYyMFH31ktjYWBLgdwAwkgvwLdXnVIDoMysrK2hqakoCfCcPnrABgD+pVtdnEhISKPjTp/4/QN/M2J6ePiLbk/xGDFzOT+SkbwlRFnyLGLAIJYafyZmmy9zcnHguRbG0tIRJSUk8+F8A4C0x4Kegj6X7lBhpj46+6NBe/tjYGKukdNXGx8fZnzqqqqowLi6Ob8cdUpIXA3wOTgBQCAC/AcA/L1RUitE6vwQAxQDgJgamDlQ0UOVE5SPV0LpqND4q7b3EAFTxL8ndXwQ4yTUUAAAAAElFTkSuQmCC",
  35: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUoSURBVGhD7ZpZSDRHEMfLk4AHBo1XUDEiimIkxgs1JprDC0FQo/gUifig0eiXhyj64J3ggWJAUOKbEhGFz1vUeMWHeMQDokg8EA9MghgN3qIVqtkZZnr3+5B1V9YlP6iH7a6e7fpPT3VN7wL8z4N5CwDeBYAPAOATAPhYB43mRfPzBQBbPgB1MAOAbAD4GQCOAACfkR0DwCQA5AKAJR/YQ0gFgD/oYgEBAVhQUIBtbW3Y19eHo6OjODIyonNG8+rv78f29nYsKirCkJAQQYwtAPiCD/B1fEcDQ0NDcWhoCJ8z4+PjGBERIQjxAwAY8MHyfE/OOTk5eHd3x1/v2VJYWCiI0MQHLOVzIXh9pLi4WBDhSz5wwhwAdgMDA/lxekVkZCQJ8BcAvMkL8BWpMzExwY/RK+bn59HIyIhE+JYX4Bd/f3/eXy9RJMUlaUKkouFfeka0wcHBAe7t7elMUq2uriYBbgDARRDgPVr+HR0dvO+jaGhoQF9fX7SwsEAzMzP09vbG2tpa3k2Eki+twqCgIJlRXgoPD8fDw0N+iFoMDg4KyZCqRsZH1DA8PMz7qk1GRobwJUqWnp7Ou+P19TXa29sr+Upte3ubH6YWk5OTwjU/EwSgGppVUpqAEqkwaQcHB2xqamKrwdraWmwfGBiQjdnY2EBjY2PWR6vFxcUFnZycRHN3d8f9/X3ZGHUZGxsT5vGpIAC9SLByUhOUlJSggYEB+5Kenh6xvaamRhTgxYsXsjG9vb1iX1VVFd7c3LBVcXFxgWdnZ3h+fo739/eyMeqidQFOT09xfX2dPVJXV1die11dnRhkRUWFbAx9pnYSbnZ2lrVpK2lqXQAeSl4vX75kjwN9j7m5OW5tbcl8kpOTWZ+JiQnGxcWhp6cnurq6YmxsrMbfR55UALqLPj4+4p13c3PDpaUlJR8vLy/RR5VRHtEUTyoA1QBCciOju9rY2CjzoRpBUZ0x39zcXGxtbcXU1FRxnKGhIa6trcnGqcuTCnBycsK2nZmZGUxISBADys/PF30uLy+ZT0tLC/OTkpKSIo6hMwlN8KQCSKGE6OjoKN7RnZ0d3kWJ7u5uUYD4+Hi+Wy20LgBtV8fHxyoDlBxOKNUdt7e3Spmf6gXBPyYmRtanLloXICkpCe3s7FhBs7m5KeujHCAEtLi4yNq6urowODiYJcjy8nKZv6JuZ5adnS3rUxetC5CZmSlOmra03d1dPDo6wry8PLGdKjuhRpBMCK2srHBqaoqtouXlZfGRIePzg7poXQAqWW1tbcWJW1paoo2NjfiZjA4vpQh1ABkVQx4eHmhqaiq2ZWVlyfwfg9YFIGh5+/n5yYImc3Z2xs7OTt6dlbtpaWlK/rQt0srRVBlMPIkABCU1KofpOa6srGTPOm2Lr2NhYYHVCaWlpdjc3Iyrq6u8y6NRJYBG3wZ1HVUCRFCDJs8DdBlV5wHvU4OqZ1MfoRutEOBDQQB7ADgrKyvjffUSOpwBgFsAeEcQgPg1LCyM99VLoqOjSYDfAcBIKsA3VJ8LhxH6Cu0qivqiVBo8YQUAf1Ktrs/QCxUA/POq/w/Qb2bsTE8fkZxJfs0HLqWFnPQtIUqC7+AD5qHE8CM503KZm5vjr/WsWFlZwcTERCH4nwDgDT7gV0E/lh5SYoyKimK/6NC7Ob2pUSWlqzY9Pc0OUevr69mBquI47m9K8nyAD8EOAIoB4DcAuFao+FyM9vkVACgDgLf5wNSBigaqnKh8pBpaV43mR6W9Ox+AKv4DrwNSXIOzBnUAAAAASUVORK5CYII=",
  36: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUmSURBVGhD7ZptKPdXGMcvD2l5aus2TysyeXPLmOEFZmMzIlEkbxRNxGbD1IgXiLbumzyVF2tFIQ/dL27hJszTxNBIURpKHrclQ8Lm4VrXyTn9fgd3+j/o77996hTXuc7v/7u+zrnOdc4fwP/cm7cB4D0A+BAAPgWATwyw0XvR+3kDgL0cgCZYAcAXAPATAOwBAD6itg8AowDwFQDYyoHdhyQA+I0e5u/vjwUFBdjS0oLd3d04ODiIAwMDBtfovXp6erC1tRWLioowMDCQi7EGAClygK/jOxoYFBSEfX19+JgZHh7G0NBQLkQ9AJjIwcp8T87Z2dl4eXkpP+/RUlhYyEVokANWksiDN0aKi4u5CJ/LgRPWALAREBAgjzMqwsLCSIA/AOAtWYAvSZ2RkRF5jFExOzuLZmZmJMK3sgA/+/n5yf5GyXVSnFcmRCoajmiN6IPt7W3c3Ny8d1Ilv62tLTZGHzx79owE+AcAXLkA79P0b29vl321oqamBr29vdHGxgatrKzQ09MTKysrZTfB1dUVVlVVoZeXF1pbW7MxPj4+2NjYKLtqxatXr3gypKqR8TEZ+vv7ZV+NSUtL4x9yo6WmpsrueH5+jnFxcTd8eautrZWHaMzo6Ch/7mdcAKqhWSWlCyiR8hd3cnLChoYGNhuePHki7L29vaox19OStadPn2JTUxOWlZXxhIUmJia4s7OjGqMpQ0ND/LPCuQB0kGDlpC4oKSlhL0zP7OrqEvbnz5+LIPPy8oT9+PgYnZ2dmd3S0hJXV1dFX2ZmJrq4uLClND09LezaoHcBDg8PcXl5mS2ps7MzYaf1zQUoLy8XdsWUZPu0zMXFhWzSCr0LILO7u4svX75ky4E+hxLc2tqa6K+vrxcCUMk6NzeHWVlZmJyczPqOjo5Uz9OWBxWAtjTK6jxAd3d3nJ+fV/nQyY33U9bnPyvHLC4uqsZow4MKQDWAubm5CMbNzQ3r6upUPrm5uaqAHRwcWI6IiYkRNkqMJycnqnGa8qACHBwcsDU+MTGh2uYoaE5+fr6w096v/GsnJSWJPl3VKQ8qgBJKiDzbm5qa4vr6OrOXlpaKIENCQlRjaLvkfbo6pepdAKro9vf3RYBKFJcTou5obm4WtvDwcJX/+Pi46EtJSVH1aYreBUhISGDrmKazck8nKAfwgCjbE0tLS6JucHR0xNPTU+Hf1tYm/OlaThfoXYD09HTx0tHR0bixsYF7e3uYk5Mj7B4eHqoaQXF/hxkZGWwGraysqHYFXR3V9S4AneTs7e3Fi9va2qKdnZ34nRpdXiqZnJwUZS81mkFUFfLfY2NjVf7aoHcBCJrevr6+qqCpUVnb2dkpuzPoxtnV1fXGmMTERJ0WQw8iAEEnPCqH6aBTUVGBL168YNvi66D+jo4OVipXV1fj1NSU7KI1twmg09OgoXObAKFk0OV9gCFz233AB2S4a20aG/SHvhbgIy6AIwAc0wXEfwG6nAGAcwB4lwtA/BIcHCz7GiWRkZEkwCIAmCkF+Ibqc13duhgqVHVaWFiQAKXK4Ik3AeB3qtWNmesj9l93/f8AfWfG7vSMEcWd5Ndy4Ep+ICdjS4iK4NvlgGUoMfxIzjRdZmZm5Gc9KhYWFjA+Pp4H3wYAb8gB3wV9WbpLiTEiIoJ9o0OXE2NjY6ySMtRG9wf0Tx1URkdFRfHruD8pycsB3gcHACgGgF8B4O9rFR9Lo31+AQDKAOAdOTBNoKKBKicqH6mGNtRG70elvYccwG38C+dIOt2sfKeOAAAAAElFTkSuQmCC",
  37: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAS7SURBVGhD7ZpZSD11FMe/LojgGpqmgWEiPohJ7qhR2qIogqJI4EuSKFKW2kOKPpgo5YZL4EOEiKiJ+KC4L3+3fEglRUgQN8QFM8U0erBEf3EGZ7hzvIZc58q9tz7wQ/yd85uZ870zZ87v3Av8z6N5GcAbAN4C8B6Ad01w0HXR9QUD8OABGIIDgE8AvABwBkCY0TgHMAvgMwDOPLDH8CGATTpYeHi4KCkpEZ2dnWJwcFBMTk6KiYkJkxt0XUNDQ6Krq0uUlZWJ6OhoWYwdAB/xAP+Nr2lhTEyMGB0dFebM9PS0iIuLk4X4FoAVD5bzDTkXFBSIm5sbfjyzpbS0VBahlQesS6YcvCVSXl4ui/AxD5xwBLAfERHB11kU8fHxJMAJgJe4AJ+SOjMzM3yNRbG8vCxsbGxIhC+5AD+GhYVxf4vkLimu6iZEKhr+oGfEGBwdHYmDgwOTSaq1tbUkwN8AXpMFeJNu/56eHu77JJqamkRwcLBwcnISDg4OIjAwUNTX13M3UVNTI0JDQ0VkZOSDg+zd3d18qUGMjIzIyZCqRol3aGJsbIz7GkxOTo58knsjOztb5ZuVlXXPR9+oqKhQrTOU2dlZ+ZgfyAJQDS1VUlpAiVS+aC8vL9Ha2irdDW5ubsr88PCw4k+BBQQE3BtBQUFywpL+UlGjBVNTU/J1vC8LQBsJqZzUAgrIyspKOsnAwIAyX1dXpwhQXFysWqOPjo4ORYDm5mZuNhijC3B5eSk2NjakR+rq6kqZb2hoUASoqqpSreFQ4nR1dZV8k5KSuPlJGF0AzvHxsejv75ceBzqPo6Oj2NnZ4W4qMjMzJV97e3uxtbXFzU/iWQWgVx89y/In7+fnJ1ZXV7mbisXFRcU/Ly+Pm5/MswpAt7Ktra0SkK+vr2hpaeFuKtLT05VPf3Nzk5ufzLMKcHFxIb12FhYWRGpqqiJEUVERd5XY3d0VdnZ2kk9aWho3a8KzCqALJURvb2/p5NbW1mJvb4+7SIWSLBI1NoyB0QW4vb0V5+fnegPUaU7orTtkO1WPJycn3KwJRhcgIyNDeHp6SuXv9va2ykY5QBZgZWVFZTs9PRUuLi6SLSoqSmXTEqMLkJubqwSZnJws9vf3xdnZmSgsLFTm/f39VTUCMTc3p9jz8/NVNi0xugCHh4fCw8NDCcbZ2Vm4u7sr/9Og5iWnra1NsTc2NnKzZhhdAIJu75CQEFXQNHx8fERvby93l6isrFT82tvbuVkznkUA4vr6WiqHaf9dXV0t+vr6pNfiQ1C+IP/x8XGjJUBCnwCa7gZNHX0CxNGElv0AU0ZfPyCUJh56Ni0N+qDvBHhbFuAVAH9SEvovQM0ZANcAXpcFIH6KjY3lvhZJYmIiCfALABtdAb6g+py2opbM+vq6vNn6Sjd4whXAr1SLWzIpKSkU/O8P/X6AvjPTrPtqauj0JD/ngevyHTlZWkLUCb6HB8yhxPA9OdPtsrS0xI9lVqytrSkdJgA/ALDnAT8EfVl6TIkxISFBalRQL592bFRJmeqYn5+XftRBmynqJt+1436jJM8DfAyeAMoB/AzgrzsVzWXQe34NQCWAV3lghkBFA1VOVD5SDW2qg66PSnt/HoA+/gHPqaJgDb0HMAAAAABJRU5ErkJggg==",
  38: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAVfSURBVGhD7Zp9LJ9XFMePEpZ4p0MtuiBeEkHWtZKW2XQzgnoJWSb+WbuSNFu36v6YpogJtrRFmSDL4h8vkyIhXoOVmjajQTSaykKIaox6XSStSZ3l3Lg3v+f2105+L/Lz2z7J/cO553mee773ec5znvMD8D975k0ACACA9wDgIwD40AAHrYvWFwgATnIAmmAJAF8AwK8AsAwAeIDGKgD0A8BXAGAjB7YXPgWAP+hkJ06cwIyMDKypqcHW1lbs6enB7u5ugxu0rra2NqytrcWrV6/iqVOnuBjTAPCZHODr+J4ODA4Oxs7OTjzI3L59G8PCwrgQPwKAiRyszA/kfPHiRXzx4oV8vgPLlStXuAjlcsCqfMKDN0YyMzO5CJ/LgRNWADAXFBQkH2dUnD59mgRYBAB7WYAvSZ2+vj75GKPi/v37aGpqSiJ8Kwvw2/Hjx2V/o2Q3KY6pJkQqGv6iZ0QfPHnyBB8/frznpLq4uIgzMzO4trYmT+mEa9eukQB/A8DbXIB36Pavr6+XfbXi5s2bGBgYiNbW1mhpaYl+fn5448YN2U1QV1eHJ0+eRHt7ezQ3N0cXFxeMi4vDsbEx2VUrOjo6eDKkqpHxARm6urpkX405f/48v8hL4+zZs7I7lpWVveTHBwmoSxH6+/v5uT/mAlANzSopXUCJlC/+yJEjWF5ezu4GR0dHYW9vbxf+KysraGdnx+y087m5uaziTElJEf4xMTGKa2hDb28vP284F4A+JFg5qQtycnLQxMSEXaSlpUXYr1+/LgK6fPmysN+9e1fYk5OThZ3w8PBgdjc3N9ze3lbMaYreBdjY2MDJyUn2SD1//lzYCwsLRaB5eXnC/uDBA2FPTU0VdsLX15fZvb2995xE/w29CyCzsLCAzc3N7HGg61hZWeH09LSYp50NDQ1lc7a2tiwZP3r0CLOystQKpi37KgDtmr+/vwjE09NTbUJ7+vQpxsbGCj8+zMzMWB2vS/ZVAKoBKAgekLu7O5aWlspuuLS0hAkJCS8JQFUb5Ytnz57Jh2jMvgqwvr7OXjuDg4MYHx8vAktPTxc+m5ubGBAQIHY8Ozsbm5qa2OuS+9OxOzs7inNryr4KoAolRFdXV3bxQ4cO4ezsLLNXV1eLQC9duqQ4Jjw8XMzdu3dPMacpeheAdmp1dVUEqIpKc0LUHfT5zW1yA6agoEDMVVZWKuY0Re8CJCUlobOzMyt/p6amFHOUA3hAo6OjzEa7zm0VFRUK/wsXLoi5qqoqxZym6F2AtLQ0sejo6Gicm5vD5eVlRaBeXl6iRmhoaBB2ektQXUB30cDAADo4OIi5iYkJ+VIaoXcB5ufn0cnJSSzcxsYGDx8+LP6mQc1LztbWlmoDEy0sLNDHx4flCW47d+6c4hraoHcBCLq9jx07pgiaxtGjR/HWrVuyO3sN0t0i+1NJTcGrVpTasi8CEFThUTlM39/5+fnY2NjIXouvgzJ9UVER69+VlJTgyMiI7KI16gTQ6degoaNOgDAy6LIfYMio6we8SwZ1z6YxQhu9K8D7XAAXANikRsR/AWrOAMA2AHhwAYjfQ0JCZF+jJDIykgSYAABTVQG+offu0NCQ7G9UPHz4kLXdAOA71eAJOwD4k2p1Y+bMmTMU/Nqr/n+AfjNjPT1jRKUn+bUcuCo/kZOxJUSV4OvlgGUoMfxMznS7DA8Py+c6UIyPj2NiYiIP/hcAeEMO+FXQj6ULlBgjIiLYLzrUy79z5w6rpAx10Bck9RSKi4sxKiqKt+OWKMnLAe4FZwDIBIARANjaVfGgDHrPjwNALgC8JQemCVQ0UOVE5SPV0IY6aH1U2nvJAajjH36/NiFF+im0AAAAAElFTkSuQmCC",
  39: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUjSURBVGhD7ZpZLP1HFMePNU1srbU0UUu8EKSKB5TShRCJhFQTiZQKD60WfagtoZa0sYRo4qEaHhAiHog1qP2hSIkgkVpjr4ilIdFapjkTM/n9xvWP3C3X/feTTOKeOfP7/c73/ubMmbkA/ufZ2AGADwB8AAAfA8BHOtjwufD5fAHAXgxAGcwA4CsA+A0ATgCAvKB2CgDjAPANAFiKgT2HzwHgT7xYQEAAyc3NJS0tLaSnp4cMDw+ToaEhnWv4XL29vaS1tZUUFBSQoKAgJsYGAHwhBvgqfsSBwcHBZGBggLxkRkdHSXh4OBPiZwAwEIMV+QmdMzMzyd3dnXi9F0teXh4ToV4MWMpnLHh9pLCwkInwpRg4Yg4AO4GBgeI4vSIiIgIF+AsA3hIF+BrVGRsbE8foFXNzc8TIyAhF+F4UYMrf31/010sekuKCNCFi0fA3zhFNsL+/T3Z3d5+dVA8ODqj/7e2t2KUWKioqUIB/AeBdJsB7+Pq3t7eLvipRW1tLfH19iYWFBTEzMyNeXl6kqqpKdOM0NTXRmsPS0pL6e3p6vtJfWfr7+1kyxKqR8iEaBgcHRV+lSUtLYzd51FJSUkR3kp+f/8iPtaSkJHJ/fy8OUZrx8XF27U+ZAFhD00pKHWAiZQ/v6OhI6uvr6dtgY2PD7X19fdx/enqa262trem3jq+plZUVtzc0NMjuoQojIyPsup8wAXAjQctJdVBcXEwMDAzoTbq7u7m9srKSB5STk8PtGRkZ3I7lNqOtrY3bfXx81JYTNC7AxcUFWV1dpVPq+vqa26urq3lAZWVl3B4WFkZtuDxtbW1x+9XVFbG3t6d9hoaGZG1tjfepgsYFEDk8PCRdXV10OuB9zM3NycbGBu8PDQ3lAmxvb3M7fuNubm5cNHXtS7QqAC593t7ePAh3d3eysLAg80lOTub9mKEZuBTiasD6cJenDrQqANYAxsbGPAhXV1dSV1cn88Gpwvpx6ZuamiLLy8skKiqK27E1NjbKximLVgU4Pz+nyw5m+ri4OB5Mdna2zC81NVUWLGumpqb87+bmZtkYZdGqAFIwITo5OdGbY1KTznektLSUznkTExOa/EpKSmSiSVcUVdC4AFi0nJ6ePgoQkRxOPFl3bG5u0jcHkZzqkKWlJdFVKTQuQEJCAnFwcKAJbH19XdaHOYAFND8/T20437HQwb2INEEeHR3xJIjjpEuqKmhcgPT0dB5kTEwM2dnZIScnJyQrK4vbPTw8eEBY/DB7SEgIOT4+pv6JiYncjmd86kLjAuzt7fECBhtubmxtbflnbHh4ybi8vJS9GehrZ2cnE+vs7Ex2D1XQuAAIvt5+fn6yoLE5OzuTjo4O0Z3Mzs4SFxeXR/5YJEmLJnWgFQGQm5sbusbjxqa8vJx0dnby5KYI/Jax2CkqKqJjnkqSqqJIALXuBnUdRQKEo0Gd5wG6jKLzgPfRoGhu6iOS0juMCfA2AFxi5fU6gIczAHADAG5MAOR3XINfBx42WcsAYCQV4Dusz2dmZkR/vWJlZYVtsH6QBo+8CQBHWKvrM7GxsRj82VP/P4C/mdEzPX1Ecib5rRi4lF/QSd8SoiT4djFgEUwMv6Izvi5Yor5kFhcXSXx8PAu+DQDeEAN+Cvyx9BATY2RkJD2rx7P8iYkJWknpapucnKSHpzU1NSQ6Opodxx1jkhcDfA4OAFAIAH8AwD8PKr6Uhuv8IgCUAMA7YmDKgEUDVk5YPmINrasNnw9Lew8xAEX8B+sJOnHsO92RAAAAAElFTkSuQmCC",
  40: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAATxSURBVGhD7ZpZSHVVFMeXA44pJQ5JaiD0JsrnPKWpOSD4IOpHj0EPPnxZSoIjavpQqKCY+KCJCEoiguKMs4bgSAg5ICYaaBZiDoFm6oq1uftwzuZek+M9ftf79YP9stee1t991l57XwH+58G4AYA/AHwIAB8DQIIJFloXrS8AANxFB9TgCACvAGASAI4BAJ9ROQGAGQD4AgCcRccewicAsE2DhYSEYGFhIXZ0dODAwACOj4/j2NiYyRVa1+DgIHZ2dmJJSQlGRkZyMX4BgE9FB+/jG+oYFRWFIyMj+JyZmprCuLg4LsR3AGAhOivyLTXOycnB29tbcbxnS1FRERehSXRYzkvuvDlSWlrKRfhMdJx4CwB+DQ0NFfuZFfHx8STA7wDwjijA56TO9PS02MesWF5eRisrKxKhQBTgx+DgYLG9WaILij/JAyIlDef0jWjJ4eEhbm1t4f7+vmhScHx8jHt7e3hxcSGajEJ1dTUJcA0A73MBXtD27+rqEtsajd3dXXR3d2dBKDo6WjQzSJz09HR0dXVFOzs79PLyYgH57OxMbPoohoeHeTCkrJHxEVWMjo6KbY1GQkICnxQjIiJEMxPIw8NDaiMvsbGxeH19LXZRzczMDB87iQtAOTTLpLSgvr5e4ZC+HZCVlSXZMzMzsbm5Gf39/aW6xsZGsYtqJiYm+LiJXAC6SLB00thsbGyw7XyfAAcHB1Ibb29vvLm5YfUrKys8YmNQUBDe3d0p+qnlyQSgTDI8PJxN5ubmhk5OTnoFoPsFF4d2Aocc9vX1ZfUkEAllDJ5MgKqqKjaRjY0NC7Cenp56BaitrZUEyMvLU9joPsJt8/PzCptankSA1dVVtLa2ZhOREFdXV+jg4KBXAFmejmVlZQpbYmKiZKObnjHQXACK2AEBAWwSCmT0KRwdHRkUIDc3V3KyvLxcYZML0NPTo7CpRXMBCgoK2AQuLi7seCPOz8/R0dGR1VMmJic/P9+gAElJSZKtr69PYVOLpgJQrm1hYcEmCAwMxJaWFmxoaMCKigoWC6ieAhsdjfytgZzmThYXFyvGoxyA24x1TGsqAAU7vuD/KqmpqaxPa2urVJedna0Yz8/Pj9VbWlri5uamwqYWTQWg71R01FDhR97S0pJUFxYWJo11cnKCzs7OrN7HxwcvLy9lM6lHUwFo0QsLC4pCDvb390vJDgXIxcVF3NnZYX0oaPLznv7S3d3deHp6KsUSKuLOeAyaCmAIEsbe3p5NHBMTI5qxra1NsTt4zkDF1tYWt7e3xS6qeS0C0BWYO2To3YHyBXJWLgRdkHp7e8Wmj+K1CEDf7+TkJLtx0idhCAp0TU1NWFlZie3t7Sx/MDb6BND0Nmhq6BMgjiq0fA8wJfS9BwRRBUXfNwH6Q+sEiOUCvAsAf9F39yage6D5BwB8uQDEgnhJMVdSUlJIgJ8BwEouwFeUhFCCYs6sr6/zO8nXcueJtwHgSLypmRtpaWnk/J+G/n+AfjNjNzdzpKamhge/L0XH5TRTI3MLiDLnu0SHRSgwfE+Nabvcl7U9B9bW1jAjI4M7/wMA2IkOG4J+LP2NAmNycjJ7uBwaGsLZ2VmWSZlqmZubYw8tdXV17L1B9y75BwV50cGH4AEApQCwCgB/61R8LoXO+TUAqASA90TH1EBJA2VOlD5SDm2qhdZHqf0HogP6+Bfr5TLx8AzBLgAAAABJRU5ErkJggg==",
  41: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAQ1SURBVGhD7ZpJSGRXFIZ/B8Sh1cQpLVmIQkBcKHHCKXbUJA4gCGLoZUOWadOLXqioYHSRoIJDoBdtL5WI6MYZ52HjgAShFQkq2AuHEGwQwaHVE86lrlQdqyzproJ6j3xwN++e++77/3rvvHPvK+B/HkwkgEQA3wD4DkCBBza+Lr6+JABRUsDHEATgZwDTAP4FQAZqxwDmAPwCIEQKewhPAfzNJ0tLS6Pq6mrq7u6moaEhmpycpImJCY9rfF3Dw8PU09NDtbW1lJWVpc3YAfBMCryP33hgdnY2jY2NkZGZmZmhvLw8bcQfALykWMnvHFxZWUnX19fyfIalpqZGm/BKCrbmRy3ejNTV1WkTfpLCmUcA3qWnp8txpiI/P58NOALwuTTgObszOzsrx5iK1dVV8vHxYROqpAGLqampMt6UWJLiX9YJkYuGE35G3Mn+/j5tbW3R3t6e7LrDycmJit3Z2aGbmxvZ/Uk0NzezAZcAYrQBX/Pt39vbK2Ndxu7uLkVFRakklJOTI7vvUFxcrGLj4uLo8vJSdn8So6OjOhly1aj4lg+Mj4/LWJdRUFCgJ6XMzEzZbUNHR8dtbExMjMsNmJub0+f/QRvANbSqpNxBe3v7raD77oCrqyuqr6+3iXXHHTA1NaXP/702gBcSqpx0NZubm+Tv7+/UgIGBAUpJSbGJM7wBXElmZGSoySIjIyk4ONihAQkJCbeiy8rKKCIiwvgGNDU1qYn8/PxUgo2OjrZrAGf5+Ph4JZbjTk9PKTw83NgGrK2tka+vr5qIjTg/P6fAwEC7BvCzPz8/r159zNHREYWEhBjXAL7gpKQkNUliYqJ6FA4PDx0aIDk4OKDQ0FDjGlBVVaUmCAsLU+9/hn/doKAgdZwrsfswtAFca3t5eakJkpOTqaurizo7O6mhoUHlAi2KX42O9hoMbQAnMcvJnbaSkhI5XGFoA/r7++8IddQqKirkcIWhDTg+PqalpSWbtrKyQoODg7cFESfI5eVl2t7elsMVhjbAEWxMQECAmjg3N1d228AG6KIpNjbWHAbwEljf+s72HXjZrBMmF0+mMODs7Iymp6fVipMfifu4uLhQu7kcu7i46PL9AHsGuHU16GnYMyCPD7hzP8CTsLcfkMIH+vr6ZKwp4R/aYsATbcBjAKeNjY0y1pRYNmg+AIjTBjBLzhYpZqGoqIgNeAvAx9qAl97e3qpAMTMbGxv6FfurtXjmMwCHzlZqRqe0tJTFv3f0/wH+ZqZWbmakpaVFJ78XUrg1rznIbAnRSnyvFCzhxPCGg/l2cVa1eTrr6+tUXl6uxf8JwF8KdgR/LD3gxFhYWEitra00MjKi9u64kvLUtrCwoDZa2tra1H6DZV/yH07yUuBD+AJAHYA1ABcWF43S+D2/DqARwJdS2MfARQNXTlw+cg3tqY2vj0v7r6QAe/wHon6kNzCJPVgAAAAASUVORK5CYII=",
  42: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAASsSURBVGhD7ZpLLHVXFMf/HvGsagmtdNBEwhBVxFtRJRIjUWUgTTosbaIDhMRrUG+iiUE1RihiIN7v1zfxiDSSEmmU6EBpRBNpQjWsZu3YN+fu7171Xfdw7/36S9bkrL3PPut/z1l77QXwPw8mAEAYgCQAHwNIt0Hj5+LnCwcQqAZgCd4AvgSwAOAMANmRnQNYBvAVgDfVwB7CZwB+4ZtFR0dTeXk59fb20tjYGM3NzdHs7KzNGT/X+Pg49fX1UWVlJcXHx0sxfgXwuRrgfXzLExMSEmhqaorsmcXFRUpNTZVCfAfASQ1WpYEHl5SU0M3NjXo/u6WiokKK0KUGrOVTGbwjUlVVJUX4Qg2ceQPAbzExMeo8hyItLY0FOAXwtipAMauztLSkznEoNjc3ycXFhUUoUwV4ERUVpY53SO6S4k/ahMhFwwV/I3pyfHxMe3t7dHR0pLqMOD09pcPDQzo7O1NdVqGpqYkFuAbwvhTgA379BwYG1LFW4+DggAIDA0USSkxMVN2CoaEhSklJIT8/P3J3d6eAgADKyMig6elpdeijmJyclMmQq0bBR3zB2gtpSU9Pl4tSXFyc6qbGxkaD35T19/erUyxmeXlZ3vcTKQDX0KKS0oOOjg6jYNQ3YH9/n5ydnYXPycmJiouLqaenh/Lz8w1z/P396fz83GiepczPz8v7ZkgB+CAhyklrs7u7Sx4eHvcKUF9fb/AVFhYa+WJjYw2+mZkZI5+lPJkAXEnKAPh79vHxMSkAr1taWkoFBQXi9dRSVFRkEGBkZMTIZylPJoD8Zd3c3ESCDQoKMimAOVjA0NBQgwD8NlmDJxFga2uLXF1dxUIsxNXVFXl5eb2SAHyik8EnJSXR7e2tOsQidBfg+vqawsPDxSJhYWHilzw5OXklARoaGgzBc4Lc2NhQh1iM7gKUlZWJBXg/5/2fubi4IG9vb3GdK7H70CZFtq6uLnXIo9BVAK61eSvj+0VGRlJ3dzd1dnZSTU2NyAV8PTg4WGyNpnoNmhObsNbWVnXIo9FVAE522gDus+zsbKO5dXV1Rn4WTw90FWB4ePilQM1ZXl6eYR632+R13i65XNULXQXgam1tbc3IOIGNjo4aCiJOkOvr66ICZPjw4+vraxAgJCSE2traqLa2Vnw6bNXV1fa1DaqwMJ6enmLh5ORkI5/66pszbnZag2cRgI/AMhBt34H39oiIiJeCNWWDg4NG97SUZxHg8vKSFhYWxIlTu6ezACsrK6LOl+1tU8Z+/lSsgSkBdD0N2hqmBEjlC3r2A2wJU/2AD/kCd2ReB/iHvhMgRQrwLoC/OBu/Dtw1aP4BECwFYNYeckhxBLKysliAnwG4aAX4hk9dXKA4Mjs7O/JMUqsNnnkLwMl/ndTsnZycHA7+T3P/P8B/MxOlpyPS3Nwsk9/XauBavudBjpYQNcEPqAGrcGL4gQfz62LNTsxzsL29Tbm5uTL4HwF4qAGbg/9Y+jsnxszMTGppaaGJiQlRtnIlZau2uroqGi3t7e2i33DXl/yDk7wa4EN4B0AVgC0Af9+paC/G+/w2gDoA76mBWQIXDVw5cfnINbStGj8fl/YhagCm+BetpUhX1V0ZegAAAABJRU5ErkJggg==",
  43: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAThSURBVGhD7ZpbKLVZGMf/DslhCMXQKKWmXIjGIGeDmSHFjUxzo6bm8hvjYsohikFmQpGpUYy4IJIr55xP5ZSGGtJkHObCYUxOkcP4PNOzsnZ7rzajffBtu/nVunnXs973ff57vc/6rwXwP8/GC0AwgDgAnwJItsDG78XvFwLAW03AEFwAvAEwBuBvAPSK2jGASQDfAnBTE3sOXwL4nW8WHh5OBQUF1NbWRr29vTQyMkLDw8MW1/i9+vr6qL29nYqKiig6OlqK8QeAr9QEn+IHHhgTE0ODg4P0mhkfH6fExEQpxE8AbNRkVX7k4JycHHr79q16v1dLYWGhFOFnNWFtvpDJWyPFxcVShK/VxJn3APwZERGhjrMqkpKSWIBDAB6qAN+wOhMTE+oYq2JpaYns7OxYhHxVgJmwsDA13ip5KIq/ahdENg3n/I2Yk729PdrY2KDd3V21S4fDw0MRc3l5qXaZhKqqKhbgFoC/FOAjnv6dnZ1qrMnY2toib29vUYRiY2PVbkF3dzfFxcWRh4cHOTk5kb+/P+Xm5tLp6akaahQDAwOyGLJrFHzCF4aGhtRYk5GcnCwfSlFRUWo3NTQ0aPrVxoX57OxMHWIwk5OT8t6fSwHYQwsnZQ7q6up0ElJnwP7+vvjFuc/R0ZHKysqoubmZAgMDNWNKSkp0xhjD6OiovO9nUgDeSAg7aWrW19dFUk8JwPba3t5e9LHllkxPT2vG8KdhKl5MAHaSkZGR4mFeXl7k6uqqV4Crqyva2dkRL8YFUDI7O6sRID09XWeMMbyYAOXl5eJBDg4OosD6+vrqFUDl5OSEZmZmxCZMCsCbHFPxIgIsLy9rpjULcX19Tc7Ozs8SIDs7W5O4u7s7dXR0qCFGYXYBbm9vKSQkRDwkODhYfAoHBwfPEuD+/p6CgoI0Ari5uVFeXp5JVwGzC5Cfny8e4OnpKdZ/5vz8nFxcXMR1dmKPcXd3J6b/wsKCSFwKkZCQQDc3N2q4QZhVAPbaNjY24gGhoaHU1NRE9fX1VFpaKmoBXw8ICBBL43POGrT9Q09Pj9ptEGYVgIudfOH/amlpaWIMT/uLiwsxW/jz0UbOJm6VlZU6fYZiVgHYzqqJPtaysrLEGDY5fn5+YoZ0dXXp3C8jI0MT39jYqNNnKGYV4Pj4mObn53Xa4uKimL7SEHGB5G98c3NTjGltbdUkye5vZWVF+H/+fORKwltYGW8sZhXgMVgYaXfj4+N1+njas9OTInDSPj4+OrOF7bGpeCcC8PZWJqPv3OHo6EhnusvG7rGiokINN4p3IgDb3bGxMbHj5E/iMebm5sQKwb94S0sLbW9vqyFGo08As+4GLQ19AiTyBXOeB1gS+s4DPuYL6hJkrfAP/SBAghTAB8CFKSutJfNwQPMPgAApADP/1CbFmkhNTWUBfgNgpy3Ad7a2tsKgWDNra2tyT/K9dvKMO4CDp3Zq1gCfLAE4eez/B/hvZmLnZo1UV1fL4perJq5NIwdZW0HUSr5TTViFC8MvHMzT5SnX9hpYXV2lzMxMmXwHAEc14cfgP5buc2FMSUmhmpoa6u/vp6mpKeGkLLXxUToftNTW1orzhofd5F9c5NUEn8P7AIoBLAO4eVDxtTRe51cBlAH4QE3MENg0sHNi+8ge2lIbvx9b+w/VBPTxL1kZXfCN0IgZAAAAAElFTkSuQmCC",
  44: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAASXSURBVGhD7ZpLSPxVFMe/PvGRaT4TwUBRd2q+8JWlVorgSox2Bi3NWrRQ0YWpUKjgI0zIVqIkohvf+H4sfCGhpEiSaAszTUVJNB+cOBfvMF5n/A8/ZvjP/OgDZ3Pv+f3uPd+5c+65dwb4H4sJAhAL4D0AHwLItUPjefH84gAEqwFowRtAKYBpAH8DIAeyMwBzAL4E8KYamCV8CuA3fllycjJVVFRQd3c3DQ0N0eTkJE1MTNid8byGh4epp6eHqqqqKD09XYrxO4DP1ABf4lt+MCMjg8bGxsiRmZmZoezsbCnE9wCc1GBVvmPnsrIyenh4UN/nsFRWVkoRflADNuYTGbweqa6uliJ8rgbOvAHgj5SUFPU5XZGTk8MC/AXgLVWAL1id2dlZ9Rldsba2Ri4uLixCuSrAYlJSkuqvSx6T4i/GCZGLhkv+jtiSw8ND2tnZoYODA7XLJBcXF8J/d3eX7u/v1W7NNDQ0sAC3AN6RArzLy7+3t1f1tRp7e3sUHBwsklBmZqba/Yzb21tRf7B/QEAAnZycqC6aGR0dlcmQq0bBB9wwPj6u+lqN3NxcOSilpaWp3c8oLy83+Pv5+VlVgLm5Ofnuj6UAXEOLSsoWtLS0GIKxZAUsLi6Sk5OTwd/aK2Bqakq++yMpAB8kRDlpbba3t8nDw8NiAa6urig6OvqJv8MKwJVkamqqGCwoKIh8fHxeKUBpaanwYV9+xqEFqKurEwO5u7uLBBsaGvqiAJyD5Kfe0dEhixbHFGB9fZ1cXV3FQCzEzc0NeXl5mRXg/PycwsPDRX9eXp5oS0xMdEwBeAuLi4sTg8TGxoqvwtHR0YsClJSUiD5e+vv7+6KNn3VIAeQW5u/vL/Z/5vLykry9vUU7V2LGDAwMiHY3Nzfq7+83tMsVEBgYaNWTqU0F4FpbbmEJCQnU2dlJbW1tVFNTI3IBt0dERIitkSfCWV/mhpCQEGpvb6fW1lZhYWFhop2Fq6+vp66uLnU4TdhUAE52jy9/pRUXF9Pp6emzdnPGu4I1sKkAvITViZsz/t6fnZ09KXpessjISHU4TdhUAA5oeXn5ia2urtLg4KChIOIEubKyIvLD3d2d6Df25z62qKgo4e/r6yu2yM3NTXU4TdhUAHOwMJ6enmLgrKwstdsk8fHxwp+T6fX1tdqtmdciAB+B5VK29N4hJiZG+PP2eXx8rHZr5rUIwJ/g9PS0WMq85C1haWlJ+PMtFdcW1sKUADY9DdobpgTI5gZb3gfYE6buAxK5oa+vT/XVJUaHrvelAG8D+Ke2tlb11SWPFzR3ACKkAMyyqUOKHsnPz2cBfgXgYizA187OzqIA0TNbW1vyTPKNcfCMH4Aj9aSmNwoLCzn4c3P/H+DfzMTJTY80NjbK5PeVGrgxP7KT3hKiUfC9asAqnBh+YmdeLpZWbfbKxsYGFRUVyeB/BuChBmwO/rH0T06MfEfX1NREIyMjND8/Lyope7WFhQXxp47m5mYqKCiQ95LHnOTVAC0hBEA1gHUA/z6q6CjG+/wGgFoAYWpgWuCigSsnLh+5hrZX4/lxaR+lBmCK/wDuKk+tDbgQXwAAAABJRU5ErkJggg==",
  45: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAATsSURBVGhD7ZpbLKVXFMf/xiUu1aiEohFBhHggVffRadF2zIgnUeKpSd9GqaQPCA9u0QYJmSYS1BupCC8MI2aMSz24piQlUjqJErRpNBHGLaxm7Zz95Zw958zImfMZTvpL1su39/72Xn/7W3utfQD/c2l8AEQB+BjAZwDSr6Hxunh90QB8VQeswQNAAYBRAP8AoBtkewDGARQBeFd17DLkAfidXxYXF0elpaXU2dlJAwMD9OTJExoZGbl2xut69OgRdXV1UXl5OSUnJ0sx/gDwlergq/ieB96+fZseP35MN5lnz55RamqqFOJHAA6qsyo/cOfCwkI6Pz9X33djKSsrkyK0qA4b86V03h6pqKiQInytOs68A+DP+Ph4dZxdkZaWxgL8BeA9VYBvWJ2xsTF1jF0xNzdHjo6OLEKJKsAvsbGxan+7xBAUfzUOiJw07PM3oifb29u0urpKGxsbatOVUl9fzwKcAgiSAnzI27+7u1vtazOeP39Ovr6+IgilpKSozQIOvrwLExISTIzj0p07d2hnZ0cdYhVDQ0MyGHLWKPiUHwwPD6t9bUZ6erqclJKSktRmOjk5IT8/P62POWMRbcH4+Lh85xdSAM6hRSalB83NzSaOmNsBa2tr5OTkJNo9PDwoKCiIAgMDNQsLC6OtrS11mFU8ffpUruVzKQAXEiKdtDUrKyvk6ur6WgH6+/u19rq6Ojo9PRW74sWLF3RwcECHh4d0cXGhDrOKKxOAM8nExEQxmY+PD3l6eloUoLa2VrQ5ODjQzMyMNl4PrkyAmpoaMZGLi4sIsP7+/hYFyMnJEW3Ozs6UmZlJERERFBwcTPfv37d5PXIlAiwsLGjfNAtxfHxM7u7uZgXgv3RkZKT2CZizlpYWkzFvgu4C8PcbHR0tJomKihIO7u7uWhRgc3NTZmdCtKKiIuro6KC8vDxNgFu3bol4Ygt0F6CkpERM4O3trR1d+/v7Irrzc87EjDk6OhJHU1tbG01NTZm05ebmaiLwnYQt0FUAzrU5kPH7YmJiqL29nR4+fEiVlZUiFvDzkJAQcTRe5tvu6+vTBMjKylKbrUJXATjYyQW/zjjAGXN2dvZS5B8cHNT637t3z6TNWnQVoLe39yVHLRlHfjmGj8vQ0FARMI0x5O3CCgoKTNqsRVcB9vb2aHp62sRmZ2dFoiMTIg6QfNavr6+LMaOjo5qTXl5eNDExIZKexcVFCggI0NrU+GAtugpgCRbGzc1NTMyFjYrMA9g4hoSHh2sxg+3BgwfqEKt5KwJwCSydMXfvwOlufn6+1kcaH4vFxcU2S4OZtyIAH3W81bni5E/CEvPz8+LUqKqqotbWVlpeXla7vDHmBNC1GrxumBMglR/oeR9wnTB3H/ARP+jp6VH72iX8hzYI8IkUwA/AQXV1tdrXLjFc0JwBCJECMNNqkWKvZGRksAC/AXA0FuA7rrjkZYS9wqeKIb+oMnae8QKwq1Zq9gYXVAD+tfT/A/ybmajc7JGGhgYZ/L5VHTemjTvZW0A0cr5bdViFA8NP3Jm3y6uytpvA0tISZWdnS+d/BuCqOmwJ/rF0hwPj3bt3qbGxUdTmXKlxJnVdbXJyUly0NDU1ifsGw73k3xzkVQcvw/sAKgAsADgxqHhTjM/5JQDVAD5QHbMGTho4c+L0kXPo62q8Pk7tw1QHzPEfvnBC57vMo0kAAAAASUVORK5CYII=",
  46: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAT6SURBVGhD7Zp7KL5nGMcvh+S4NuW4omj/KIcZcpoNM1KiRP5Ran+I388WU4g/DLXlkFPtj1lRyCF/OefMyjGWQhpTRmZrmSTM6VrX3Xs/Pc+99zW9ntfv9W6fukr38bm+z/1c93XfL4D/eTROAOAHAB8CwCcAEGuERs9Fz+cPAM6iA/pgBwCvAGASAP4AAHxBdgIAMwDwOQC8JTr2GDIA4CcaLDg4GIuLi7GjowMHBgZwfHwcx8bGjM7ouQYHB7GzsxNLS0sxPDyci/EzAGSJDj7E19QxIiICR0ZG8CUzNTWF0dHRXIhmADATnRX5hhrn5eXh3d2dON6LpaSkhIvwreiwnHTuvClSVlbGRfhMdJywB4BfQkJCxH4mRUxMDAnwGwC8IwrwmtSZnp4W+5gUKysraGFhQSIUiQL8EBQUJLY3STRB8Ud5QKSk4Yy+EUNydHSE29vbuL+/L1YpoOB7eHiIBwcHYpUqVFdXkwDXAODJBXifln93d7fYVjX29vbQ2dmZBaHIyEixmnF/f491dXXo6+uL9vb2aGdnhwEBAdja2io2fRLDw8M8GFLWyPiYCkZHR8W2qhEbG8snxbCwMLEab25uMCUlRWojWmNjo9hFb2ZmZvi4n3IBKIdmmZQhaGhoUDijbQVoliUzHx8fbGtrw4qKCh6w0MzMjH1CajAxMcHniuMC0EGCpZNqs7W1hdbW1g8KcH5+ju7u7qzO1tYWd3d3pbqcnBz08PBAf39/XFpaUvTTl2cTgIJZaGgom8zJyQkdHBy0CiBbkmyfFrm9vRWLnsSzCVBZWckmsrKyYgHWzc1NqwDNzc2SAJSyrq2tYW5uLmZmZrK6s7MzRfun8iwCrK6uoqWlJZuIhLi6umLLW5sAdHLjAlDU539z8/b2xo2NDUWfp2BwAa6vr9k3S2P6+fmxT+H4+FinAPn5+QqHXVxcsKCgAJOSkqQyCowXFxeKfvpicAGKiorYBI6Ojmz/J2gZ075O5ZSJySksLJQcpTbyt52RkSHVqZWnGFQAyrVpy6LxAgMDsaWlBZuamrC8vJzFAir38vJiWyO/a6DtjjsZFRWlGG9oaEiqU+uUalAB6C3xB/43S0xMZH3a29ulsri4OMV4c3NzUl1WVpaiTl8MKkBfX98/HNVlaWlprM/m5qa0alxdXfHy8lIar6urS2pP13JqYFABTk5OcHFxUWHLy8vY398vJUQ8qZEnPLL7O8zOzmbj7OzsKHYFtY7qBhVAF+SQjY0Nm1j8zon5+Xkp7SWjnYDvGmTJycliF715IwLQEZg7o+vegW6cPT09pXbc0tPTVU2G3ogA9F1PTk6yEyd9Ero4PT3Fnp4erKqqwvr6elxYWBCbPBltAhj0NGhsaBMgmgoMeR9gTGi7D/iACnp7e8W2Jgm9aI0AH3EBXAHgnDKy/wKaC5obAPDiAhCL4iHFVElISCABNgDAQi7Al+bm5qrduhgrlHVqziRfyZ0n3gaAY/GkZmpojth/6vr/AfrNjJ3cTJGamhoe/L4QHZfzHTUytYAoc75bdFiEAsP31JiWy0NZ20tgfX0dU1NTufNdAGAtOqwL+rH0VwqM8fHxWFtbyy4nZmdnWSZlrEb3B3TRQmk03Tdo7iV/pyAvOvgYXACgDABWAeAvjYovxWifXweACgB4V3RMHyhpoMyJ0kfKoY3V6PkotX9PdEAbfwP/vitotZ7+yAAAAABJRU5ErkJggg==",
  47: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAARvSURBVGhD7ZpLSHVVFMf/PhAfmSZoSYNACEcq+X5lqZUiOBBRAidBA5GyBg1UdGA6KB/gI2iQISJqIo58v19NfCAhpIiaYKBphoEEmiE71ubuyz3ru8fvcj3H795TP1gIZ+1991n/e87aa68r8D8uEwkgHsDbAN4DkO+BRvdF95cAIIoH4A4hAD4BsAjgDwDCi+wKwAqAzwC8zANzhQ8BHNCHpaSkiNraWjEwMCDGx8fF/Py8mJub8zij+5qYmBCDg4Oivr5eZGZmKjF+AfARD/AhvqKJWVlZYnp6WngzS0tLIjc3VwnxDQAfHiznaxpcXV0t7u/v+ed5LXV1dUqEb3nAjpSr4K1IQ0ODEuFjHjjxEoBfU1NT+TxLkZeXRwJcAHiFC/ApqbO8vMznWIqtrS3h5+dHItRwAX5MTk7m4y2JLSn+5JgQqWi4pnfETM7OzsT+/r44OTnhrieltbWVBLgD8IYS4C16/IeHh/lYwzg+PhZRUVEyCWVnZ2t8LS0tIikpSaSlpeka+YeGhjTz3GVqakolQ6oaJe/ShZmZGT7WMPLz89WiIiMjQ+OrqKiw+x6yxsZGzTx3WVlZUZ/5gRKAamhZSZlBZ2enJhD+BFBgsbGxz1hcXJxKWPIvFTVGsLCwoO7lfSUAHSRkOWk0e3t7IjAw8EEB9Ojv77cL0NXVxd1u82QCUCWZnp4uF4uMjBShoaEuC3B6eirCw8Pl+KKiIu5+FE8mQHNzs1woICBAJtjo6GiXBSgvL5dj6ek5PDzk7kfxJAJsb28Lf39/uRAJcXt7K4KDg10SYGNjQ92gqKys5O5HY7oAd3d3IiEhQS4SHx8vX4Xz83OXBSgtLbV/+wcHB9z9aEwXoKamRi4QEREh93/i+vpahISEyOtUielB4+mVoXElJSXcbQimCkC1to+Pj1wgMTFR9PT0iO7ubrnVqcBiYmLk1uis19De3m5//KmxYQamCkDJTgXwPHOW3VXzgnaMi4sL7jYEUwUYHR19JlA9Kysr08y9vLwUYWFh0kfbp1mYKsDV1ZVYX1/X2ObmphgbG7MXRJQgKdMfHR1p5q6urtrFqaqq0viMxFQB9CBhgoKC5MI5OTncLent7bUL0NHRwd2G8UIEoCOwCk6v79DU1GQf09fXx92G8UIEuLm5EYuLi/LESa+EM+iVIP/s7KxpCZBwJoCpp0FPw5kAuXTBzH6AJ+GsH5BEF0ZGRvhYS0JftE2Ad5QArwH4i5LQfwFbg+YfADFKAGL9eYcUq1BYWEgC/AzAz1GAL3x9fWWBYmV2d3fVmeRLx+CJcADnD53UrEBxcTEF/6fe/w/Qb2aGdV89jba2NpX8PueBO/IdDbJaQnQIfpgHzKHE8D0NpsdFr2rzFnZ2duwdJgA/AAjkAetBP5b+RomxoKBANiomJyfliY0qKU+1tbU12WihwxT1G2x9yd8pyfMAXeFVAA0AtgH8bVPRW4z2+R0ATQBe54G5AxUNVDlR+Ug1tKca3R+V9m/yAJzxL611kugAwXqiAAAAAElFTkSuQmCC",
  48: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUVSURBVGhD7ZpZLK1XFMeXIcaqIaGkkSakeCJVxFQtqsQscpuKl9J4kFYr+oAYoog2iKlSUo0nVAwvrinuNTbEEBFiSqNEH9RQJCIxhdWsHfvL+fY9VI5zXE77S/bLXnv41v/7ztprLwD+585YA4ArAHwAAB8DQPAjbPRc9HxuAGAjOqAKpgDwJQAMAsDfAIBPqB0AwAgAfA0Ab4qO3YXPAOB3WszT0xOzsrKwqakJnz9/ji9evMCBgYFH1+i5uru7sbm5GXNyctDX15eL8QcAfC46eBvf00Q/Pz/s6+vDp8zQ0BAGBgZyIX4EAB3RWZEfaHBaWhpeXl6K6z1ZsrOzuQg/iQ4r8il3XhvJzc3lInwhOk68AQB/enl5ifO0iqCgIBJgBwAsRQG+InWGh4fFOVrFzMwM6unpkQiZogC/eXh4iOO1kuugOKcYEClpOKLfiCbZ2trC1dVV3NzcFE0ydnZ2cGNjAw8PD0WTWigtLSUBzgHgHS7Ae/T5t7a2imPVxvr6OtrY2LAg5O/vL5oZLS0t6OPjg5aWlmhgYIC2trYYExODc3Nz4tB70dvby4MhZY2Mj6ijv79fHKs2goOD+abMSZHa2lrJLjYzMzO1ijAyMsLX/oQLQDk0y6Q0QVVVlcwh8QvY399HCwsLZqM3X1hYyDLOxMREaU5kZKRszn14+fIlXzeEC0AXCZZOqpvl5WU0MjK6VYDx8XHJlpCQILM5ODiwfnt7e7y4uJDZVOXBBKBM0tvbm21mbW3NPmVlAiwsLEgCpKSkyGwuLi6s38nJSW2Z6YMJUFRUxDaiz5oCrJ2dnVIB6M0GBAQwm7m5ORu7srKCeXl5kjDFxcWyOffhQQSYnZ1FfX19thEJcXp6iiYmJkoFIPb29jA6OlpymDdag/J4daJxAc7Pz9HNzY1t4urqyj7d7e3tWwXY3d3FuLi4VwSgrC0jIwNPTk7EKSqjcQEyMzPZBlZWVuz8J46OjtDU1JT1UyamyPHxMROKv/H8/Hzs7OzEpKQkSYjY2Fi8urqSzVMVjQpAubaOjg7bwN3dHRsaGrCmpgYLCgpYLKB+iux0NPJaAxVauKPp6emy9UJCQiTbxMSEzKYqGhWAAhh/4H9r4eHhbA5dv3mfWIApKSmRbPX19TKbqmhUgI6Ojlccvak9e/aMzaG3zvvq6upk66Wmpkq2xsZGmU1VNCrAwcEBTk5Oytr09DR2dXVJCREFyKmpKVxbW2Nz2tvbJScdHR1ZXkC/97GxMRZHuG1xcVHcTiU0KsBNkDDGxsZsYzrzFTk7O1MsYKKhoSE6Ozujrq6u1JecnCybcx9eiwB0BebOKKs70DEYEREhjeGNAio5T3mEungtAtA5Pjg4yG6c9JO4CYr0FRUVrH5XXV3NEip1o0wAjd4GHxvKBAikDk3WAx4TyuoB71NHW1ubOFYroRd9LcCHXABbADimQsR/gesCzQUAOHABiElllxRtJCwsjARYBAA9RQG+pXOXEhRtZmlpid9JvlN0nrAAgG3xpqZtREVFkfOHN/3/AP3NjN3ctJGysjIe/L4RHVfkZxqkbQFRwflW0WERCgy/0GD6XG7L2p4C8/PzGB8fz53/FQCMRIdvgv5Y+hcFxtDQUCwvL8eenh4cHR1lmdRjbXSDpJpCZWUlqzdc1yV3KciLDt6FtwAgFwBmAeDsWsWn0uicnweAQgB4W3RMFShpoMyJ0kfKoR9ro+ej1P5d0QFl/AOXNSas4eEo3QAAAABJRU5ErkJggg==",
  49: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAT3SURBVGhD7ZpJSLVVGMf/jjiklqglLVLBjQslU3FKUytFEAQxA0HCFi7KWrRwBE0XhQqKQYsMXago4sZZnKeFAyLigOScYCaiIopm6onn4Hm593Tvh1zv63fvWz84cD3De+7zf895zvOcK/A/T8YbQDCADwF8DCDJAgt9L/p+IQB8ZANMwRXAVwBGAZwAYFZUTgFMAPgGgLts2FP4HMBv9LDw8HBWWFjIWlpaWE9PDxseHmZDQ0MWV+h79fb2stbWVlZSUsKio6OFGNsAvpANfBU/0MCYmBg2MDDArJmxsTGWkJAghPgJgI1srMyP1Dk/P5/d39/Lz7NaioqKhAg/ywbr8pkwXouUlpYKEb6UDSfeAPB7RESEPE5TJCYmkgB/AnhLFuBrUmd8fFweoykWFhaYnZ0diVAgCzAdFhYm99ckj05xSdchUtBwQXtETQ4PD9nGxgbb39+Xm/SgfgcHB+zu7k5uMgtVVVUkwC2A94QA79Pyb29vl/uajZ2dHebj48OdUGxsrNzMaWpq4jGHu7s7c3V1ZUFBQaympkbu9mz6+/uFM6SokfMRVQwODsp9zUZSUpKYlEVFRcnNrLi4WGmXS3Z2Nnt4eJCHmMzExIR49qdCAIqheSSlBnV1dXoGyStgZmZGafP09ORvnZaph4eHUt/Q0KA35jmMjIyI534iBKBEgoeT5mZ9fZ05OTm9UoC8vDyljcJtQVtbm1IfHBxsNp/wYgJQJBkZGckn8/b2Zm5ubgYFiI+P5/V0PO3u7ir1V1dXit+wtbVlm5ubeuNM5cUEqKys5BM5OjpyB+vr62tQgLi4OEWAvb09pZ7eeEBAgLIKzJWXvIgAi4uLzN7enk9EQtzc3DAXFxeDAuTk5ChGkocW0FFIp4FooyzPHKguwO3tLQsJCeGT0N6lrXB0dGRUADp9hJF09E1PT7PV1VWWkpKi1FNpbGzUG2cqqgtQUFDAJyCPTuc/cXFxobxNisRkcnNz9YwVhbaP+Nzc3CwPMwlVBaBY28bGhk8QGhrKj6/6+npWXl6uGEP7mo5GeU/TVqE2BwcH7vwqKipYenq6IkBXV5def1NRVQBydvJbNFZSU1Pl4RxaNefn5/yzzq0OW1lZkbuahKoCdHZ2/stQYyUzM5OPof1OK4VykaWlJeVZ5DfEtvH39+eO1ByoKsDp6SmbnZ3VK/Pz86y7u1sJiMhBzs3Nsa2tLT6Ggh8hCjnI4+NjdnJywrKyspR6uuMzF6oKYAwSxtnZmU9M574ul5eX/A0LY728vHjgJP4ODAxkZ2dnemOew2sRgFJgYZChewdaJX5+fkofUUis7e1tufuzeC0CXF9fs9HRUX7mk7GGoLdMwU5ZWRlPhtRKzgwJoGo2aGkYEiCBKtS8D7AkDN0HfEAVHR0dcl9NohN6xwsB3gFwSZHXf4HHC5q/AQQIAYhZOUnRKo9J1ioAO10BvqNLBwpQtMza2prISb7XNZ54E8CRoUxNS6SlpZHxZ8b+f4B+M+OZmxaprq4Wzu9b2XBdfqFOWnOIOsa3ywbLkGP4lTrTcjEWtVkLy8vLLCMjQxjfBsBJNtgY9GPpH+QYk5OT+V19X18fm5yc5JGUpZapqSl+0VJbW8vvGx7vJY/JycsGPoW3AZQCWATw16OK1lLonF8GUAHgXdkwU6CggSInCh8phrbUQt+PQvtA2QBD/API1Sr5kSpOggAAAABJRU5ErkJggg==",
  50: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUzSURBVGhD7Zp5SGdVFMdP7lCKypRK7iKIDEa4wWiYlguKwrgMgQhBf/hHWUqKjnuKFCoo5IakIjokOsMILjPuSwi5RAiJoigSSJZLM+oorifOxft476Ymv0V+/qYPXJBzz3vvfr+/9zvvvPsT4H+uzdsA4AUAHwDAxwDwkQ4OWhet7z0AeEcUoApvAsDnADAEAJsAgLdobAPAKAB8CQAWorDr8AkALNLJfH19MSsrC1tbW7GrqwsHBgawv79f5watq7u7Gx89eoQ5OTl47949bsYyAHwqCryKb+nAgIAAfPbsGd5mhoeHMTg4mBvxPQC8IYoV+Y6SU1JS8PT0VDzfreXhw4fchBpRsJwHXLw+kpuby034TBROvAUAv/v5+YnH6RUhISFkwJ8AYCUa8AW5MzIyIh6jV0xPT6OhoSGZkCka8JOPj4+Yr5ecF8Vf5QWRmoYd+o7oApubm7i6uoq7u7vilEYoLS0lA44AwIkb8D7d/m1tbWKuyuzv72NUVBTrIfz9/RWDYomJieIhuLCwgPfv38c7d+6gmZkZ2tvbs4L88uVLMVUtent7eTGkrpHxIQWeP38u5qrM3Nwcv8iFg8TJWVlZQRsbm3/l0QgKCsKjoyNFvjqMjo7yc4dxA6iHZp2Upujo6JAEWFtbo5OTEzo4OLBB4sPCwhT5CQkJUn58fDzW19ejl5eXFKuqqlLkq8Pg4CA/byg3gF4kWDupKfLy8qTF9/X14fHxMR4eHuKrV69wb28PDw4OpNy1tTV2y1MuGXRycsLiMzMzvGKjt7c3np2dya6gOjdiQExMDLuIlZUVE33V4un9gptFdwKHjnF1dWVxMoiM0gRaN4A+bXd3d3YRKmjR0dFMiJubGyYlJbH6IKe8vFwyIC0tTTFH7yN8bmJiQjGnKlo3gAqasbGxtHBxWFhY4Pj4uJQv69MxPz9fca7Q0FBpjt70NIHWDejp6ZEWbWtri8XFxVhbW4uBgYFS3NnZmdUCIjU1VYoXFBQoziU34PHjx4o5VdG6AVtbW6zw1dXV4fLyshSnIujp6SkJ6uzsZPH09PRLDaCnhZivLlo34CoyMjIkQUVFRSxGonksOztbkU89AJ/T1GP6xgygT1xELpZ/3xsaGqRYcnKyIv/u3bssbmBggPPz84o5VdG6AfSdpxcrFxcX1nbKiY2NlcS2tLSw2OTkpBSjVpmzvb3NCibFHR0dFb2DOmjdgJKSEkkQdXNUB2h36cmTJ9LTwdLSEjc2Nlg+3Sn8eU+fdHt7O7548QIzMzMvvTPUQesG7OzsoIeHh7R4c3NzqS/go7GxUXFMU1OTYt7Ozk7629TUFBcXFxX56qB1A4ilpSXFY48Pei+orq4W0xn01SGx8nx6QXr69KmYqhY3YgBnaGiIdXokjrbV/6udpUJXU1PDnhDNzc24vr4upqjNRQZo/G1Ql7nIgGAKaHI/QJe5aD/AmwJUfV8H6IM+NyCIG2ALAHu8M9N3KisrSfwxALhyA4ifqWq/DkRERJABvwGAodyAr6kJoa5Mn6G9CBMTEzLgG7l4whIA1mnPXJ+hzRkA+Puy/x+g38ywsLBQPE4vKCsr48XvK1G4nHpK0reCKBPfJgoWocLwAyXT7TI1NSWe61YxOzuLcXFxXPyPAGAmCr4M+rH0DyqM4eHhrJ2lra6xsTHWSenqoL1G+qeOiooKjIyMRCMjIxL+FxV5UeB1sAGAXAD4BQAOz128LYOe87MAUAQA74rCVIGaBuqcqH2kHlpXB62PWnt3UcBF/AOA1CeBYTEiywAAAABJRU5ErkJggg==",
  51: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAARwSURBVGhD7ZpbSC5VFMf/eSMo71ZKijcE8cEIb6BGaeUVRRQOgQgHeizrISTFa4oUKqiEKRG+qCSKIHi/315SiRCOIIk+BGpmFogXvK5YG/cws/s8xzxf8M3UD9bLzNoz8//PzNpr9vcB/3NvXgEQDeAtAO8BeNcBg6+Lr+8NAK+qAh7CSwA+AjAD4HcAZKL4A8A8gE8AeKjC7sMHAH7mg8XFxVFpaSl1d3fT0NAQTU1N0eTkpMMFX9fw8DD19PRQeXk5JSYmSjO2ADxWBT6NL3lgUlISjY2NkZmZnZ2llJQUacTXAF5Qxap8xcnFxcV0fX2tHs+0lJWVSRO+UQXreSTFW5GKigppwoeqcOZlAL/Ex8er4yxFamoqG7APwFs14GN2Z25uTh1jKVZXV8nZ2ZlN+Fw1YCk2NlbNtyS3RfEnfUHkpuGI3xFH4ejoiDY2Nmhra4tubm7U3c9FQ0MDG3ABIFga8CY//r29vWrugzk9PaXs7GzRQyQkJBiCtxUWFqpDDGRmZoqCFRYWRhcXF+ru52J0dFQWQ+4aBe/whvHxcTX3wayvr8uT2IzAwEB1iEZra6uWFxwcbHcD5ufn5fHTpAHcQ4tOyl709/drInx8fISQoKAgESw+LS1NHUJXV1dUWVlpMOrfeAKmp6fl8d+XBvCHhGgn7YVeyMTEBF1eXtL5+TmdnJzQ8fExnZ2dGfIHBgYoJibGIN7UBuTm5oqTeHt7C9HPKmRRUVGa6Ly8PPLz8zOvAXy3IyIixElYSE5OjhASHh5ORUVFoj7oYXMiIyNFDhdifkJ8fX3Na8D29ja5uroaHmV9eHh40OLiopbP7/7CwoKY+pj9/X2RY1oDRkZGNLH+/v5UV1dH7e3tlJycrG0PCQkRd9oWe3t75OnpaV4DDg8PReHr6OgQjYyEi6D+XR8cHDSMk5jegKdRUlKiGVBbW6vuFljGAL7jKtXV1ZoBVVVV6m6B6Q3gd54/rEJDQ0XbqSc/P18zoKury7BPYnoD6uvrNZHR0dGiDvDqEjc7cnbw8vKig4MDdajA9AbwdMbzujTB3d1d6wtkdHZ2qsM02AAew3n8FJnOAGZzc9Mw7cng74K2tjY13cDu7i65ubmJ/ICAAHMaIJmZmaGmpiZRF3hZfWdnR035G1w8eTWXv06Xlpae2Ub/U2wZYPevQUfGlgEpvMGe6wGOjK31gBje0NfXp+ZaEr7Rtwa8LQ3wB3B8V2dmNVpaWlj8JYAwaQDzA1ft/wIZGRlswBMAznoDPnNycqLl5WU131LwWsTtFPuFXjzjBeBXXjO3Mrw4A+DPu/4/wL+ZUU1NjTrOEjQ2Nsri96kqXM+3nGS1gqgT36sKVuHC8B0n8+OysrKiHstUrK2tUUFBgRT/PYAXVcF3wT+W7nFhTE9PF+0sL3Xx2h13Uo4avNbIf+pobm6mrKwscnFxYeG/cZFXBd6H1wBUAPgRwPmti2YJnufXANQCeF0V9hC4aeDOidtH7qEdNfj6uLWPUAXY4i+dspjNoRNXYwAAAABJRU5ErkJggg==",
  52: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAT1SURBVGhD7ZpZSHVVFMeXc1DOWkKKE4r4YEZOOGBaDij6oJCpKEGPaYEhKc6K5KwQDmT44pADguA8D72kEiGkDzlBIGUOgThgoivWxn04Z3sVu95r1/v1g/Wyz9pn7/U/56y99r4X4H8ejT0A+ABAGAB8CAAf6KDRvGh+7wDAm2IA6vA6AHwGALMAcAgA+ILsGAAWAOBzALAQA3sMHwPAr3Qzf39/zMvLw66uLhweHsbp6WmcmprSOaN5jYyMYHd3NxYUFGBwcDAXYxsAPhEDfIivqWNISAiOj4/jS2Zubg4jIiK4EN8AgIEYrEgVOWdnZ+P19bV4vxdLfn4+F6FFDFjORzx4faSwsJCL8KkYOPEGAPwWEBAg9tMrIiMjSYB9ALAWBcgidebn58U+esXq6ioaGRmRCF+JAvzg5+cn+uslt0nxZ3lCpKLhhL4RXWB/fx93d3fx8PBQvKQRampqSIC/AcCZC/Auvf69vb2ir9qcn59jfHw8qyECAwMVRm3p6eliF+zv78fw8HC0sbFBMzMztLe3x6ioKJyYmBBdn8TY2BhPhlQ1Mt6nBk0OtL6+zgdRaY6Ojgr/6urqOz5y6+npUfg/hYWFBX7faC4A1dCsktIUAwMD0uTpiTo7O6OTkxMzCj46Olry3draQkNDQ+ZrYGCAWVlZ2NHRgSkpKdI9bG1t8fj4WDGGuszMzPD7RnEBaCPByklNUVRUJE1+cnISr66u8PLyEs/OzvD09BQvLi4k34qKCsk3LS1NcZ+goCDFfTTBswiQmJjIBrG2tmZB39zciC4SNG5OTg6mpqay11NOZmamJMDQ0JDimrpoXQB62h4eHmwQOzs7TEhIQDc3N3R3d8eMjAyWHx4DleKenp6SABsbG6KLWmhdgJ2dHTQxMZEmLpqFhQUuLS2J3e5AOzreJyws7MG36N+gdQFGR0eliTs4OLBvvLW1FUNDQ6V2FxcXlgvuo6qqSvKlBLmysiK6qI3WBTg6OmIJq62tDbe3t6V2SoLe3t5SYPd90/KkSNbS0iK6PAmtC/AQubm5UmDl5eXiZfmOjVl9fb3o8mSeTQB64iIlJSVScMXFxYprJIg8+Pb2dsV1TaF1AegVpo2Vq6srKzvlJCUlSQF2dnZK7XTcxtvNzc3v9NMkWhegsrJSCsbHx4flAVrSBgcHpdXBysoKDw4OmD9tfiwtLaU+tIQ2NDRgWVkZlpaWMqM358UsgycnJ+jl5aV4orwu4EalLkd89e8zOuzUBFoXgNjc3FQse9xoX9Dc3Cz50dru6+t7x0+V9fX1KcZQl2cRgDM7O4t1dXUsL9B3vre3p7hOAiwuLrJlkx9vqzK6Tp+KJlAlgMZ3g7qMKgEiqEGT5wG6jKrzgPeogU5kXgXoQd8KEM4FcACAU1WVmT7S1NREwV8BgBsXgPiRsvarQGxsLAnwCwAYyQX4knZdy8vLor9eQWcRpqamJECZPHjCCgD+oDNzfYYOZwDgr/v+P0C/mbHSUx+pra3lye8LMXA535KTviVEWfC9YsAilBi+I2d6XTR5EvNfsLa2hsnJyTz47wHgNTHg+6AfS3+nxBgTE8PKWTrqorKVKildNTprpD91NDY2YlxcHBobG1Pgf1KSFwN8DG8BQCEA/AQAl7cqvhSjdX4NAMoB4G0xMHWgooEqJyofqYbWVaP5UWnvIQagin8AqNk87WZRc7wAAAAASUVORK5CYII=",
  53: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUhSURBVGhD7ZppSK1FGMf/Xa+ilKJiJaS4IYiIEbmAXjG1VBTvB4UIRAj6WOaHuC645lZ4LyiEKd5cQEVxgcu97vsGphIp5JdcbyBmZi644HYnnsF5Oe+kJmeR46kfzJd5n5l3nv8788wzcw7wPzfmTQC+AEIAfAggwggLjYvG9y6At2QHtOF1AJ8DGATwJwB2h8pfAEYAfAnARnbsJnwC4FfqzN/fn6Wnp7OGhgb24sUL1t/fz/r6+oyu0Lg6OjpYY2Mjy8zMZEFBQUKMJQCfyg5exzfUMDg4mHV3d7O7zNDQEAsLCxNCfAfgNdlZmW/JODk5mZ2fn8v93VkyMjKECN/LDmvysXDeFMnKyhIifCY7TrwB4LeAgAC5nUkRHh5OAmwAsJMF+ILUGR4eltuYFDMzM8zMzIxESJMFGPfz85PtTZKLoPizZkCkpGGP1ogxsLGxwV6+fMkODg7kR3qhpKSEBDgB4CIEeI+mf3Nzs2yrNYeHhyw2NpbnEIGBgapCdYmJiXIT1tbWxkJCQpidnR2zsrJiLi4uLCUlhe3s7MimOtHV1SWCIWWNnA+ooqenR7bVmvn5efGSS4uTk5PKvqKi4h82olBg3t3dVdnrwsjIiOg7UghAOTTPpPRFa2ur4oC9vT3/ms7OzryQ85GRkYrt+vo6/+Jka2lpyfLz81l1dTXz8vJS+sjNzVX1rwsDAwOi34+EAHSQ4OmkvsjOzlYG39vby05PT9nx8TFf1/v7++zo6EixpfT6/v373JZSbsHY2JjSBy0NfXErAjx8+JC/hNYzOf3q1SvZRIHEWF1d5QOjACiYmJhQBIiLi1O10QWDC0Bf29PTk7/EwcGBD97d3Z15eHiwpKQkHh+uY3t7m42Pj/NgKQSgQ46+MLgAy8vLzNzcXBm8XGxsbPj0vgoSSdja2tqypqYm2UQnDC5AZ2en4oCjoyMrKCjgUf7BgwdKvaurK48FMrRUfHx8VGKlpqbqdRcwuABbW1s88FVWVrKlpSWlnoKgt7e34tyzZ89U7YizszM+/aemprjjwjY0NJS31wcGF+A6Hj16pDhF292/ERERodg/f/5cfqwVtybAZV+M9nPhUE5ODq+jaU/LgWLHycmJyj4tLU2xLy4uVj3TFoMLQGueDlZubm487dQkPj5ecai+vp7XkSiUHFlYWLCWlhaVvdhOqVRVVameaYvBBSgqKlIG7evry+MA3S61t7cruwNF983NTW5fV1en2FP2Nzs7y/P/p0+fKgkSHWEXFxflV2mFwQXY29tTpbHW1tZKXiBKTU2NYk/TnjI98Yycpt1D0/4m8eKmGFwAYmFhQbXtiULngvLyctmczwbN6S4KiVdYWCib68StCCAYHBxkT5484XGBrtXX1tZkExWTk5OsrKyMf/Ha2lq2srIim+jMZQLo/TRozFwmQBhV6PM+wJi57D7gfaqQtyBThT70hQChQgBHAPv6jLTGDMUYAKcA3IUAxI8Utf8LREdHkwC/ADDTFOCre/fu8UOIKUN3EZRxAvha03nCFsDvdGduytDlDIDtq/4/QL+Zsby8PLmdSfD48WMR/FJkxzWpIiNTC4gazjfLDstQYPiBjGm6TE9Py33dKebm5lhCQoJwvgmApezwVdCPpesUGKOiong6S1ddo6OjPJMy1kJ3jfSnjtLSUhYTEyNOk39QkJcdvAlvA8gC8BOA4wsV70qhfX4OQD6Ad2THtIGSBsqcKH2kHNpYC42PUntP2YHL+BtUTVKGs+WbxgAAAABJRU5ErkJggg==",
  54: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAATkSURBVGhD7ZprSGVVFMf/vobQfGsKjm984AczfIEapZWKoh8UIhAx6KNZH2JQURnzQaIDKmEzRMwHUZJRQRjf78cHX0QoCZI4SCCVZqL4wMe4Ym1mH+7dXSe7XId7b/1gfTj7rH3OWf+zz9pr73uB/7kx3gCiAbwN4H0A75mh8XPx870J4A01AGNwAlAMYALAHwDIguxPANMAPgPgogZ2Ez4C8DNfLD4+nsrKyqijo4OePn1KY2NjNDo6anbGz9Xf30+dnZ1UUVFBSUlJUoxNAB+rAb6Mr7hjcnIyDQ0NkSUzOTlJqampUoivAdiowao0sHNJSQk9f/5cvZ7FUl5eLkX4Rg1Ylw9l8NZIZWWlFOETNXDmdQC/JCQkqP2sirS0NBbgdwDuqgCfsjpTU1NqH6tieXmZ7OzsWIRSVYC5uLg41d8qeZEUf9RNiFw0HPI3Yk4cHBzQ+vo6bWxs0OXlpXraaBobG1mAcwCBUoC3ePh3dXWpvkZzcnJC2dnZooZITEzUM24rKChQu+hxfn4u/Pi5PD09aXd3V3UxmsHBQZkMuWoUvMsNw8PDqq/RrK2tyZsYtLt376pd9CgtLdV83dzcTCrA9PS0vHa6FIBraFFJmYru7m4tAA8PDwoMDCR/f39hHHx6erraRWNubo5sbGy0/qYeAePj4/LaH0gBeCEhyklTUVVVpQUwMjJCFxcXdHZ2RsfHx3R0dESnp6dqFwGfDw8P1xstFilAbm6uuIm7u7sI6urqSnUxSHFxsejn7OxM3t7elikAv+2wsDBxEy8vL8rJyaGQkBAKDQ2lwsJCkR8MwTlIvvWHDx/KosXyBHj27Bk5ODhowajm4uJCs7Ozen329/cpICBAnM/IyBBtsbGxlinAwMCAFqyvry/V1taKN5qSkqK1BwUFiVwgKSoqEu089Le2tkRbdHS0ZQqwt7cnEt+jR49oc3NTa+ckGBUVpYnQ19cn2nt7e8Uxj5qenh7NX44A/oxMuTK9dQFexr179zQBGhoahCg8SvjYx8eH2traqLW1VZifn59od3Jyorq6Ompvb1cvZxSvTAAOTuX+/ft6AvC3L4//yXhWMAW3LgB/87ywCg4OFmWnLnl5eVpAPPT/jQA8i5iCWxegvr5ee2hOZJwH+BvmgOXswOUt5wpe5CwuLtLCwoJmfMwmp1JXV1cxRa6urqq3MopbF+Dw8JAiIyM1ETizy2CkPX78WO32N2JiYoQvl9LXVY7GcOsCMLyE1Z32pHEwnOhuQkREhOjj6OhIOzs76mmjeSUCSCYmJujBgwciL/C2+vb2tupyLfPz82Lo8y4VL49NhSEBTL4aNGcMCZDKDabcDzBnDO0HxHLDkydPVF+rRGfR9Y4UwBfAUU1NjeprlbS0tHDwFwBCpADMAmft/wKZmZkswE8A7HQF+MLW1lYUINYM70XcuXOHBfhSN3jGDcBvvGduzfDmDID96/4/wL+ZUXV1tdrPKmhqapLJ73M1cF2+ZSdrS4g6wXepAatwYviOnXm4LC0tqdeyKFZWVig/P18G/z2A19SAr4N/LP2VEyPv0XE5y1tdMzMzopIyV+O9Rv5TR3NzM2VlZZG9vT0HvsNJXg3wJvgAqATwA4CzFypaivE8vwKgBoCfGpgxcNHAlROXj1xDm6vx83FpH6YGYIi/AAwrRFJxFWD+AAAAAElFTkSuQmCC",
  55: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUaSURBVGhD7ZpZSHVVFMeXI4EpKlZiihOi+GBETqhZWjmiD4opihD0pmlCSIrzWKigGI6kLyqJA4jzPOVDDpFC+pDoQypWZoH4KSq6Ym3ch3O29/uy671yvV8/WC/7rH3OXv99ztpr73sB/ufevAYAXgDwLgB8CAAf6KDRuGh8bwHA62IA6mAGAOkAMAMAfwIAPiH7CwDmASATACzEwO5DEgD8Qjfz8fHBnJwc7OzsxKGhIZyamsLJyUmdMxrX8PAwdnV1YV5eHgYEBHAxdgDgEzHAF/EVdQwMDMSxsTF8yszOzmJISAgX4hsAMBCDFfmanDMyMvD6+lq835MlNzeXi9AoBiznYx68PpKfn89F+FQMnHgVAH719fUV++kVoaGhJMDvAGAlCvAZqTM3Nyf20StWV1fRyMiIRPhSFOB7b29v0V8vuU2KP8kTIhUNJ/SNvAxUVVWRAJcA4MgFeJte/+7ubtFXbc7OzjA6OprVEH5+fgqjtpSUFLELS770For+lJeCg4Px8PBQ7KIWo6OjPBlS1ch4nxrGx8dFX7XZ3NzkD1Fp9vb2Cv+Liwu0tbW94ye33d1dRR91mZ+f5/cM4wJQDc0qKU3R29srDdza2hodHR3RwcGBGQUfFham8N/e3kZjY2Pmb2ZmpvAnc3Nzw/39fUUfdZmenuZj+4gLQBsJVk5qioKCAkmAiYkJvLq6YrP87NkzPD09xfPzc4X/4OCg5F9ZWYmXl5fMnz4l8qd+Nzc3ij7q8igCxMbGsodYWVnda/Dl5eXM38DAAJeXl1mbtipRrQtAs02vLN3TxsYGY2Ji0MXFBV1dXTE1NZXlB5GEhATmb2JiwpKnh4cHOjs7Y1RUlMb3I1oXgJIVBcJfadEsLCxwcXFR8qeZ9vT0vOMnt8bGRsUzHoLWBRgZGZEGTpm9rKwMm5qaMCgoSGp3cnJi3zaxt7fHqzOWCDMzM7GtrQ2TkpIkf0NDQ9za2hIfpRZaF+D4+JglvubmZtzZ2ZHaKanJZ3pgYIC1U0Kkpam1tRWXlpZkd0JMTEyU/OlMQhNoXYAXkZ2dLQVUWloqXr5Df3+/5E+5RBM8mgA04yJFRUVSQIWFhYprlDzFzC//nCIjIxXX1EXrAtA3TyUtZXEqO+XExcVJAXV0dLC2vr4+9Pf3Z6sE9ZVzW7czS09PV1xTF60LUFFRIQ3ay8uL5QGaWXqd+epgaWmJR0dHzF82INa+sLDA6ob19XW0s7OTron5QV20LsDJyQlbx/nAzc3NpbqAW3t7u6IPrwPIqBhyd3dHU1NTqS0tLU3h/xC0LgBBtb182eNG+4KGhgbRnS2JycnJd/xpWczKyvrXSvK/8CgCcGZmZrCmpoZ923SsfnBwILooWFtbw/r6eiwpKcGWlhaVVeNDUSWAxneDuowqAUKoQZPnAbqMqvOAd6ihp6dH9NVLaKJvBXiPC2ALAKf3qcz0gbq6Ogr+CgBcuADED5S1XwYiIiJIgJ8BwEguwBe04+KHEfoKrSq39UWJPHjCEgB+ozNzfYY2VADw9/P+P0C/mWFxcbHYTy+orq7mye9zMXA5reSkbwlRFny3GLAIJYZvyZlel5WVFfFeT4qNjQ2Mj4/nwX8HAK+IAT8P+rH0kBJjeHg4K2dpb047NaqkdNXorJEOUWtra9mB6u1vDX9QkhcDvA9vAEA+APwIABe3Kj4Vo3V+AwBKAeBNMTB1oKKBKicqH6mG1lWj8VFp7yYGoIp/AIgDN3qiLmZ0AAAAAElFTkSuQmCC",
  56: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAVFSURBVGhD7ZppSF1HFMePS6S4oWKrpDFuCCK4lGqEaNOa1gXFBQURRBD6QbS1RavUEDVxo0UNiYiNlKKCihuCoFHjFk1Bq1KLoF/q8qEuXVKtiCZYjaecwRnunWqQt4Tna38wH96ZM/fO+d/7zpyZ9wD+59y8CQB+APAeAHwEAB8aYKN50fz8AeAtOQBNsAKATwBgFAD+BAC8QG0bAMYB4DMAsJUDOw8pAPAzXSwoKAgLCgqwpaUFe3t7cXh4GIeGhgyu0bz6+vqwtbUVb9++jdevX+dirABAuhzgq/iKBoaEhODAwABeZMbGxjAsLIwLUQsAJnKwMl+Tc3Z2Nr58+VK+3oXl1q1bXIRv5ICVJPPgjZHCwkIuwsdy4IQ1APxy7do1eZxRcfPmTRLgdwCwlwX4lNR58uSJPMaomJ2dRTMzMxLhS1mA7wMDA2V/o+QkKf6kTIhUNOzSd8QQoOS7vr6Oa2trcpdOqKysJAH+BgBXLsA79Pq3t7fLvhrz/PlzjImJYTVEcHCwqpEtNTVVHoLHx8d479499PX1RWtra7SyssKAgABsbGyUXbWiv7+fJ0OqGhkfkGFwcFD21ZjFxUV+k1PblStXVP6Hh4eYkJDwLz/eampqVP7aMD4+zq8bwQWgGppVUrqiq6tLTN7BwQFdXV3RxcWFNQo+IiJC5X/yWrLm4+ODTU1NWFpayhMWmpiY4ObmpmqMpoyMjPB7hXMBaCPBykldUVRUJAJ6/Pgxe8IHBwe4v7+Pe3t7+OLFC+FLny9fvsx8LS0tcXl5WfRlZmbi1atX0d/fH6enp4VdG16LAHFxcewm9vb2LGj6fp+F4pVk67TM0dGRbNIKvQtAT9vLy4vdxNHREWNjY9HDwwM9PT0xLS2N5QcltbW1QgAqWefm5jArK4v5Ut/u7q7KX1v0LsDq6ipeunRJBCU3W1tbfPr0qfCnnRvvo6wv+5NwCwsLqntog94FePTokZi8s7MzlpWV4cOHDzE0NFTY3dzc2HefyMnJUQXs5OSEubm57M3hNkqMtLTqAr0LsLW1xRJffX09rqysCDslQQqEB9XT08PseXl5wkZrv/Jpp6SkiD5d1Sl6F+BV5Ofni4BomSNKSkqE7caNGyp/5dukq13qaxOAnrjMnTt3REDFxcXM1tzcLGzh4eEqf8oVvC89PV3Vpyl6F4C+87Sxcnd3Z2WnksTERBEQBU7QqkCFDtkoZyhrhLa2NuFPx3K6QO8CVFRUiEn7+fmxPEAbnO7ubrE62NnZ4bNnz8QYxfkdZmRk4Pb2Ni4tLalWBV1t1fUuAK3b3t7eYuI2NjaiLuCtoaFBNWZyclKUvdRoJaCqkH+Oj49X+WuD3gUg6Okplz3eaF9QV1cnuzPoxJn2DPKY5ORknRZDr0UAzujoKFZXV7O8QMfqGxsbsouKnZ0d7OjowPLycrx//z5OTU3JLlpzmgA63w0aMqcJEEYGXZ4HGDKnnQe8S4bOzk7Z1yihB30iwPtcAGcA2OOVmbHz4MEDCv4QADy4AMQPlLX/C0RFRZEACwBgphTgC1NTU52duhgqVHVaWFiQACXK4Ak7APiNzsyNmZMt9l9n/X+AfjPDu3fvyuOMgqqqKp78PpcDV/ItORlbQlQE3y4HLEOJ4TtyptdlZmZGvtaFYn5+HpOSknjwbQDwhhzwWdCPpb9SYoyMjGTlLB1OTExMsErKUBudH9CfOqiMjo6ORnNzcwr8D0rycoDnwQkACgHgRwA4OFHxojRa5+cBoBQA3pYD0wQqGqhyovKRamhDbTQ/Ku295ABO4x/ASB/7IELc0wAAAABJRU5ErkJggg==",
  57: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAASySURBVGhD7ZpZSHVVFMf/OSGUI1ZCihOiCBrhiBqllYqioIIEIgTfg0hZDyEpjilSDqASpkSIOCWKIDjPQy+pfISQL04PgZSZBeKAie5YG/fm3u3V5HqOXO/XDxYX9lr7nLP+59x11t73Av9zb14FEAbgbQDvA3jPAo2ui67vTQCvqQmYw8sAPgYwB+BPAOwJ2V8AFgF8CsBZTew+fAhgkw4WGRnJiouLWU9PDxsZGWEzMzNsenra4oyua3R0lPX29rLS0lIWGxsrxNgB8JGa4F18RRPj4uLYxMQEe8rMz8+zhIQEIcQ3AF5Sk1X5moILCwvZ5eWlerwnS0lJiRDhWzVhQ3JE8tZIWVmZEOGZmjjxCoBfo6Ki1HlWRWJiIgmwD8BNFeATUmdhYUGdY1Wsra0xW1tbEuELVYAfIyIi1Hir5Loo/mxYEKlpOKLvyItAfX09CfAPAB8hwFv0+Pf396uxZnN6esrS0tJ4DxEdHW1kNJabmytj6+rqWHh4+I04QyN/X1+f0TnMZXx8XBRD6ho579LA5OSkGms2Gxsb4iQmzcvLS8aSGKrflFVVVRmdw1wWFxfFMZOEANRD805KKwYHB+WFu7u7Mx8fH+bt7c2Nkk9KSpKxlFhQUNANCw0NFQWLf1JTowWzs7Pi2j4QAtBCgreTWlFeXi4FmJqaYhcXF+z8/JydnJyw4+NjdnZ2pk65QVdXlxSgpaVFdZvNowiQkZHBT+Lm5saTvrq6UkPuZG9vj7m6uvJjpKamqu4HobsAdLcDAwP5STw8PFh6ejrz9/dnAQEBLC8vj9eH/yInJ4fPd3R0ZFtbW6r7QeguwO7uLrO3t5dfAdWcnZ3Z8vKyOk2ysrIiY/Pz81X3g9FdgLGxMZmAp6cnq6mpYW1tbSw+Pl6O+/r68lpgiuzsbB5Dd39zc1N1PxjdBTg8POSFr729ne3s7MhxKoIhISFShOHhYaN5BD09Dg4O3J+Zmam6NUF3Ae6iqKhIClBdXa26WWNjo/TTxoYePJoAdMdVKisrZYIVFRWqW25eODk5sf39fdWtCboLQN95Wlj5+fnxttOQrKwsKUB3d7eR7+DggLm4uHBfTEyMkU9LdBegtrZWJhkWFsbrAO0uDQ0NybcDveMpYUOWlpbkvIKCAiOfluguwNHREQsODpbJ0OMs+gJhHR0d6jQ+JvxNTU2qWzN0F4Cg5sXwtSeM1gWtra1qOIeKoojr7OxU3ZrxKAII5ubmeGWnukDb6tTi3sb29jZfkdIrVK8CSJgSQPPVoCVjSoAEGtByP8CSMbUfEE4DAwMDaqxVQjf6WoB3hACeAI5NdWbWSHNzMyV/AcBfCED8RFX7RSAlJYUE+AWAraEAn9vY2PClqDVDexHXi60vDZMnXAH8Tr24NUObMwD+vu3/A/SbmWa7r5ZGQ0ODKH6fqYkb8h0FWVtBNEi+X01YhQrD9xRMj8vq6qp6rCfF+vq63GEC8AMARzXh26AfS3+jwpicnMzbWdrqohUbdVKWarTXSH/qoMUU7Sbb2dlR4n9QkVcTvA+vAygD8BzA+bWKT8XoPb8OoBrAG2pi5kBNA3VO1D5SD22pRtdHrX2gmoAp/gWoqYd+VKZZwAAAAABJRU5ErkJggg==",
  58: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAVBSURBVGhD7ZprSFxHFMePmmih9RVsK9RgVDTBD5ZSH6iprbY+UKMmgVKRQLXkQ2htxSKNaEzqqyUJeVCtUopf1FQSBMGYh28tkZogohgCDSoUpFVrBTEmVpJT/oMz7J1sxK67sm77g/kyc2bunP/cOXPu7BL9z6Z5lYjCiOgdIvqAiN63w4J5YX5vEtFrugOW8DIRfUpEPUT0JxHxDip/EVE/EX1ORB66Y5vhIyL6FYNFRETwyZMnuampidvb27mrq4s7OzvtrmBe169f5+bmZi4pKeGYmBgpxiQRfaw7uBHfoGNsbCzfvHmTdzK9vb0cHx8vhfiOiJx0Z3W+hXF+fj4/ffpUH2/HUlxcLEX4XnfYlA+l845IaWmpFOET3XHwChH9FhkZqfdzKBISEiDALBF56wJ8BnX6+vr0Pg7FvXv32MXFBSJ8pQvwc3h4uG7vkKwHxVHTgIikYQl7xB6YnZ3l6elpXlxc1JuswtmzZyHA30TkLwV4C69/S0uLbmsxKysrnJaWJnKIqKgoQ0FdTk6O3oWvXLnC0dHR7O3tza6uruzr68uZmZk8Ojqqm26JGzduyGCIrFHwHipu3bql21rM/fv35UPMFj8/P4N9TU3NczayuLu7W1WE/v5+OXaSFAA5tMikrMW1a9eUA3v27GF/f3/eu3evKHA+KSlJ2S4sLLCXl5ewxcqXl5eLjBNviRwjPT3dMP5W6O7uluMmSgHwISHSSWtx6tQpNfnbt2/z2toar66u8qNHj3h5eZkfP36sbO/cuaNss7OzDeMEBgaKegiHMazBtgiQkZEhHoL9DKefPXummyjGx8eVAMePHze0HThwQNSHhIRYLTO1uQBYqeDgYPEQHx8fPnTokFjJoKAgPnbsmIgPun1cXJyw9/T0FMH4wYMHhreosrLS0Gcr2FyAqakp3r17t5q8Xjw8PHhwcNDQZ35+Xr01pmXXrl0ij7cmNhego6NDOYCjrKKiguvq6vjgwYOqft++fSIWSObm5vjw4cPPCYCsrbCw0BAztorNBUBUR+Crr6/nyclJVY8gGBoaqpxra2sT9RAiLCxMrXhZWRm3trZybm6uss3KytowjvwbbC7ARhQVFSmncNyBxsZGVVdQUGCwT0xMVG1DQ0OGNkvZNgGw4jqnT59WDmGlAT6/ZZ1+AVNdXa3a8EZZA5sLgD2PD6uAgACRdppy5MgR5RBWHmDVZR1ihSknTpxQbQ0NDYY2S7G5AFVVVWrS2NuIAzjDsa/l6YDMD5EfmGaNOCqRF2C/46RAFinbJiYm9EdZhM0FWFpaUgkMCnJ5mReYW01sFZMLTHZzc+P9+/ezs7OzqsvLyzM8YyvYXADw8OFDw7EnC1a0trZWNxfHIL4edXsnJyfh/JMnT/QuFrMtAkh6enr4/PnzIi7gWn1mZkY3MYBIf+HCBXF/d/nyZR4ZGdFNtow5Aaz+NWjPmBMgHhXWvA+wZ8zdB7yNiqtXr+q2DgkWel2Ad6UAvkS0LDMzR+fSpUtwfo2IAqUA4BdE7f8CKSkpEGCCiFxMBfgS5+7w8LBu71DgLgLXbkT0tanzwIuI/sCduSODyxkiWnzR/wfwmxmfOXNG7+cQnDt3Tga/L3THTfkBRo4WEE2cb9Ed1kFg+BHGeF3u3r2rj7WjGBsb46NHj0rnfyKil3SHXwR+LP0dgTE5OVmks7jqGhgYEJmUvRZ8QeJO4eLFi5yamipumIhoDkFed3AzvE5EpUQ0QkSr6yrulIJzfoyIyonoDd0xS0DSgMwJ6SNyaHstmB9S+2DdAXP8A1e/Gz+fm+rcAAAAAElFTkSuQmCC",
  59: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAU0SURBVGhD7ZpbSF1HFIZX44WC9RbUCo14CYLkwVLqBRKr1bYqihIi1IAoxT740KZ9KFI1GlNFWlRQKDZSiz6oKIpgSIx34+2hKkVEfakYpWJq1VQRo7GarPJvnGHviRE5nhOOp/1g4DizZvasf89Zs2aORP9zYjyJKJiIPiCij4noIyssmBfm9y4ReakOmIITEX1BRP1EtE5EfIbK30Q0SERfEZGL6thJuE5Ev2Ow0NBQzsnJ4YaGBr537x739vZyT0+P1RXM6/79+9zY2Mg3b97ky5cvCzHmiegz1cHj+B4dr1y5wp2dnXyWGRgY4OjoaCHEj0T0huqsyg8wvnHjBj9//lwd78ySm5srRPhJdVjPp8J5WyQ/P1+I8LnqOHiLiP4ICwtT+9kUMTExEOAvInJXBfgS6jx8+FDtY1NMTEywnZ0dRPhWFWAkJCREtbdJDoPipD4gImnYwnfEGnj8+DEvLS3xwcGB2mQWSktLIcA/ROQrBHgPy7+5uVm1NZmdnR1OTEzUcojw8HBDQV1aWprahevq6rQ2FxcXdnJy4kuXLnF5eblqdmoePHgggiGyRo0PUdHV1aXamszs7Kx4yJHlwoULBvu8vLyXbESBWC9evDDYn4bBwUExdqwQADm0lkmZi9bWVunA+fPn2dfXl318fLQC52NjY6Xt6OiowRZvHcvU1dVV1tfU1BjGPw19fX1i3E+EADhIaOmkuSgoKJCT7+7u5v39fd7b2+OnT5/y9vY27+7uStusrCxpi3Rb0NTUJOuDg4PNFhNeiwDJycnaQ9zd3TWnj1vCUVFRmi22p4WFBVmPfl5eXlrbuXPneG5uztDPVCwuAN52YGCg9hAPDw9OSkrigIAAvnjxIqenp2vxQU9kZKQUYHFxUdbjjaOfWAXmOpdYXIBHjx6xg4ODnLhaEOWHh4elfUZGhmxDhBZgK8RuINpwyjMHFhego6NDTtrb25uLi4v5zp07HBERIev9/Py0WACw+4h6bH0jIyM8MzPD8fHxBuFqa2vVR5mExQV48uSJFviqq6t5fn5e1iMIwkHhUHt7u2zLzMw0OCuKo6Oj/FxfXy/tT4PFBTiO7Oxs6VBRUZGhDSsF33l8fRD80H716lVpf/fuXYO9qbw2AfDGVQoLC6VDt27dUps1EEM2Nze1z7pbHZ6enlZNTcLiAuBN4mDl7+9vCGrg2rVrLy1pfN+R6OAsMjk5KW1XVlZkEMRYz549041kOhYXoKSkRDqJBAZxALdLbW1tcndwc3PjtbU1zR7Jj7BHoFxdXeX19XVOTU2V9bjjMxcWF2Bra4uDgoLk5J2dnWVeIIo+omM3wBsWbcgdPD095d/ou7GxYXjGabC4AABZm37bEwW5flVVlWrO4+Pj2tao2iNJ0u8k5uC1CCDo7+/XDjeIC1jqy8vLqokEbxnJDgIlDkPmPJzpOUoAs58GrZmjBIhGhTnvA6yZo+4D3kdFS0uLamuT6FLvKCGANxFtq5mZrVJZWQnn94koQAgAfkXU/i9weMiaISI7vQDf4NJhbGxMtbcpcBdxeMD6Tu88cCOiFdyZ2zK4nCGijVf9/wB+M+Pbt2+r/WyCsrIyEfy+Vh3X8zOMbC0g6pxvVh1WQWD4BcZYLkhRzzJTU1OckpIinG8iojdVh18Ffiz9E4ExLi5OS2dx1TU0NKRlUtZacNeIy9OKigpOSEhge3t7OL6KIK86eBLeJqJ8IvqNiPYOVTwrBfv8FBEVEdE7qmOmgKQBmRPSR+TQ1lowP6T2gaoDR/EvxAkfjzJmHtcAAAAASUVORK5CYII=",
  60: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAVTSURBVGhD7ZppSF1HFMePUeLSqi3YWm20KAQFxSIuAddq6xb0Q4wphRAsVvBDjVUU3JeaQIuKCnVDFBEVRYIG1CQkatQSaNRSJBVFaIwF0RaxLo1GjZ5yBme4d/Js5S1BX/qDAT3n3Llz/m9m7pn7HsD/HJv3AMALAIIB4DMA+PQENhoXje9jAHhfTkAb3gKArwFgCABWAABPUVsFgBEASAMAGzmx4/AFAMxRZ35+fpiTk4Pt7e3Y19eHDx48wPv375+4RuPq7+/Hjo4OzM/Px4CAAC7GbwDwpZzgv/EdXRgYGIh3797F08zw8DCGhYVxIX4AABM5WZnvKfj69eu4v78v93dqyc3N5SLUyQkr+Zwnb4wUFBRwEb6SEyfeBoDf/f395euMivDwcBLgDwB4VxYgldR5+PChfI1RMTExgaampiRCtizAj76+vnK8UXK4Kf6i3BCpaNigNWIItra2cGFhAZeWlmSXRlZWVvDZs2e4ubkpu/RCWVkZCbALAB9xAbxp+nd1dcmxOkGJZGRkoIuLC1paWqKtrS2Ghobi0NCQHMqYnZ3FS5cuoZ2dHVpYWOC5c+fYhry+vi6H6sSdO3f4ZkhVI+MTMty7d0+O1Rr6tD08PPiNVI3W4MjIiCr+6dOnaG9v/0osNRJtd3dXFa8LdO/DviO5AFRDs0pKX8THx4sEYmJiWBWZnJwsbJ6enri3tyfir1y5InwJCQnY2NiIXl5ewlZTU6PqXxcGBwd5vxFcADpIsHJSHzx58kQM3N3dXVVQBQcHsyUREhLClgixuLjIpjzFOzk54cuXL5l9cnKS79jo4+ODBwcHoh9dMLgA5eXlQoC8vDzZ/UqFSecLHk8zgUMJu7q6MjsJRELpA4MLcO3aNXYDExMTvHXrFtt0EhMTMSkpCTs7O+VwrKioEALQpqmEziPc9+jRI5VPWwwuQGRkJLsBTV+aujwB3qgaW1tbE/GKOh2LiopUfUVERAgfnfT0gcEFoPWtTNjb2xszMzORCi1uu3r1qohPT08X9uLiYlVfSgFoNukDgwugOH6im5sbPn/+nNlfvHjBdn/uo+c+kZWVdaQAfDZRu337tsqnLQYXIDY2VgyaPnklhYWFwkePRoKS5jZ506QagPv09Zg2uACpqali0PRmRkllZaXw1dfXM1tzc7OwpaSkqOL5jDlz5gzOzMyofNpicAHa2tpEQrQclKSlpQlfT08Psz1+/FjYLly4IGJXV1fRxsaG2Z2dnXF7e1vRk/YYXAAqcKju50nV1tbixsYGjo6OsjqfbFZWVuJwtLOzI5739El3d3ezp0R2drboQ54ZumBwAQia3nzw1BwdHVX/37x5UxXf0tKi8js4OIi/zc3NcW5uThWvC69FAIIKHGtra1VilAw99zVx48YN5lfG0wGpt7dXDtWJ1yYAMT8/j01NTVhaWooNDQ3/uZGRv66ujsW3trbi8vKyHKIzmgTQ+2nwJKNJgDAy6PN9wElG0/sAHzLQ7vsmQB/0oQChXIAPAOBvWndvAtXV1ZT8HgC4cgGIn4KCguRYoyQ6OpoE+BUATJUCZFIRQlWZMTM9PY1nz54lAb5VJk+8AwDLculqbMTFxVHyfx31+wH6zgxLSkrk64wCxau6b+TElTRSkLFtiIrku+SEZWhjaKJgmi7j4+NyX6eKqakpvHz5Mk++EwAs5ISPgr4sXaKNMSoqitX1AwMD7ERHldRJbWNjY+xHHVVVVXjx4kU0MzOjxP+kTV5O8DjYA0ABAPwMADuHKp6WRs/5KQAoBYAP5cS0gYoGqpyofKQa+qQ2Gh+V9uflBDTxDxt4DzZhvjmAAAAAAElFTkSuQmCC",
  61: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAASFSURBVGhD7ZpJSC1HFIZ/BzSaqAmaQbMQhaCoJIjDwjGaxAkFUXgERAzGXdQgLuIIRoQEFQeCUYI7FUXUjSPOwyYq4SE+UQIiZqExSCKi4nzCKaymb73rQ/Q+uH2TD2ph16nu/v/uOnW6rsD/PJh3AXwMIBbA5wA+s8LG98X39wmA91QBj+FNAN8AmAFwCIAM1P4GMA+gGIC7KuwhfAngdz5ZREQElZWVUXd3Nw0PD9PU1BRNTk5aXeP7GhkZoZ6eHqqsrKSoqChpxjaAr1SBr+IHHhgdHU3j4+NkZGZnZykhIUEa8RMAO1Wsyo8cXFRURDc3N+r5DEt5ebk04WdVsJ5nUrwtUlVVJU34WhXOvAXgj8jISHWcTZGYmMgGHAB4RzWgkN2Zm5tTx9gUq6ur5ODgwCZ8pxqwFB4ersbbJHdJ8bk+IXLRcMxz5HVwdnZGu7u7tL+/r3bdy/HxMW1tbdH29jbd3t6q3U+ivr6eDbgE4CsNCOXXv6+vT419EoeHh1RSUkJ+fn7k4uJCHh4eFB8fTzMzM2roS6SmpoqE5e/vT5eXl2r3kxgbG5PJkKtGwad8YGJiQo19NPy0g4OD5YVMGs/B+fl5dYhGa2urFuvr62txA/jad+dPkgZwDS0qKUuRlZWlieCnyVVkQUGBdiwkJISurq5MxlxfX1N1dbWJWa/jDZienpbn/0IawB8Sopy0BOvr65qAwMBAk4IqNjZWTIm4uDgxRSSDg4MUFhZmIt6wBjQ0NGgCKioq1G6zFWZQUJA2JjMzk7y8vIxrQG5urriAnZ0dDQwMiKSTl5dH+fn51Nvbq4aLLM9vCovlRHxyckKenp7GNSApKUlcgJOdudeaq7GjoyMtnuf+wsKCWPqYg4MDcnd3N64BPL/1gkNDQ6m0tJS40JLHcnJy1GEavILwkmlYA3SfnxQQEECnp6fi+Pn5ucj+so8LHXMY3oD09HRNJD95PfpljpdGcxjegMLCQk0k78zoaWpq0vra29tN+iSGN6Crq0sTydNBT3FxsdY3NDRk0icxvAFc4EgB3Nra2kSG50wv13dXV9d7P44MbwDDr7c0gJuPj4/J33V1deoQDTbAzc1NxHHVaEgDmMbGRk2IbM7OzmJ/7lXs7e2Rk5OTiPf29jauAczOzg51dnZSbW0tdXR00ObmphryEhcXF2I3l79Ol5aWLL4fYM4Ai38NWjPmDEjgA5bcD7BmzO0HhPGB/v5+NdYm4Qd9Z0C8NOADACc8T/8LtLS0sPgrAP7SAObXmJgYNdYmSUlJYQNeAHDQG1Bqb29Py8vLarxNsbGxIZfY7/XimbcB/KmWrrZGRkYGi//nvv8f4N/MqKamRh1nE+i26r5Vhev5hYNsLSHqxPepglU4MXRyML8uKysr6rkMxdraGmVnZ0vxvQDeUAXfB/9Yus+JMTk5WdT1o6Oj4ouOKylrbYuLi+KfOpqbmyktLY0cHR1Z+F+c5FWBD+F9AFUAfgNwceeiURqv82sAagF8qAp7DFw0cOXE5SPX0Nba+P64tP9IFWCOfwE4VoCCAN9iSQAAAABJRU5ErkJggg==",
  62: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUKSURBVGhD7ZpZSHVVFMeXA5qWWQ4V9hAGDqCU4oA4puWIvih8pqKG+ZYa6kOOOELOCuFAii/OIiI4z8P38qlEiDkQhtiLGWIiKZboirVxH87ZXkOu935d79cP1oN7rXPOXv+z9z5r7yvA/zwYWwD4CAACAOAzAPhUB436Rf37GADeERNQh9cB4CsAmAeAYwDAJ2QnALAEAFkA8KaY2EP4HAB+ppt5eXlhXl4ednd34+joKM7OzuLMzIzOGfVrbGwMe3p6sLCwEH19fbkYvwDAF2KC/8a3dKGfnx9OTk7iU2ZhYQGDg4O5EN8BgIGYrEgVBWdmZuL19bV4vydLfn4+F6FFTFjOM568PlJUVMRF+FJMnHgDAH719vYWr9MrQkJCSIAjAHhbFCCD1FlcXBSv0SvW19fRyMiIRPhGFOC5p6enGK+X3C6KP8oXRCoazmiOaIOLiws8ODjAw8ND0aWSo6Mj3N/fx+PjY9GlEWpqakiAvwHgAy6AOw3//v5+MfZRUALZ2dlob2+PZmZmaGlpiUFBQTg/Py+GMgYHB5nfysoKTU1N0dbWFkNDQ3FqakoMfRQTExN8MaSqkfEJNWjyQfS2XVxc+IMURnNwaWlJEV9dXX0nTm69vb2K+MdAz769bxgXgGpoVklpitjYWKnzkZGRrIpMT0+X2lxdXfHq6orF7u3toaGhIWs3MDDAjIwM7OzsxPj4eCne2toaT05OxMeoxdzcHL9vKBeANhKsnNQEm5ubUsednZ0VBVVAQACbEoGBgdIcr6iokOITExNld0L08fGRfNPT0wqfumhdgNraWqnTBQUFovtOhUnPzcnJwYSEhDtTIyUlRbrXyMiIwqcuWhcgOTmZPYCG89DQEFt0UlNTMS0tDfv6+sTweyGhHB0dJQG2t7fFELXQugBhYWHsAbTYeXh4SAlwo2rs9PRUvOwOtKPj19DUubm5EUPUQusC0PyWJ+zu7o65ublIhRZvS0pKEi9TUFVVJcXSArm2tiaGqI3WBZBtP9HJyQnPz89Z++XlJVv9uW93d1e8lCFfFMlaWlrEkEehdQGio6OlztObl1NcXCz56NMoItuxMauvrxdDHo3WBaDvOE+A5rGchoYGydfa2qrwlZeXK5Jvb29X+DWF1gXo6uqSkqDpICcrK0vyDQ8PS+00Gni7hYUF+3JoC60LQAUO1f08oebmZjw7O8Pl5WW0sbFhbebm5tLmiDY/8ngHBwc2UsrKyrC0tJRZSUnJ0/kMEjS8eUJkdnZ2ir8rKyulWHHo32d02KkJXooARF1dHRvO8iRol0fncxz6tru5ud1JVpUNDAwo7q8uL00Agvb1HR0d7C23tbXhzs6Owk8C0NSgOp8fb6sy8tNU0QSqBND4blCXUSVAMDVo8jxAl1F1HuBBDXQi8ypAL/pWgCAuwHsA8CfN01eBpqYmSv4KAD7kAhAv/P39xVi9JCIiggT4CQCM5ALk0q5rdXVVjNcrtra20MTEhAQokydPvAUAv4mlq74RExNDyf9x3/8P0G9mrPTUR2RHdV+Licv5noL0bUGUJd8vJixCC0MHBdNw0eRJzH/BxsYGxsXF8eT7AOA1MeH7oB9LD2lhDA8PZ3X9+Pg4K1upktJVW1lZYf/U0djYiFFRUWhsbEyJ/06LvJjgQ3gXAIoA4AcA+OtWxadi9J3fAIByAHhfTEwdqGigyonKR6qhddWof1TaO4gJqOIfQ30koopYRFQAAAAASUVORK5CYII=",
  63: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUySURBVGhD7ZpZSCVHFIaPC26JG5hEjCAGgoKSIOo8uIxREzfUB4UwIGIwvsUFEUbHBY0MJqioEIzigoKKIuLDuOK+gVEJUYwowaiTF2OU6Ij7dsIprKa7dCbm2ne4XvJBPVh1qrvO392n/+orwP/cmfcA4BMA8AeAzwEgWAcbrYvW9ykAvC8moAnvAMA3ADAMADsAgA+o/Q0AYwCQCgBWYmJ34QkA/EYH8/b2xqysLGxubsauri4cHBzEgYEBnWu0ru7ubmxpacGcnBz08fHhYvwOAF+JCb6J72iir68v9vX14UNmZGQEAwMDuRA/AICBmKzI9xSckpKCl5eX4vEeLM+ePeMi/CgmLOdLnrw+kpuby0X4WkyceBcA/nj06JE4T68ICgoiAbYAwFYUIJnUGR0dFefoFXNzc2hkZEQiZIoCTHp5eYnxesl1UfxFXhDJNOzTM6INjo6O8OXLl7i5uSkO3crW1haLPzw8FIdUobi4mAQ4AwAnLoAH3f5tbW1i7L3Y2dnB9PR0dHZ2RnNzc7S2tsaAgAAcHh4WQxkdHR3o7++Ptra2LN7JyQnT0tJwb29PDL0Xvb29vBiSa2R8Rh39/f1irMbQ1XZzc+MnUjR6BsfGxhTxVVVVN+J4o8L86tUrRfx9oHNfHzuEC0AemjkptYiJiZESCA8PZy4yKSlJ6nN3d8fz83MWS2LRFad+MzMzLCwsxPr6enR1dZXi8/PzxVNozNDQED/uF1wA2kgwO6kGi4uL0sIpCbmholucHonHjx+zR4Qge21sbMziyXJzJiYmpOPQPLXQugAlJSXSwrOzs8XhGw7z+PgYNzY22MKoAHKmpqak40RFRSnm3AetCxAfH89OYGBgwAobFZ2EhARMTEzE1tZWMfwGu7u7ODk5yTZhXADa5KiF1gUICQlhJ6Bi5+npKSXBG7mxN1V2LiA1GxubO4n2X9C6APR8yxP28PDAjIwMJKPF++Li4sRpjKurK1YgeZyVlRU+ffpU1beA1gWQbT/RxcVFMjQnJyeK5FZWVsSpeHFxwW7/mZkZljiPJf9wenoqhmuE1gWIjIyUFk5XXk5eXp40Rq/GfyM4OFiKf/HihTisEVoXIDk5WVo0fZmRU1ZWJo2R+SHotj84OMC1tTU8OztTxGdmZkrxRUVFijFN0boATU1N0qLpcZCTmpoqjXV2drI+MjmOjo5oYmKC7e3tivjo6GgpvqamRjGmKVoXgAwO+X6+8MrKStzf38fx8XG0s7NjfRYWFtLmqLGxUYol4zQ/P8/eErW1tZJBojfK6uqqeCqN0LoAhOjtHRwcFH8/f/5ciqXbnpweH6Ok7e3tFfFkj9XirQhAlJaWoqWlpSIRU1NT9n1OZHt7W3G780bz5WKpwVsTgFhfX8e6ujp2Baurq3F5eVkMUTA9PY0VFRUsvqGhgc1Xm9sEUH03qMvcJkAgdaj5PUCXue17gCd1iK8gfYUu9LUAAVwAewA4ULPS6jJUYwDgHAA+4gIQP/n5+YmxeklYWBgJ8CsAGMkFyDA0NGSbEH1maWmJOU4A+FaePGEDAH+K1lXfoC9LALD7uv8foN/MsKCgQJynF8g+1aWJicupoSB9K4iy5NvEhEWoMNRRMN0us7Oz4rEeFAsLCxgbG8uTbwUAMzHh10E/lm5SYQwNDWW+vqenh+3oyEnpaqNP6fRPHeXl5RgREcF3k39RkRcTvAsfAEAuAPwMAKfXKj6URu/5BQAoBIAPxcQ0gUwDOSeyj+ShdbXR+sjafywmcBv/AO7iOjtGihfOAAAAAElFTkSuQmCC",
  64: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAT4SURBVGhD7ZpJSC1HFIaPc9QYjUMUBYMRB1CM4rBwjCZxQjcKj4CID+POIYgLZzAiJKg4EIwSXDmgiK4ccR4WToSgRpQnQc3GRMWJOMSBE05hNX3rqc/c3Pu4fV8+qEVXnaqu83fV6dN1L8D/PBsHAPADgAgA+AIAPtfBQvOi+X0KAB+JDqiDJQBkA8AEABwCACqoHAHANADkAcAHomPP4SsAeEWDBQcHY1FREXZ0dGB/fz+OjY3h6OiozhWa18DAAHZ2dmJpaSmGhoZyMX4DgJeig0/xHXUMCwvD4eFhVDKTk5MYHR3NhfgBAAxEZ0W+J+Pc3Fy8u7sTx1MsxcXFXIQfRYflvODO6yNlZWVchK9Fx4n3AeD3kJAQsZ9eERMTQwL8CQAfigLkkDpTU1NiH71ieXkZjYyMSIRCUYC5oKAg0V4vuQ+Kv8gDIiUNZ7RHtMHFxQXu7u7i3t6e2PQkp6enuLm5iVtbW3h7eys2q011dTUJcA0AH3MBAmj5d3d3i7b/icPDQ8zPz0c3Nzc0NzdHa2trjIqKwomJCdH0Na6vr1n+QfOys7PDg4MD0URthoaGeDCkrJHxGVWMjIyItmpDT9vHx4ffSKXQHpyenha7qFBYWCjZ29jYaFQAuvf92LFcAMqhWSalKVJSUiQHEhISWBaZlZUl1fn6+uLNzY3YjTE3N4cGBgaSraZXwPj4OB/7Sy4AfUiwdFITrK2tSZP39vZWSagiIiLYloiMjGRbROT8/Bw9PT2l/ooUoKamRpp8SUmJ2Pxkhpmdnc36WVlZoYODgzIFSE9PZzegZdzb28uCTkZGBmZmZmJXV5doLkExiAvX3NzMkxblCRAbG8tuQMEuMDBQcooXcuzk5ESlz/HxMbq6urL2uLg4Vsf7Kk4A2t9yhwMCArCgoAAp0eJ1aWlpKn1ohVA9Lf2dnR1W5+fnp0wBZJ+f6OXlxQIbcXV1xaI/b6Mkh+jr62PXJiYmbMtw+Aqwt7d/Mm78W7QuQFJSkuQkPXk55eXlUltPTw9LeJycnNi1o6MjNjU1YWNjIysuLi6s3tLSEquqqrCtrU1lLHXRugA5OTmSk3QyI6eurk5qI4coFvDrNxV6K2gCrQvQ3t4uTZq2g5y8vDypjaL+0dHRa44+Vtzd3VXGUhetC0AJDuX9fOK0rM/OznBmZobtZ6qzsLDA/f19lg0uLi7iwsKCVOiaioeHB7OlsUis1dVV8VZqoXUBCHqPy5+es7OzyjXt6Tfh7+/PbG1tbfHy8lJsVpu3IgBRW1vLXmtyx83MzNj53HOgN4h8tWiKtyYAsb29ja2trVhZWYktLS24sbEhmjzK/Pw8W/p0SkVvC03xkAAa/xrUZR4SIJoqNHkeoMs8dB4QSBWUmLwLyD66orgATgDwF+3Td4GGhgZy/gYAPuECEAvh4eGirV4SHx9PAvwKAEZyAQoMDQ1ZAqLPrK+vo6mpKQnwrdx5wgYA/hBTV30jOTmZnD9+7P8D9JsZVlRUiP30AtlR3Tei43J+IiN9C4gy57tFh0UoMLSSMS2XpaUlcSxFsbKygqmpqdz5LgB4T3T4MejH0j0KjHRGR3n94OAg+6KjTEpXy+zsLPtTR319PSYmJqKxsTE5vk9BXnTwOTgCQBkA/AwAf9+rqJRC7/kVAKgEABfRMXWgpIEyJ0ofKYfW1ULzo9TeQ3TgIf4BHFMsDeHk7lwAAAAASUVORK5CYII=",
  65: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAVISURBVGhD7ZprSGVVFMeX79RMAyuZEJlEHBwpxCfjK618znxxmJQRMcxPvlL8kE9QRyxUVAwfI/pJRRH1g+NjcHznh3xEivkgSTEFK6REfIyKs2Jt7tmcs0dLrvdO11s/WF/2Xuecvf5373XW3ucC/M+leQsA3gcAfwD4GAA+0kGjcdH4PgCAt8UA1MESAJIBYBgAdgAAr5H9AQBjAJAGAG+IgV2GGAD4iW7m6emJWVlZ2NLSgk+ePMFnz57h4OCgzhmNq7e3F1tbWzE3Nxfv3LkjifEzAHwmBvh3fEUX+vr64sDAAF5nRkZGMCgoSBLiGwAwEIMV+ZqcU1NT8ezsTLzftSU7O1sSoVYMWM6nUvD6SF5eniTC52LgxOsA8IuXl5d4nV4RHBxMAvwGAG+KAqSQOqOjo+I1esXMzAwaGRmRCF+KAnzr4eEh+uslqqT4gzwhUtGwR2tEGxweHuLGxgZub2+LXf8KpaWlJMAJADhIArjR9G9vbxd9r8TOzg5mZGTgzZs30dzcHK2trTEwMBCHh4dFVwYlX5qF3t7eCqO8FBAQoDEB+/v7pWRIVSPjQ2p4+vSp6Ks2NNjbt29LD1IYrcGxsTGF//HxMdrZ2b3kK7e1tTXFNepCz1bdM0QSgGpoVklpiqioKD7w8PBwVkUmJibyNldXVzw9PeX+q6uraGxszPosLS3RwcEB7e3tuTk5OeHW1pbiGeoyNDQkjeMTSQDaSLByUhMsLCzwQG/duqUoqPz9/dmSoClNS0Sip6eHX1NSUoInJydsVlD+2N/fx4ODA3zx4gX3vwpaF6CsrIwHk5OTI3afW2EWFxczfwMDA5yamrrQTxNoXYC4uDgeTGdnJ0s68fHxmJCQgG1tbaI748GDB+waExMTjIyMZDOHZkpERITG9yNaFyAkJIQ9gJKdu7s7nw2SUTW2u7vL/emXdnFxeclPbrW1tYpnXAWtC0DrWz54Nzc3zMzMZK84qS02Npb7b25uStUZS4RpaWnY1NSEMTEx3N/Q0BCXlpYUz1EXrQsg236is7MzS2DE8+fPWfaX+lZWVlj70dERezU1NDTg5OSk4l7R0dHcn84kNIHWBbh79y4fNP3ycvLz83kfvRr/ia6uLu5/7949sVsttC5ASkoKHzSdzMipqKjgfXV1dYo+qgvEzN/X18f9qZ7QBFoXoLm5mQ+aloMcWt9SX3d3N2ujN4WPjw86Ojrio0ePFP6qup1ZcnKyok9dtC4AFThU90sDr6mpwb29PRwfH0dbW1vWZmFhwWt72YDQxsaG+VHRMzc3hzdu3OB9Yn5QF60LQND0lgZOJg+EjAofOVIdQEb1AyVPU1NT3paUlKTwvwqvRACivLwcraysFIGbmZmx8zkRKncfPnyo8CWj12J6errGymDilQlArK+vY2NjIxYVFWF9fT0uLy+LLgpmZ2exuroaCwsL8fHjx7i4uCi6XJnzBND4blCXOU+AIGrQ5HmALnPeeYA7NXR0dIi+egn90CoBAiUB7ABgn9bpf4GqqioK/hQA3pMEIL7z8/MTffWSsLAwEuBHADCSC5BJOy7pMEJfobeKqr4olAdP2ADAr2Lpqm/QhgoA/rzo/wP0zQwLCgrE6/QC2VHdF2LgchrISd8Soiz4djFgEUoMjeRM02V6elq817Vifn4e79+/LwXfBgCviQFfBH0s3abEGBoayup62pvTTo0qKV21iYkJdohaWVnJDlRV3xp+pyQvBngZ3gGAPAD4HgCOVSpeF6P3/DwAFAHAu2Jg6kBFA1VOVD5SDa2rRuOj0t5JDOA8/gIipx8vj5WOrgAAAABJRU5ErkJggg==",
  66: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAVUSURBVGhD7ZppSF1HFMePC1q32kJaxYJipShRq6Lxg0ustq4oEQMiiChWCLFqMQp1BRWhxSVRxCpFUFBxQfyiiWLcYsFUQ1NExVBTglViW8SKuLVGTzmDM9w7eSnB9154vvYH88Ez596Z87+znJknwP+8Nu8BwMcAEAoAnwHApwZYqF/UPx8AeF8O4DzYAMAXADABAFsAgBeobAPANADkAcDbcmCvQwoA/Ewvu3LlChYVFWFXVxcODQ3h/fv3cWxszOAK9Wt4eBi7u7uxtLQUg4KCuBi/AECGHOC/8TU9GBwcjCMjI3iRmZycxPDwcC5EEwCYyMHKfEPOubm5eHJyIr/vwlJcXMxF+FYOWEkyD94YKSsr4yJ8LgdO2ALAr4GBgfJzRkVERAQJ8DsAvCsLkEPqTE1Nyc8YFY8ePUIzMzMS4StZgO8DAgJkf6PkbFH8SbkgUtKwS3NEHxwcHODa2hpubm7KVRqhxXdjYwPX19flKp1QU1NDAvwNAC5cAD8a/r29vbKvVmxtbWF+fj66urqilZUV2tvbY1hYGE5MTMiujNPTU6yvr0dvb2+0tbVFGxsb9PX1xfb2dtlVK+7du8cXQ8oaGZ+QYXR0VPY9N/S1PT09eUOqQnNwenpa5X98fIyJiYkv+fLS2Nio8tcGavvsvVFcAMqhWSalK5KSkkTnY2NjWRaZlZUlbF5eXixoztmwZOXy5cvY0dGBVVVVfMFCExMTfP78uaqN8zI+Ps7biuQC0EGCpZO6YHFxUQTj4eGhSqhCQ0PZlLh69SqbIsTe3h46OTkxf2tra3z69Knwv3nzJjo7O6OPjw/Ozc0JuzboXYDa2lohQElJiVz9UoapGJJsn5Z58eKFbNIKvQuQlpYmhu3AwABbdNLT0zEzMxN7enpkd2xqahICUMr6+PFjzM7OZu+hut3dXfkRrdC7AFFRUawBmr/+/v4iOOVX3tnZEf50cuN1tOrL/m5ubri0tKRqQxv0LgDNb2UAfn5+WFBQgJRocVtqaqrwp61S6e/g4IC3bt3ChIQEYaOFkfIJXaB3ARTHT3R3d8f9/X1mPzo6Yqs/r3vy5AmzFxYWChvt/cqvnZKSIup0lafoXYD4+HjRafrySsrLy0UdbY1EZWWlsNHoUXL37l1Rp6tTqt4FyMnJEZ2m+a3k9u3boq6lpYXZOjs7hS0yMlLlPzMzI+oyMjJUdedF7wIoA6LpoCQvL0/UDQ4OMtvy8jLbMcjm6OiIh4eHwp92De5P13K6QO8CUIJDeT/veHNzM9vKHjx4gJcuXWI2SniUhyPF/R3euHEDt7e3cXV1VbUr6OqorncBCBrevONUeKbHS3V1tcp/dnZWpL1UaCcgkfjf165dU/lrwxsRgKirq0M7OztV4JaWlizZ0QTdOLu4uKj8qSQnJ+s0GXpjAhDPnj3DtrY2drBpbW3FlZUV2UUFJUh9fX1shNy5cwcfPnwou2iNJgF0fho0ZDQJEE4GXd4HGDKa7gP8ydDf3y/7GiX0oc8ECOMCOALAHs3T/wINDQ0U/DEAfMgFIH4ICQmRfY2SmJgYEmAJAMyUAhSYmprq7NbFUKGs08LCggSoVAZPvAMAv8mpq7FxdsT+81X/P0C/mWFFRYX8nFGguKr7Ug5cyXfkZGwLoiL4XjlgGVoY2siZhsv8/Lz8rgvFwsICXr9+nQffAwBvyQG/CvqxdJMWxujoaJbX0+UEnegokzLUQvcH9E8dlEbHxcWhubk5Bf4HLfJygK+DAwCUAcCPAPDXmYoXpdA+vwAAVQDwgRzYeaCkgTInSh8phzbUQv2j1P4jOQBN/ANa7AewltBygwAAAABJRU5ErkJggg==",
  67: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAATISURBVGhD7ZppSH1FGMYfF9zKNLAF+xAGoaAW4oprWrmhH1SQQMQwQSQ1xA+5golQLrgQpoSIuCPiF/d96UsqEWKiqCkGYqaUSIolMvEOzuHc8fpHrvf8u976wSCe951zzvOcM3PfmXuB/3kwrwB4B0AYgA8AvG+Cje6L7u9dAK/KAgzhBQCfApgFcAqAPaH2O4AFAPkAXpKFPYSPAGzTyfz9/VlRURHr7u5mw8PDbHp6mk1NTZlco/saGRlhPT09rLS0lAUHBwszfgbwsSzwWXxJHUNCQtj4+Dh7yszNzbHIyEhhxNcALGSxMl9Rcl5eHru5uZHP92QpLi4WJnwjC1aTKsSbI2VlZcKET2ThxIsAfgkICJD7mRVRUVFkwDGAl2UDcsmd+fl5uY9Zsbq6yqysrMiEz2UDvvPz85PzzZLbSfFH9YRIRcM5jREtuLy8ZAcHB+zo6EgO/SvU1NSQAX8DeFMY4EOvf39/v5z7KE5PT1lBQQFzc3Nj9vb2zMnJiUVERLDZ2VmdvOrqaubr68sCAwPvbRTv7e3V6WcoY2NjYjKkqpHzHh2YmJiQcw2Gnranp6e4kE6jMbiwsKDkpqWl3cnR1yoqKnSuYSh07dtzRgsDqIbmlZSxSE5OVm48Li6OV5FZWVnKMS8vL3Z9fc1zSZi7u/ud5u3tLSYs/peKGmMwMzMj7uNDYQAtJHg5aQzW19cVoR4eHjoFVVhYGB8S4eHhfIg8i87OTsWApqYmOWwwmhtQW1urGFBSUiKHH1RhHh4eMmdnZ36O+Ph4OfwoNDcgPT2dX8DCwoINDg7ySScjI4NlZmayvr4+OV0vqamp/Bx2dnZsZ2dHDj8KzQ2Ijo7mF6DXl2Zv8TaIRtXY2dmZ3E1heXlZyc3OzpbDj0ZzA2h8qwX7+PiwwsJCRoWWOEYz/32kpKTwHHr629vbcvjRaG6AavnJZ/OLiwt+/Orqis/+Ira1tSV3ZXt7e8zGxobHk5KS5LBR0NyAhIQERSQ9eTXl5eVKjD4aZerq6pQ4bWxogeYG5ObmKiJoZ0ZNfX29EmtpadGJEeLtcXR0ZMfHx3LYKGhuQFdXlyKSBKnJz89XYkNDQzqxk5MTXi5TLCgoSCdmTDQ3gAocIYRac3MzOz8/Z4uLi8zFxYUfc3BwuLM4orjok5OToxMzJpobQNDrLcRQc3V11fm/qqpK7sLa29uVeENDgxw2Gs/FAIImNBrLauG2trZ8f04flZWVSl5HR4ccNhrPzQBif3+ftbW1cXGtra1sc3NTTlHY3d3lK9LJyUnNJkBCnwFGXw2aMvoMiKQDxtwPMGX07Qf40oGBgQE51yyhB31rQIQw4HUAf9I4/S/Q2NhI4q8BvCUMIL4PDQ2Vc82S2NhYMuAnAFZqAwotLS35UtSc2djYEIutL9TiCWcAv8qlq7mRmJhI4v+47/cD9J2Z0XZfTQ3VVt1nsnA131KSuU2IKvH9smAZmhjaKJlel5WVFflcT4q1tTVlhwlAHwA7WfB90JelRzQxxsTE8Lp+dHSUr9iokjLVtrS0xH/UQYsp2k22trYm4b/RJC8LfAivASgD8AOAv25dfCqNPufXAFQCeEMWZghUNFDlROUj1dCm2uj+qLR/Wxagj38AQ01vMzEx65gAAAAASUVORK5CYII=",
  68: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAVeSURBVGhD7ZpZSJ1HFMePC1ptrQu2lRQpKo2CYhGN4F5t1bhFE6FURKxWH6RqXR7qjhWxRcWFaiNFfHFFE19cIm5Ri9IoIm4YCkXSF6vRKuJa0VPO4Hx8d7wJcr036G1/MA85c+b75vznfmfOjAH4n0vzHgA4A4APAHwOAJ9dw0bzovl9AgDviwGowtsA8A0AjADAJgDgDWp/A8AYAKQDwLtiYJfhSwD4nR52584dzMnJwZaWFuzp6cGhoSEcHBy8do3m1dvbi62trZifn4+enp5cjD8A4CsxwNfxAw308vLCJ0+e4E1mdHQU/f39uRA/AYCOGKzIj+SclpaGp6en4vNuLLm5uVyEn8WA5XzBg9dGCgoKuAhfi4ET7wDAn+7u7uI4rSIgIIAEWAcAc1GAVFLn6dOn4hitYmZmBvX09EiE70QBfnVzcxP9tZLzpDgnT4hUNOzSN6IJDg4O8MWLF7i2tiZ2KWV9fR1XV1dxe3tb7FIL5eXlJMA/APARF8CFfv4dHR2i75XY3NzEzMxMtLGxQSMjIzQ1NUU/Pz8cGRkRXRltbW3o4eGB5ubmaGBggFZWVhgZGYlzc3Oi65Xo7+/nyZCqRsanZBgYGBB9VYZW29HRkb9IodE3ODY2puBfV1d3wY83ExMTtYpA7z5/dhAXgGpoVkmpiwcPHkgBhISEsCoyKSlJsjk5OeHJyQnz3draQjMzM2anlS8pKWEVZ2xsrOQfHh4uvkJlhoeH+XMDuQB0kGDlpDpYXFyUJu7g4KBQUPn4+LBPwtfXl30ixOTkpOQfExMjexKira0ts1tbW0uCXRWNC1BRUSEFlJeXJ3ZfqDAXFhYk/+TkZIU+EpDst2/fvjBOVTQuQFxcHHuBjo4OPnr0iCWd+Ph4TExMxPb2dtGdrSz9ImgMJUpKxisrK1hYWCgJU1paKg5TGY0LEBQUxF5Ayc7V1VUKgjeqxnZ2dhTGvHz5Eu/du3fBV19fn9Xx6kTjAvDV5M3FxQWzs7ORCi1uowQnZ2NjA+/fv39BABIxKysLDw8PFfyvgsYFkB0/0d7eHvf395n96OiIZX/e9/z5c2bf29tDZ2dnZqMVLyoqwsePH2NCQoLkGxUVhWdnZ8KbVEPjAtCWxSdOKy9H/l3T1kg0NzdLtoyMDAX/wMBAqW9qakqhT1U0LkBqaqo0abqZkVNVVSX1PXz4kNno+M1t4gVMWVmZ1NfQ0KDQpyoaF0C+ovQ5yElPT5f6uru7mY1WXRSFk5KSIvU1NTUp9KmKxgWgAoe2Mz7x+vp63N3dxfHxcbS0tGQ2Y2Nj6XDU1dUl+drZ2bG6gL73iYkJtLCwkPqWlpbEV6mExgUgaCX5xKndunVL4d/yff34+Fh+gYmGhoYseerq6ko2qiHUxRsRgKisrGQHGXngFJyyfZ22wbCwMAVfalRMUfC0g6iLNyYAQef6xsZGdsChJEYV3uugTE+Jku4mamtrcXZ2VnS5MsoEUPtp8DqjTAB/MqjzPuA6o+w+wJUMnZ2doq9WQgt9LoAfF8AKAPboO/0vUFNTQ8GfAIAtF4D4zdvbW/TVSu7evUsCLAGAnlyAbNp3nz17JvprFcvLy+zaDQC+lwdPmAHAX2Lpqm1ERERQ8Nuv+v8D9DczLC4uFsdpBbKrum/FwOX8Qk7alhBlwXeIAYtQYmgkZ/q5TE9Pi8+6UczPz2N0dDQPvh0A3hIDfhX0x9I1SozBwcGsru/r62MnOqqkrmujEyTdKVRXV2NoaCi7YQKADUryYoCX4QMAKACAWQA4PlfxpjTa5+cBoAQAPhQDUwUqGqhyovKRaujr2mh+VNp/LAagjH8B8lQC9HqujYQAAAAASUVORK5CYII=",
  69: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAVFSURBVGhD7Zp/KO9XGMcfv3cZNmHL/hhKdN0scf3h59jmV5SoWQnLlD+GJX/MzxhqC6GW0Yg/EOl2S1zkt8sfQ0tCNPk1hckdCd3Nj2c9J+f0+Zy5S1/f743v3atOcc7z+XzO8z7nPOc5B4D/uTF2AOABAAEA8CkAfHIHC/WL+vcRANjLDmiCBQB8DQDDALAPAHiPyp8AMAYAmQBgJTt2E74AgN/oZY8fP8acnBxsbW3F7u5uHBwcxIGBgTtXqF89PT3Y1taG+fn56Ovry8VYBYAvZQf/i+/pQT8/P+zr68P7zMjICAYHB3MhfgQAA9lZmR/IOCMjAy8uLuT33Vtyc3O5CD/JDiv5nDuvjxQUFHARvpIdJ94GgN99fHzk5/SKkJAQEuAPAHhXFiCd1BkdHZWf0StmZmbQyMiIRPhWFmDC29tbttdLroLirDIgUtJwRGtEF5yenuLm5ibu7OzITdeyvb2NW1tbeH5+LjdphfLychLgbwD4kAvgSdO/o6NDtr0V+/v7mJWVhU5OTvjgwQO0trbGoKAgHB4elk0Zzc3NLOewsrJCCwsLfPjwIVZWVspmt6a3t5cHQ8oaGR9TRX9/v2yrMTTa7u7u/EOqQmtwbGxMZZ+Xl/cvO14SEhLw8vJSZX8b6NtX7w7lAlAOzTIpbREbGysciIiIYFlkamqqqHv06BGenZ0x28nJSVFvY2PDRp2mKc0YXt/Q0CB/QmOGhob4ez/jAtBBgqWT2mB+fl503M3NTZVQBQQEsCURGBiIL168YHVpaWnCnoTitLe3i3oPDw+txQSdC1BRUSE6TlNbRs4wKS6QLS2N9fV1UX9ycoL29vaszdDQEFdWVlTPaYrOBUhMTGQfMDAwwCdPnrCgk5ycjCkpKWxUZWg2cAE2NjZEPY24s7OzEFNb5xKdCxAaGioc8vLyEg7wQtnY4eGhsE9KShJtJBaHtkLaDXgbnfK0gc4F4CPKi6enJ2ZnZyMlWryOIjuHdh9eT1vfxMQELiwsYHh4uOo9TU1Nqu9ois4FUBw/0dXVla1l4uXLlyz687bl5WXxDC0PpbO8mJqaip9bWloUX9EcnQsQFRUlOk0jr6SwsFC0KSM+UVpayta8iYkJC34lJSUYExMj7Lu6ulT2mqJzAdLT00Wn6WZGSVVVlWirq6tTtXHW1tZEjFDc6rDtVRvoXACaqrzTtByUZGZmiranT5+yOlrvlOjQWWR2dlbY7u7uiiBIuQMtIW2gcwHoDKDM4mpra/Ho6AjHx8fR1taW1Zmbm4vDES0Fbuvv7497e3vsHfHx8aJenkm3QecCEDS9eeepODg4qH4vKysTtsfHx2yEeRuJZGdnJ353cXHBg4MD1ftvw2sRgKCc3tLSUuW4mZkZu5+TmZ6eRkdHR5UtFdpSV1dXZfNb8doEICi1bWxsZBG9vr4el5aWZBMBjTIlO0VFRewwpM3DmZLrBND6afAuc50AwVShzfuAu8x19wFeVNHZ2Snb6iWK1DuIC/A+ABzTOn0TqKmpIefPAMCZC0D8Qnvwm8DVIWsBAIyUAmTTpcPU1JRsr1csLi7yA9Z3SueJdwBgV05d9Y3o6Ghy/uBV/z9AfzPD4uJi+Tm9QHFV943suJKfyUjfAqLC+Q7ZYRkKDI1kTNOFUtT7zNzcHMbFxXHn2wHgLdnhV0F/LN2hwBgWFsby+mfPnrETHWVSd7U8f/6cXZ5WV1djZGQkGhsbk+N7FORlB2/CewBQAAC/AsBfVyrel0L7/BwAlADAB7JjmkBJA2VOlD5SDn1XC/WPUnsX2YHr+AergQdHIQj/TAAAAABJRU5ErkJggg==",
  70: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAATLSURBVGhD7ZpbSGxVGMf/qahQSIGlkhp4fxAjvIEapuUFxQfxGD0eDDwPaSoJXlHTh/KGQuqDJCooiQgKXo63vCKkxwwxUYREA8lCTDO8pX7xbVyb2euMYuM0jXP6wUJZ31pr7/9/7/n2t/YM8D935nUAAQDeBfABgPfNsPF58fm9DeANWYAhvAzgEwDfAtgDQA+o7QOYAvApAAdZ2F34CMAGLxYcHEz5+fnU0dFB/f39NDY2RqOjo2bX+LwGBgaos7OTioqKKCwsTJjxE4DHssDb+IInhoeH09OnT+khMzExQVFRUcKIrwC8JIuV+ZIHZ2Zm0uXlpbzeg6WgoECY0CQL1uVDId4SKS4uFiZ8LAtnXgHwc0hIiDzPooiOjmYDfgXwmmxABrszOTkpz7Eonj17RtbW1mxCnmzAbFBQkDzeIrlOij/oJkQuGv7gz4g5sLe3R1tbW3R0dCSHjEJVVRUbcA7gLWHAO3z7d3V1yWMNIi0tjfhuCg0N1ds4z3BtsbS0pJm3vr5OycnJ5OjoSPb29uTq6qok5MPDQ824+zI0NCSSIVeNCu9xx/DwsDzWILy8vMQBbm0jIyPqnM3NTXJycnpuDLfIyEg6Pz/XHOM+TE1NibVjhQFcQyuVlDFITU0lX19fTfPz81P+ClHOzs60vb2tmSNijx49oubmZgoICFD7GhoaNMe4D+Pj42LdGGEAbySUcvLfJCsrSzkwZ+GZmRm1f2dnR7nlOebm5kYXFxdK/+LiosjYFBgYSFdXVzqrGc5/YgCvLa5maWmpJsb7CxHjO0HAgj08PJR+NoiNMgYmN+Dk5IR8fHyUg3p7e9PZ2ZkmXlNToxqQk5OjifF+RMTm5uY0MUMxuQGNjY2qiLa2NjmsW6dTSUmJJhYTE6PGeKdnDExqAF9tT09P9eqfnp7KQyg7O1sVKX88dA3o6enRxAzFpAb09vaqAiorK+WwQm5u7o0GxMbGqrG+vj5NzFBMaoB4vHE2X1tbk8MKLFqILCws1MS4BhAxYz2mTWYAV3Bc1fHat+0zWlpaVJFPnjzRxPz9/ZV+KyurGw38p5jMAJ2KizIyMuSwyvz8vDqOS2XB/v4+OTg4KP3u7u7K08QYmMyA2tpaVVhra6scVuFEKZ73fKW7u7vp4OCA8vLybrwz7oPJDEhPT1cFzM7OymENbJAYy83FxUX9387OjjY2NuQpBmMyAxITE1URKysrcvg5KioqFLG6RvAGiZ8kxsRkBvB2l3eYnL2Pj4/lsF440TU1NVF5eTm1t7fT7u6uPOTe6DPAqLtBc0efAVHcYaz3AeaOvvcBgdzB2fdFgC/0tQGRwgBnAH/y5+5FoL6+nsX/BcBDGMB8FxERIY+1SOLj49mAHwFY6xrwGRchXJVZMqurq2Rra8sGfK4rnnkVwC6/M7dkkpKSWPzvN/1+gL8zo7KyMnmeRVBdXS2SX5YsXJdmHmRpCVFHfJcsWIYTw9c8mG+XhYUFea0HxfLyMqWkpAjx3wCwlwXfBH9Z+gsnxri4OOXF5eDgIE1PTyuVlLk2ft3OP+qoq6ujhIQEsrGxYeG/cZKXBd4FJwDFAL4HcHbt4kNp/JxfBlAO4E1ZmCFw0cCVE5ePXEOba+Pz49LeWxagj78Bk5F21Guy4/wAAAAASUVORK5CYII=",
  71: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAQSSURBVGhD7ZpJSB1ZFIb/qEigoyZoJ0oW4iwuuglOoIZEMzghCGKTZWPATWtcZBFFReMm7YRKMIumQRdKRBAXzkMcN63SBrFFaXBhLzSGkKCIU9QTTuEt3jt5L9jmBd675IPLgzrn3qrz161T5956wHfOzI8AfgJwE8BdAHecsPF18fX9DOCqDOA8/ADgNwCvALwDQC7U3gOYAPAIgLcM7Cw8APAvDxYbG0vFxcXU1tZGPT09NDIyQsPDw07X+Lp6e3upvb2dSktLKSEhQYmxCuBXGeCXeMYdExMTaWBggFyZsbExSk5OVkI8B3BBBiv5nZ0LCwvp+PhYjueylJSUKBFeyIAt+UUFryNlZWVKhIcycOYSgP/i4uJkP61ISUlhATYBXJECFLA64+Pjso9WzM3Nkbu7O4vwRAowHRMTI/215DQpvrZMiFw0bPMz4ixsb2/TysoKra6u0snJiTR/FTU1NSzAIYBAJcANnv4dHR3S91zk5eURz6b4+HibjfMM1xbz8/Oyq0l6erqRsIKDg+nw8FCav4r+/n6VDLlqNLjNBwYHB6XvuQgNDVUn+GIbGhqSXQ2amppMn8DAQIcLMDExoca/rwTgGtqopBxBbm4uRUREWLXIyEjjVwXm7+9Pa2trVv2Ojo6ovLzcSqRvMQNGR0fV+PeUALyQMMrJb0lRUZFxYs7CU1NTVrauri6Kjo7+bJZoIwCPrYKqqKiQZoqKijLt2dnZ5Ofnp48Ae3t7FB4ebpw0LCyMDg4OrOyc5fkR4WA5Ee/s7JCvr68+AjQ3N5t3t7W1VZqNZ39yctJ49TGbm5vk7e2thwB8t0NCQsy7v7+/L10+Y2Njg3x8fPQQoLu727z71dXV0mwTrQTgVyKPzZl/eXlZmm2ijQBbW1tmNv8/6wxtBLCouKigoECa7aKNAPX19aYALS0t0mwXbQTIz883BZienpZmu7AAXl5eRr+goCDXFSAzM9MUYHFxUZrtsr6+Tp6enka/gIAA1xWAl7u8wuRF1u7urjTbhWsH3s3lvjxzHL0fYEsAh64GnR1bAiTzAUftBzg7tvYDovlAZ2en9NUSvtGnAtxSAvgD2KmqqpK+WtLY2MjBfwQQrARg/kpKSpK+WpKWlsYC/APA3VKAx25ubjQzMyP9tWJpaUm9Yp9aBs9cBvCG98x1Jisri4P/YO//A/zNjCorK2U/LaitrVXJr0gGbskf7KRbQrQIvkMGLOHE8Cc783SZnZ2VY7kUCwsLlJOTo4J/CeCiDNge/LF0gxNjamoq1dXVUV9fn7F3x5WUszbebuc/dTQ0NFBGRgZ5eHhw4G85ycsAz8I1AGUA/gZwcKqiqzR+zy8AqAJwXQZ2Hrho4MqJy0euoZ218fVxaR8mA7DFJ7Bv6CARvv1kAAAAAElFTkSuQmCC",
  72: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAASGSURBVGhD7ZpLSHVVFMf/PtCgzPJRQoMgn6My8oUPTMsUxZGY6UDCwEmaYIMUBV+DfCuEDjLUgZo6UfD9yOcklQwxFcIQm5QhhhE+Cl2xDu7DucuryL33u9/1fv1gIZy99tnn/7/bfdbe9wL/82D8AbwJIAHA+wDec8Dg5+LnewvAK1KAJTwP4FMA3wE4BkCPKE4ALAH4DMCLUthD+AjAz3yzyMhIKisro76+PhobG6O5uTmanZ11uODnGh8fp/7+fqqoqKDY2Fhlxi8APpYC7+NL7hgXF0dTU1P0mFlYWKCkpCRlxFcAXKRYST0nFxcX09XVlbzfo6W8vFyZ0CkFG/lQiXdGKisrlQmfSOHMCwB+jYqKkv2ciuTkZDbgCMDL0oAidmdxcVH2cSo2NjbIzc2NTfhCGrAaEREh852Sm0XxR+OCyEXDX/w/4ggcHR3RwcEBHR8fyyab0NjYyAb8A+B1ZcDbPP0HBwdlrkUUFBQQz6bo6GizwesM1xabm5sm/YaHhykxMZF8fHzI09OT/P39KSUlhaanp03yrGVyclIthlw1arzLF2w1UFBQkBrg3piZmdH7NDQ03Go3xsDAgMkY1rC0tKTu+4EygGtorZKyBdnZ2RQaGmoSYWFh2l8lKCAggA4PD7X8/f19cnV11a67uLhQUVERdXd3U05Ojp7v6+tLJycnciiLmJ+fV/dNUQbwRkIrJ58kJSUl2sC8Cq+srOjX6+rqdKF5eXkmfWJiYvQ244yxhqdiAN9bCamqqrrVVlpaSrm5udr0NJKfn6/3Gx0dNWmzFLsbcH5+TiEhIdqgwcHBdHl5KVPMwqW46sexu7srUyzC7gZ0dHToInp7e2XznfCOTvVLSEig6+trmWIRdjWAP+3AwED907+4uJApZqmvr9fF8wK5vr4uUyzGrgaMjIzoQvhV9xCMiyJHZ2enTLEKuxrAr0S+N6/8e3t7svkWhh2bFi0tLTLFauxmwOnpKfn5+WmDPWSfUVtbayK+q6tLptgEuxlgqLi04uY++LhN5Xp5eWnl6pPCbgbw9FWienp6ZLMOb368vb31XF4sW1tbqaamhqqrq7Xg2uHRvQYLCwt1Uaurq7JZR079u4IPO22B3QzIyMjQH357e1s2a/C7PTw8/JZYczE0NCS7W4TdDODtLu8weZN1dnYmmzXYgOXlZa3O5/HvCm7nfxVbYM4Am+4GHR1zBiTxBVudBzg65s4D3uELfCLzLMAf9I0BicqAAAB/82r8LNDe3s7i/wXwhjKA+T4+Pl7mOiVpaWlswE8A3IwGfM67rrW1NZnvVOzs7JCHhwcbUGMUz7wE4Hc+M3dmMjMzWfyfd/1+gL8z00pPZ6SpqUktfiVSuJGvOcnZFkSD+EEpWMILwzeczNPFlicxT4OtrS3KyspS4r8F8JwUfBf8ZelvvDCmpqZSc3MzTUxMaGUrV1KOGnzczj/qaGtro/T0dHJ3d2fhf/AiLwU+hFcBVAL4AcDljYuPJfg9vwWgFsBrUpglcNHAlROXj1xDO2rw83FpHywFmOM/u5aMQMLvIzYAAAAASUVORK5CYII=",
  73: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAS8SURBVGhD7ZpbSHRVFMf/qYhSioKVkCD0eXuQIvICXjC1UhR9EaMXIfzAlzQfAi8oaCKWNxTCBE30QVFEELzf8g6mkikmQpiXXsyM1PCeumId3Ic529FPZqbTOPSDzTBnrb3PWf+zz9prnxngfx7NqwDeAhAJ4H0AsVbY+Lr4+t4G8JocgCm8DOBTAN8B+AMAPaH2J4BJAJ8BcJUDewwfA/iZBwsODqa8vDxqbW2l3t5eGh0dpZGREatrfF19fX3U1tZGBQUFFBYWJsT4BcAncoAP8SV3DA8Pp8HBQXrKjI+PU3R0tBDiawAvycHKfMXOWVlZdH19LY/3ZMnPzxcifCMHbMhHInhbpLCwUIjwXA6ceQXAryEhIXI/myImJoYF2APgLguQyepMTEzIfWyKxcVFsre3ZxFyZQFmgoKCZH+b5DYp/miYELlo+IufEWtgb2+PdnZ26OTkRDZZhIqKChbgEoC3EOAdnv4dHR2yr0mkp6cTz6bQ0FCjjfMM1xZLS0uafl1dXRQZGUnu7u7k7OxM3t7elJ2dTYeHhxo/cxkYGBDJkKtGhff4wNDQkOxrEj4+PuIED7bh4WG1T319/R27aCzY0dGR5hzmMDk5Kcb+UAjANbRSSVmC1NRU8vf317SAgADlUwTl6empTHNmd3dXueN83MnJiUpKSqipqUnpI/yLiork05jM2NiYGPcDIQBvJJRy8t+EpzOfh7Pw9PS0epzLawcHB8XGJbeAfYQA/GhYiv9EAB77vrt5dnZG29vbyoVxAhTMzs6qfZKSkjR9zEF3AThAPz8/5aS+vr50cXEhu2g4ODigmZkZJVEKAXiTYyl0F6Curk4NpKWlRTbfIS0tTfV3c3Oj9vZ22cUsdBWA7/azZ8/Uu39+fi67aLi5uaHAwEBVAFdXV8rJybHoKqCrAN3d3Wow5eXlsvkOV1dXyvSfn59XAhd9o6KiXvjoPBZdBeAlkcfmzL++vi6bX0hsbKwqQk9Pj2w2Cd0E4Gnr4eGhnOyhfQZP++PjY9rc3KTLy0uNLTc3VxWgrKxMYzMV3QQwqLgoMzNTNqvwsujl5UWOjo7U2dmpsSUnJ6tjNDQ0aGymopsA1dXV6sU3NzfLZhVeGYQfV3/Ly8tK/d/Y2KgWSPwIbWxsyF1NQjcBMjIy1MA4sd0HT3uu9IQvB82lsvjOjctjS6GbAImJiWoAq6urslnD/v6+ZrqL5uLiQqWlpbK7WegmAG93eYfJm6zT01PZbJS5uTmqra1V7jg/NltbW7KL2RgTwKK7QWvHmADRfMBS7wOsHWPvA97lA/ISZKvwjb4VIEoI4Ang2JKZ1prhHAPgbwBvCgGY7yMiImRfmyQ+Pp4F+AmAvaEAn9vZ2SmbEFtmbW1NqTgBfGEYPOMG4Dd+Z27L8JslAAf3/X+AfzOj4uJiuZ9NUFlZKZJfthy4IQ3sZGsJ0SD4DjlgGU4M37IzT5eFhQV5rCfFysoKpaSkiODbATjJAd8H/1i6y4kxLi6OqqqqqL+/n6amppRKylobv0rnP3XU1NRQQkKC2E3+zkleDvAxvA6gEMAPAC5uVXwqjdf5FQAlAN6QAzMFLhq4cuLykWtoa218fVza+8oBGOMfZwqh2YqIjoAAAAAASUVORK5CYII=",
  74: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAASBSURBVGhD7ZpZSGRHFIZ/VwaN0XGJgmAgrk+OwQ3UkGgWRfFJDHkzOOCLGh/yoKKgI0LihkowQggookRE8cF9ieuDGzGoEUUNaF6SaFAM4xqlwimmLj0n7Wiam6b7kg8OTVedqr7nv9Xnnqpu4H8ejB+ASADvAPgAwPs2aHRddH1PALzBA7AEdwD5AL4H8AcAYUd2DGAGwGcAXueBPYRPAOzQZLGxsaKkpER0dnaKgYEBMTExIcbHx23O6LoGBwdFV1eXKCsrEwkJCUqMnwF8ygN8FV/QwMTERDEyMiLsmampKZGcnKyE+AqAAw+W8yU5FxYWitvbWz6f3VJaWqpE+JoHbMrHKngjUl5erkR4ygMnXgPwS1xcHB9nKFJSUkiA3wE85gIUkDrT09N8jKFYWVkRTk5OJEIxF2A+JiaG+xuSF0nxR9OESEXDn/QdsSVOT0/F9va22N3dFTc3N7zbYmpra0mAawBvKgHepuXf3d3NfS0iNzdX0GqKj483a5RnqLZYXV3lQzWur6+lD12Xj4+PODo64i4WMzw8rJIhVY2S96hhdHSU+1pESEiI+oBX2tjYGB+qUVxcrPl5eXnpKsDMzIya+yMlANXQspLSg+zsbBEeHv6SRUREyFcVVEBAgDg4OOBDJfPz88LBwUHz1XsFTE5Oqrk/VALQRkKWk/8lRUVF8oMpC8/NzfFuydnZmQgLC3tppRhCAJpbBVRRUcG7NfLz86WPh4eH8PPzM4YAFxcX2l0NDQ0VV1dX3EVCOUiJ1NraqooW+xegpaVFC6y9vZ13S05OTkRQUJD0SU1NlW3R0dH2LwDd7eDgYO3uX15echdJTk6O9KGlv7+/L9siIyPtX4D+/n7t7tfU1PBuSV9fn+x3cXERvb29WrtaAb6+vrruTK0qAD0SaW7K/FtbW7xbHB8fy0ci+fj7+8uvS3Nzs7TAwEDZ7u7uLqqrq0VHRwcfbhFWE4BKWbp7NPdd+4ydnR1thdxn9FTQA6sJYFJxiYKCAt4t2dvb+0egdxnlEj2wmgANDQ3axbe1tfFuCSXFpaUlsbi4qBm9J6OkSWM9PT3lI3J9fZ0PtwirCZCXl6cJQOXtvyUqKkqO9fb2lrWEXlhNgIyMDE2AjY0N3n0vau/g5uYmDg8PebfFWE0A2u7S0qVN1vn5Oe++l4WFBTmeTqloe6wX5gTQdTdo65gTIJka9DoPsHXMnQdEU0NPTw/3NSQmm653lQABAJ5XVVVxX0PS1NREwf8F4C0lALGYlJTEfQ1JWloaCfATACdTAT53dHSUBYiR2dzcFK6uriTAM9PgCS8Av9GZuZHJzMyk4E/u+v8A/WYmKisr+ThDUFdXp5JfEQ/clG/IyWgJ0ST4bh4whxLDt+RMy2V5eZnPZVesra2JrKwsFfx3AB7xgO+Cfiz9lRIjndHV19eLoaEhMTs7KyspWzU6bqc/dTQ2Nor09HTh7OxMgR9SkucBPgR/AOUAfgBw9UJFezF6zq8BqAIQyAOzBCoaqHKi8pFqaFs1uj4q7UN5AOb4G1XJkqw1AqGEAAAAAElFTkSuQmCC",
  75: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAS+SURBVGhD7ZpZSHVVFMf/jgRGGFiJIWKO+FBETqhZ2uCET2KKT2HggzP0kKLgiOWEYigoiT4oiSiC8/Q5P+SQKaZCkmAKVkRhOIuuWIe7D/fs735fcvK76a0fLISz1j5nr/85Z5219xX4n1vzEoDXAbwN4H0A791D43nx/N4A8LKcgB4cAGQAeATgNwD0gOx3ADMAsgG8ICd2G5IB/MAnCwgIoLy8POro6KCBgQGamJig8fHxe2c8r8HBQers7KSCggIKCQkRYvwI4GM5wafxOQ8MDQ2lkZEReshMTU1RRESEEOJLAFZysjJfcHBWVhZdX1/L53uw5OfnCxGa5ISN+Ugkb4kUFhYKET6RE2eeB/BTYGCgPM6iiIyMZAF+AfCiLEAmqzM9PS2PsSiWl5fJxsaGRfhMFmDe399fjrdIDEXxO+OCyE3Dn/yO/BeoqqpiAS4BuAkB3uTHv6urS47VRWpqKvHTFBQUZNK4znBvsbq6qhnHxdfUOI4PDw+nw8NDTbxehoeHRTHkrlHhXT4wOjoqx+rC09NTXOCpNjY2po65uLggZ2fnx2KMbXd3V3MdvczMzIhzfigE4B5a6aTugsTERPLx8dGYr6+v8lckw8nu7e2pY3Z2dsjW1lbxOTg4kJubG7m6uqrm5eVFBwcHmuvoZXJyUszjAyEALySUdvJZkpOTo1yYq/Dc3JzG19/fr4pTUVFBl5eXylNxenpKx8fHdHJyQjc3N5oxevlXBOBziwSLiopkN5WXlys+KysrWlxcVI49q07U7AKcnZ2Rt7e3clF+lPnOyvBrw347OzuKi4tTXhl3d3eKjY298/WI2QVobGxU7357e7vsVu60n5+fGmPKmpqa5GG6MasAfLc9PDzUu39+fi6H0P7+vujOlEKYnZ1Nra2tlJycrApgbW1NW1tb8lBdmFWAvr4+NYnKykrZrcCvCH+aWlpaaGFhQeNLSkpSx/OexF1gVgHEu813eHt7W3b/Lb29vaoA8fHxslsXZhPg6OiInJyclIvdZp1xdXX1WOUfGhpSBYiJidH49GI2AYw6LsrMzJTdKj09PRQcHKzUirKyMo3P0LcrlpGRofHpxWwC1NbWqpNva2uT3SpGEyJHR0eanZ1Vmp61tTVycXFRfXJ90IvZBEhLS1MnPz8/L7s1iFrBxs0Qt8329vbqsfT0dHmIbswmADc0IoGNjQ3ZrYHb3ZSUFDVeGH8Wc3Nz76wNZswmAC93eYXJiyzu6W/DysoKNTQ0UElJCTU3N9Pm5qYc8o8xJcCdrgbvO6YEiOADd7UfcN8xtR/wFh/o7u6WYy0SvtEGAd4RAjgDOC4tLZVjLZL6+npO/grAa0IA5puwsDA51iKJjo5mAb4HYGMswKe84hKbEZYKf1UM/UWJcfKMI4Cfec/ckuEFFYA/nvT/A/ybGRUXF8vjLILq6mpR/HLkxI1p4SBLK4hGyXfJCctwYfiKg/lxWVpaks/1oFhfX6eEhASR/NcAnpMTfhL8Y+khF8aoqCiqqalR1ua8UuNO6r4ab7fzJmpdXZ2yoWr4reFXLvJygrfhFQCFAL4FcGFQ8aEYf+fXAZQCeFVOTA/cNHDnxO0j99D31Xh+3Np7yQmY4i+awIbNn3kDawAAAABJRU5ErkJggg==",
  76: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAATXSURBVGhD7Zp7KPZnGMe/DmmhtZWNrKg5prysORRmYzMiUSL/KFkRY2JqRDmkNrxCsj/WikIO//jD2WuOxdAsvaTG/jCyQ8teksMcrnX9cv96fvfrNT3Ps2ePZ/vUHX7Xdd/37/r+fs91X/f9AP7n3rwG4BGAdwB8AOB9M2x8X3x/AQBelwPQBwcAHwP4BsDvAOgBtQMAMwA+AfCyHNh9SAfwAw8WHBxMpaWl1NXVRYODg/TkyROamJgwu8b3NTQ0RN3d3VReXk5hYWFCjB8BZMoB3sXn3DE8PJxGR0fpITM1NUVRUVFCiFYAVnKwMl+wc0FBAV1dXcnjPVjKysqECF/KAeuSJoK3RCoqKoQIH8mBM44AfgoJCZH7WRTR0dEswK8AXpUFyGd1pqen5T4WxcrKCtnY2LAIn8kCzAcFBcn+FslNUvxeNyFy0XDEnxFzgJPv3t4e7e7uyiajUF9fzwL8CcBdCPAWv/69vb2yr15kZWURv02hoaG3Ns4zXFusrq5q+l1fX1NjYyP5+/uTo6MjOTg4UGBgILW3t2v8DGVkZEQkQ64aFd7jC2NjY7KvXnh6eooJ7mzj4+Nqn4uLC0pOTn7OR7SWlhbNHIYwMzMjxv1QCMA1tFJJGYPU1FTy8fHRNF9fX+WnCMjFxYV2dnbUPjevpdL8/Pyoo6ODampqRMIiKysr2t/f18yjL5OTk2KuGCEAbySUcvKfpLCwUJmYg5qbm1OvHx8fk6urq2Kzt7en7e1t1Zabm0tubm4UEBBAS0tL6nVD+FcE4LHFE66srNTYdF5JZZ2Wuby8lC8ZhMkFOD09JW9vb2VSLy8vOj8/19hbW1tVAbhk5eSYl5dHGRkZiu3o6EjjbygmF6CtrU0NkD/bMrxzE3bO+uJ30Tw8PGh9fV3upjcmFYCfNgcgnv7Z2ZnsQkVFRZqAnZ2dqbi4mBITE9VrnBhPTk7krnphUgEGBgbUIOrq6mSzQklJierDa7/u005PT1dtxqpTTCoAL4k8Nmf+zc1N2axQXV2tBhkZGamxDQ8PqzZj7VJNJsDh4SE5OTkpk921z+js7FSDjImJ0dh4uRS2zMxMjU1fTCaA7vKWn58vm1U2NjaUQof9uEDiVUPQ09OjjsHHcsbAZAJwXS9u/u/qeZ3zO8rJyaGDgwPa2trSrArG2qqbTIDs7Gz15ufn52WzhoWFBbXs5cYrAVeF4u+kpCS5i96YTICEhAQ1gKdPn8rm5+ATZ3d3d7WPaGlpaUYthkwmAFd0vMPkTdZ91/Bnz55RX18f1dbWUlNTEy0uLsouBnObAEbdDZo7twkQxReMdR5g7tx2HvA2X+jv75d9LRJ+0DcCvCsEcAFwzAcQ/wWam5s5+AsAbwoBmG8jIiJkX4skLi6OBVgHYKMrwKfW1tZGO3UxV7jqtLOzYwGqdYNnXgHwC5+ZWzI3W+w/XvT/A/ydGVVVVcn9LIKGhgaR/ArlwHX5ip0sLSHqBN8rByzDieFrdubXZXl5WR7rQbG2tkYpKSki+B4AL8kBvwj+svRnToyxsbH0+PFj5XBidnZWqaTMtfH5Af9TB5fR8fHxZGtry4H/xkleDvA+OAOoAPAdgPMbFR9K43V+DUANgDfkwPSBiwaunLh85BraXBvfH5f2XnIAt/EX0wVvTlNzNsIAAAAASUVORK5CYII=",
  77: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAARBSURBVGhD7ZpLSB1nFMf/UREhsRqwjdBFoT43tgSfqKXVNlEUFyJKwE0xoEi1LrKIoqBx0/pCpdhFKCqiRty48P2Iz01VYhArglrBgLTGkGAIPhLkC+fDb7g9c2Mv13F679gfHC7MOfPNnP8398z5vnuB/3GYjwF8AeArAN8B+NYFje6L7u9LAJ/wBJzhKoAfADwG8AKAcCN7CWAawI8APuKJOcIdAOs0WHR0tCgtLRWdnZ2iv79fjI+Pi7GxMZczuq+BgQHR1dUlysvLRXx8vBLjTwDf8wTP4ic6MSEhQQwPDwt3ZnJyUiQlJSkhfgFwhSfL+ZmCi4uLxcnJCR/PbSkrK1Mi/MoTtiVHJW9FKioqlAh3eeLENQDPYmJi+HmWIjk5mQTYBXCdC1BE6kxNTfFzLMXi4qLw9PQkEe5zAeaioqJ4vCU5LYpPbQsiNQ2v6TtyGaitrSUB3gL4TAlwkx7/np4eHusUeXl5gp6m2NhYu0Z1hnqLpaUlGV9TUyMiIyN1cbZG/u7ubn4ppxgaGlLFkLpGyTd0YGRkhMc6RXBwsLrAmTY6Oirjc3NzdT57VlVVxS/lFNPT02rM20oA6qFlJ2UE2dnZIiws7B8WHh4uP1UygYGBYnt7W8ZTYjyeLCIiQhUs+UlNjRFMTEyo+7ilBKCFhGwnL5KSkhItmdnZWe7W0dHRoQnQ3NzM3U7znwhAY6vZr6ys5G4dOzs7wt/fX8anpaVx97kwXYDDw0MRGhoqLxoSEiKOj495iI6cnBwZ7+PjIzY2Nrj7XJguQEtLizb77e3t3K1jfn5eiy8oKODuc2OqADTbQUFB2uwfHR3xEB1ZWVna7K+vr3P3uTFVgL6+Pm026X3/b2xtbQlvb28Zn5mZyd2GYKoA9Eqksamar62tcbeO+vp6TTDa2LgITBNgf39fBAQEyIs5us5Qmxe+vr5id3eXuw3BNAFsOi5RVFTE3Tr29vaEn5+fjI+Li+NuwzBNgIaGBk2AtrY27tYxMzOjxRcWFnK3YZgmQH5+vpbQ3Nwcd+tobW3V4hsbG7nbMEwTID09XUtoZWWFu3VUV1dr8Y70C85imgC03KUVJi2yDg4OuFvH5uamjKdV4kUVQMKeAIauBl0dewIk0QGj9gNcHXv7AZF0oLe3l8daEproUwG+VgIEAnhDRegy0NTURMm/A/C5EoD4PTExkcdaktTUVBLgDwCetgLc8/DwkEtRK7O6uqoWWw9skyf8AfxNvbiVycjIoORffej/A/SbmWG7r65GXV2dKn4lPHFbHlKQ1QqiTfI9PGEOFYbfKJgel4WFBT6WW7G8vKztMAF4BMCHJ/wh6MfSv6gwpqSkyI2KwcFBuWKjTspVjbbb6U8dtJii3WQvLy9K/DkVeZ6gI9wAUAHgCYDjUxXdxeg9vwygGsCnPDFnoKaBOidqH6mHdlWj+6PWPoQnYI/3u2bW0ZyiPVgAAAAASUVORK5CYII=",
  78: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAATiSURBVGhD7ZppSHVFGMcfFzQoc8FKCAncPxmRCy5ZWqa4ixiJX9JQkDTFPqS4JmK5K2lKhAouiSJ+cF9yRUlFxQ0hCLEvpWZiiEuiE8/gDPfMexU53k73nvrB8MI8z5xz/v85d84z4wvwP4/mJQBwB4C3AOA9AHhXDxs+Fz7f6wDwsihADs8DwCcA8AMA/A4AxIDaHwAwAwCfAsCLorDH8CEA/IQX8/T0JDk5OaSjo4MMDAyQiYkJMj4+rncNn2twcJB0dnaSvLw84uvry8z4GQA+EgU+xJc40M/Pj4yMjBBDZmpqigQGBjIjvgYAI1GsyFeYnJGRQW5ubsTrGSy5ubnMhG9EwZp8wMSrkfz8fGbCx6Jw5AUA+MXLy0scpyqCgoLQgAMAsBYNSEd3pqenxTGqYmVlhZiYmKAJn4sGzHt4eIj5quRuUVzXXBCxaPgTfyP6wMHBAdnb2yMnJydiSCdUVFSgAX8BwGvMgDfw9e/u7hZzZZGcnEzwbfL29tbacJ3B2mJtbU0yrquri/j4+BBra2tiZmZG7OzsSHR0NFlfX5fkPZXh4WG2GGLVSHkHO0ZHR8VcWTg5ObEbPNjGxsb4mIaGhmfirFlYWOjUhJmZGXbt95kBWEPTSkoXxMfHE1dXV0lzc3Oj/zJROLv7+/s0//j4mFhZWdF+nPmSkhJacSYmJvL8iIgI8TaymZycZNcNZgbgRoKWk/8kmZmZ9Ma4Cs/NzfH+hYUFLjQhIUEyxsHBgfbb29uT6+trSUwu/4oBeG0msqioSBLb3NzksZSUFEkM3xzsd3Fx0VllqrgBFxcXVADew9nZmVxdXUniOLMBAQE0bmlpSRfj3d1dUlBQwI0pLS2VjHkKihvQ2NjIhbS1tYlhytHREYmKiuJ5rJmamtI6XpcoagDOtqOjI5/9y8tLMYVyeHhIYmNjnzEA14vs7Gz6FukKRQ3o7+/nYsrLy8Uw5ezsjLi7u/MZLywsJH19fSQpKYmPjYmJIbe3t+JQWShqAH4S2Uzi71ob7e3tXGhWVpYkFhwczGOLi4uSmFwUM+D09JTY2trSmz20z8DtNxMpHsCUlZXxWHNzsyQmF8UM0Ki4SHp6uhjm4KyzvKamJkksLS2Nx1paWiQxuShmQHV1NX/41tZWMczp7e3lebhgYl2Av3cslmxsbHhse3tbHCoLxQxITU3lDz8/Py+GOfil0DjAJObm5rRsNjY25n24wdIVihkQHh7OBWxtbYlhCfgZ1MxnzcjIiIq/7/MpB8UMwO0u7jBxk3V+fi6GtYIrfU1NDT2/q6+vJ6urq2LKk9FmgE53g/qONgMCsUNX5wH6jrbzgDexo6enR8xVJTjRdwa8zQywA4AzPIj4L1BXV4firwHAgRmA/Ojv7y/mqpLQ0FA0YBsATDQN+Ay/u0tLS2K+qtjZ2aHHbgDwhaZ4xAoAfsMzczUTGRmJ4k/u+/8D+DczUlxcLI5TBZWVlWzxyxSFa/ItJqltQdQQ3y0KFsGF4TtMxtdleXlZvJZBsbGxQeLi4pj47wHgOVHwfeAfS3/FhTEkJIRUVVWRoaEhMjs7SyspfW24g8QzhdraWhIWFkZPmADgEBd5UeBjeAUA8gFgFQCu7lw0lIbf+Q0AKAGAV0VhcsCiASsnLB+xhtbXhs+Hpb2zKEAbfwNqfGqSozMbrwAAAABJRU5ErkJggg==",
  79: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAATSSURBVGhD7ZpZSG1VGMf/OREYWqIm9JCz4oOXSw6gpmmDogiCmBcECQN9SPOhh5xA04dyQiHsIUMFFcUXEech54dUUkRFSJwSzERSw7HUFd/GtThn3ePN9t3tjuf2g4XnfOtbe+/vv/f+1rfWEfife+MCIBDA2wDeA/CuGTa6Lrq+RwBc5QDUYA/gEwDfAzgAwB5Q+w3AOIBPATjIgd2HJwB+ooMFBwezvLw81tLSwrq7u9nw8DAbGhoyu0bX1dPTw1pbW1lhYSELCwvjYqwD+EgO8Fl8SQPDw8NZf38/e8iMjo6y6OhoLsTXAF6Sg5X5ipxzcnLY9fW1fLwHS35+PhfhGzlgQz7kwVsiRUVFXISP5cCJVwD8HBISIo+zKGJiYkiAXwG8JguQTeqMjY3JYyyKubk5Zm1tTSJ8LgswFRQUJPtbJLdJccEwIVLR8Du9I+bA7u4u29nZYVdXV3KXJlRUVJAAfwB4kwvwmB7/9vZ22VcVGRkZjJ6m0NBQk43yDNUW8/PzRuMaGxsVu4ODA7O3t2cBAQGsqqrKyEcL+vr6eDKkqlHhHTIMDAzIvqrw9vbmJ3hmGxwcFGMKCgqe6uctLS2N3dzcGJ3jeRgfH+fH/oALQDW0UklpQUpKCvPz8zNq/v7+yl8elJubG9ve3lb8p6enhd3JyUm56/SYOjo6Cnt9fb18GtWMjIzw477PBaCFhFJO/pvk5uYqJ6YsPDk5KexZWVkiUCq3OW1tbcIeGBioWU74TwSgY/NgiouLjfqioqKEMJubm8J+enrKXF1dlT4rKyu2trZmNE4tugtwfn7OfH19lZP6+Piwy8tLo/7IyEghwNbWlrDTHff09BTCabUu0V2Auro6EURTU5PczdLT00U/ZWgOTYU0G/A+WuVpga4C0N328vISd//i4kJ2UWYfHiRNfVNTU2x5eZnFxcUJO7WGhgZ5qCp0FaCzs1MEUF5eLncLqHYwDJY3Ozs78bm5uVkepgpdBaApkY5N7/fq6qrcbURZWZnyztva2irJr7S0lCUlJQkBurq65CGq0E2A4+Nj5uzsrJzsn6wzNjY22NHRkfLZYFeHLS0tya6q0E0Ag4qLZWdny90Cet+p0KG1yMLCgrDv7e2JJOjh4WEyf6hBNwGqq6uFAFTn3wUVP9wvIiKC7e/vs4ODA5aamirstMenFboJkJmZKQKgzH4XJycnyh3mvvTauLi4iO80exweHsrDVKObAAkJCSKIv3t/Z2dnmbu7u/DnjYqk9fV12f250E0AWu7SHE+LrLOzM7n7KeguU7FDpTIthrRanMmYEkDT1aC5Y0qAaDJotR9g7pjaD3iLDB0dHbKvRWJQekdxAdwAnFDl9SJQW1tLwf8JwJMLQPxAc/CLwO0iaxmAtaEAn9Gmw8zMjOxvUaysrPAF1heGwROvAtijPXNLJjExkYI/vOv/B+g3M1ZSUiKPswgqKyt58suVAzfkW3KytIRoEHy7HLAMJYbvyJkeFypRHzKLi4ssOTmZB98G4GU54LugH0t/ocQYGxur7NX39vayiYkJpZIy10bb7bR5WlNTw+Lj45mNjQ0Fvk9JXg7wPrwOoAjAjwAub1V8KI3m+UUApQDekANTAxUNVDlR+Ug1tLk2uj4q7X3kAEzxF9bGbuL7hyd8AAAAAElFTkSuQmCC",
  80: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAVsSURBVGhD7Zp9LJ9XFMePqpfESy2KybBE7I+mQlAa1cXYqDSkadou63+TRfxhNjIJbQXTitcEmZAIEQ2dIEHUS1pTLNJ560gmLY0G8bpguhJhbc9ybtyb57n9WeT30vCzT3KTn3POc597vr/nOc+5zw/A/+wbBwDwAoBPAeALAPj8AA5aF63PGwAc5QS0wQoA4gDgFwBYAQA8RGMNAHoA4DsAsJUT2w9fAcAkTebv748pKSlYU1ODra2t+PDhQ3zw4MGBG7Su+/fvY21tLd66dQvPnTvHxZgCgK/lBP+LbDowKCgIOzo68DDT3d2NISEhXIifAMBETlYmh4Lj4+PxzZs38nyHlhs3bnARSuWElXzJkzdGUlNTuQjfyIkT1gAwGxAQIB9nVISGhpIAywDwgSzAt6TOo0eP5GOMiqGhITQ1NSURkmUBfj1z5owcb5TsFsXflQWRmoa/6R7RN2/fvsW5uTmcmZnBzc1N2a2RlZUVnJ6exlevXskuvZCXl0cC7ADAx1wAH7r86+rq5FidyM3NRU9PT7S2tkZLS0t0dXXFmJgYXF5elkMZz549w8uXL+PJkydZvIuLCyvIL1++lEN1or29nRdD6hoZn5Ghs7NTjtWa6OhofpJ3xunTp3F1dVUV/+LFC3RycnonlkZwcDDu7Oyo4nWhp6eHzx3OBaAemnVS+uDx48di8Y6OjlhRUYEtLS3Kzgzv3LmjOubatWvCd/XqVSwvL0cvLy9hKykpUcXrQldXF583jAtAGwnWTuqDoqIisfDi4mJhf/78OZqZmTF7ZGSksM/Pz7NLnux0m7x+/ZrZh4eHecVGPz8/Vk/0gcEFqKqqEgLQZ87i4iKam5sz+/Xr14Wd9hc8nq4EDiXs7u7O7CQQCaUPDC4AJWpvb89OcurUKdZbjI2NqS5z5R6joKBA2BMTE1Vz0X6E+/r7+1U+bTG4AMTo6Ch6e3uLxfNhZ2eHd+/eVcUq+nRMS0tT+cLCwoSPdnr64L0I8OTJE/T19dUogFzQEhIShD89PV3lUwrQ2Nio8mmLwQWYmJhAGxsbdhK6FcrKyrChoUG5LcWsrCwRn5SUtKcA4eHhwtfc3KzyaYvBBYiLixOLvnfvnrBvbGywKk/2EydOiF6AkubxN2/eVMyErAfgPn09pg0uQGBgIDsBPfIWFhZUvkuXLomEaGNCVFZWCltsbKwqnjpJsh87dgyfPn2q8mmLwQU4f/48O4GJiQmOjIyofPR6jSdLTwZiYGBA2M6ePSti19bW0NbWltnd3Nxwa2tLMZP2GFwA5T1N9zBdBdTcUHdH3yTZqc/nm6Pt7W3xvCd/fX09rq+vY3JysphHvjJ0weACzM7OooODg1g8FUIPDw/xNw0qjEqUzRMNZ2dn8dnCwgInJydV8bpgcAEIamN9fHxUSdGwsrLC7OxsOZxx+/ZtlqwynjZITU1NcqhOvBcBCNrBUceXk5PDGhzaFE1NTclhKqjQlZaWYmZmJlZXV+PS0pIcojOaBNDrbvCgo0mAEDLo833AQUbT+wA/MlD1PQrQF70rQDAX4EMA2KD77iiw+77iHwBw5wIQv1EDcxSIiIggAf4AAFOlAD9QE0JdmTEzPj7OX8r8qEyesAOAJdqxGTNRUVGU/F97/f8A/WaGGRkZ8nFGQX5+Pi9+38uJKymnIGMriIrk6+SEZagwVFAwXS6Dg4PyXIcK2nFeuXKFJ/8zAFjKCe8F/Vi6SIXxwoUL7MVlW1sb9vb2sk7qoI6+vj7WfhcWFuLFixfx+PHjlPifVOTlBPeDEwCkAsAIAGzvqnhYBj3nxwAgEwA+khPTBmoaqHOi9pF66IM6aH3U2n8iJ6CJfwG7zgqM6iac5AAAAABJRU5ErkJggg==",
  81: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAASfSURBVGhD7ZpZSHVVFMf/jgiO4VCSIoqBE4kjOERp5YSiIEa+JeGTWQ8JKQ6YKI7wKUmCCIKoieKDOKJ+jg/mmEKCJj4oohlaIPqgqSvWwX24d3/XEr8b3HPrBxu8+6x9zvn/7z7rrL2vwP88GXcA7wJ4D8BHAD40wcb3xfcXCsBDFvAc7AEUAHgJ4AwAaaj9DmAOwJcAnGRhT+FTAL/wyaKioqi4uJi6u7tpeHiYpqamaHJy0uQa39fIyAj19PRQaWkpxcbGCjP2AXwmC/w7anlgXFwcjY+Pk5aZmZmhhIQEYcR3ACxksTJ1HFxYWEh3d3fy+TRLSUmJMOF7WbAunwjx5khZWZkw4XNZOOMA4DA6OloeZ1YkJiayAacA3pAN+ILdmZ2dlceYFaurq2RlZcUmfCMbsBgZGSnHmyUPSfEn3YTIRcMFPyPG5v7+no6Ojujg4ICurq7kw49ycXFBOzs7tL+/r5zDmDQ0NLABNwB8hAFhPP37+vrk2Neivr6eQkJCyMHBgezs7Mjb25vy8/Pp9PRUDn2F1NRUJWH5+fnRzc2NfPi1GBsbE8mQq0aFD7hjYmJCjn02eXl54iKvtODgYDo/P5eHqLS0tKixPj4+Rjdgbm5OnD9JGMA1tFJJGYOlpSVVgIeHB3V0dNDQ0JBuZUbV1dXyMLq9vaXy8nI9s/6NGTA9PS3O/7EwgBcSSjlpDJqbm1UB/G0K9vb2yMbGRulPT0/XGzM4OEgRERF64jVrQGdnpyqA/xacnJyQra2t0p+bm6s3JigoSB2TlZVFbm5u2jWAhbq6uioXCQwMVGqLra0tysnJUUXqrjE4ywcEBChiORFfXl6q4zVpALO5uUmhoaGqYNFcXFyoq6tLL5af/fn5eeXVx/BbwsnJSdsGbGxsUHh4uEEDWltb5XA9eAY5Oztr14Dd3V1ydHRULsJTua2tjQYGBnSXpVRTUyMPU9G8AQUFBarQ3t5etZ+fbS6GuJ8Fnp2d6Y0TaN6AmJgY5QL8yjs+PtY7lpmZqZrDCxNDaN6A+Ph45QIWFha0vr6ud4y314QB/GYwhOYNKCoqUkUmJSUps4AzfXt7O1laWir9Xl5ejy6ONG/A4eEhubu7qyZwIvT391c/c+PE+BhsgEiivr6+2jOAWVtbo7CwMD3R3Ozt7am2tlYO14NnjKgYPT09tWkAwzfOFV9dXR1VVFQoiyJe3/8T19fXym4ur04XFxeNvh9gyACjrgZNHUMGJHCHMfcDTBlD+wER3NHf3y/HmiX8RT8Y8L4w4C0Al1VVVXKsWfKwX/EnAD9hAPMjFzD/BVJSUtiAnwFY6RrwNRcpy8vLcrxZsb29LV6x3+qKZ1wA/MorNnMmIyODxf/x2P8P8G9mVFlZKY8zCxobG0Xy+0oWrks7B5lbQtQR3ycLluHE0MHBPF1WVlbkc2kKXnFmZ2cL8T8AsJMFPwb/WHrCiTE5OZmamppodHRU2bvjSspU28LCglJ+v3jxgtLS0sja2pqF/8ZJXhb4FN4EUAZgHcD1g4taafye3wJQBeBtWdhz4KKBKycuH7mGNtXG98el/TuyAEP8BZ3qe9Wg582SAAAAAElFTkSuQmCC",
  82: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUdSURBVGhD7ZprKP5nGMcvf8fCGLHVaLVQojkMOWZsDom8kL/xQtPyCluZGpGznE9NFKKEObwh58Oc9sIcR40yJBKzsFq8mOFa1y/3r99zeyx7PM9/j8c+dRX3fd2/+76+fs99X9f9APifR2MBAB8CgB8AfAoAn6ih0bpofU4AYMkHoAiGAJAEAN8DwCkA4DOycwCYAYAvAeAtPrDH8BkA/EIPc3d3x/T0dGxvb8eBgQGcmJjA8fFxtTNa1+DgIHZ0dGBmZiZ6e3szMXYB4HM+wH+imAb6+PjgyMgIPmempqYwICCACfEtAGjxwfKUkHNKSgre3Nzwz3u2ZGRkMBHq+YClvGbBayJZWVlMhC/4wAkjADjw8PDgx2kUgYGBJMAJALzNC5BM6kxPT/NjNIqlpSXU1tYmEb7hBfjBzc2N99dI7jbFn6QbIiUNf9BnRNnc3t7i4eEh7u/v4+XlJd8tl5OTE9zb28PT01O+SymUlZWRAFcA8D4TwIVe/66uLt73SZSWlqKjoyMaGRmhgYEBWltbY2JiohCgPHp6etDf3x/NzMxQX18fLSwsMCgoCEdHR3nXJzE8PMw2Q8oaBT6mBmVOlJCQwCa5Zw4ODnh2dibjT2LxflLr7OyU8X8KMzMz7LnBTADKoYVMShnMz8+LC7e0tMTm5mbs7++XZmZYWFgo+u/s7OCrV6+Edi0tLUxOTsaWlhaMiYkR/c3NzfH8/FxmHkWZnJxkzw1iAlAhIaSTyqCmpkZceG1trdi+vb2Nurq6Qnt4eLjYXlBQIPrHxcWJ7YSnp6fYNzY2JtOnKCoXoLW1VVw0/cw4Pj5GPT09oT02NlZsp3lTU1OFNno9pcTHx4vP6uvrk+lTFJULQIHSK0vPtLe3F3KL9fV1jI6OFoN5TI1BqbidnZ04ZnNzk3dRCJULQKytraGTk5O4eGampqbY1tbGu8uFKjo2zs/PTzhSlcEbEWB1dRVdXV3lClBXV8e736OkpEQcQxvk4uIi76IwKhdga2sLjY2NhUnoo9DQ0IC9vb3SshSLior4YSLSTZGsvr6ed3kSKhcgKSlJXLz0/L64uBCSIWo3MTGRm+lJKjbBKisreZcno3IBvLy8hAnoyDs6OpLpi4yMFIOjwkRKfn6+TPBNTU0y/cpC5QL4+voKE1BSs7KyItNH12ssQDoZGHTdxtrp40PpqqpQuQBpaWliMMHBwcJbcH19jY2NjWLGZ2VlJRZHVBvQR4KNsbW1xaqqKszLy8Pc3FzBcnJyns8xeHBwIBQyLCDaCG1sbMTfyWhjZPCv/kNGl53KQOUCEMvLy+ji4nIvCENDQywuLhb96Gx3dna+5yfPuru7ZeZQlDciAHF1dSVkfHSmZ2dnC0XR7u6ujA8JMDs7K+T57HpbnlH/Q2X0v0WeAEqtBtUdeQIEUIMy7wPUGXn3AR9RA93IvAToD30ngD8T4F0AuKDd+CVwd1/xFwB8wAQgfqQE5iUQGhpKAvwMANpSAb6mJGVhYYH31yg2NjbYpUyeNHjCFAB+pYpNk4mIiKDgf3/o/wfoOzMh9dREysvL2eb3FR+4lEZy0rQNURJ8Fx8wD20MzeRMr4syb2L+C6jijIqKYsF/BwAGfMAPQV+WHtPGGBISghUVFTg0NCSkrZRJqavNzc0J6Xd1dTWGhYWhjo4OBf4bbfJ8gI/hHQDIAoAVAPjzTsXnYnTOrwNAPgC8xwemCJQ0UOZE6SPl0OpqtD5K7W35AOTxN6kRH/UzQ/TeAAAAAElFTkSuQmCC",
  83: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAVRSURBVGhD7ZpZSF5HFMePS4KiMYpohSpCsaBEKloVkihWbY2Imgex1AehUnyy1oeKC0qSiruCSqUBoxjEjeiLe1DrDta1CvUhdRfRWrUmxRC06ilncIbvTr4E8y3h82t/MKBz/3funP+9d+acqwD/c2EcAOATAAgEgM8BINQAG82L5ucFAI5yAJpgBQBJAPAzAOwBAF6i9hcADAHAdwBgIwd2Eb4CgN9pMD8/P8zIyMD6+nrs6OjAvr4+7O3tNbhG8+rs7MSGhgbMysrCW7ducTOWAeBrOcC3UUAn3r59G3t6evAyMzAwgMHBwdyIHwHARA5WppDEycnJeHp6Ko93acnMzOQm/CQHrMqXPHhjJDs7m5vwjRw4YQ0AG/7+/vJ5RkVISAgZsAMAdrIB35I7g4OD8jlGxdTUFJqZmZEJ6bIBo76+vrLeKDlfFH9VXRApafib3hFdc3Z2hpubm7i+vo4vX76UD6tlZ2fnnfTvSnFxMRlwDACu3ABvevybm5tlrVYUFRWhp6cnWltbo4WFBbq4uGBiYiILUB2tra0YGBiIdnZ2aGlpia6urpiSkoLPnz+XpVrR3d3NF0PKGhmfUcfTp09lrcYkJCTwi7zWbty4gfv7+wr9w4cPX9PxRgvzixcvFHptGBoa4mOHcQMoh2aZlC4YHx8Xk3d0dMTq6mpsa2tTzcwwNzdX6Le3t9kdp356UnJycrCmpgbd3d2F/v79+4praEN/fz8f9wtuABUSLJ3UBeXl5WLiFRUVon9xcRGvXLnC+iMjI0U/pdfm5uasn1JuzsjIiBiHXg1doXcDamtrxcTpZw7d6atXr7L+uLg40f/q1StcW1tjE1NdH8bGxsQ4UVFRol9b9G4ABWpvb88u4uHhwXKL+fl5jI2NFQG9rcY4ODjA0dFRVoRxPRU5ukLvBhBzc3Po5eUlAuDN1tYW6+rqZLmC+Ph4hb6pqUmWaMV7MWB2dhZ9fHzUGlBZWSnLBZQ30NbJ9TY2NpiWlqbTXUDvBjx79gyvXbvGLkKvAm1xLS0tqmUp5uXlyacxTk5O2OM/MTHBAuf6oKAgPDo6kuUaoXcDkpKSxMQbGxtF/+HhIUuGqP/69eu4t7enOE8doaGhYqz29nb5sEbo3YCbN2+yC9CWt7W1pTh29+5dERAVJgQ99mTOysoKHh8fK/Tp6elCn5+frzimKXo3ICAggF3AxMQEZ2ZmFMdUV3baGQhKcpydndkW+eTJE4U+Ojpa6KuqqhTHNEXvBqSmpopJh4WFsaeA3m0KwNTUlPVTwLzYefz4sdBT9kc7COX/jx49EgkSlbBLS0vypTRC7wZsbGygg4ODCIoWQjc3N/E7NVoYOfTYU6bHj1HQTk5OCj2lx7pC7wYQ09PT6O3trQiCmpWVFRYUFMhy3N3dVTzuvNFuolo36IL3YgBBd5YyvsLCQrx37x4ripaXl2WZAiqkqJagO05p9OrqqizRGnUG6LQaNHTUGRBMHbr8HmDIqPse8Cl1yFuQsUI3+tyAIG6AEwAc6nKlNWTOv1f8AwAfcQOIXyiB+S8QHh5OBvwGAGaqBnxPSQoVIcbMwsIC/yjzg2rwhC0A/EEVmzFDX5YA4OBN/z9AfzPDBw8eyOcZBSUlJXzxS5EDV6WKRMa2IKoE3ywHLEMLQzWJ6XGZnJyUx7pUUMUZExPDg28CAAs54DdBfyzdpoXxzp07WFpail1dXTg8PMwyKUNt9Cmd0u+ysjKMiIjg1eSftMjLAV6EDwAgGwBmAODo3MXL0mifnweAHAD4UA5MEyhpoMyJ0kfKoQ210fwotf9YDkAd/wJUhTWOwKSFeQAAAABJRU5ErkJggg==",
  84: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUOSURBVGhD7ZpZLK1XFMeXMWIoNSeGJnWRG2ImMTQtbREhHkRTT7fSeFLtQyUIERViTJBSiUgkgorhQYxBjQ+m6xYpoSJBxFgVRaSm1awd+8s5+x736uk5jXPaX7ISZ++197fX/+xvfWt/B8D/PBobAPACgA8A4BMA+PgJGq2L1ucNALZiAMpgAgCpAPATAPwGAKhB9jsAjAHA1wDwjhjYY/gcAH6lyQIDAzEzMxObmpqwu7sbh4aGcHBw8MkZraunpwebm5sxOzsbQ0JCuBgbAPCFGOCbKKKBoaGh2N/fj5rMyMgIhoeHcyG+BwAdMViRYnJOS0vD29tbcT6NJSsri4vwgxiwLJ/x4LWRnJwcLsKXYuCEKQBsBwUFieO0ioiICBLgAADeFQX4itQZHR0Vx2gVc3NzqKenRyJkiAJMBgQEiP5ayX1S/Fk2IVLR8AfdI6rm7u4Od3Z2cGtrCy8uLsTuN3J6eoqrq6u4vr6ONzc3YrfSlJaWkgBXAPAeF8CXtn9ra6vo+48oKSlBT09PNDU1RSMjI3RycsKUlBQ8ODgQXV/j6uqK1R+0LisrKzw6OhJdlKavr48nQ6oaGR9Rw8DAgOirNMnJyfwir5mHhwceHx+LQ+TIyMiQ/C0sLFQqwNjYGJ87kgtANTSrpFTB1NSUtHhbW1usr6/Hrq4u2coMCwoKxGESk5OTqKOjI/mqegcMDw/zuT/lAtBBgpWTqqCyslJafFVVldRO97KBgQFrj42NlRvDoTzh5uYmjddIARoaGqTF09+cvb09NDQ0ZO1JSUlyYzipqams38zMDG1sbDRTAAqUFk1zPn/+nNUWi4uLmJiYKAmj6IxBOYj319bW8qJF8wQgFhYW0NvbWwqIGyW0xsZG0R1PTk7Q2dmZ+URFRbE2f39/zRXg1atX6Ofnp1CA6upq0R1fvHjB+mnrb25usjYvLy/NFGBtbY0FwhdP27m9vV32WIqFhYWSf2dnJ2ujBNnR0SG18x1gbW2t0pOp2gXgiYyspaVFaj8/P2fFELWbm5vj2dkZM3t7e9ZmZ2eHNTU17MlB5uDgwNpNTEzYY1PRraMMahcgODhY+kZ3d3fl+uLj4yVxVlZWcH9/X/r8NqOngipQuwBhYWHsAlTMzM/Py/Xx8paM6oK/I4CLi4vcXMqidgHS09OlRUdGRrJdQIeZuro61NXVZe2Ojo54eXmJ19fXODMzg9PT05LRZzJXV1fmS7cLPSKXlpbESymF2gXY3t6WihgySoTPnj2T+zYpMb4NHx8f5mtpacnEUhVqF4B4+fIl+vr6ygVNRgmtqKhIdFeIu7s7G2NsbIyHh4dit9L8KwIQdKSliq+4uBhzc3PZoWhjY0N0exA6VNHWp0qS5lIVigRQ6WnwqaNIgHBqUOX7gKeMovcB/tTQ1tYm+molMoeuD7kA9gBwnp+fL/pqJffvK64B4H0uADFNBcx/gejoaBLgFwDQkxXgWypSqADRZpaXl/lLme9kgycsAGCfTmzaTFxcHAV/8tD/D9BvZpiXlyeO0wrKysp48vtGDFyWOnLStoQoE3yrGLAIJYZ6cqbtMjs7K86lUdC7yISEBB78jwBgJAb8EPRj6R4lRnpHV15ejr29vTg+Ps4qqadqExMTrPyuqKjAmJgY1NfXp8APKcmLAT4GOwDIAYB5APjzXkVNMXrOLwJAPgA4iIEpAxUNVDlR+Ug19FM1Wh+V9q5iAIr4C/XzJyGjcQ/3AAAAAElFTkSuQmCC",
  85: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAVhSURBVGhD7ZppSF1HFMeP0YigNhbRilXEqriFiqm71lbbukb8IFaTT5XiF61WqNCIIu47uFQUxSCIWnH5oMYFtWrSQGpcGqEJWFFwQWvRFooLrqec4d3LvfNMked7qb72Bwd8M2funfN/c8+cuU+A/7kwZgDwPgB8CACfAsAnV9BoXjQ/NwAw5wNQBUMASAaAHwBgGwDwGtkfADAJAKkA8BYf2EWIB4Bf6WKenp744MEDbG1txf7+fhwdHcWRkZErZzSvR48eYVtbG2ZmZqKfn58gxhIAfMEH+E8U00B/f38cGhrC68z4+DgGBQUJQnwHADp8sDwl5JySkoKnp6f89a4tGRkZggh1fMBSPheC10aysrIEEb7kAyeMAGDVy8uLH6dVBAcHkwBbAPA2L8BXpM7ExAQ/RquYnp5GXV1dEuFbXoAfPTw8eH+tRJEUf5YmRCoa/qJnRN2cnZ3h+vo6rqys4N7eHt/9r1BWVkYCHAGAjSCAOy3/jo4O3vdSlJaW4u3bt9HIyAgNDAzQ2toaExMTcWtri3dlUPKlVejt7S0zykuBgYG4ubnJD1GJwcFBIRlS1cj4mBqGh4d5X5VJSEgQbqJkrq6uuLOzI/M/PDxECwsLJV+pLS8vy8aoyuTkpHDNEEEAqqFZJaUOnj17Jk7a3Nwcm5qasLe3V1qZYUFBgWzM4uIi6unpsT5DQ0O0sbFhK0YwBwcH9iipg7GxMWEenwkC0EGClZPqoKqqSgy0urpabKcgb968ydrv3r0rG9PX1yeOKSoqwqOjI7Yq9vf3cXd3l+UPyifqQOMCNDc3i8HQ3wL0DOvr67P2e/fuycbQiqB2HR0dnJqaYm2aqkQ1LgAFampqym7i7OzMaov5+XmMjY0VheHPGEIfrZDIyEh0cnJCW1tbjIiIUPK9LBoXgHjx4gW6ubmJAQtmYmKCLS0tMl/6pl1cXJR8pVZXVycbcxneiABzc3N4584dpUBIgNraWpnv2tqaUJ2xRJiamooPHz7E+Ph4cdyNGzfw1atXsnGqonEBFhYW0NjYmN2EHoX6+nrs6uqSHkuxsLBQ9D84OGBbU2NjIz59+lR2rbi4OHEMvZNQBxoXIDk5WZx0e3u72E7ZnLY0ar916xZub2/Lxp1HT0+PeK2oqCi+WyU0LoCvry+7ASW0jY0NWV90dLQYEB1MpBwfHytl/oGBAdE/PDxc1qcqGhcgICCA3YC2tNnZWVkfvV4TAqKdgeju7kYfHx+0s7PD/Px8mb+ibmdGK0sdaFyA9PR0cdIhISFsFZycnLBnnJIZtVtZWYmHI8mEWJJ8/PgxK3poJ7G0tBT7+PygKhoXYHV1Fc3MzMSJUyK0t7cXP5NRYpQirRFo5Tg6OopFE1lSUpLM/zJoXABiZmYG3d3dZUGTUZ1fXFzMu7MEef/+fSV/2hbT0tLUVgYTb0QAgup5quJKSkowOzubHYqWlpZ4NxkkXE1NDebm5mJDQwO+fPmSd7k05wmg1tPgVec8AYKoQZ3vA64y570P+IAaOjs7eV+thL5ohQAfCQJYAMBuXl4e76uVKN5XHAPAe4IAxE9UwPwXCAsLIwF+AQBdqQDfUJEivIzQVmhXUdQXudLgCRMA+I1ObNoMHagA4M/X/f8A/WaGOTk5/DitoLy8XEh+X/OBS2kkJ21LiJLgO/iAeSgxNJEzLZfnz5/z17pW0IkzJiZGCP57ADDgA34d9GPpJiXG0NBQrKioYGdzOqlRJXVV7cmTJ6z8rqysZC9UFb81/E5Jng/wIrwDAFkAMAsAhwoVr4vRPj8PAHkA8C4fmCpQ0UCVE5WPVENfVaP5UWnvwAdwHn8DiDsaglrBOvgAAAAASUVORK5CYII=",
  86: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAVmSURBVGhD7ZpZSB5XFMePS4K4NBZxwbpAtQ9qrEtVcGla02pEFAVRzIOgFBG1tmgDVRRRSVCjIQmGCiJEXOL2JGoiat0KWtcq6INVH1zQWqoVcaFup5zLN8PM1RTxmy98fukP7oPnnjtzz39mzj33fgL8z6WxBIBPAeBzAPgaAL7Swkbzovl5AIAVH8BVMAGAdAD4GQD+AgC8Rm0bAAYA4DsA+IAP7DLEA8DvdDFfX1/Mzs7G+vp6bG9vx56eHuzu7ta6RvPq6OjAhoYGzM3NxYCAAEGMJQBI5AP8L4ppYGBgIL558wavM319fRgcHCwIUQEAenywPCXknJGRgaenp/z1ri05OTmCCD/xAUuJE4LXRfLy8gQRvuEDJ0wBYMXPz48fp1PcvXuXBNgEgA95Ab4ldfr7+/kxOsX4+DgaGBiQCD/yAvzi4+PD++skqqT4mzQhUtGwS9+I0pydneHa2houLy/j/v4+330hlHxpzOrqKt+lCI8fPyYBjgDAURDAi17/pqYm3lctSktL8fbt22hqaopGRkZob2+PycnJuLm5ybsySKwnT56gu7s7G2NiYoKenp748uVL3lUtXr9+LSRDqhoZX5Khq6uL970ySUlJwk3ONTc3N9za2pL5Hx8fY3R09DlfoT1//lzmrw4DAwPCdUMFAaiGZpWUEoyMjIgTt7Kywurqamxra5NWZvjw4UPZGNVryZqrqyvW1NRgUVGRkLBQT08P19fXZWOuSm9vr3CvEEEA2kiwclIJnj17duGTW1hYwBs3bjB7RESEaN/b20NbW1tmNzY2xsXFRbEvNTUVHRwc0MPDA0dHR0W7OmhcAPpmBQGk3+/GxgbevHmT2e/fvy/aJa8kW6d5Tk5OeJNaaFwACtTCwoLdxMXFhdUWMzMzGBsbKwYq3WNUVFSIdipZp6amMC0tDRMSEljf7u6u7PrqonEBiOnpafbaCoEJzdzcHGtra2W+tHMT+inr82OcnJxwdnZWNkYd3okA9BS9vb3PBUMCvHjxQuabmZkp87G2tsasrCyMjIwUbZQYDw4OZOOuisYFmJ+fRzMzM3YT+hQqKyuxtbVVui3FR48eif4PHjwQ7bT2S592fHy82KdUnaJxAdLT08VJv3r1SrRTtqdiiOy3bt0Sa4HCwkLR/86dO5IrIXZ2dop9Su1SNS6Av78/uwEtefzaHRUVJQZEGxOirq5OtIWEhMj8h4aGxL7ExERZ31XRuABBQUHsBlS8TE5OyvroeE0IiFYGYm5ujvmSzcbGBg8PD0X/xsZG0Z+O5ZRA4wJIv+nQ0FD2FtBaXlVVhfr6+sxuZ2cn2xxJq8SUlBTc3t5mhZN0VVBqq65xAVZWVtDS0lKcOCVCZ2dn8W9qlBilDA8Pi2UvNVoJqCoU/qZPRyk0LgAxMTGBXl5esqCpUZYvLi7m3Rl04uzo6HhuTFxcnKLF0DsRgDg6OmIVX0lJCebn57NN0dLSEu8mY2dnB5ubm9lm6enTp2xjpTQXCaDoblDbuUiAYDIoeR6gzVx0HvAZGVpaWnhfnYQetEqALwQBbABgjw4g3gdU5xXHAPCxIADxKxUw7wNhYWEkwCwAGEgF+IGKFKVOXbQVqjpVhzKF0uAJcwD4g3Zsuoxqi/332/5/gH4zw4KCAn6cTlBWViYkv+/5wKVUkZOuJURJ8E18wDyUGKrJmV6XsbEx/lrXCtpxxsTECME3AoARH/DboB9LNygx3rt3D8vLy9nhxODgIKuktLXR+QGV31RGh4eHo6GhIQX+JyV5PsDLYA0AeQAwCQD/qFS8Lo3W+RkAKAKAj/jArgIVDVQ5UflINbS2Npoflfaf8AFcxL/7QgMGgieFmgAAAABJRU5ErkJggg==",
  87: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAATXSURBVGhD7Zp7KL5nGMcvxxTGEluNVsuSMDkf18Y2JJJEk39++iXJbH9MjUgmciw0TUlImPhLzoef0/4wx1Hzh2ORwwxbLf6Y4VrXnfvpeW6v38Tz/PZ6f/vUFe7rup/nub7v/V7vdd8vgP+5N/YA8AEAfAgAnwLAJ3po9Fz0fJ4A4CAm8BAsASATAF4AwAkA4BOy3wFgEgC+BIA3xMTuw+cAsE4X8/Pzw5ycHGxra8Pe3l4cHR3FkZERvTN6rr6+Pmxvb8e8vDwMDg7mYmwBwDMxwZdRShNDQkJwcHAQnzLj4+MYFhbGhfgOAIzEZEXKKDgrKwuvrq7E6z1ZcnNzuQjfiwnLSeLJGyL5+flchOdi4oQVAOz6+/uL8wyK8PBwEuAIAN4UBfiC1JmYmBDnGBTz8/NoYmJCInwjCvCjr6+vGG+Q3BTFn+UFkZqGP+k9ojbX19e4t7eHOzs7eH5+Lrr/EyoqKkiACwB4lwvgRcu/s7NTjH0U5eXl6O7ujlZWVmhhYYFOTk6YlpaGR0dHt+J8fHwwICDgTiN/R0eHYt5DGRgY4MWQukbGxzQwNDQkxj6Y1NRUfpNb5ubmhqenp1JsSkrKrRhdVlhYqLjHQ5mcnOTXjOACUA/NOik1mJmZkR7awcEBGxsbsaenR96ZYXFxsRRPibm4uNwyDw8PXrDYT2pq1GBsbIw/x2dcANpIsHZSDWpqaqREa2trpfGNjQ00MzNj4zExMYo5umhtbZUEkF/nsWguQHNzsyQA/c45PDxEc3NzNp6cnKyYI7K/v4+2trYsNjo6WnQ/Cs0FoETt7OzYTVxdXVlvsbKygomJiZIw/7bHSEpKYnFUPGnlqInmAhDLy8vo6ekpJcyNXlVa2i9jdnZWik9PTxfdj+aVCLC0tITe3t46BairqxPDFSQkJEiv/vr6uuh+NJoLsLa2htbW1uwm9Faor6/H7u5u+bYUS0pKxGmM7e1tqU7Ex8eLblXQXIDMzEwpUXnzcnZ2xpohGrexscGTkxPFPKKqqkqaSwcbWqC5AEFBQewG9JF3cHCg8MXFxUkJ0sZEhK8SWkFix6gWmgsQGhrKbmBkZISLi4sKHx2vcQHok0HO8fExWxnkCwwMVPjURHMBsrOzpSQjIiLYKri8vMSGhgY0NjZm446Ojrc2R1NTU9K8jIwMhU9NNBdgd3cX7e3tpWSoEDo7O0t/k1FhFGlqapL81dXVols1NBeAWFhYQC8vL0XSZJaWllhaWiqGM4qKiqS4lpYW0a0ar0QA4uLignV8ZWVlWFBQwDZFW1tbYpjE5uYm25EODw9rVgAJXQKouhvUd3QJEEYDap4H6DO6zgN8aKCrq0uMNUjohb4R4CMuwNsAcEZF6HXg5rzibwB4jwtA/EQNzOtAVFQUCfALAJjIBfiamhTaihoyq6urfLP1rTx5whYAfqVe3JCJjY2l5P+46/8H6Dsz1U5f9Y3Kykpe/L4SE5fTQEGGVhBlyXeKCYtQYWikYFouc3Nz4rWeFLTj5CdMAPADAFiICd8FfVl6SIUxMjKSHVT09/ezHRt1Uvpq09PTrP2mzRSdJpuamlLiv1GRFxO8D28BQD4ALALAXzcqPhWjz/kVACgCgHfExB4CNQ3UOVH7SD20vho9H7X274sJ6OIfqOFqhm8+9qEAAAAASUVORK5CYII=",
  88: "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",
  89: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAVtSURBVGhD7Zp/SJVXGMcfNUPQ1Gm6wZKp6B9mKLYKzJazbSpSFIhzEcRsRH+4tj8WTFOsWeGPgpLFAhUFf6TYX2G/MFeaQas0k+yPFlqZ5A9sxjBwLnvG9+A5vO/xFnLve+N62wcOXJ/znPec53vf9znPea9E/7NgQogojog+I6IviegLF2xYF9YXT0ShegD24EtEuUT0OxFNEBEvovYXEXUQ0Q9E5K8HthC+IaI/cbG1a9dyXl4eNzQ0cGtrK1++fJnb2tpcrmFd586d48bGRi4oKOD169dLMQaI6Fs9wLdRgoFJSUl88eJFXsxcuXKFU1JSpBC/EpGHHqxOKZz37t3Ls7Oz+vUWLfn5+VKE3/SAjXwtg3dHCgsLpQjf6YEDPyIaWrdunT7Ordi0aRMEGCOiD3QBvoc6V69e1ce4Fbdv32YvLy+I8LMuQNeaNWt0f7dkLin2GhMiioa/8YxYzevXr3l4eJifPHnCL1++1Ltt8uzZM3769Cm/evVK77KE8vJyCDBDRJ9IARJw+zc3N+u+DlFWVsarVq1iPz8/9vHx4bCwMN69ezePjY3proLa2lpRc/j7+7Ovry+vXLmSjx07prs5zIULF2QyRNUo+ByGS5cu6b52k5OTIyeZ12JjY/n58+cm//3798/zk23Hjh3iTrKKjo4Oee1UKQBqaFFJWcGNGzfU4kNDQ7m6uprPnj1rrMz48OHDyv/69evKHhQUJL513KYBAQHKXlVVZZrDEdrb2+V1v5IC4CAhykkrOHHihFp4RUWFsj98+JC9vb2FffPmzcq+Z88e5Y9yW9LU1KTscXFxluUEpwuAZ1kuHJ8lIyMjvHTpUmHfvn27sicnJwsbtqdHjx4pO5Im7iD0eXp6CgGtwOkCINDg4GAxSUxMjKgt+vr6OCsrSwljPGNs3LhRCfD48WNlxzceGRlpc4wjOF0AcPfuXY6Pj1eLly0wMJDr6upMvjt37lT9yNASbIXYDWQfTnlW8E4EuHPnDq9evdqmACdPnjT5YveR/dj6urq6uL+/n9PT001ja2pqTOPsxekCPHjwgJctWyYmwaNw6tQpPnPmjPFYykeOHDGN2bVr1zyx0GTOQKuvrzeNsRenC5Cbm6sWffr0aWWfmpoSxRDs2OImJiZM4w4dOiSeeewUSH7FxcW8bds2dS1spVbgdAESExPFBAgEZa2RrVu3qoBwMLHF4OAgv3jxQnw21g737t3TXe3C6QJs2LBBTODh4cE9PT2mPpS6MiDsDADPOwodnEV6e3uV7+joqEqCERERPD09bbiS/ThdgH379qkgU1NTxV2ALa2yslLs57CvWLFCHY5Q/Eh/iDc+Pi4ej+zsbGXHOz6rcLoAQ0NDHBISohaPRBgVFaX+RkNilCA34BuWfcuXLzeNj46O5snJSdMcjuB0AUB3dzcnJCSYgkbDLV1SUqK7861btzg8PHyeP4qkgYEB3d0h3okAYGZmRlRvpaWlXFRUJA5FbwsG3zKKnQMHDojDkFWHMx1bAlh6GnR1bAmQAoOV7wNcGVvvAz6FoaWlRfd1Swyld7IU4CMimkLl9T4w977iXyKKlAKAP7AHvw/MHbL6icjLKMBPKFJu3ryp+7sV9+/flwesX4zBg0AiGsWJzZ3ZsmULgp980/8P4DczPnjwoD7OLTh69KhMfj/qgRuphJO7JURD8M16wDpIDNVwxu2CEnUxgxNnZmamDL6JiHz0gN8EfiwdQWJMS0sT7+rPnz/PnZ2dopJy1Xbt2jVRfh8/fpwzMjJ4yZIlCHwcSV4PcCF8SESFRNRDRP/MqbhYGvb5PiIqJqKP9cDsAUUDKieUj6ihXbVhfSjto/UAbPEfxEECl5/TkvwAAAAASUVORK5CYII=",
  90: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAVpSURBVGhD7Zp9SJdXFMfPUlLSdMOc05UDQQkSZfgW6nC2aWEEiTr2V0yGGG1thYLma7M/NlLQWFjIJARjFkEvvlKttCEsdUzDXpTMEEU3xJUZ5usZ34v38jw33eL3EvprH7jw89zz3Oc53+f+zj33/iT6n1fGm4hCiOgjIvqUiD5ZhQ3PhecLJaJ39QAswY2IviKiX4honIh4DbUJImolom+IyEMP7FX4nIj6MVhERATn5uZybW0t19fX87Vr1/jq1aurruG5Ghoa+OzZs5yfn8/R0dFSjAEi+kIP8N/4HhfGxMRwc3Mzr2Vu3LjB8fHxUogfiegtPVidH+B88OBBXlhY0Mdbsxw5ckSKUKkHbOQzGbwjUlBQIEX4Ug8cuBPRUGRkpH6dQ7Fjxw4I8CcRvaML8DXUuXnzpn6NQ9HZ2clOTk4QIUcX4Nfw8HDd3yFZSop/GBMiioZJfEdszdzcHA8NDfHo6KjetSLj4+P8+PFjfvbsmd5lE44fPw4BZonoAynAh5j+dXV1uq/FzM7OcklJCW/dupU3bNjAnp6evH37dj537pzuqnjw4AEnJyfzpk2b2NXVlTdv3iwS8tOnT3VXq2hqapLJEFWj4GMYWlpadF+LQPC7d++WN3mp4Q3oPHr0iH18fF7yRYuLixNj2orW1lY5dqIUADW0qKRswYkTJ9TDYwacOnWKs7KyTEF1dXWZrklLS1N9qampXFVVxSEhIcp28uRJk781XL9+XY6bIAXARkKUk9ayuLjIYWFh4gaYxvfu3VN9hw8fVgGlp6cr+8jIiPCFfcuWLTw/Py/sEGkpY4sxMbYtsKsAT548YW9vb3GDwMBAUx8CkgKgTwaK/YW0YyZIEHBAQIASE0LZArsKMDExwV5eXuIGmP5G+vr6VKBubm48PDws7GVlZcqOWWIE+xHZ197ebuqzFLsKgP3Dtm3bxA3c3d1Ny9/ly5dVMGj3798XdkOdzkVFRYbRmBMSElQfdnq2wK4CgGPHjqmHTkpK4jt37oi3FxQUZBKgu7tb+B86dEjZiouLTWMZBbhw4YKpz1LsLgAKGOwpjMHKtn79evW5t7dX+GdnZ68oQGJiouq7dOmSqc9S7C4AQC7IyMgQa7uzs7NIelgO5dfDxcWFBwcHhS+ClkHm5eWZxkENIPtstUy/FgEkU1NTosiZmZkRZTEqQtzL39+fX7x4IXyqq6tVkJmZmabrg4ODhX3dunUqZ1iL3QVApVVZWSneJpZFCU6XZKB79+5V9tu3byt7VFSUsmMWeXh4KMGmp6dVnzXYXYD9+/ergA4cOMCTk5P88OFDDg0NVfYrV64of8wOud7jTZ8/f14Il5OTo/z1mWENdhcA1R8CkQ/v5+cn1n35NzY8OmfOnFH9aL6+vuoz8kV/f79+icXYXQBQU1PDGzduNAWFtm/fPn7+/LnuLsDyiWCN/kiiFy9e1F2t4rUIAAYGBvj06dNcWFjIFRUV4jTmv0CiQ/7AVhoijo2N6S5Ws5wANt0NrnaWEyAeBludB6x2ljsPCIMB2fdNAC96SYA4KcB7RDSF792bAPIREc0RUYAUAPwWGxur+zoku3btggC9RORkFCALazeqMkfm7t27ckP2nTF48DYRjeHM3JHZs2cPgv97pf8fwG9mfPToUf06h6C0tFQmv2/1wI1UwcnREqIh+Do9YB0khp/gjOnS0dGhj7Wm6Onp4ZSUFBn8z0Tkqge8EvixdBSJcefOneLgsrGxkdva2kQltVrbrVu3xLa7vLxcHMfhMIaI/kKS1wN8FXyIqICIfieimSUV10rDOt9DRCVE9L4emCWgaEDlhPIRNfRqbXg+lPaBegDL8Q+/oQ5teNCxwQAAAABJRU5ErkJggg==",
  91: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAASeSURBVGhD7Zp7KL5nGMe/Q8hxCzNWRJGoaTkkrI0dnFKkfu0vTUvJZkv+mGMZf2yh0Aqt/aNoKH84C3NMzWENOS6Hkpi1TA7lfK3rzv30vveQ/N6V590+dZf3ea77eZ7v97nf677u+wX8z6NxA/AWgHcAfADg/WfY+Ln4+YIBvK4KeAr2AD4D8BOAPwGQjtoBgFEAXwBwUoU9ho8B/MYXCwsLo/z8fGpqaqKuri4aHBykgYGBZ9f4ubq7u6m5uZmKioooMjJSmrEB4BNV4EN8wx2joqKor6+P9Mzw8DDFxMRII74D8IoqVuVbDs7JyaHr62v1erqloKBAmlCnCjbkhRRvjhQXF0sTPlWFMw4AtsPDw9V+ZkVsbCwbsA/gNdWAz9mdkZERtY9ZMTMzQ5aWlmzCV6oBE6GhoWq8WXKbFH81TIhcNBzxd8TUXF5e0vb2Nu3t7amnHuTo6IhWV1dpY2ODbm5u1NMvRUVFBRtwAcBbGvA2D/+WlhY19slcXFxQWVkZBQQEkJ2dHTk7O1NERAS1traqoXeSkJAgEpavr6+4linp7e2VyZCrRsF7fKC/v1+NfRL8wElJSfIm/2j8Bh6itrZWi/X29ja5AaOjo/L6H0kDuIYWlZQpMBTAI6C+vp7y8vKMTJidnVW70dXVFZWUlBjF/RsjYGhoSF7/Q2kALyREOfmy8Pc1JCRE3MDW1paWl5e1c7m5uZqwjIwMo37t7e1aP10bcHh4SG5ubuIGfn5+Ruf4rUthfI7fuCQwMFA7l5KSQq6urvo04ODggFxcXMQNePgbsra2pom0t7ennZ0dcZxHDceyWE7EJycn2jV0ZwCvH4KCgsQNHBwcjKa/jo4OzQBuKysr4jiPhLGxMTH1Mfv7++Tk5KRPA5jy8nJNZGJiIi0sLNDk5CT5+/sbGTA3N6d2FbBpPG3q1oDj42PiNYWhWNmsra21vxcXF9WuAt0bwHAuyMzMJHd3d7KyshJJj6dD+fWwsbGhra0ttZvALAyQcELb3Nyk8/NzURZLYV5eXnR2dqaGC3RvAFdadXV1VFhYKKZFCe8uyeHPU9196N6ArKwsTWh2drbI7uvr6xQcHKwd7+zsVLtp6N4Arv4sLCw0sZ6enmLel59TU1PVLkawAY6OjiLWx8dHfwYwjY2NmgjDlp6eTqenp2q4Ebu7u9ps4eHhoU8DGF7LNzQ0iAVOTU2N2I15DJwweTeXV6cTExMm3w+4ywCTrgafO3cZEMMHTLUf8Ny5az8ghA+0tbWpsWYJv+hbA96VBrwB4IS3sP4LcD4CcAnAVxrA/BwdHa3GmiXx8fFswCIAS0MD8njunpqaUuPNiqWlJTnFfm0onnkVwO+8Z27OJCcns/i/7vv/Af7NjEpLS9V+ZkFlZaVMfl+qwg35noPMLSEaiG9RBatwYviBg3m4TE9Pq9fSFfPz85SWlibF/wjAVhV8H/xj6R4nxri4OKqqqqKenh6xd8eV1HNt4+PjYtldXV0ttuN4MwbAH5zkVYGPwR1AMYBfAJzfuqiXxvP8PIAyAG+qwp4CFw1cOXH5yDX0c238fFza+6kC7uJv3H9/uWdCsnsAAAAASUVORK5CYII=",
  92: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUiSURBVGhD7ZppSHZFFMePC77ikqWYYWAQKIKQhgvigmmZogiKkOkHSUIQ00L9kCtuH3JfCFQyBEFN/eS+5xZCLpGKC4YLiGCGmLiA+4kzOJd7x8cXe3yeN33efnBAZ87cued/ZzkzCvA/D8YaAD4AAF8A+AQAPn6CRu9F7+cMAG+LAaiDKQB8BQA/A8A+AOAzsgMAGAeArwHgDTGwh/A5APxBD3N3d8f09HRsamrC7u5uHB4exqGhoSdn9F49PT3Y3NyMWVlZ6OXlxcXYAIAvxABfxnfU0NvbG/v7+/E5Mzo6iv7+/lyI7wFATwxWpIick5OT8fr6WnzesyUjI4OLUCMGLOczHrwukp2dzUX4UgycMAOAbQ8PD7GdThEQEEAC7AHAW6IASaTO2NiY2EanmJ2dRQMDAxLhW1GAX9zc3ER/neR2UfxdviBS0nBEc0TTXF5e4vb2Nu7u7opV97K3t4dbW1u4v78vVmmEkpISEuACAN7jAnxIw7+1tVX0VZuLiwssKChAR0dHNDExQQsLC/T09MS2tjbRVaK9vR39/PzQ0tISX7x4gdbW1hgYGIgDAwOi66Po6+vjiyFljYyPqEBTHVHwoaGhvJM7Rl9ApLi4+I6f3FpaWsQmajM+Ps6f+ykXgHJolklpgurqaunFaQTU1tZiWlqaIqC5uTnJf319HfX19Vm5np4eJiUlYUNDA0ZFRUn+VlZWeHBwoOhHXUZGRvhzA7kAdJBg6eRjubm5QVdXV9aBsbExrqysSHUpKSlSQHFxcVJ5YWGhVB4TEyOVEzRteN3g4KCiTl20KsDh4SGbu/Q8e3t7RR19dR4M1V1dXbFy6jc1NRWjo6PZ8JQTGxsrteno6FDUqYtWBaBhSsOVnkfDX87a2poUjKmpKe7s7CjqRSgVd3BwkNrIR9Nj0KoA9NJOTk6sAzMzM8X219nZKQVDtrq6qmgrQic67uvr68umlybQqgCEfE6HhITg4uIiTk1NKb4m2fz8vNhUoqioSPKjBXJmZkZ0URutC3B8fIx0ppAHy83IyEj6eWlpSWzKkAtIVlNTI7o8Cq0LQNBaEB8fjzY2NmhoaMgWPdoO+fSgRIeyPRHZiY1ZeXm56PJoXokAnJOTE9zc3MTz83OWFlNGSH3Z2dnh2dmZwpcyR3nw9fX1inpNoXUBaCujYZuZmcm2RQ7dLvHgwsPDFW3ouo3XmZubs3RVW2hdgISEBCmYxMREPDo6Ytmes7OzVN7V1SX50+GHjwwymi4VFRWYn5+PeXl5zHJzc5/HNkjQi/LUlszW1pbt+/z3iIgIhb849O8zuuzUBFoXgGhsbGRDWQyCMrvT01PJj/Z2FxeXO36q7GUnyX/DKxGA2NjYwLq6OszJycGqqip2GyNCAkxMTLA8n19vqzKqp6miCVQJoNHT4FNHlQD+VKCp+4Cnjqr7AFcqoBuZ1wH60LcC+HEB3gGAE1qNXwdoPQKASwB4nwtA/Orj4yP66iTBwcEkwBIAGMgFSKO9e3p6WvTXKZaXl/mBLF8ePPEmAPxJd+a6TFhYGAX/933/P0B/M2Oppy5SWlrKF79vxMDl/EBOurYgyoJvFQMWoYXhR3Km4aLJm5j/goWFBYyMjOTB/wQAxmLA90F/LN2lhTEoKAjLysqwt7eXpa2UST1Vm5ycZMfuyspKdh1HlzEA8Bct8mKAD8EGALIB4DcAOL9V8bkY7fMLAFAAAO+KgakDJQ2UOVH6SDn0UzV6P0rt7cUAVPEP56Yj2R7VNtgAAAAASUVORK5CYII=",
  93: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAVhSURBVGhD7ZppSF1HFMdPoxj3BbFKKgoFg2CoFE0QjVi1rWIwGISSL5FKEcTWShA0bmhVbFFBpaCiFRUMLvjFXdQaV6zGUhUXLDEuCNZaNHEhbskpZ3CGdyfa2ud74fnoDwZ05szcOf8798yZUYD/OTd2APARAPgCwKcAEKiDheZF83MHgPdlB9TBDAC+BoCfAeAvAMBLVDYBoA8AvgUAS9mx83AfAH6nwW7evImPHj3CmpoabGlpwe7ubuzq6tK5QvNqbW3Fx48fY0pKCnp7e3MxFgDgS9nBf+J76ujj44MdHR14ment7UV/f38uxI8A8J7srMwPZBwbG4uvX7+Wx7u0JCUlcRGKZYdV+YI7r4+kpqZyEb6SHSfMAWDl1q1bcj+9IiAggARYBwAbWYBvSJ0nT57IffSKp0+fooGBAYmQKAsw6OnpKdvrJSdB8TfVgEhJwzZ9I5rm6OgIV1ZWcG1tTW46k/X1dVxeXsa9vT25SSPk5uaSAIcA4MwF+JiWf11dnWyrNoeHh5iZmYmurq5oamqKVlZW6OXlhfX19bKpoLGxEX19fdHGxgZNTEzQ2dkZ4+Li8MWLF7LphWhvb+fBkLJGxidU0dnZKduqBTl/584d/pC3Cr0BmZKSkrfseKHA/PLlS7mL2vT19fGxP+cCUA7NMilNUFRUJCZPK4Cci4+PVzg1Pj4u7OnzoDdO9cbGxmzlVFRUsL7cPj09XfGMi9DT08PH/YwLQAcJlk5elDdv3qCHh4dwZnZ2VrQ9fPhQOBQZGSnqKb02NDRk9ZRycwYGBoQ9fRqaQqsC0PdqZ2fHHuDi4qJoo7fOHaK24+NjVv/q1StcWlpiE6MAyBkaGhL2oaGhKiNdDK0KsLm5iba2tuwBtIRVmZ+fFw6ZmZnh6uqqop2ztbWFg4OD7BDG7emQoym0KgCdH9zc3NgDzM3NFdtfU1OTcIjK3Nycoi/nwYMHwsba2hpra2tlkwuhVQGIrKws4UBISAhOTU3h8PAwXr9+XSHAxMSE3JXFkBs3bggbS0tLTEhI0OguoHUBdnZ22Nal6iwvRkZG4ufp6Wm5K4sLtPxHR0eZ49zWz88PDw4OZHO10LoABMWCqKgotLe3ZxGegh5th/zzuHr1Ki4uLsrd3iIwMFCI0NzcLDerxTsRgLO7u4vPnz9nb4/SYsoI6VlOTk64v7/PbGjZcztKolRJTEwUAuTk5Cja1EXrAlCmVVxcjMnJyYo0lm6XuDNhYWGinpIcR0dH9nk0NDSIeuLu3buiT1lZmaJNXbQuQHR0tJh0TEwMbm9v47Nnz9Dd3V3Uqy7nqqoqUU9bJwVHEq68vFwkSHSEpTE0gdYFoOzvypUrwqlr166xfZ//fu/ePYU9LXvK9Hg7Oe3g4CB+p0LpsabQugBEdXU1WlhYKJygEhERceoxd2NjQ7HceaExsrOzZfML8U4EIBYWFrC0tBTT0tKwsLCQ3cb8GyMjI8yW3nhlZeW5dor/ymkCaPQ0qOucJoA/VWjqPkDXOe0+wIMq5C1IX6EXfSKAHxfAAQB2NRlpdRmKMQBwBAAfcgGIX27fvi3b6iXBwcEkwDQAGKgKEE97Nx1C9JmZmRl+IPtO1XnCGgD+oDtzfYZulgBg66z/H6C/mWFGRobcTy/Iy8vjwS9OdlyVMjLSt4Co4nyd7LAMBYafyJiWy9jYmDzWpWJychLDw8O587UAYCw7fBb0x9I1CoxBQUGYn5+PbW1t2N/fzzIpXS10lU7H7oKCAnYdd3Ka/JOCvOzgebAHgFQA+BUADk5UvCyF9vlJAMgEgA9kx9SBkgbKnCh9pBxaVwvNj1J7F9mB0/gbkxo5croC/qwAAAAASUVORK5CYII=",
  94: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAUbSURBVGhD7Zp7KP5XHMc/7nIZcxlRVpNLKSyXhLWxC5EitfYX01JituQP9zLUFgotl9b+IRriD3dhrsl1DbnOJUlsLHe5+6zPyfn2OB772bPHL8+zveqU55zPuXzez7l8znkA/M+TsQQAVwB4DwA+AoAPX2CicdH43ADgLdEBRTAEgAQA+BkA/gQAVKG0BwD9APAVALwhOvYUPgOA36gxLy8vTE1NxerqamxpacHu7m7s6up6cYnG1draijU1NZiRkYG+vr5cjFUA+Fx08O/4lir6+flhR0cHqjK9vb0YEBDAhfgeADREZ0W+I+PExES8ubkR21NZ0tLSuAhlosOyfMqdV0cyMzO5CF+IjhNGALDh7e0t1lMrAgMDSYA/AOBNUYAvSZ2+vj6xjloxMTGBWlpaJEKKKMCQp6enaK+W3G2Kv8puiBQ0HNEaUTZXV1e4sbGB29vbYtErOTw8xMXFRVxeXsbr62uxWGHy8/NJgEsAeJsL8C5N/9raWtFWYS4vLzEnJwednZ3RwMAATUxM0MfHB+vq6kRTuVB9ij9oXObm5ri7uyuaKEx7ezvfDClqZHxAGZ2dnaKtQtDgQ0NDeScPEn0DryIlJUWyNzU1VaoA/f39vO1PuAAUQ7NIShmUlJRIg6cZUF5ejsnJyfdEmJycFKtJDA0NoYaGhmSr7BnQ09PD2/6YC0AXCRZO/ltub2/Rw8ODdaCvr4/z8/NSWVJSkuRUTEzMvXqc09NTdHR0vCeWSglwcHCAlpaWrAMHB4d7ZfStc6eoTN7GlpCQwMqNjY2ldlRKgL29PTZgao+mvyxLS0uSAIaGhri5uXmvnPYgXk7L5i5oUS0B6P7g4uLCOjAyMrp3/DU1NUkOUlpYWJDK9vf30c7OjuUHBQWxPL6UVEoAIjc3V3IyJCQEZ2ZmcHh4+MHanpqakupER0ezPJr66+vrLM/V1VU1BTg+Pka6U8g6y5Ourq709+zsLLNvbGxkn3V0dLChoUFqh88ACwsLpd5Mn10AgvaC2NhYtLKyQm1tbbbp0brmy0NPTw+3traYWNbW1iyPbEtLS9kxSsnW1pbl036Rl5eHVVVVYjcK8VoE4JycnODa2hpeXFywsJgiQuqL1jsdmTTdxVnyWKJTQRk8uwAUaZWVlWF6ejo7Fjn0usSdCQ8PZ3mrq6sPHH0s2dvby/SiOM8uQFxcnDTo+Ph4PDo6wpWVFXRzc5Pym5ubme35+TmOjY3h6OiolOgzJVo2ZEuzho5I2kyVwbMLQNGfpqam5KyNjQ1bx/xzRESEWEUu7u7uzN7MzAzPzs7EYoV5dgGIyspKdqRxp3mKiopi4e5TcHJyYnXoNrmzsyMWK8xrEYCg9V1RUYFZWVlYXFzMXmP+CSMjI2zq0ysV3TCVhTwBlHobfOnIEyCAMpT1HvDSkfce4EEZ9fX1oq1aInPpep8LYA0AJ/SE9V+A9iMAuAKAd7gAxKi/v79oq5YEBweTALMAoCUrQDKd3RSAqDNzc3P8QvaNrPOEKQD8Tm/m6kxYWBg5v//Y/w/Qb2aYnZ0t1lMLCgoK+Ob3tei4LD+QkbptiDLO14oOi9DG8CMZ03QZHx8X21IppqenMTIykjv/EwDoiw4/Bv1Yuk0bI73RFRYWYltbGw4MDLBI6qWmwcFBdu0uKipiz3H0GAMAO7TJiw4+BSsAyASAXwDg4k5FVUl0zk8DQA4A2IqOKQIFDRQ5UfhIMfRLTTQ+Cu0dRAfk8RcUayqQ8/UkdQAAAABJRU5ErkJggg==",
  95: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAVbSURBVGhD7ZppSF1HFMePS6K4VItYJcWIFRcUKsVERa2tdomJCIpYxQ9SKYJotRE/1BXcsEUDLgU36gchUiP6wV3UulWhLqUqLlg30IBtKVbUuMdTzuBc7h1NG57vpfraHxzwzZy5d87/zjtz7jwB/uelsQSAtwHgXQD4EAA+uIJG86L5uQHAG2IAqmAMAAkA8D0A/AEAeI1sEwAGACAJAF4TA3sZIgHgF7rY3bt3MTU1FR8/foytra3Y09OD3d3dV85oXm1tbVhXV4cZGRno7e3NxVgGgE/FAP+Or2igj48PdnZ24nWmr68P/f39uRDfAICOGKzI1+ScmJiIz58/F693bUlLS+MilIsBy/mEB6+NZGZmchE+EwMnTABgzcPDQxynVQQEBJAAvwHA66IAn5M6/f394hitYnx8HPX09EiEL0UBfrhz547or5WcJcWf5QmRioZt+o6om+PjY1xbW8ONjQ2x61+jsLCQBDgCAFsuwDu0/Ovr60VflTk6OsLc3Fx0dnZGIyMjNDMzQy8vL3zy5InoKkHJl1ahp6enwigv+fn5qU3Ejo4OngypamS8Tw1dXV2ir0pQ8EFBQfwm54yegMjh4SFaW1uf85XbysqKOEwlBgYG+DU/5gJQDc0qKXVQWloqTZpWQEVFBaakpCiCmZiYUIxZXFxEfX191mdsbIy2trZoY2MjmYODAz59+lQxRlV6e3v5PD7iAtCLBCsnL8vp6Sm6u7uzGxgaGuLc3JzUl5ycLAkQExOjGNfS0iL1FRQUsFVEq2Jvbw93d3fx2bNn7NrqQKMCbG1toaWlJbsBPTU59NR5kNR3cnIi9eXn57N2HR0dHB0dZW2aqkQ1KsDm5iZaWFiwG9Dyl7OwsCAJQMtcvqTDw8NZ+40bN1j+oLF2dnb44MEDtb+PaFQAemqurq7sBiYmJorM3dzcLAlANj8/L41xcXFR9IlWXl4uu8vl0KgARF5enjRxeoLT09M4MjKCjo6OiqAmJyeZ//r6Oq/OWCJMSkrCmpoajIyMlHx1dXUV+eQyaFyAnZ0dtneLT5Hs5s2b0t8zMzPMf39/n21N1dXVODw8rLhWRESE5E9nEupA4wIQlAtiY2PRysqKPVVKerQd8q+HgYEBrq6uisPO0dTUJAkQHBwsdqvEKxGAQ1sYFTC0pVFZTBUh3ev27dt4cHCg8KV+MfO3t7dLAty/f1/RpyoaF4CWMyWt9PR0ti1yKJvzYEJCQqT2xsZGVibb29uz/CHnrG5nlpCQoOhTFY0LEBcXJ006Pj4et7e3cWlpCd3c3KR2Knw4sgmhubk5Dg4OsqKHkuStW7ekPjE/qIrGBaBsTVmbT5yCoH2ffw4NDRWHSHUAGRVDTk5OioRJQqoLjQtA1NbWoqmpqRQAt+joaFbWilCuiIqKOudPCfThw4dqK4OJVyIAsby8jJWVlZiVlYUlJSXsNOafoHK5rKwMc3JysKqqCmdnZ0WXS3ORAGp9G7zqXCSAPzWo6zzgqnPReYA7NTQ0NIi+Wgk96DMB3uMCWAPALh1h/RegfAQAxwDwFheA+NHX11f01UoCAwNJgBkA0JMLkEJ7Nz+M0FZoVzmrL3LkwRPmAPArnZlrM/RCBQB/vuj/B+g3M8zOzhbHaQVFRUU8+X0hBi6nmpy0LSHKgq8XAxahxPAtOdNyGRsbE691rZiamsKwsDAe/HcAYCgG/CLox9INSoz37t3DR48esXdzelOjSuqq2tDQEHvtLi4uZsdxZ781/E5JXgzwZbACgEwA+AkADs9UvC5G+/wUAOQCwJtiYKpARQNVTlQ+Ug19VY3mR6W9gxjARfwFxtAeZq0cS+sAAAAASUVORK5CYII=",
  96: "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",
  97: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAATsSURBVGhD7Zp5SGV1FMe/qTjikoaYYWAQuICYhCtqlNPi4CAoggQDkoS4lIX4R65g+ke5gCOBSoSoOObMf+7rjEsIuUQqLriDjJgZJi7gfuJc/F3e+6mD6LvT89UHDuo753fvPd/7u+ee3+8J/M+VcQLwDoD3AHwE4EMjNL4uvj4fAK/LCVwHGwBfAHgK4C8AdItsE0AfgK8AvCondhU+BTDHB/P396eMjAyqq6uj5uZm6u7upq6uLqMzvq6WlhZ69OgRZWdnU3BwsBBjEcBncoIv4jseGBISQu3t7XSbefbsGYWFhQkhfgDwipyszPccnJqaSicnJ/Lxbi2ZmZlChHI5YV1iRfKmSE5OjhDhczlxxhbASkBAgDzOpLh79y4LsA7gNVmAL1md3t5eeYxJMTIyQubm5izCN7IAv/j5+cnxJslZUfxdtyBy07DNz4ihOTo6opWVFVpbW5Nd/xpFRUUswCGAt4QA7/L0b2hokGOvzeHhIeXn55OnpydZW1uTvb09BQUF0ePHj+VQKiwsJF9fXwoMDLzU2F9fXy8PvRZtbW2iGHLXqPABf9DR0SHHXgtO/v79++Ik54zvgC4PHjw4F3OR5eXl6Y27Ln19feKYnwgBuIdWOilDUFZWpl40z4CKigpKT0/XS2Z0dFSN58Q8PDzOmbe3tyhYyk9uagxBT0+PuI6PhQC8kFDayZtyenqqTFc+npWVFU1PT6u+tLQ0VYD4+Hi9cRdRW1urCsCiGgpNBdja2iInJyflBG5ubno+vutCAPYdHx/r+XVZXV0lBwcHJTYiIkJ23whNBdjc3CRHR0flBDz9dZmdnVUFsLGxoefPn+v5dYmNjVVn0fz8vOy+EZoKwOsHLy8v5QS2trZ6r7/GxkZVALaZmRm9sYKhoSE1JjExUXbfGE0FYAoKCtQEePpOTEzQ4OAgubu76wkwNjYmD1WIiYlR7/7c3JzsvjGaC7Czs0O8ptBNVpilpaX6++TkpDyUlpaW1Jjo6GjZbRA0F4DhWpCQkEDOzs5kYWGhFD1+HYrH486dO7S8vCwPo5KSElUg3tjQgpcigGB3d1e5qwcHB0pbzB0hn8vV1ZX29/flcHXzws7OjtbX12W3QdBcAO60ysvLKSsrS3ktCnh3SdzdqKgovTHMxsaGKhC3zVqhuQBJSUlqoikpKbS9vU0LCwvk4+Ojft7U1CQPo/7+ftWfnJwsuw2G5gJw92dmZqYm4+Liorz3xd+XFbeqqio1prS0VHYbDM0FYGpqapTnWCQkLC4ujvb29uRwBV49irjq6mrZbTBeigDM4uIiVVZWUm5uLj18+FDZjXkR/JjwirSzs1OzAshcJIBBV4PGzkUChPEHhtoPMHYu2g/w5Q+ePHkix5okfKPPBHhfCPAGgF0uQv8FuB4BOALwthCA+TU0NFSONUnu3bvHAkwCMNcVIJ3f3bwUNWWmpqbEYutb3eQZBwB/cC9uykRGRnLyf1/2/wP8nZnBdl+NjeLiYlH8vpYT1+VHDjK1gqiTfIOcsAwXhp84mKfL8PCwfKxbxfj4uLrDBOBnAFZywpfBX5aucWEMDw9XNipaW1uVFRt3UsZqAwMDyrKbF1O8HcebMQD+5CIvJ3gVnAHkAPgNwMGZirfF+D0/DiAfwJtyYteBmwbunLh95B7aWI2vj1t7NzmBi/gH53Zual0dMA4AAAAASUVORK5CYII=",
  98: "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",
  99: "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAoCAYAAABOzvzpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAVoSURBVGhD7Zp7SNZXGMefeUfNNq+jgStBCQVleCG05XQXpQiSYA4EmY1Q2twQ/5ilYtM/NiqoGJisoUSK2n/dVMyppcK8DPMazkuhoE5EwwxS02d8D57D+x412uv7xuvbPnBAnt/z+53zfM/v95znnFei/3ltfIgolIg+JqLPiOhTK2wYF8YXRkS+egCm4EZE3xLRH0Q0S0S8g9ocETUT0fdE5KEH9jp8RUR/42GRkZGck5PD5eXlfPv2bb537x7X19dbXcO47ty5wxUVFZybm8vR0dFSjFEi+loP8FX8jBtjYmK4traWdzKNjY0cFxcnhfiViN7Rg9X5Bc6ZmZm8urqqP2/Hcvr0aSlCsR6wIV/K4G2RvLw8KcI3euDAnYjGo6Ki9Ptsivj4eAjwDxG9pwvwHdRpamrS77EpOjs72d7eHiL8qAvQEhERofvbJOtJsdswIaJoWMA3Ym5WVlZ4fHycp6am9EtbMjk5yRMTE/zy5Uv9klk4d+4cBFgmog+lAB/h9a+qqtJ9TWZ5eZkLCwt5//797Orqyrt37+YDBw5wdXW17qooKysTNYeHhwe7ublxcHAwX7hwQXfbNjU1NTIZomoUfAJDXV2d7msSCP7IkSOykw0NM6Bz5syZDX6ypaSk8Nramn6LyTQ3N8tnfyEFQA0tKilzcPnyZTV4vAFXrlzh7Oxso6C6urqUf2trq7J7enqKWYdIeGuk/erVq0Z9bIeGhgb53M+lANhIiHJyu2CmwsPDRQcuLi48ODiormVlZamA0tLSlD09PV3ZUW5LKisrlT00NNRsOcGiAjx9+pR9fHxEB4GBgUbXMOsyIFyTlWZsbKywYXl6/Pix8n/+/Dn7+vqKa3Z2djw8PGzwNNOxqABzc3Ps5eUlOsDrb8jQ0JASAEkO2R4cOnRICfDkyRPljxkPCAhQ95hrX2JRATCrISEhogN3d3ej5e/mzZsqGLRHjx4Je2pqqrIhQ0uwFEIoeQ27PHNgUQFAUVGRGvThw4e5t7eX29raOCgoyEiAhw8fCn+sPtKGpa+lpYX7+/s5MTHRyL+0tFTvyiQsLsCzZ88YewrDwcvm5OSk/kaQkhMnTmzw1f2vX79u1I+pWFwAgFxw8uRJ9vPzYwcHB5H0sBzKz8PZ2dko4QG8OfjmHR0dRfJDIXXs2DElAD4hc/BGBJAsLi7y2NgYLy0tibJYru3+/v784sUL3V0Af6wmwOBUh/v6+nRXk7C4AKi0iouLRXUnAwHI4jIYzKwEnwIKHexFuru7lX16elolwX379m0p2H/F4gJkZGSoQE+dOsULCws8MjLCYWFhyn7r1i3lj+JH2g8ePMgzMzM8OzvLycnJyo4zPnNhcQFQ/aFwkYPfs2eP0XKWlJRk5I/PBDMsr3t7e6tiCg35Y35+3uie7WBxAcC1a9d4165dKgjZsOajwtPp6OjgvXv3bvBHkTQ6Oqq7b4s3IgDAwEtKSjg/P58vXbokTmNeBWYZxU5BQYHYDJlrc6azmQBm3Q1aO5sJEAeDuc4DrJ3NzgPCYbhx44bua5MYlN6xUoD3iWgRldfbAPIREa0QUYAUAPyJNfhtYH2T1U9E9oYCZGPtbm9v1/1tioGBAbnB+skwePAuEU3jzNyWOXr0KIKf3+r/B/CbGZ89e1a/zyY4f/68TH4/6IEb8hucbC0hGgRfpQesg8TwO5zxuqBE3cn09PTw8ePHZfCVROSiB7wV+LF0CokxISFBnNXfvXuX79+/Lyopa20PHjwQ2+6LFy+K4zgcxhDRDJK8HuDr4EdEeUT0FxEtrau4UxrW+R4iKiSiD/TATAFFAyonlI+ooa21YXwo7QP1ADbjX9pSBoQgOZBrAAAAAElFTkSuQmCC"
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
 * 겹쳐 올립니다. anchorCol/anchorRow/photoOffsetX/photoOffsetY는
 * 방금 넣은 사진과 똑같은 칸 기준이라, 그 사진의 실제 가로/세로(px)만
 * 더하면 사진 안에서의 좌표를 그대로 계산할 수 있습니다.
 **************************************************************/
function overlaySeqBadge_(targetSheet, seq, anchorCol, anchorRow, photoOffsetX, photoOffsetY, photoWidth, photoHeight) {
  const badgeInfo = recoveryBadgeInfoForSeq_(seq);
  if (!badgeInfo) return;

  const badgeX = Math.max(photoOffsetX, photoOffsetX + photoWidth - RECOVERY_BADGE_MARGIN_PX - badgeInfo.width);
  const badgeY = Math.max(photoOffsetY, photoOffsetY + photoHeight - RECOVERY_BADGE_MARGIN_PX - badgeInfo.height);

  const badge = targetSheet.insertImage(badgeInfo.blob, anchorCol, anchorRow, badgeX, badgeY);
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
        const anchor = pixelXToColumnOffset_(targetSheet, pixelX);

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
        placedPhotos.push({
          item: item, anchorCol: anchor.column, offsetX: anchor.offsetX, width: finalWidth, height: finalHeight
        });
        insertedCount++;
      } catch (error) {
        failedCount++;
      }
    });

    // 2단계: 번호 뱃지는 이 줄의 사진을 전부 넣은 뒤에 올립니다 — 먼저
    // 넣은 사진 위에 뱃지가 가려질 일이 없게 순서를 분리했습니다.
    placedPhotos.forEach(function(p) {
      overlaySeqBadge_(targetSheet, p.item.seq, p.anchorCol, blockStartRow, p.offsetX, 4, p.width, p.height);
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
