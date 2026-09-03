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
 *   시트로 자료를 정리해 옮깁니다. 자세한 내용은
 *   cleanMonthlyOutsourceListAction_ 참고.
 *
 * 배포 방법
 * ------------------------------------------------------------
 * 1. 월마감 외작건정리 스프레드시트 > 확장 프로그램 > Apps Script에 이 파일 추가
 * 2. 배포 > 새 배포 > 유형: 웹 앱, 실행 계정: 나, 액세스 권한: 필요 범위
 * 3. 배포 후 나오는 웹 앱 URL을 대시보드의 "월간업무" 페이지 >
 *    "월간업무 데이터" > "상세" > "월마감 외작건정리 연결"에 입력
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
const GRAND_TOTAL_LABEL = "합계";


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
      return jsonOutput_(cleanMonthlyOutsourceListAction_());
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
 * 기준으로 묶어서 행이 많은 그룹부터(내림차순) 정렬해 "종합" 시트에
 * 씁니다. 같은 그룹 안에서는 시트1의 원래 순서를 유지합니다.
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
function cleanMonthlyOutsourceListAction_() {
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

  // 행이 많은 그룹부터(내림차순). 개수가 같으면 먼저 나온 순서 유지(안정 정렬).
  groupOrder.sort(function(a, b) { return groups[b].rows.length - groups[a].rows.length; });

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

  return {
    ok: true,
    resultSheet: SUMMARY_SHEET_NAME,
    rowCount: parsedRows.length,
    groupCount: groupOrder.length
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


function normalizeText_(value) {
  return String(value === null || value === undefined ? "" : value).trim();
}


function jsonOutput_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
