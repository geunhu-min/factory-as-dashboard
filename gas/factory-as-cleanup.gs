/**************************************************************
 * 1공장 / 2공장 AS 현황 정리
 *
 * 실행 메뉴
 * ------------------------------------------------------------
 * 1공장정리(추가건)
 * 2공장정리(추가건)
 *
 * 완료 후 이동
 * ------------------------------------------------------------
 * 1공장 실행 완료 → 1공장추가건 시트
 * 2공장 실행 완료 → 2공장추가건 시트
 *
 * 공통 회수내역 시트
 * ------------------------------------------------------------
 * 회수내역
 *
 * 주요 기능
 * ------------------------------------------------------------
 * - 기존 시트를 삭제하지 않음
 * - 새로운 시트를 생성하지 않음
 * - 시트 순서와 gid 유지
 * - 시트 이름으로 찾아서 작업
 * - 누적데이터와 정리 시트의 열 수가 달라도 자동 보정
 * - 추가건 정렬:
 *   일반 → 감성 → 취급 → VN → 미회수
 **************************************************************/


const FACTORY_CONFIG = {
  FACTORY_1: {
    name: "1공장",

    originalSheet: "1공장 금주AS현황분석시트",
    accumulatedSheet: "1공장 누적데이터",
    cleanSheet: "1공장정리",
    addSheet: "1공장추가건",
    deleteSheet: "1공장삭제건",

    manufacturer: "충주1"
  },

  FACTORY_2: {
    name: "2공장",

    originalSheet: "2공장 금주AS현황분석시트",
    accumulatedSheet: "2공장 누적데이터",
    cleanSheet: "2공장정리",
    addSheet: "2공장추가건",
    deleteSheet: "2공장삭제건",

    manufacturer: "충주2"
  }
};


const RECOVERY_SHEET_NAME = "회수내역";


/**************************************************************
 * 스프레드시트 메뉴
 **************************************************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("업무 매크로")
    .addItem("1공장정리(추가건)", "runFactory1")
    .addItem("2공장정리(추가건)", "runFactory2")
    .addToUi();
}


/**************************************************************
 * 1공장 실행
 **************************************************************/
function runFactory1() {
  runFactory_(FACTORY_CONFIG.FACTORY_1);
}


/**************************************************************
 * 2공장 실행
 **************************************************************/
function runFactory2() {
  runFactory_(FACTORY_CONFIG.FACTORY_2);
}


/**************************************************************
 * 공장별 전체 실행
 *
 * 작업 완료 후 해당 공장의 추가건 시트로 이동합니다.
 **************************************************************/
function runFactory_(config) {
  try {
    runFactoryCore_(config);

    /*
     * 작업 완료 후 해당 추가건 시트로 이동
     */
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const addSheet = getRequiredSheet_(config.addSheet);

    ss.setActiveSheet(addSheet);
    addSheet.getRange("A1").activate();

    SpreadsheetApp.flush();

    SpreadsheetApp.getUi().alert(
      config.name + " 정리가 완료되었습니다.\n\n" +
      "정리 시트: " + config.cleanSheet + "\n" +
      "추가건: " + config.addSheet + "\n" +
      "삭제건: " + config.deleteSheet + "\n\n" +
      "추가건 정렬 순서\n" +
      "일반 → 감성 → 취급 → VN → 미회수\n\n" +
      "'" + config.addSheet + "' 시트로 이동했습니다."
    );
  } catch (error) {
    showError_(config.name, error);
    throw error;
  }
}


/**************************************************************
 * 공장별 정리 파이프라인만 실행 (메뉴 얼럿/시트 이동 없음)
 *
 * 대시보드 Web App(factory-as-webapp.gs)처럼 UI 컨텍스트가 없는
 * 곳에서 호출하기 위해 분리했습니다. SpreadsheetApp.getUi()는
 * 이런 컨텍스트에서 호출하면 오류가 나므로 여기서는 사용하지 않습니다.
 **************************************************************/
function runFactoryCore_(config) {
  cleanAndRemoveDuplicates_(config);
  compareAddDelete_(config);
  matchRecovery_(config);
  sortAndFormat_(config);

  SpreadsheetApp.flush();
}


/**************************************************************
 * 필수 시트 찾기
 *
 * 시트 순서는 상관없으며 이름으로 찾습니다.
 **************************************************************/
function getRequiredSheet_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(
      "'" + sheetName + "' 시트를 찾을 수 없습니다.\n\n" +
      "시트 이름과 띄어쓰기를 확인해 주세요."
    );
  }

  return sheet;
}


/**************************************************************
 * 시트 크기 확보
 **************************************************************/
function ensureSheetSize_(sheet, requiredRows, requiredColumns) {
  const rows = Math.max(Number(requiredRows) || 1, 1);
  const columns = Math.max(Number(requiredColumns) || 1, 1);

  const currentRows = sheet.getMaxRows();
  const currentColumns = sheet.getMaxColumns();

  if (currentRows < rows) {
    sheet.insertRowsAfter(
      currentRows,
      rows - currentRows
    );
  }

  if (currentColumns < columns) {
    sheet.insertColumnsAfter(
      currentColumns,
      columns - currentColumns
    );
  }
}


/**************************************************************
 * 고정 시트 초기화
 *
 * 시트 자체는 삭제하지 않습니다.
 * 시트 이름, 위치, gid는 유지됩니다.
 **************************************************************/
function clearFixedSheet_(sheet) {
  const filter = sheet.getFilter();

  if (filter) {
    filter.remove();
  }

  const dataRange = sheet.getDataRange();

  if (dataRange) {
    try {
      dataRange.breakApart();
    } catch (ignore) {
      // 병합된 셀이 없는 경우 무시
    }
  }

  sheet.clear();
}


// 1공장정리/2공장정리 시트가 항상 이 순서로 나와야 합니다. 원본
// ("N공장 금주AS현황분석시트")의 열 순서가 컴퓨터/변환 과정에 따라
// 달라져도, 아래처럼 헤더 이름으로 열을 찾아서 이 순서대로 뽑아
// 쓰므로 흔들리지 않습니다.
const CLEAN_SHEET_TARGET_HEADER = [
  "최종조치일", "브랜드", "지역센터", "접수번호", "구분", "형태",
  "고객명", "부품명", "제품코드", "색상", "수량", "금액", "조치결과",
  "회수구분", "서비스요구내역", "반납일자", "제품공급처"
];

// 색상 열의 고정 위치(1-based). 정리/추가건/삭제건 시트 모두 위
// CLEAN_SHEET_TARGET_HEADER 순서를 그대로 이어받으므로 항상 같은 자리입니다.
// "061"처럼 앞자리 0이 있는 값을 텍스트로 유지하려면 이 열에는 절대
// "General" 서식을 적용하면 안 됩니다 — 이미 텍스트로 저장된 값이라도
// "General"로 서식을 바꾸는 순간 시트가 숫자로 재해석해서 앞자리 0이
// 사라집니다(applyBasicFormats_에서 이 열만 빼고 General을 적용).
const COLOR_COLUMN_POSITION = CLEAN_SHEET_TARGET_HEADER.indexOf("색상") + 1;

/**************************************************************
 * 원본 헤더에서 CLEAN_SHEET_TARGET_HEADER 각 이름의 열 위치를 찾습니다.
 * 원본 열 순서와 무관하게 항상 같은 순서로 뽑아 쓰기 위함입니다.
 **************************************************************/
function resolveCleanSheetColumnIndexes_(header, sheetLabelForError) {
  return CLEAN_SHEET_TARGET_HEADER.map(function(label) {
    const idx = header.indexOf(label);

    if (idx === -1) {
      throw new Error(
        "'" + sheetLabelForError + "'에서 '" + label + "' 열을 찾을 수 없습니다.\n\n" +
        "헤더(1행)에 이 이름이 정확히 들어있는지 확인해 주세요."
      );
    }

    return idx;
  });
}


/**************************************************************
 * 1단계
 * 정리 시트 작성 및 중복 제거
 **************************************************************/
function cleanAndRemoveDuplicates_(config) {
  const original = getRequiredSheet_(config.originalSheet);
  const cleanSheet = getRequiredSheet_(config.cleanSheet);

  const sourceLastRow = original.getLastRow();
  const sourceLastColumn = original.getLastColumn();

  if (
    sourceLastRow < 2 ||
    sourceLastColumn < 1
  ) {
    throw new Error(
      "'" + config.originalSheet + "' 시트에 정리할 데이터가 없습니다."
    );
  }

  const sourceRange = original.getRange(
    1,
    1,
    sourceLastRow,
    sourceLastColumn
  );

  const sourceValues = sourceRange.getValues();

  // 색상 열은 "061"/"072"처럼 앞자리 0이 사용자 정의 숫자 서식(예: "000")으로
  // 채워져 표시되는 경우가 있는데, getValues()는 서식이 아니라 실제 값(숫자
  // 61/72)을 가져와서 그 0이 사라집니다. getDisplayValues()로 화면에 보이는
  // 그대로("061"/"072", 문자열)를 따로 읽어서 색상 열만 그 값으로 바꿔치기합니다.
  const sourceDisplayValues = sourceRange.getDisplayValues();

  const header = sourceValues[0];
  const columnIndexes = resolveCleanSheetColumnIndexes_(header, config.originalSheet);
  const receiptNoColIndex = columnIndexes[CLEAN_SHEET_TARGET_HEADER.indexOf("접수번호")];
  const colorColIndex = columnIndexes[CLEAN_SHEET_TARGET_HEADER.indexOf("색상")];
  const sourceRows = [];

  /*
   * 제목 행을 제외하고 A열이 비어 있지 않은 행만 사용
   */
  for (let i = 1; i < sourceValues.length; i++) {
    const aValue = sourceValues[i][0];

    if (
      aValue !== "" &&
      aValue !== null
    ) {
      sourceRows.push({
        originalIndex: i,
        values: sourceValues[i]
      });
    }
  }

  if (sourceRows.length === 0) {
    throw new Error(
      "'" + config.originalSheet + "' 시트에서 " +
      "A열이 입력된 자료를 찾을 수 없습니다."
    );
  }

  /*
   * 접수번호 열(이름으로 찾은 위치) 기준 중복 자료 중 남길 행 선정
   */
  const selectedRows = {};

  for (let i = 0; i < sourceRows.length; i++) {
    const rowInfo = sourceRows[i];
    const row = rowInfo.values;

    const cellValue = normalizeText_(row[receiptNoColIndex]);
    const splitData = cellValue.split("-");

    if (splitData.length < 2) {
      continue;
    }

    const code = normalizeText_(splitData[0]);
    const sequence = Number(splitData[1]) || 0;

    if (!code) {
      continue;
    }

    if (!selectedRows[code]) {
      selectedRows[code] = {
        sequence: sequence,
        sourceIndex: rowInfo.originalIndex
      };

      continue;
    }

    const saved = selectedRows[code];

    /*
     * 순번이 작은 행 우선
     */
    if (saved.sequence > sequence) {
      selectedRows[code] = {
        sequence: sequence,
        sourceIndex: rowInfo.originalIndex
      };
    }
  }

  const cleanedRows = [];

  /*
   * 선택된 행만 정리 결과에 포함
   */
  for (let i = 0; i < sourceRows.length; i++) {
    const rowInfo = sourceRows[i];
    const row = rowInfo.values;

    const cellValue = normalizeText_(row[receiptNoColIndex]);
    const splitData = cellValue.split("-");

    if (splitData.length >= 2) {
      const code = normalizeText_(splitData[0]);

      if (
        selectedRows[code] &&
        selectedRows[code].sourceIndex !== rowInfo.originalIndex
      ) {
        continue;
      }
    }

    cleanedRows.push(rowInfo);
  }

  /*
   * 원본 열 순서와 무관하게, 이름으로 찾은 위치에서 고정된 순서로 뽑음
   * (색상 열만 표시값을 써서 앞자리 0을 보존)
   */
  const outputHeader = CLEAN_SHEET_TARGET_HEADER.slice();

  const outputRows = cleanedRows.map(function(rowInfo) {
    return columnIndexes.map(function(idx) {
      if (idx === colorColIndex) {
        return sourceDisplayValues[rowInfo.originalIndex][idx];
      }
      return rowInfo.values[idx];
    });
  });

  const result = [
    outputHeader,
    ...outputRows
  ];

  /*
   * 정리 시트 삭제 없이 내용만 초기화
   */
  clearFixedSheet_(cleanSheet);

  ensureSheetSize_(
    cleanSheet,
    result.length,
    result[0].length
  );

  // 색상 열은 값을 쓰기 전에 먼저 텍스트("@") 서식으로 고정해둡니다.
  // 서식이 "General"인 채로 "061"처럼 숫자처럼 보이는 문자열을 쓰면
  // 그 즉시 숫자 61로 바뀔 수 있어서, 쓰기 전에 서식부터 지정합니다
  // (applyBasicFormats_도 이 열은 절대 General로 안 바꾸도록 되어 있어서
  // 이후 서식 적용 과정에서 다시 숫자로 바뀌는 일은 없습니다).
  if (result.length >= 2) {
    cleanSheet
      .getRange(2, COLOR_COLUMN_POSITION, result.length - 1, 1)
      .setNumberFormat("@");
  }

  cleanSheet
    .getRange(
      1,
      1,
      result.length,
      result[0].length
    )
    .setValues(result);

  applyBasicFormats_(
    cleanSheet,
    result.length,
    result[0].length
  );

  cleanSheet.setFrozenRows(1);
}


/**************************************************************
 * 2단계
 * 추가건 / 삭제건 비교
 **************************************************************/
function compareAddDelete_(config) {
  const cleanSheet = getRequiredSheet_(config.cleanSheet);
  const accumulatedSheet =
    getRequiredSheet_(config.accumulatedSheet);

  const addSheet = getRequiredSheet_(config.addSheet);
  const deleteSheet = getRequiredSheet_(config.deleteSheet);

  const cleanData = cleanSheet
    .getDataRange()
    .getValues();

  const accumulatedData = accumulatedSheet
    .getDataRange()
    .getValues();

  if (cleanData.length < 2) {
    throw new Error(
      "'" + config.cleanSheet + "' 시트에 비교할 자료가 없습니다."
    );
  }

  // 누적데이터 시트가 헤더만 있고 데이터 행이 0개인 것(달이 바뀌어서
  // "1,2공장 누적데이터 매칭"이 아직 아무것도 못 채운 상태)은 정상
  // 상황입니다 — 이 경우 금주 정리 시트의 모든 행이 그대로 추가건이
  // 되어야 하고, 아래 cleanMap/accumulatedMap 비교 로직이 이미 그렇게
  // 처리합니다. 헤더 행조차 없는(완전히 빈) 시트일 때만 에러로 알립니다.
  if (accumulatedData.length < 1) {
    throw new Error(
      "'" + config.accumulatedSheet + "' 시트에 헤더가 없습니다."
    );
  }

  /*
   * 정리 시트의 제목 열 수를 기준으로 통일
   */
  const header = cleanData[0];
  const targetColumnCount = header.length;

  // 두 시트 모두 "접수번호" 열 위치를 이름으로 직접 찾습니다(예전엔 D열
  // 고정이었는데, 두 시트가 항상 같은 순서라는 보장이 없어서 — 순서가
  // 어긋나면 엉뚱한 열로 비교하게 되어 추가건/삭제건이 조용히 잘못
  // 계산될 수 있었습니다).
  const cleanReceiptNoIdx = header.indexOf("접수번호");

  if (cleanReceiptNoIdx === -1) {
    throw new Error("'" + config.cleanSheet + "'에서 '접수번호' 열을 찾을 수 없습니다.");
  }

  const accumulatedHeader = accumulatedData[0];
  const accumulatedReceiptNoIdx = accumulatedHeader.indexOf("접수번호");

  if (accumulatedReceiptNoIdx === -1) {
    throw new Error("'" + config.accumulatedSheet + "'에서 '접수번호' 열을 찾을 수 없습니다.");
  }

  const cleanMap = new Map();
  const accumulatedMap = new Map();

  /*
   * 정리 시트 접수번호 열을 비교 키로 사용
   */
  for (let i = 1; i < cleanData.length; i++) {
    const key = normalizeText_(cleanData[i][cleanReceiptNoIdx]);

    if (key) {
      cleanMap.set(
        key,
        normalizeRowLength_(
          cleanData[i],
          targetColumnCount
        )
      );
    }
  }

  /*
   * 누적데이터 접수번호 열을 비교 키로 사용
   */
  for (let i = 1; i < accumulatedData.length; i++) {
    const key = normalizeText_(accumulatedData[i][accumulatedReceiptNoIdx]);

    if (key) {
      accumulatedMap.set(
        key,
        normalizeRowLength_(
          accumulatedData[i],
          targetColumnCount
        )
      );
    }
  }

  const addRows = [];
  const deleteRows = [];

  /*
   * 금주에는 있고 누적에는 없는 자료
   */
  cleanMap.forEach(function(row, key) {
    if (!accumulatedMap.has(key)) {
      addRows.push(row);
    }
  });

  /*
   * 누적에는 있고 금주에는 없는 자료
   */
  accumulatedMap.forEach(function(row, key) {
    if (!cleanMap.has(key)) {
      deleteRows.push(row);
    }
  });

  writeCompareSheet_(
    addSheet,
    header,
    addRows
  );

  writeCompareSheet_(
    deleteSheet,
    header,
    deleteRows
  );
}


/**************************************************************
 * 행의 열 개수를 목표 열 수에 맞춤
 *
 * 열이 부족하면 오른쪽에 빈칸 추가
 * 열이 많으면 초과 열 제거
 **************************************************************/
function normalizeRowLength_(row, targetColumnCount) {
  const result = row.slice(0, targetColumnCount);

  while (result.length < targetColumnCount) {
    result.push("");
  }

  return result;
}


/**************************************************************
 * 추가건 / 삭제건 결과 입력
 **************************************************************/
function writeCompareSheet_(sheet, header, rows) {
  const columnCount = header.length;

  const normalizedHeader =
    normalizeRowLength_(header, columnCount);

  const normalizedRows = rows.map(function(row) {
    return normalizeRowLength_(row, columnCount);
  });

  const result = [
    normalizedHeader,
    ...normalizedRows
  ];

  clearFixedSheet_(sheet);

  ensureSheetSize_(
    sheet,
    result.length,
    columnCount
  );

  // 색상 열(정리 시트와 같은 고정 위치)도 쓰기 전에 텍스트("@")로
  // 고정해서, 정리 시트에서 넘어온 "061" 같은 값이 여기서도 숫자로
  // 바뀌지 않게 합니다.
  if (result.length >= 2 && columnCount >= COLOR_COLUMN_POSITION) {
    sheet
      .getRange(2, COLOR_COLUMN_POSITION, result.length - 1, 1)
      .setNumberFormat("@");
  }

  sheet
    .getRange(
      1,
      1,
      result.length,
      columnCount
    )
    .setValues(result);

  applyBasicFormats_(
    sheet,
    result.length,
    columnCount
  );

  sheet.setFrozenRows(1);
}


/**************************************************************
 * 날짜 및 숫자 표시 형식
 **************************************************************/
function applyBasicFormats_(sheet, rowCount, columnCount) {
  if (
    rowCount < 1 ||
    columnCount < 1
  ) {
    return;
  }

  // 색상 열(COLOR_COLUMN_POSITION)은 General 범위에서 제외합니다. 이미
  // 텍스트로 저장된 "061" 같은 값도 서식을 General로 바꾸는 순간 시트가
  // 숫자로 재해석해서 앞자리 0이 사라지기 때문입니다 — 그 열 앞/뒤만
  // 나눠서 General을 적용하고, 그 열 자체는 항상 텍스트("@")로 고정합니다.
  const hasColorColumn = columnCount >= COLOR_COLUMN_POSITION;

  if (hasColorColumn) {
    if (COLOR_COLUMN_POSITION > 1) {
      sheet
        .getRange(1, 1, rowCount, COLOR_COLUMN_POSITION - 1)
        .setNumberFormat("General");
    }

    if (columnCount > COLOR_COLUMN_POSITION) {
      sheet
        .getRange(
          1,
          COLOR_COLUMN_POSITION + 1,
          rowCount,
          columnCount - COLOR_COLUMN_POSITION
        )
        .setNumberFormat("General");
    }

    sheet
      .getRange(1, COLOR_COLUMN_POSITION, rowCount, 1)
      .setNumberFormat("@");
  } else {
    sheet
      .getRange(
        1,
        1,
        rowCount,
        columnCount
      )
      .setNumberFormat("General");
  }

  if (rowCount < 2) {
    return;
  }

  /*
   * A열 최종조치일
   */
  if (columnCount >= 1) {
    sheet
      .getRange(
        2,
        1,
        rowCount - 1,
        1
      )
      .setNumberFormat("yyyy-mm-dd");
  }

  /*
   * K열 수량
   */
  if (columnCount >= 11) {
    sheet
      .getRange(
        2,
        11,
        rowCount - 1,
        1
      )
      .setNumberFormat("0");
  }

  /*
   * L열 금액
   */
  if (columnCount >= 12) {
    sheet
      .getRange(
        2,
        12,
        rowCount - 1,
        1
      )
      .setNumberFormat("#,##0");
  }

  /*
   * P열 반납일자
   */
  if (columnCount >= 16) {
    sheet
      .getRange(
        2,
        16,
        rowCount - 1,
        1
      )
      .setNumberFormat("yyyy-mm-dd");
  }
}


/**************************************************************
 * 3단계
 * 회수내역 매칭
 *
 * 1공장과 2공장 모두 동일한 회수내역 시트 사용
 **************************************************************/
function matchRecovery_(config) {
  const targetSheet = getRequiredSheet_(config.addSheet);
  const recoverySheet =
    getRequiredSheet_(RECOVERY_SHEET_NAME);

  const targetLastRow = targetSheet.getLastRow();
  const recoveryLastRow = recoverySheet.getLastRow();
  const recoveryLastColumn = recoverySheet.getLastColumn();

  /*
   * 추가건이 없으면 제목 처리를 위해 26열만 확보
   */
  if (targetLastRow < 2) {
    ensureSheetSize_(
      targetSheet,
      1,
      26
    );

    return;
  }

  if (recoveryLastRow < 2) {
    throw new Error(
      "'" + RECOVERY_SHEET_NAME + "' 시트에 매칭할 자료가 없습니다."
    );
  }

  if (recoveryLastColumn < 15) {
    throw new Error(
      "'" + RECOVERY_SHEET_NAME + "' 시트의 열 수가 부족합니다.\n" +
      "최소 O열까지 자료가 있어야 합니다."
    );
  }

  ensureSheetSize_(
    targetSheet,
    targetSheet.getMaxRows(),
    26
  );

  const targetData = targetSheet
    .getRange(
      1,
      1,
      targetLastRow,
      26
    )
    .getValues();

  const recoveryData = recoverySheet
    .getRange(
      1,
      1,
      recoveryLastRow,
      recoveryLastColumn
    )
    .getValues();

  const recoveryExactMap = {};

  /*
   * 회수내역 C열을 기준으로 매칭표 생성
   *
   * "I202606230210-01"/"I202606230210-02"처럼 같은 접수번호에 순번이
   * 다른 여러 건이 있으면 접수번호 앞부분만으로는 서로 구분이 안 되어
   * 뒤섞일 수 있어서, 전체 문자열(접수번호+순번)이 정확히 같을 때만
   * 매칭합니다. 정확히 일치하는 회수내역이 없으면 미회수로 처리됩니다.
   */
  for (let i = 1; i < recoveryData.length; i++) {
    const sourceValue =
      normalizeText_(recoveryData[i][2]);

    if (!sourceValue) continue;

    const exactKey = sourceValue.toLowerCase();

    if (!recoveryExactMap[exactKey]) {
      recoveryExactMap[exactKey] = recoveryData[i];
    }
  }

  /*
   * 추가건 D열과 회수내역 C열 매칭 (전체 문자열 정확히 일치하는 것만)
   */
  for (let i = 1; i < targetData.length; i++) {
    const targetValue =
      normalizeText_(targetData[i][3]);

    const matchRow = recoveryExactMap[targetValue.toLowerCase()];

    /*
     * R~Z열 이전 결과 초기화
     */
    for (
      let columnIndex = 17;
      columnIndex <= 25;
      columnIndex++
    ) {
      targetData[i][columnIndex] = "";
    }

    if (matchRow) {
      targetData[i][17] = matchRow[10]; // R 유형
      targetData[i][18] = matchRow[9];  // S 세부유형
      targetData[i][19] = matchRow[11]; // T 하자상세
      targetData[i][20] = matchRow[12]; // U 로트

      targetData[i][21] = matchRow[6];  // V 원인
      targetData[i][22] = matchRow[6];  // W 포장

      targetData[i][24] = matchRow[13]; // Y 유형분류
      targetData[i][25] = matchRow[14]; // Z 유형분류

      /*
       * 회수내역 G열 값이 포장/원인 형태이면 분리
       *
       * 앞부분 → W열 포장
       * 뒷부분 → V열 원인
       */
      const combinedValue =
        normalizeText_(targetData[i][21]);

      if (combinedValue.includes("/")) {
        const splitData = combinedValue.split("/");

        targetData[i][22] =
          normalizeText_(splitData[0]);

        targetData[i][21] =
          normalizeText_(splitData[1]);
      }
    } else {
      targetData[i][17] = "미회수";
    }
  }

  targetSheet
    .getRange(
      1,
      1,
      targetData.length,
      26
    )
    .setValues(targetData);
}


/**************************************************************
 * 4단계
 * 제목 입력 및 정렬
 *
 * 최종 정렬 순서
 * ------------------------------------------------------------
 * 일반 → 감성 → 취급 → VN → 미회수
 **************************************************************/
function sortAndFormat_(config) {
  const sheet = getRequiredSheet_(config.addSheet);

  ensureSheetSize_(
    sheet,
    sheet.getMaxRows(),
    28
  );

  /*
   * AB열이 숨겨져 있으면 잠시 표시
   */
  try {
    sheet.showColumns(28);
  } catch (ignore) {
    // 이미 표시 상태이면 무시
  }

  const headers = [
    "등록일",     // Q
    "유형",       // R
    "세부유형",   // S
    "하자상세",   // T
    "로트",       // U
    "원인",       // V
    "포장",       // W
    "제조자",     // X
    "유형분류",   // Y
    "유형분류"    // Z
  ];

  sheet
    .getRange(
      1,
      17,
      1,
      headers.length
    )
    .setValues([headers]);

  const lastRow = sheet.getLastRow();

  /*
   * 추가건이 없으면 제목만 유지
   */
  if (lastRow < 2) {
    sheet
      .getRange(
        1,
        28,
        sheet.getMaxRows(),
        1
      )
      .clearContent();

    sheet.hideColumns(28);
    sheet.setFrozenRows(1);

    return;
  }

  /*
   * 현재 월 주차 계산
   *
   * new Date()의 getFullYear()/getMonth()는 스크립트 프로젝트에 설정된
   * 시간대를 따르는데, 그 설정이 Asia/Seoul이 아니면(기본값이 다르게
   * 잡혀 있는 경우가 있음) 자정 근처에서 주차가 하루씩 밀릴 수 있어서,
   * Utilities.formatDate로 한국 시간 기준 "일(day)"만 직접 구합니다.
   */
  const dayOfMonth = Number(Utilities.formatDate(new Date(), "Asia/Seoul", "d"));
  const weekNumber = Math.floor((dayOfMonth - 1) / 7) + 1;

  /*
   * Q열 주차
   */
  sheet
    .getRange(
      2,
      17,
      lastRow - 1,
      1
    )
    .setValue(weekNumber + "주");

  /*
   * X열 제조자
   *
   * 1공장: 충주1
   * 2공장: 충주2
   */
  sheet
    .getRange(
      2,
      24,
      lastRow - 1,
      1
    )
    .setValue(config.manufacturer);

  const temporaryColumn = 28;

  sheet
    .getRange(1, temporaryColumn)
    .setValue("정렬우선순위");

  let data = sheet
    .getRange(
      1,
      1,
      lastRow,
      temporaryColumn
    )
    .getValues();

  /*
   * R열 유형이 미회수인 경우
   *
   * S열 세부유형 = .
   * T열 하자상세 = .
   * U열 로트 = .
   * V열 원인 = 미회수
   * W열 포장 = 미회수
   * Y열 유형분류 = .
   * Z열 유형분류 = .
   */
  for (let i = 1; i < data.length; i++) {
    const typeValue =
      normalizeText_(data[i][17]);

    if (typeValue === "미회수") {
      data[i][18] = ".";
      data[i][19] = ".";
      data[i][20] = ".";

      data[i][21] = "미회수";
      data[i][22] = "미회수";

      data[i][24] = ".";
      data[i][25] = ".";
    }
  }

  const header = data[0];

  const normalRows = [];
  const emotionRows = [];
  const handlingRows = [];
  const vnRows = [];
  const uncollectedRows = [];

  /*
   * 그룹 분류
   *
   * 감성 및 취급은 원인/포장과 관계없이
   * 각각 감성, 취급 그룹으로 이동합니다.
   */
  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const typeValue =
      normalizeText_(row[17]); // R 유형

    const causeValue =
      normalizeText_(row[21]); // V 원인

    const packageValue =
      normalizeText_(row[22]); // W 포장

    if (typeValue === "미회수") {
      uncollectedRows.push(row);
    } else if (typeValue === "감성") {
      emotionRows.push(row);
    } else if (typeValue === "취급") {
      handlingRows.push(row);
    } else if (
      causeValue === "VN" ||
      packageValue === "VN"
    ) {
      vnRows.push(row);
    } else {
      normalRows.push(row);
    }
  }

  /*
   * 각 그룹 내부 정렬
   */
  sortRowsWithinGroup_(normalRows);
  sortRowsWithinGroup_(emotionRows);
  sortRowsWithinGroup_(handlingRows);
  sortRowsWithinGroup_(vnRows);
  sortRowsWithinGroup_(uncollectedRows);

  /*
   * 최종 정렬 순서
   */
  const sortedData = [
    header,
    ...normalRows,
    ...emotionRows,
    ...handlingRows,
    ...vnRows,
    ...uncollectedRows
  ];

  sheet
    .getRange(
      1,
      1,
      sortedData.length,
      temporaryColumn
    )
    .setValues(sortedData);

  /*
   * 이전 실행 자료가 아래쪽에 남지 않도록 제거
   */
  const maxRows = sheet.getMaxRows();

  if (sortedData.length < maxRows) {
    sheet
      .getRange(
        sortedData.length + 1,
        1,
        maxRows - sortedData.length,
        temporaryColumn
      )
      .clearContent();
  }

  /*
   * AB 임시열은 삭제하지 않고 내용만 제거
   */
  sheet
    .getRange(
      1,
      temporaryColumn,
      sheet.getMaxRows(),
      1
    )
    .clearContent();

  sheet.hideColumns(temporaryColumn);

  /*
   * 날짜 및 숫자 형식
   */
  applyBasicFormats_(
    sheet,
    sortedData.length,
    Math.min(sortedData[0].length, 26)
  );

  sheet.setFrozenRows(1);
}


/**************************************************************
 * 그룹 내부 정렬
 *
 * 1차 W열 포장
 * 2차 S열 세부유형
 * 3차 V열 원인
 * 4차 R열 유형
 * 5차 D열 접수번호
 **************************************************************/
function sortRowsWithinGroup_(rows) {
  rows.sort(function(a, b) {
    let result = compareKoreanText_(
      a[22],
      b[22]
    );

    if (result !== 0) {
      return result;
    }

    result = compareKoreanText_(
      a[18],
      b[18]
    );

    if (result !== 0) {
      return result;
    }

    result = compareKoreanText_(
      a[21],
      b[21]
    );

    if (result !== 0) {
      return result;
    }

    result = compareKoreanText_(
      a[17],
      b[17]
    );

    if (result !== 0) {
      return result;
    }

    return compareKoreanText_(
      a[3],
      b[3]
    );
  });
}


/**************************************************************
 * 한글 및 숫자 포함 문자 정렬
 **************************************************************/
function compareKoreanText_(valueA, valueB) {
  const textA = normalizeText_(valueA);
  const textB = normalizeText_(valueB);

  return textA.localeCompare(
    textB,
    "ko",
    {
      numeric: true,
      sensitivity: "base"
    }
  );
}


/**************************************************************
 * 문자 정리
 **************************************************************/
function normalizeText_(value) {
  return String(
    value === null || value === undefined
      ? ""
      : value
  ).trim();
}


/**************************************************************
 * 오류 알림
 **************************************************************/
function showError_(factoryName, error) {
  const message =
    error && error.message
      ? error.message
      : String(error);

  SpreadsheetApp.getUi().alert(
    factoryName + " 작업 중 오류가 발생했습니다.\n\n" +
    message
  );
}
