/**
 * 품질 고객클레임 대시보드 2026년 마감자료 자동화
 *
 * 설치 방법:
 *  1) 편집 가능한 목적지 구글시트(https://docs.google.com/spreadsheets/d/1lr-KbRyQ3zmcnT-RvRS2-krVnMl1cGndHRxEVpD6ZWM/edit)를 엽니다.
 *  2) 확장 프로그램 > Apps Script 를 클릭합니다.
 *  3) 기본 생성된 코드를 지우고 이 파일 전체를 붙여넣습니다.
 *  4) CONFIG 값들을 확인/수정합니다 (특히 SOURCE_WEBAPP_URL, TARGET_SHEET_GID).
 *  5) 함수 선택 드롭다운에서 updateCurrentMonthWeeklyCounts 를 선택해 한 번 실행하고
 *     권한 요청을 승인합니다. 실행 로그(보기 > 로그)로 결과를 확인하세요.
 *  6) 문제 없으면 installMonthlyTrigger 를 한 번 실행해 매달 1일 자동 실행 트리거를 등록합니다.
 *
 * 대시보드 연동
 * ------------------------------------------------------------
 * claim-dashboard-webapp.gs를 이 파일과 같은 프로젝트에 추가하면,
 * factory-as-sample/index.html 대시보드의 "클레임 대시보드" 탭에서
 * 이 시트를 읽기 전용으로 보고, "이번달 주간 데이터 채우기" /
 * "다음달로 표 넘기기" 버튼으로 아래 함수들을 원격 실행할 수 있습니다.
 * "이번달 주간 데이터 채우기"는 대시보드가 "26년 마감자료" 웹앱 URL을
 * body.sourceWebappUrl로 매번 같이 보내주므로, 아래 CONFIG.SOURCE_WEBAPP_URL은
 * 이 함수를 스프레드시트 메뉴에서 직접 실행할 때만 쓰는 예비값입니다.
 *
 * 원본 데이터 소스 (2026-08 변경)
 * ------------------------------------------------------------
 * 예전에는 "가져올 N월 자료" cost 시트를 공개 CSV로 읽었는데, 그 cost
 * 시트가 달이 바뀌어도 지난달 데이터를 지우지 않고 계속 누적되는
 * 방식이라(등록일 "N주" 값이 지난달 것까지 뒤섞여 있어서), 이번달 1주만
 * 있어야 하는데 지난달의 2주~5주까지 잘못 채워지는 문제가 있었습니다.
 * 그래서 "26년 마감자료" 스프레드시트(달마다 "N월마감(건수)" 탭으로
 * 완전히 분리되어 있음)를 웹앱으로 배포해서 쓰기로 바꿨습니다. 이
 * 웹앱은 action 파라미터와 무관하게 요청할 때마다 { months: { "1월":
 * [...], ..., "8월": [...] } } 형태로 전체 월 데이터를 한 번에 돌려주므로,
 * fetchAndAggregateSource_는 그중 이번달(monthNum) 배열만 꺼내 씁니다.
 */

const CONFIG = {
  // "26년 마감자료" 웹앱 URL의 예비값(대시보드 버튼은 항상 자기 값을
  // 보내주므로 보통 안 쓰임 — 메뉴에서 이 함수를 직접 실행할 때만 씀).
  SOURCE_WEBAPP_URL: "",

  // 목적지 시트(전체/퍼시스/일룸/데스커/외작/내작 표가 들어있는 탭)의 gid
  TARGET_SHEET_GID: 1130212422,

  BRANDS: ["전체", "퍼시스", "일룸", "데스커", "외작", "내작"],
  GUBUN_SET: new Set(["시공미결", "고객불만", "기타"]),
  WEEK_LABELS: ["1주", "2주", "3주", "4주", "5주"]
};

/** 시트를 열 때 수동 실행용 메뉴 추가 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("클레임 대시보드 자동화")
    .addItem("이번달 주간 데이터 채우기", "updateCurrentMonthWeeklyCounts")
    .addItem("다음달로 표 넘기기 (수동 실행)", "reshapeForNewMonth")
    .addItem("매달 자동 표 넘기기 트리거 설치", "installMonthlyTrigger")
    .addToUi();
}

/* ------------------------------------------------------------------ */
/* 1) 이번달 주차별 값 채우기                                            */
/* ------------------------------------------------------------------ */

function updateCurrentMonthWeeklyCounts(sourceWebappUrl) {
  const effectiveSourceUrl = sourceWebappUrl || CONFIG.SOURCE_WEBAPP_URL;
  if (!effectiveSourceUrl) {
    throw new Error(
      "26년 마감자료 웹 앱 URL이 없습니다. 대시보드에서 입력/저장했는지 확인하거나, " +
      "메뉴에서 직접 실행하는 경우 CONFIG.SOURCE_WEBAPP_URL을 채워주세요."
    );
  }

  const sheet = getTargetSheet_();
  const grid = sheet.getDataRange().getValues();

  const { headerRowIdx, subHeaderRowIdx } = findHeaderRows_(grid);
  const monthNum = currentMonthNum_();
  const monthLabel = `${monthNum}월`;
  const monthCol = grid[headerRowIdx].indexOf(monthLabel);
  if (monthCol === -1) {
    throw new Error(`헤더 행에서 "${monthLabel}" 컬럼을 찾을 수 없습니다.`);
  }

  const isExpanded = String(grid[subHeaderRowIdx][monthCol]).trim() === "1주";
  if (!isExpanded) {
    throw new Error(
      `${monthLabel} 칸이 아직 주차별(1주~5주)로 펼쳐져 있지 않습니다. ` +
      `reshapeForNewMonth()를 먼저 실행해서 표를 펼친 뒤 다시 시도하세요.`
    );
  }
  const weekCols = [0, 1, 2, 3, 4].map((i) => monthCol + i); // 1주..5주

  const leafRows = getLeafRows_(grid, headerRowIdx);
  const aggregate = fetchAndAggregateSource_(effectiveSourceUrl, monthNum);

  const presentWeeks = new Set();
  Object.values(aggregate).forEach((byKey) =>
    Object.values(byKey).forEach((byWeek) =>
      Object.keys(byWeek).forEach((w) => presentWeeks.add(w))
    )
  );

  let updatedCells = 0;
  leafRows.forEach((leaf) => {
    const brandTable = aggregate[leaf.brand];
    const key = `${leaf.gubun}|${leaf.type}|${leaf.detail}`;
    const byWeek = brandTable ? brandTable[key] : undefined;

    // weekCols[0..4]는 항상 붙어있는 5개 열이라, 한 행씩 5칸을 모아서
    // 한 번에 씁니다(전에는 주당 최대 5번 clearContent/setValue를 따로
    // 호출해서 leaf 행마다 API 호출이 5번씩 들었음).
    const weekValues = CONFIG.WEEK_LABELS.map((weekLabel) => {
      if (!presentWeeks.has(weekLabel)) {
        // 아직 시작 안 된 주차는 빈 칸으로(0으로 채우면 "확인해봤는데
        // 0건"처럼 보여서 미래 주차와 구분이 안 됨). 이미 0이 들어가있던
        // 셀도 이 함수를 다시 실행하면 여기서 지워짐(자동 복구).
        return "";
      }
      updatedCells++;
      return byWeek && byWeek[weekLabel] ? byWeek[weekLabel] : 0;
    });

    sheet.getRange(leaf.row + 1, weekCols[0] + 1, 1, weekCols.length).setValues([weekValues]);

    // 해당 월의 합계(주차 5개 합) 칸은 항상 SUM 수식으로 고정
    const totalCol = monthCol + 5;
    const weekRangeA1 = `${columnToLetter_(weekCols[0] + 1)}${leaf.row + 1}:${columnToLetter_(weekCols[4] + 1)}${leaf.row + 1}`;
    sheet.getRange(leaf.row + 1, totalCol + 1).setFormula(`=SUM(${weekRangeA1})`);
  });

  Logger.log(`${monthLabel} 데이터 반영 완료: leaf row ${leafRows.length}개, 셀 ${updatedCells}개 갱신`);
}

/* ------------------------------------------------------------------ */
/* 2) 월 전환: 지난달 접기 + 이번달 펼치기                                 */
/* ------------------------------------------------------------------ */

/** 매달 1일에 트리거로(또는 메뉴에서 수동으로) 실행: 지난달을 접고 이번달을 주차별로 펼친다. */
function reshapeForNewMonth() {
  const sheet = getTargetSheet_();
  const grid = sheet.getDataRange().getValues();
  const headerRows = findAllHeaderRows_(grid);

  const monthNum = currentMonthNum_();
  const currentMonthLabel = `${monthNum}월`;
  const prevMonthNum = monthNum === 1 ? null : monthNum - 1;
  if (prevMonthNum === null) {
    throw new Error("1월 전환(연말->연초)은 이 스크립트가 자동 처리하지 않습니다. 수동으로 확인하세요.");
  }
  const prevMonthLabel = `${prevMonthNum}월`;

  const currentMonthCol = grid[headerRows[0]].indexOf(currentMonthLabel);
  const prevMonthCol = grid[headerRows[0]].indexOf(prevMonthLabel);
  if (currentMonthCol === -1) throw new Error(`"${currentMonthLabel}" 헤더를 찾을 수 없습니다.`);
  if (prevMonthCol === -1) throw new Error(`"${prevMonthLabel}" 헤더를 찾을 수 없습니다.`);

  const subHeaderRowIdx = headerRows[0] + 1;
  const currentIsExpanded = String(grid[subHeaderRowIdx][currentMonthCol]).trim() === "1주";
  const prevIsExpanded = String(grid[subHeaderRowIdx][prevMonthCol]).trim() === "1주";

  if (currentIsExpanded && !prevIsExpanded) {
    Logger.log(`${currentMonthLabel}은(는) 이미 펼쳐져 있고 ${prevMonthLabel}은(는) 이미 접혀 있어 할 일이 없습니다.`);
    return;
  }

  // 리프 행(실제 유형상세 데이터 행)은 컬럼 삽입/삭제와 무관하게 행 번호가 그대로이므로 미리 한 번만 계산해도 된다.
  const leafRows = getLeafRows_(grid);

  // --- 1) 이번달 칸 펼치기: 아직 6칸(1주~5주,합계)으로 안 펼쳐져 있으면, 지난달의 서식/수식을 통째로 복사해 온다 ---
  if (!currentIsExpanded) {
    const maxRows = sheet.getMaxRows();
    const templateRange = sheet.getRange(1, prevMonthCol + 1, maxRows, 6);

    sheet.insertColumnsAfter(currentMonthCol + 1, 5); // 이번달 칸 뒤에 5칸 추가 -> 총 6칸
    const destRange = sheet.getRange(1, currentMonthCol + 1, maxRows, 6);
    templateRange.copyTo(destRange); // 값+수식+서식+병합까지 그대로 복사 (상대참조라 열 위치는 자동 보정됨)

    // 복사돼 온 헤더 라벨("N월")을 이번달 라벨로 교체 (표 6개 전부에 대해 한 번에 처리)
    headerRows.forEach((hr) => {
      sheet.getRange(hr + 1, currentMonthCol + 1).setValue(currentMonthLabel);
    });

    // 리프 행의 1주~5주 값은 지난달 실측치가 복사되어 온 것이므로 지워서
    // 빈 칸으로 초기화. 0으로 채우면 "그 주차에 실제로 확인해봤는데
    // 0건이었다"처럼 보여서, 아직 시작도 안 한 미래 주차와 구분이 안 되는
    // 문제가 있었음 — updateCurrentMonthWeeklyCounts가 실제 그 주차
    // 데이터를 만나면 그때 값(0 포함)을 채워 넣음. 합계 칸은 그대로 둔다
    // (지난달 서식이 수식이었다면 빈 칸 기준으로 0으로 자동 재계산됨).
    const weekCols = [0, 1, 2, 3, 4].map((i) => currentMonthCol + i);
    leafRows.forEach((leaf) => {
      weekCols.forEach((col) => sheet.getRange(leaf.row + 1, col + 1).clearContent());
    });

    Logger.log(`${currentMonthLabel} 칸을 ${prevMonthLabel} 서식을 복사해 1주~5주로 펼쳤습니다.`);
  } else {
    Logger.log(`${currentMonthLabel}은(는) 이미 펼쳐져 있어 건너뜁니다.`);
  }

  // --- 2) 지난달 칸 접기: leaf/소계/합계 행 전부의 5주 합을 구해 단일 칸에 남기고 나머지 5칸은 삭제 ---
  if (prevIsExpanded) {
    const weekCols = [0, 1, 2, 3, 4].map((i) => prevMonthCol + i);
    const contentRows = getContentRows_(grid); // leaf + 소계 + 합계 행 전부 (헤더/서브헤더 행 제외)

    contentRows.forEach((r) => {
      const total = weekCols.reduce((sum, col) => sum + (Number(grid[r][col]) || 0), 0);
      sheet.getRange(r + 1, prevMonthCol + 1).setValue(total);
    });

    headerRows.forEach((hr) => {
      sheet.getRange(hr + 1, prevMonthCol + 1, 1, 6).breakApart();
      sheet.getRange(hr + 1, prevMonthCol + 1).setValue(prevMonthLabel);
      sheet.getRange(hr + 2, prevMonthCol + 1).clearContent(); // 서브헤더의 "1주" 잔여 텍스트 제거

      // breakApart 후 남은 옛 병합 셀의 자투리 서식(테두리 등)을 지우고,
      // 바로 왼쪽(이미 접혀 있는 이전 달) 헤더/서브헤더 칸의 서식을 그대로 복사해 시각적으로 통일한다.
      sheet
        .getRange(hr + 1, prevMonthCol, 1, 1)
        .copyFormatToRange(sheet, prevMonthCol + 1, prevMonthCol + 1, hr + 1, hr + 1);
      sheet
        .getRange(hr + 2, prevMonthCol, 1, 1)
        .copyFormatToRange(sheet, prevMonthCol + 1, prevMonthCol + 1, hr + 2, hr + 2);

      // 1~6월/8~12월처럼 데이터가 없는(단일 칸) 달은 헤더 행과 서브헤더 행이 세로로 병합돼 있으므로
      // 이번에 접은 달도 똑같이 세로 병합해서 모양을 맞춘다.
      sheet.getRange(hr + 1, prevMonthCol + 1, 2, 1).merge();
    });

    // 열 너비: 6칸으로 펼쳐져 있던 시절의 좁은 폭이 그대로 남아있으므로,
    // 바로 왼쪽(이미 단일 칸인 이전 달) 컬럼과 같은 폭으로 맞춘다.
    sheet.setColumnWidth(prevMonthCol + 1, sheet.getColumnWidth(prevMonthCol));

    // 나머지 5칸(2주,3주,4주,5주,합계) 삭제 -> 뒤 컬럼들이 자동으로 왼쪽으로 당겨짐
    sheet.deleteColumns(prevMonthCol + 2, 5);

    Logger.log(`${prevMonthLabel} 칸을 접었습니다 (합계만 남김).`);
  } else {
    Logger.log(`${prevMonthLabel}은(는) 이미 접혀 있어 건너뜁니다.`);
  }
}

/** 매달 1일 오전 6시에 reshapeForNewMonth 자동 실행되도록 트리거 등록 */
function installMonthlyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === "reshapeForNewMonth")
    .forEach((t) => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("reshapeForNewMonth")
    .timeBased()
    .onMonthDay(1)
    .atHour(6)
    .create();

  Logger.log("매달 1일 06:00 자동 실행 트리거를 등록했습니다.");
}

/* ------------------------------------------------------------------ */
/* 공통 유틸                                                            */
/* ------------------------------------------------------------------ */

function getTargetSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets().find((s) => s.getSheetId() === CONFIG.TARGET_SHEET_GID);
  if (!sheet) throw new Error(`gid=${CONFIG.TARGET_SHEET_GID} 시트를 찾을 수 없습니다.`);
  return sheet;
}

/** "구분" 텍스트가 있는 메인 헤더 행과 그 바로 아래 서브헤더 행 인덱스(0-based)를 찾는다 (첫 번째 표 기준) */
function findHeaderRows_(grid) {
  const headerRowIdx = grid.findIndex((row) => String(row[1]).trim() === "구분");
  if (headerRowIdx === -1) throw new Error('헤더 행("구분")을 찾을 수 없습니다.');
  return { headerRowIdx, subHeaderRowIdx: headerRowIdx + 1 };
}

/** 시트 안에 반복되는 표(전체/퍼시스/일룸/데스커/외작/내작) 각각의 "구분" 헤더 행 인덱스(0-based)를 전부 찾는다 */
function findAllHeaderRows_(grid) {
  const rows = [];
  grid.forEach((row, r) => {
    if (String(row[1]).trim() === "구분") rows.push(r);
  });
  if (!rows.length) throw new Error('헤더 행("구분")을 찾을 수 없습니다.');
  return rows;
}

/**
 * 헤더/서브헤더 행을 제외한, 실제 값이 있는 모든 행(leaf 행 + 소계 행 + 합계 행)을 찾는다.
 * reshapeForNewMonth의 "지난달 접기"에서 소계/합계까지 포함해 합계를 다시 계산할 때 사용.
 * 반환: 0-based row 인덱스 배열
 */
function getContentRows_(grid) {
  const rows = [];
  let active = false;
  for (let r = 0; r < grid.length; r++) {
    const b = String(grid[r][1] || "").trim();
    const c = String(grid[r][2] || "").trim();
    const d = String(grid[r][3] || "").trim();

    if (CONFIG.BRANDS.includes(b) && !c && !d) {
      active = false;
      continue;
    }
    if (b === "구분") {
      active = true;
      continue;
    }
    if (!active) continue;
    if (b || c || d) rows.push(r);
  }
  return rows;
}

/**
 * 시트 전체를 스캔해 브랜드 블록(전체/퍼시스/일룸/데스커/외작/내작)별로
 * leaf row(실제 유형상세 값이 채워지는 데이터 행)를 찾아낸다.
 * 반환: [{ row, brand, gubun, type, detail }]  (row는 grid 배열 기준 0-based 인덱스)
 */
function getLeafRows_(grid, firstHeaderRowIdx) {
  const leaves = [];
  let currentBrand = null;
  let currentGubun = null;
  let currentType = null;
  let sawHeaderForBlock = false;

  for (let r = 0; r < grid.length; r++) {
    const row = grid[r];
    const b = String(row[1] || "").trim();
    const c = String(row[2] || "").trim();
    const d = String(row[3] || "").trim();

    if (CONFIG.BRANDS.includes(b) && !c && !d) {
      currentBrand = b;
      currentGubun = null;
      currentType = null;
      sawHeaderForBlock = false;
      continue;
    }
    if (b === "구분") {
      sawHeaderForBlock = true;
      continue;
    }
    if (!sawHeaderForBlock || !currentBrand) continue;
    if (b === "소계" || b === "합계") continue; // 소계/합계 행은 건너뛴다 (수식이 있을 것으로 가정)

    if (CONFIG.GUBUN_SET.has(b)) currentGubun = b;
    if (c) currentType = c;

    if (!currentGubun || !currentType) continue;

    const detail = d || currentType; // "기타" 카테고리는 유형상세가 비어있고 유형명을 그대로 씀
    if (!d && !c) continue; // 유형/유형상세 둘 다 없는 행은 leaf가 아님

    leaves.push({ row: r, brand: currentBrand, gubun: currentGubun, type: currentType, detail });
  }
  return leaves;
}

/**
 * "26년 마감자료" 웹앱은 이 저장소의 다른 웹앱들과 다른 전용 구조입니다.
 * action 파라미터를 받지 않고, 요청할 때마다 매달 데이터를 한 번에 전부
 * { ok, months: { "1월": [...], "2월": [...], ..., "8월": [...] } } 형태로
 * 돌려줍니다. 각 월 배열은 0행이 안내문("※붉은글씨 : 등록된 후 삭제건"),
 * 1행이 헤더, 2행부터가 실제 데이터입니다(직접 확인함).
 *
 * 예전에는 누적되는 cost 시트 CSV 전체를 읽어서, 등록일(N주) 값만으로
 * 이번달 것인지 판단했는데, 그 cost 시트가 지난달 데이터를 지우지 않고
 * 계속 쌓이는 방식이라 지난달의 2주~5주까지 뒤섞여 잘못 집계되는
 * 문제가 있었습니다. "26년 마감자료"는 달마다 완전히 분리된 배열이라,
 * 이번달 배열(months[monthLabel])만 읽으면 그 안의 등록일(N주) 값은
 * 전부 이번달 것만 남습니다.
 *
 * 참고: 안내문에 따르면 "등록된 후 삭제된 건"은 빨간 글씨로 표시된다고
 * 하는데, 이 JSON 응답에는 셀 색상 정보가 없어 삭제건을 구분해 뺄 수
 * 없습니다. 삭제건이 실제로 존재한다면 집계 건수가 그만큼 더 많이
 * 잡힐 수 있습니다.
 **/
function fetchAndAggregateSource_(sourceWebappUrl, monthNum) {
  const monthLabel = `${monthNum}월`;

  const response = UrlFetchApp.fetch(sourceWebappUrl, { muteHttpExceptions: true });
  const data = JSON.parse(response.getContentText());
  if (data.error) throw new Error(`26년 마감자료: ${data.error}`);

  const monthRows = (data.months && data.months[monthLabel]) || null;
  if (!monthRows) {
    throw new Error(`26년 마감자료에서 "${monthLabel}" 데이터를 찾을 수 없습니다.`);
  }

  const header = monthRows[1] || [];
  const idx = {
    brand: header.indexOf("브랜드"),
    week: header.indexOf("등록일"), // 실제로는 "N주" 값이 들어있는 주차 컬럼
    type: header.indexOf("유형"),
    line: header.indexOf("포장"), // 사용자 확인: 이 컬럼 값으로 외작/내작 판정
    cls: header.indexOf("유형분류") // "유형분류"가 2개면 앞쪽(제조자 바로 다음)을 사용
  };
  ["brand", "week", "type", "line", "cls"].forEach((k) => {
    if (idx[k] === -1) throw new Error(`"${monthLabel}" 데이터에서 "${k}" 컬럼을 찾을 수 없습니다.`);
  });

  const agg = {}; // brand -> key -> week -> count
  const add = (brand, key, week) => {
    agg[brand] = agg[brand] || {};
    agg[brand][key] = agg[brand][key] || {};
    agg[brand][key][week] = (agg[brand][key][week] || 0) + 1;
  };

  // deriveGubun_이 "미분류"를 반환하면(등록된 leaf 행 목록에는 절대
  // 없는 구분값) 그 건은 어떤 leaf 행의 key와도 매칭되지 않아 집계에서
  // 조용히 사라집니다. 그런 유형이 실제로 있으면 숨기지 말고 바로
  // 알려주기 위해 모아서 에러로 던집니다.
  const unclassifiedTypes = new Set();

  for (let r = 2; r < monthRows.length; r++) {
    const row = monthRows[r];
    const brand = String(row[idx.brand] || "").trim();
    const week = String(row[idx.week] || "").trim();
    const type = String(row[idx.type] || "").trim();
    const cls = String(row[idx.cls] || "").trim();
    if (!type || !CONFIG.WEEK_LABELS.includes(week)) continue;

    const gubun = deriveGubun_(type, cls);
    if (gubun === "미분류") {
      unclassifiedTypes.add(`${type}${cls ? "/" + cls : ""}`);
      continue;
    }
    const detail = cls || type;
    const key = `${gubun}|${type}|${detail}`;
    const loc = deriveLoc_(row[idx.line]);

    add(brand, key, week);
    add("전체", key, week);
    add(loc, key, week); // "외작" 또는 "내작"
  }

  if (unclassifiedTypes.size) {
    throw new Error(
      `"${monthLabel}" 데이터에 deriveGubun_이 분류하지 못하는 유형이 있어 집계를 중단했습니다: ` +
      Array.from(unclassifiedTypes).join(", ") +
      " — deriveGubun_에 이 유형(들)을 추가해주세요."
    );
  }

  return agg;
}

function deriveGubun_(type, cls) {
  if (type === "가공" || type === "포장" || type === "설계") return "시공미결";
  if (type === "자재") return cls === "불량" ? "시공미결" : "고객불만"; // 기능/외관 -> 고객불만
  if (type === "외관") return "고객불만";
  if (type === "감성" || type === "취급") return "기타";
  return "미분류";
}

function deriveLoc_(lineValue) {
  const v = String(lineValue || "").trim();
  if (/외주|구매/.test(v)) return "외작";
  return "내작";
}

/** 스크립트 프로젝트의 시간대 설정과 무관하게 항상 한국 시간 기준 "이번달"을 구합니다. */
function currentMonthNum_() {
  return Number(Utilities.formatDate(new Date(), "Asia/Seoul", "M"));
}

function columnToLetter_(column) {
  let temp = "";
  let col = column;
  while (col > 0) {
    const rem = (col - 1) % 26;
    temp = String.fromCharCode(65 + rem) + temp;
    col = Math.floor((col - rem - 1) / 26);
  }
  return temp;
}
