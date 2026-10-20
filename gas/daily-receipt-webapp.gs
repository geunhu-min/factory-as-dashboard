/**************************************************************
 * "접수데이터" 스프레드시트 Web App
 *
 * 배포 위치
 * ------------------------------------------------------------
 * 이 파일은 "접수데이터" 스프레드시트의 확장 프로그램 > Apps Script
 * 프로젝트에 이미 배포되어 있던 파일입니다(사진/영상 링크를 S열/T열에
 * 자동 반영하는 PHOTO_LINK_SHEET_SYNC_URL 기능). 이번에 doGet만
 * 추가했고, 기존 doPost/사진 동기화 로직은 그대로 둡니다.
 *
 * 하는 일
 * ------------------------------------------------------------
 * [기존] doPost:
 * - action 없음/기본: 화면(사진 링크 첨부)에서 receiptNo+seq+code로
 *   행을 찾아 S열(링크, 콤마 구분)/T열(종류, 콤마 구분)에 링크를
 *   추가합니다. 이미 같은 드라이브 파일 링크가 있으면 중복 추가하지
 *   않습니다.
 * - action="delete": 위와 같은 방식으로 행을 찾아 S/T열에서 해당
 *   링크만 제거합니다.
 * - action="syncFolder": 드라이브 폴더 링크를 받아, 폴더명(yyMMdd)과
 *   A열 날짜가 같은 행들을 대상으로 고객명이 파일명에 포함된 파일을
 *   찾아 링크를 채웁니다.
 *
 * [신규] doPost action="appendReceiptRows" — 일일접수현황 대시보드의
 * "접수데이터 누적저장" 버튼 전용. body에 rows(회의자료 화면에 보이는
 * 18개 열짜리 행 배열 그대로, 헤더/맨 아래 "일일합계" 요약행 제외)를
 * 담아 보내면, 접수번호+순번+부품코드+색상+고객명 기준으로 이 시트에
 * 이미 있는 건은 M열(하자내역)이 달라졌을 때만 그 칸을 덮어쓰고(수정),
 * 같으면 그대로 둡니다. 아직 없는 건만 맨 아래에 새로 이어 붙입니다.
 * 회의자료의 열 순서(접수일자~일일합계)가 이 시트의 A~R열과 그대로
 * 같아서 매핑 없이 붙여넣기만 합니다(S/T열 사진 링크/종류는 기존 값
 * 그대로 두고 건드리지 않음).
 *
 * [신규] doPost action="uploadFile" — "추가" 버튼에서 링크 대신
 * 사진/동영상 파일을 직접 선택했을 때 전용. body에 receiptNo/seq/code/
 * kind(사진 또는 영상)/fileName/mimeType/base64를 담아 보내면, 드라이브
 * 폴더("일일접수현황_직접첨부", 없으면 자동 생성)에 파일로 저장하고
 * "링크 있는 사람 모두 보기" 권한을 준 뒤, 그 결과 링크를 action 없는
 * 기본 doPost와 똑같은 방식으로 S열/T열에 바로 추가합니다. 응답의
 * url을 화면에서 그대로 사용합니다.
 *
 * [신규] doGet — 대시보드의 "일일업무" 페이지(포장일확인)가 이
 * 스프레드시트의 데이터를 읽을 때 씁니다:
 * - action="sheets": 모든 탭 이름 + gid + 행/열 수 목록
 * - action="read" (기본값): 탭 하나를 header + rows로 반환합니다.
 *   sheet 파라미터(탭 이름) 또는 gid 파라미터로 지정, 둘 다 없으면
 *   첫 번째 탭(시트1)을 씁니다.
 * - action="packagingKeys": "포장일확인" 버튼 전용 — sheet, date(연-월-일)
 *   파라미터를 받아 A열 접수일자가 그 날짜와 같은 행만 이 스프레드시트
 *   안에서 미리 걸러서, F열 부품코드/G열 색상/H열 생산로트/I열 고객명/
 *   M열 하자내역만 뽑아 반환합니다(전체 행/전체 열을 다 내려받지 않아도
 *   되게 하기 위한 전용 엔드포인트 — action="read"보다 훨씬 가볍고
 *   빠릅니다).
 *
 * 배포 방법
 * ------------------------------------------------------------
 * 1. "접수데이터" 스프레드시트 > 확장 프로그램 > Apps Script에서 기존
 *    코드를 이 파일 내용으로 교체
 * 2. 배포 > 배포 관리 > 기존 배포 옆 연필(수정) > 버전: 새 버전 > 배포
 *    (URL을 그대로 유지하려면 "새 배포"가 아니라 기존 배포를 "수정"
 *    해야 합니다 — 이미 대시보드에 저장해둔 URL이 그대로 유지됩니다)
 * 3. 이미 저장해둔 "접수데이터 연결" URL은 그대로 두면 됩니다(URL이
 *    바뀌지 않았으므로 다시 입력할 필요 없음).
 *
 * 주의
 * ------------------------------------------------------------
 * - 토큰 검증이 없으므로 URL을 아는 사람은 누구나 이 시트를 읽고
 *   사진 링크를 추가/삭제/동기화할 수 있습니다.
 **************************************************************/

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

    if (action === "packagingKeys") {
      const sheet = resolveSheet_(params.sheet, params.gid);
      return jsonOutput_(packagingKeysForDateAction_(sheet, params.date));
    }

    if (action === "spreadsheetUrl") {
      return jsonOutput_({ url: SpreadsheetApp.getActiveSpreadsheet().getUrl() });
    }

    return jsonOutput_({ error: "알 수 없는 action입니다: " + action });
  } catch (error) {
    return jsonOutput_({ error: error.message });
  }
}


/**************************************************************
 * D열(접수번호)+E열(순번)+F열(부품코드)+G열(색상)+I열(고객명)을
 * 합친 중복 검사 키.
 **************************************************************/
function buildReceiptRowKey_(row) {
  return [
    normalizeText_(row[3]), // 접수번호
    normalizeText_(row[4]), // 순번
    normalizeText_(row[5]), // 부품코드
    normalizeText_(row[6]), // 색상
    normalizeText_(row[8])  // 고객명
  ].join("||");
}


/**************************************************************
 * "접수데이터 누적저장" 버튼 전용. 회의자료 화면의 18개 열짜리 행
 * (헤더/요약행 제외)을 받아, 접수번호+순번+부품코드+색상+고객명
 * 기준으로 이 시트에 이미 있는 건은:
 * - EDITABLE_FIELD_COLUMNS_(생산로트/원인처/포장처/유형/하자내역/
 *   미결구분 — 회의자료 화면에서 직접 입력 가능한 칸 전부) 중 내용이
 *   달라진 칸만 새 값으로 덮어씀(수정)
 * - 전부 같으면 그대로 둠(스킵)
 * 아직 없는 건(또는 같이 보낸 행들 중 처음 나온 건)만 이 시트 맨
 * 아래에 새로 이어 붙입니다.
 *
 * A열(접수일자)/E열(순번)은 문자열로 값을 써도 시트가 자동으로
 * Date/숫자로 바꿔버려 값이 틀어질 수 있어서 텍스트 서식("@")을
 * 먼저 고정해두고, P~R열(금액/패널티/일일합계)은 이 시트에 남아있을
 * 수 있는 날짜 서식이 숫자에 씌워지는 일이 없도록 일반 숫자 서식
 * ("#,##0")으로 먼저 고정해둡니다.
 **************************************************************/
// 회의자료 화면에서 직접 입력 가능한 칸(index는 0-based, column은
// setValue용 1-based 열 번호). 여기 없는 칸(구분/접수번호/순번/
// 부품코드/색상/고객명/수량/금액/패널티/일일합계)은 유형별자료 원본을
// 그대로 옮긴 값이라 회의자료에서 편집할 일이 없어서 비교 대상에 넣지
// 않습니다. 접수일자/번호는 회의자료에서 직접 고치는 경우가 있어서
// 포함시켰습니다.
const EDITABLE_FIELD_COLUMNS_ = [
  { index: 0, column: 1 },  // A열 접수일자
  { index: 1, column: 2 },  // B열 번호
  { index: 7, column: 8 },  // H열 생산로트
  { index: 9, column: 10 }, // J열 원인처
  { index: 10, column: 11 }, // K열 포장처
  { index: 11, column: 12 }, // L열 유형
  { index: 12, column: 13 }, // M열 하자내역
  { index: 13, column: 14 }  // N열 미결구분
];

function appendReceiptRowsAction_(sheet, rows) {
  if (!Array.isArray(rows) || !rows.length) {
    return { ok: true, addedCount: 0, updatedCount: 0, skippedCount: 0 };
  }

  const columnCount = 18; // 접수일자(A) ~ 일일합계(R), 회의자료와 열 순서 동일
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  // 1행은 헤더이므로 2행부터 기존 데이터를 읽어, 키별로 실제 행 번호와
  // 그 시점의 편집 가능 칸 값들을 미리 담아둠(나중에 값이 바뀌었는지 비교용)
  const existingByKey = new Map();
  if (lastRow >= 2 && lastColumn >= 1) {
    const existingValues = sheet
      .getRange(2, 1, lastRow - 1, Math.min(lastColumn, columnCount))
      .getValues();

    existingValues.forEach(function(row, idx) {
      const key = buildReceiptRowKey_(row);
      if (!existingByKey.has(key)) {
        const fields = {};
        EDITABLE_FIELD_COLUMNS_.forEach(function(field) {
          fields[field.index] = normalizeText_(row[field.index]);
        });
        existingByKey.set(key, { rowNumber: 2 + idx, fields: fields });
      }
    });
  }

  const newRows = [];
  const seenKeysThisBatch = new Set();
  let updatedCount = 0;
  let skippedCount = 0;

  rows.forEach(function(row) {
    const fixed = (row || []).slice(0, columnCount);
    while (fixed.length < columnCount) fixed.push("");

    const key = buildReceiptRowKey_(fixed);
    const existing = existingByKey.get(key);

    if (existing) {
      const changedFields = EDITABLE_FIELD_COLUMNS_.filter(function(field) {
        return existing.fields[field.index] !== normalizeText_(fixed[field.index]);
      });

      if (changedFields.length) {
        // 바뀐 칸마다 setValue()를 따로 부르면(최대 8번) API 호출이
        // 그만큼 늘어나므로, 현재 행 전체를 한 번 읽어서 바뀐 칸만
        // 값을 바꾼 뒤 한 번에 다시 씁니다(읽기 1번 + 쓰기 1번).
        const currentRow = sheet.getRange(existing.rowNumber, 1, 1, columnCount).getValues()[0];

        changedFields.forEach(function(field) {
          currentRow[field.index] = fixed[field.index];
          existing.fields[field.index] = normalizeText_(fixed[field.index]);
        });

        sheet.getRange(existing.rowNumber, 1, 1, columnCount).setValues([currentRow]);
        updatedCount++;
      } else {
        skippedCount++;
      }

      return;
    }

    if (seenKeysThisBatch.has(key)) {
      skippedCount++;
      return;
    }

    seenKeysThisBatch.add(key);
    newRows.push(fixed);
  });

  if (newRows.length) {
    const startRow = sheet.getLastRow() + 1;

    sheet.getRange(startRow, 1, newRows.length, 1).setNumberFormat("@");
    sheet.getRange(startRow, 5, newRows.length, 1).setNumberFormat("@");

    // F열(부품코드)/G열(색상)도 "061"처럼 앞자리 0이 있는 값이 숫자로
    // 바뀌지 않도록, 값을 쓰기 전에 텍스트 서식으로 고정합니다.
    sheet.getRange(startRow, 6, newRows.length, 2).setNumberFormat("@");

    sheet.getRange(startRow, 16, newRows.length, 3).setNumberFormat("#,##0");

    sheet.getRange(startRow, 1, newRows.length, columnCount).setValues(newRows);
  }

  return {
    ok: true,
    addedCount: newRows.length,
    updatedCount: updatedCount,
    skippedCount: skippedCount
  };
}


/**************************************************************
 * "포장일확인" 버튼 전용. A열 접수일자가 dateValue("yyyy-MM-dd")와
 * 같은 행만 이 시트 안에서 걸러서, F열 부품코드/G열 색상/H열
 * 생산로트/I열 고객명/M열 하자내역만 뽑아 반환합니다. 전체 시트를
 * 그대로 내려받는 action="read"보다 데이터량이 훨씬 적어 빠릅니다.
 **************************************************************/
function packagingKeysForDateAction_(sheet, dateValue) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 2 || lastColumn < 1) {
    return { keys: [] };
  }

  const values = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  const keys = [];

  values.forEach(function(row) {
    if (formatDateOnly_(row[0]) !== dateValue) return;

    keys.push({
      partCode: normalizeText_(row[5]),    // F열 부품코드
      color: normalizeText_(row[6]),       // G열 색상
      lot: normalizeText_(row[7]),         // H열 생산로트
      customerName: normalizeText_(row[8]), // I열 고객명
      defect: normalizeText_(row[12])      // M열 하자내역
    });
  });

  return { keys: keys };
}


/**************************************************************
 * 날짜 셀 값을 "yyyy-MM-dd" 문자열로 맞춥니다(Date 타입/텍스트 모두
 * 처리). dateToFolderFormat과 비슷하지만 형식이 달라 별도로 둡니다.
 **************************************************************/
function formatDateOnly_(value) {
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, "Asia/Seoul", "yyyy-MM-dd");
  }

  const text = String(value === null || value === undefined ? "" : value).trim();
  const match = /^(\d{4}-\d{2}-\d{2})/.exec(text);
  return match ? match[1] : text;
}


function normalizeText_(value) {
  return String(
    value === null || value === undefined ? "" : value
  ).trim();
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


function jsonOutput_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}


/**************************************************************
 * 아래부터는 기존에 배포되어 있던 코드입니다(그대로 유지).
 **************************************************************/

// S열(링크)/T열(종류)을 각각 콤마로 split만 하면, 중간에 빈 링크가 하나라도
// 끼어있을 때 links만 filter(Boolean)로 빈 칸을 빼면서 links/kinds 인덱스가
// 서로 어긋난다(사진이 엉뚱한 kind를 받아 재생 아이콘이 겹쳐 보이는 등
// "이상한 모양"으로 보이는 원인). 원래 위치 기준으로 짝을 맞춘 뒤 빈 링크만
// 제거해서 항상 같은 길이로 맞춰 돌려준다.
function parseLinksAndKinds_(linkStr, kindStr) {
  var rawLinks = String(linkStr || '').split(',').map(function (v) { return v.trim(); });
  var rawKinds = String(kindStr || '').split(',').map(function (v) { return v.trim(); });
  var links = [];
  var kinds = [];
  rawLinks.forEach(function (link, i) {
    if (!link) return;
    links.push(link);
    kinds.push(rawKinds[i] || '');
  });
  return { links: links, kinds: kinds };
}

function driveFileIdFromUrl(url) {
  var s = String(url || '');
  var m = s.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) || s.match(/[?&]id=([a-zA-Z0-9_-]+)/) || s.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : '';
}

function isSameDriveLink(a, b) {
  var idA = driveFileIdFromUrl(a);
  var idB = driveFileIdFromUrl(b);
  if (idA && idB) return idA === idB;
  return a === b;
}

function fileOrderRank(fileName, customerName) {
  var idx = fileName.indexOf(customerName);
  var rest = idx >= 0 ? fileName.slice(idx + customerName.length) : '';
  var m = rest.match(/^(\d+)/);
  return m ? parseInt(m[1], 10) : 0; // 번호 없는 파일(이름.png)이 0으로 가장 먼저
}

/**************************************************************
 * D열(접수번호)+E열(순번)+F열(부품코드)로 행을 찾습니다. 이 시트는
 * 행이 수천 개라서 매번 getDataRange().getValues()로 전체를 통째로
 * 읽어와 한 줄씩 비교하면(특히 사진/영상 업로드마다 매번) 느립니다.
 * TextFinder로 D열에서 접수번호가 일치하는 후보 행만 먼저 좁힌 뒤,
 * 그 몇 개 후보에 대해서만 E/F열을 읽어 순번+부품코드까지 비교합니다.
 * 못 찾으면 -1을 반환합니다.
 **************************************************************/
function findReceiptRow_(sheet, receiptNo, seq, code) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2 || !receiptNo) return -1;

  var finder = sheet.getRange(2, 4, lastRow - 1, 1)
    .createTextFinder(receiptNo)
    .matchEntireCell(true);
  var matches = finder.findAll();

  for (var i = 0; i < matches.length; i++) {
    var row = matches[i].getRow();
    var rowValues = sheet.getRange(row, 5, 1, 2).getValues()[0];
    var rowSeq = String(rowValues[0] || '').trim();
    var rowCode = String(rowValues[1] || '').trim();
    if (rowSeq === seq && rowCode === code) return row;
  }

  return -1;
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

    if (payload.action === 'syncFolder') {
      var result = syncPhotosFromFolderCore(sheet, payload.folderUrl);
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }

    if (payload.action === 'appendReceiptRows') {
      var appendResult = appendReceiptRowsAction_(sheet, payload.rows);
      return ContentService.createTextOutput(JSON.stringify(appendResult)).setMimeType(ContentService.MimeType.JSON);
    }

    if (payload.action === 'uploadFile') {
      var uploadResult = uploadFileAndLinkAction_(sheet, payload);
      return ContentService.createTextOutput(JSON.stringify(uploadResult)).setMimeType(ContentService.MimeType.JSON);
    }

    var receiptNo = String(payload.receiptNo || '').trim();
    var seq = String(payload.seq || '').trim();
    var code = String(payload.code || '').trim();
    var targetRow = findReceiptRow_(sheet, receiptNo, seq, code);

    if (targetRow < 0) {
      return ContentService.createTextOutput(JSON.stringify({ ok: true, updated: 0 }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // S/T열을 따로 getValue() 두 번 부르지 않고 인접한 두 열을 한 번에
    // 읽어서(API 호출 수를 줄여) 속도를 아낀다 — 아래 쓰기도 마찬가지.
    var stRange = sheet.getRange(targetRow, 19, 1, 2);
    var stValues = stRange.getValues()[0];
    var parsedPair = parseLinksAndKinds_(stValues[0], stValues[1]);
    var links = parsedPair.links;
    var kinds = parsedPair.kinds;

    if (payload.action === 'delete') {
      var link = String(payload.link || '').trim();
      var newLinksArr = [];
      var newKindsArr = [];
      for (var li = 0; li < links.length; li++) {
        if (isSameDriveLink(links[li], link)) continue;
        newLinksArr.push(links[li]);
        newKindsArr.push(kinds[li]);
      }
      links = newLinksArr;
      kinds = newKindsArr;
    } else {
      var newLink = String(payload.link || '').trim();
      if (newLink) {
        var alreadyExists = links.some(function (existing) { return isSameDriveLink(existing, newLink); });
        if (!alreadyExists) {
          links.push(newLink);
          kinds.push(String(payload.kind || '').trim());
        }
      }
    }

    stRange.setValues([[links.join(','), kinds.join(',')]]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true, updated: 1 }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**************************************************************
 * "일일접수현황"에서 사진/동영상 파일을 직접 선택해 올렸을 때 호출.
 * base64로 받은 파일을 드라이브에 저장하고(공유 설정까지) 그 결과
 * 링크를 기존 doPost(action 없음)과 같은 방식으로 S열/T열에 바로
 * 추가합니다. 반환값의 url을 화면에서 그대로 사용합니다.
 **************************************************************/
// 폴더 ID를 알면 여기 채워 넣으세요 — 그러면 매번 이름으로 폴더를 검색하지
// 않고 바로 열어서(getFolderById) 그만큼 빨라집니다. 비워두면(기본값)
// 기존처럼 이름으로 찾고, 처음 한 번 그 폴더의 ID를 응답(timings.folderId)과
// 실행 기록 로그에 알려주니 그걸 복사해서 여기 넣으면 됩니다. 혹시 이
// ID의 폴더가 없어지거나 접근이 안 되면 자동으로 이름 검색으로 되돌아갑니다.
var UPLOAD_FOLDER_ID_ = '1bYZRtQYhxaZD7DAOUwQ5XFtFLbVR0K6e';

// getOrCreateUploadFolder_가 ID로 바로 열었는지, 이름으로 다시 찾았는지를
// 화면(브라우저 콘솔)에서도 바로 볼 수 있게 기록해둔다 — Apps Script
// 실행 기록까지 안 들어가도, UPLOAD_FOLDER_ID_가 제대로 먹혔는지 바로 확인 가능.
var lastFolderLookupMethod_ = '';

function getOrCreateUploadFolder_() {
  if (UPLOAD_FOLDER_ID_) {
    try {
      var folderById = DriveApp.getFolderById(UPLOAD_FOLDER_ID_);
      lastFolderLookupMethod_ = 'id';
      return folderById;
    } catch (err) {
      lastFolderLookupMethod_ = 'id-failed-fallback:' + err;
      Logger.log('UPLOAD_FOLDER_ID_(' + UPLOAD_FOLDER_ID_ + ')로 폴더를 열지 못해 이름 검색으로 대체합니다: ' + err);
    }
  } else {
    lastFolderLookupMethod_ = 'no-id-set';
  }
  var name = '일일접수현황_직접첨부';
  var folders = DriveApp.getFoldersByName(name);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
  Logger.log('업로드 폴더 ID(고정하려면 UPLOAD_FOLDER_ID_에 붙여넣기): ' + folder.getId());
  return folder;
}

function uploadFileAndLinkAction_(sheet, payload) {
  var t0 = Date.now();
  var timings = {};
  var mark = function (label) { timings[label] = Date.now() - t0; };

  var base64 = String(payload.base64 || '');
  if (!base64) return { ok: false, error: '업로드할 파일 데이터가 없습니다.' };

  var fileName = String(payload.fileName || 'upload');
  var mimeType = String(payload.mimeType || 'application/octet-stream');
  var bytes;
  try {
    bytes = Utilities.base64Decode(base64);
  } catch (decodeErr) {
    var badMatch = /[^A-Za-z0-9+/=]/.exec(base64);
    var badInfo = badMatch
      ? ('위치 ' + badMatch.index + '에 잘못된 문자 "' + badMatch[0] + '"(코드 ' + base64.charCodeAt(badMatch.index) + ') 발견, 주변: "' + base64.slice(Math.max(0, badMatch.index - 15), badMatch.index + 15) + '"')
      : 'base64 문자 자체는 전부 정상(A-Za-z0-9+/=) - 길이/패딩 문제로 추정';
    return {
      ok: false,
      error: '디코딩 실패 (받은 길이: ' + base64.length + '자, 4의 배수 여부: ' + (base64.length % 4 === 0) +
        ', 끝부분: "' + base64.slice(-20) + '") - ' + badInfo + ' - ' + String(decodeErr)
    };
  }
  mark('decode');

  var blob = Utilities.newBlob(bytes, mimeType, fileName);
  var folder = getOrCreateUploadFolder_();
  // ID로 바로 연 경우는 이미 그 ID를 알고 있으니 folder.getId()를 또 부를
  // 필요가 없다(속성 접근 한 번도 드라이브 쪽에 비용이 들 수 있어 굳이
  // 안 해도 되는 호출은 뺀다). 이름 검색으로 대체된 경우에만(진단 목적으로
  // 실제 ID를 알려주기 위해) 호출한다.
  var folderId = (lastFolderLookupMethod_ === 'id') ? UPLOAD_FOLDER_ID_ : folder.getId();
  mark('getFolder');
  var file = folder.createFile(blob);
  mark('createFile');
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  mark('setSharing');
  var link = file.getUrl();

  var receiptNo = String(payload.receiptNo || '').trim();
  var seq = String(payload.seq || '').trim();
  var code = String(payload.code || '').trim();
  var targetRow = findReceiptRow_(sheet, receiptNo, seq, code);
  mark('findRow');

  if (targetRow < 0) {
    Logger.log('uploadFile 단계별 소요시간(ms): ' + JSON.stringify(timings));
    return { ok: true, url: link, updated: 0, timings: timings, folderId: folderId, folderLookupMethod: lastFolderLookupMethod_ };
  }

  var stRange = sheet.getRange(targetRow, 19, 1, 2);
  var stValues = stRange.getValues()[0];
  mark('readST');
  var parsedPair = parseLinksAndKinds_(stValues[0], stValues[1]);
  var links = parsedPair.links;
  var kinds = parsedPair.kinds;
  links.push(link);
  kinds.push(String(payload.kind || '').trim());
  stRange.setValues([[links.join(','), kinds.join(',')]]);
  mark('writeST');

  Logger.log('uploadFile 단계별 소요시간(ms, 누적): ' + JSON.stringify(timings));
  return { ok: true, url: link, updated: 1, timings: timings, folderId: folderId, folderLookupMethod: lastFolderLookupMethod_ };
}

function dateToFolderFormat(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, 'Asia/Seoul', 'yyMMdd');
  }
  var s = String(value || '').trim();
  var m = s.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (!m) return '';
  var yy = m[1].slice(-2);
  var mm = ('0' + m[2]).slice(-2);
  var dd = ('0' + m[3]).slice(-2);
  return yy + mm + dd;
}

function syncPhotosFromFolderCore(sheet, folderUrl) {
  var idMatch = String(folderUrl || '').match(/[-\w]{25,}/);
  if (!idMatch) return { ok: false, error: '폴더 링크에서 폴더 ID를 찾지 못했습니다.' };

  var folder;
  try {
    folder = DriveApp.getFolderById(idMatch[0]);
  } catch (err) {
    return { ok: false, error: '폴더를 열 수 없습니다: ' + err };
  }
  var folderName = folder.getName().trim();
  var data = sheet.getDataRange().getValues();

  var files = [];
  var it = folder.getFiles();
  while (it.hasNext()) files.push(it.next());

  var matchedCount = 0;
  var skippedRows = [];

  for (var i = 1; i < data.length; i++) {
    var customerName = String(data[i][8] || '').trim();
    if (!customerName) continue;
    if (dateToFolderFormat(data[i][0]) !== folderName) continue;

    var matchedFiles = files.filter(function (f) { return f.getName().indexOf(customerName) >= 0; });
    if (!matchedFiles.length) { skippedRows.push(i + 1); continue; }

    matchedFiles.sort(function (a, b) {
      return fileOrderRank(a.getName(), customerName) - fileOrderRank(b.getName(), customerName);
    });

    matchedFiles.forEach(function (mf) { mf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); });
    var newLinks = matchedFiles.map(function (mf) { return mf.getUrl(); });
    var newKinds = matchedFiles.map(function (mf) { return mf.getMimeType().indexOf('video') >= 0 ? '영상' : '사진'; });

    var sCell = sheet.getRange(i + 1, 19);
    var tCell = sheet.getRange(i + 1, 20);
    var existingLinks = String(sCell.getValue() || '').split(',').map(function (v) { return v.trim(); }).filter(Boolean);
    var existingKinds = String(tCell.getValue() || '').split(',').map(function (v) { return v.trim(); });
    var addedAny = false;
    newLinks.forEach(function (link, idx) {
      var alreadyExists = existingLinks.some(function (existing) { return isSameDriveLink(existing, link); });
      if (!alreadyExists) {
        existingLinks.push(link);
        existingKinds.push(newKinds[idx]);
        addedAny = true;
      }
    });
    if (addedAny) {
      sCell.setValue(existingLinks.join(','));
      tCell.setValue(existingKinds.join(','));
      matchedCount++;
    }
  }

  return { ok: true, folderName: folderName, matched: matchedCount, skippedRows: skippedRows };
}

function debugPhotoMatch() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var data = sheet.getDataRange().getValues();
  var rowNum = 7489;

  var rawDate = data[rowNum - 1][0];
  var formattedDate = dateToFolderFormat(rawDate);
  Logger.log('원본 날짜값: ' + JSON.stringify(rawDate) + ' / 타입: ' + Object.prototype.toString.call(rawDate));
  Logger.log('변환된 날짜(폴더명 형식): [' + formattedDate + ']');

  var customerName = String(data[rowNum - 1][8] || '').trim();
  Logger.log('고객명: [' + customerName + '] 길이: ' + customerName.length);

  var folderUrl = 'https://drive.google.com/drive/folders/1HmaJWpB_czrvgWbIeqaEmn1Q0pC5tc3t?usp=drive_link';
  var idMatch = folderUrl.match(/[-\w]{25,}/);
  var folder = DriveApp.getFolderById(idMatch[0]);
  Logger.log('실제 폴더 이름: [' + folder.getName().trim() + ']');
  Logger.log('날짜 일치 여부: ' + (formattedDate === folder.getName().trim()));

  var it = folder.getFiles();
  while (it.hasNext()) {
    var f = it.next();
    var name = f.getName();
    Logger.log('파일명: [' + name + '] 포함여부: ' + (name.indexOf(customerName) >= 0));
  }
}
