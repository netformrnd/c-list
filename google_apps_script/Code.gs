// ================================================================
// 공사현황리스트 ↔ Google Sheets 양방향 동기화
// Google Apps Script - Web App (doGet / doPost)
// ================================================================

// ★ 아래 SPREADSHEET_ID를 실제 Google Sheets ID로 교체하세요
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';

const VALID_YEAR_TABS = ['2020-2021', '2022', '2023', '2024', '2025', '2026'];

// 컬럼 헤더 정의
const HEADERS = [
  'NO', '지역', '공사현장', '공사명', '공사유형', '영업담당자', '현장소장',
  '계약금액_부가세별도', '계약금액_부가세포함', '계약일', '착공일', '준공일',
  '계약금_금액', '계약금_계산서', '계약금_입금일', '계약금_미입금', '계약금_입금완료',
  '중도금1_금액', '중도금1_계산서', '중도금1_입금일', '중도금1_미입금', '중도금1_입금완료',
  '중도금2_금액', '중도금2_계산서', '중도금2_입금일', '중도금2_미입금', '중도금2_입금완료',
  '잔금_금액', '잔금_계산서', '잔금_입금일', '잔금_미입금', '잔금_입금완료',
  'updatedAt'
];

// ===== GET 요청 처리 =====
function doGet(e) {
  try {
    const action = e.parameter.action || 'readAll';

    if (action === 'readAll') {
      return jsonResponse(readAllYears());
    } else if (action === 'readYear') {
      const year = e.parameter.year;
      if (!year) return jsonResponse({ error: 'year 파라미터 필요' });
      return jsonResponse({ success: true, year: year, data: readYear(year) });
    } else if (action === 'getMeta') {
      return jsonResponse({ success: true, meta: getMeta() });
    }

    return jsonResponse({ error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ===== POST 요청 처리 =====
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;

    if (action === 'writeYear') {
      writeYear(payload.year, payload.data);
      updateMeta(payload.year);
      return jsonResponse({ success: true, message: payload.year + ' 저장 완료' });
    } else if (action === 'writeAll') {
      Object.keys(payload.data).forEach(year => {
        if (VALID_YEAR_TABS.includes(year)) {
          writeYear(year, payload.data[year]);
          updateMeta(year);
        }
      });
      return jsonResponse({ success: true, message: '전체 저장 완료' });
    } else if (action === 'migrate') {
      const data = payload.data;
      Object.keys(data).forEach(year => {
        if (VALID_YEAR_TABS.includes(year) && Array.isArray(data[year])) {
          writeYear(year, data[year]);
          updateMeta(year);
        }
      });
      return jsonResponse({ success: true, message: '마이그레이션 완료', version: getMeta().version });
    }

    return jsonResponse({ error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ===== 읽기 함수 =====
function readAllYears() {
  const result = {};
  VALID_YEAR_TABS.forEach(year => {
    result[year] = readYear(year);
  });
  return { success: true, data: result, meta: getMeta() };
}

function readYear(yearKey) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(yearKey);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return []; // 헤더만 있음

  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  return data.map(row => rowToObject(row));
}

// ===== 쓰기 함수 =====
function writeYear(yearKey, dataArray) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(yearKey);

  if (!sheet) {
    sheet = ss.insertSheet(yearKey);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    // 헤더 서식
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setBackground('#2c3e50');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(9);
  }

  // 기존 데이터 삭제 (헤더 제외)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, HEADERS.length).clearContent();
  }

  // 새 데이터 쓰기
  if (!dataArray || dataArray.length === 0) return;

  const rows = dataArray.map(obj => objectToRow(obj));
  sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);

  // 숫자 컬럼 서식 (금액)
  const numCols = [8, 9, 13, 18, 23, 28]; // H, I, M, R, W, AB (1-indexed)
  numCols.forEach(col => {
    if (rows.length > 0) {
      sheet.getRange(2, col, rows.length, 1).setNumberFormat('#,##0');
    }
  });

  // 상태 컬럼 가운데 정렬 (미입금/입금완료 O 표시)
  const statusCols = [16, 17, 21, 22, 26, 27, 31, 32]; // 각 단계의 미입금/입금완료 컬럼
  statusCols.forEach(col => {
    if (rows.length > 0) {
      sheet.getRange(2, col, rows.length, 1).setHorizontalAlignment('center');
    }
  });
}

// ===== 메타데이터 =====
function getMeta() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let metaSheet = ss.getSheetByName('_meta');
  if (!metaSheet) {
    metaSheet = ss.insertSheet('_meta');
    metaSheet.getRange(1, 1).setValue('version');
    metaSheet.getRange(1, 2).setValue(0);
    metaSheet.getRange(2, 1).setValue('lastSync');
    metaSheet.getRange(2, 2).setValue(Date.now());
    // 연도별 lastModified
    VALID_YEAR_TABS.forEach((year, i) => {
      metaSheet.getRange(3 + i, 1).setValue(year);
      metaSheet.getRange(3 + i, 2).setValue(0);
    });
  }

  const version = metaSheet.getRange(1, 2).getValue() || 0;
  const lastSync = metaSheet.getRange(2, 2).getValue() || 0;
  const yearMeta = {};
  VALID_YEAR_TABS.forEach((year, i) => {
    yearMeta[year] = metaSheet.getRange(3 + i, 2).getValue() || 0;
  });

  return { version: version, lastSync: lastSync, years: yearMeta };
}

function updateMeta(yearKey) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let metaSheet = ss.getSheetByName('_meta');
  if (!metaSheet) {
    getMeta(); // 초기화
    metaSheet = ss.getSheetByName('_meta');
  }

  // version 증가
  const currentVersion = metaSheet.getRange(1, 2).getValue() || 0;
  metaSheet.getRange(1, 2).setValue(currentVersion + 1);
  metaSheet.getRange(2, 2).setValue(Date.now());

  // 연도별 lastModified 업데이트
  const idx = VALID_YEAR_TABS.indexOf(yearKey);
  if (idx >= 0) {
    metaSheet.getRange(3 + idx, 2).setValue(Date.now());
  }
}

// ===== 변환 함수 =====
function rowToObject(row) {
  const obj = {
    no: row[0] || '',
    region: row[1] || '',
    site: row[2] || '',
    name: row[3] || '',
    type: row[4] || '',
    manager: row[5] || '',
    siteManager: row[6] || '',
    amount: toNum(row[7]),
    amountTax: toNum(row[8]),
    contractDate: toStr(row[9]),
    startDate: toStr(row[10]),
    endDate: toStr(row[11]),
    contract1: { amt: toNum(row[12]), bill: toStr(row[13]), paid: toStr(row[14]), isRed: toBool(row[15]), manualPaid: toBool(row[16]) },
    interim1:  { amt: toNum(row[17]), bill: toStr(row[18]), paid: toStr(row[19]), isRed: toBool(row[20]), manualPaid: toBool(row[21]) },
    interim2:  { amt: toNum(row[22]), bill: toStr(row[23]), paid: toStr(row[24]), isRed: toBool(row[25]), manualPaid: toBool(row[26]) },
    balance1:  { amt: toNum(row[27]), bill: toStr(row[28]), paid: toStr(row[29]), isRed: toBool(row[30]), manualPaid: toBool(row[31]) },
    updatedAt: toNum(row[32]) || Date.now()
  };
  return obj;
}

function objectToRow(obj) {
  const c1 = obj.contract1 || {};
  const i1 = obj.interim1 || {};
  const i2 = obj.interim2 || {};
  const b1 = obj.balance1 || {};

  return [
    obj.no || '', obj.region || '', obj.site || '', obj.name || '',
    obj.type || '', obj.manager || '', obj.siteManager || '',
    toNum(obj.amount), toNum(obj.amountTax),
    obj.contractDate || '', obj.startDate || '', obj.endDate || '',
    toNum(c1.amt), c1.bill || '', c1.paid || '', toLabel(c1.isRed), toLabel(c1.manualPaid),
    toNum(i1.amt), i1.bill || '', i1.paid || '', toLabel(i1.isRed), toLabel(i1.manualPaid),
    toNum(i2.amt), i2.bill || '', i2.paid || '', toLabel(i2.isRed), toLabel(i2.manualPaid),
    toNum(b1.amt), b1.bill || '', b1.paid || '', toLabel(b1.isRed), toLabel(b1.manualPaid),
    obj.updatedAt || Date.now()
  ];
}

// ===== 유틸리티 =====
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function toNum(v) {
  if (v === '' || v === null || v === undefined) return 0;
  const n = Number(String(v).replace(/,/g, ''));
  return isNaN(n) ? 0 : n;
}

function toStr(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) {
    const y = v.getFullYear();
    const m = String(v.getMonth() + 1).padStart(2, '0');
    const d = String(v.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + d;
  }
  return String(v);
}

function toBool(v) {
  // 기존 TRUE/FALSE 및 새 O/빈칸 형식 모두 지원
  if (v === true || v === 'TRUE' || v === 'true' || v === 1 || v === 'O') return true;
  return false;
}

// boolean → 시트 표시용 라벨 (TRUE/FALSE 대신 O/빈칸)
function toLabel(v) {
  return v ? 'O' : '';
}
