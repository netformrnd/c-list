// ================================================================
// 공사현황리스트 + 기술컨설팅 ↔ Google Sheets 양방향 동기화
// Google Apps Script - Web App (doGet / doPost)
// ================================================================

// ★ 아래 SPREADSHEET_ID를 실제 Google Sheets ID로 교체하세요
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';

const VALID_YEAR_TABS = ['2020-2021', '2022', '2023', '2024', '2025', '2026'];

// 공사현황 컬럼 헤더
const HEADERS = [
  'NO', '지역', '공사현장', '공사명', '공사유형', '영업담당자', '현장소장',
  '계약금액_부가세별도', '계약금액_부가세포함', '계약일', '착공일', '준공일',
  '계약금_금액', '계약금_계산서', '계약금_입금일', '계약금_미입금', '계약금_입금완료',
  '중도금1_금액', '중도금1_계산서', '중도금1_입금일', '중도금1_미입금', '중도금1_입금완료',
  '중도금2_금액', '중도금2_계산서', '중도금2_입금일', '중도금2_미입금', '중도금2_입금완료',
  '잔금_금액', '잔금_계산서', '잔금_입금일', '잔금_미입금', '잔금_입금완료',
  'updatedAt'
];

// 기술컨설팅 컬럼 헤더
const CONSULTING_HEADERS = [
  'ID', '연도', '유형', '현장명', '프로젝트명', '낙찰금액', '실행금액',
  '담당자', '업체명', '업체담당자', '하자여부', '현장담당자', '사진',
  '계약일', '착공일', '준공일', '비고', 'updatedAt'
];

const CONSULTING_SHEET_NAME = '기술컨설팅';

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
    } else if (action === 'readConsulting') {
      return jsonResponse({ success: true, data: readConsulting() });
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
    } else if (action === 'writeConsulting') {
      writeConsulting(payload.data);
      return jsonResponse({ success: true, message: '기술컨설팅 저장 완료' });
    } else if (action === 'migrate') {
      const data = payload.data;
      Object.keys(data).forEach(year => {
        if (VALID_YEAR_TABS.includes(year) && Array.isArray(data[year])) {
          writeYear(year, data[year]);
          updateMeta(year);
        }
      });
      // 기술컨설팅 마이그레이션
      if (payload.consultingData && Array.isArray(payload.consultingData)) {
        writeConsulting(payload.consultingData);
      }
      return jsonResponse({ success: true, message: '마이그레이션 완료', version: getMeta().version });
    }

    return jsonResponse({ error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ===== 공사현황 읽기/쓰기 =====
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
  if (lastRow <= 1) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  return data.map(row => rowToObject(row));
}

function writeYear(yearKey, dataArray) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(yearKey);
  if (!sheet) {
    sheet = ss.insertSheet(yearKey);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setBackground('#2c3e50');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(9);
  }
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, HEADERS.length).clearContent();
  if (!dataArray || dataArray.length === 0) return;
  const rows = dataArray.map(obj => objectToRow(obj));
  sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
  const numCols = [8, 9, 13, 18, 23, 28];
  numCols.forEach(col => {
    if (rows.length > 0) sheet.getRange(2, col, rows.length, 1).setNumberFormat('#,##0');
  });
  const statusCols = [16, 17, 21, 22, 26, 27, 31, 32];
  statusCols.forEach(col => {
    if (rows.length > 0) sheet.getRange(2, col, rows.length, 1).setHorizontalAlignment('center');
  });
}

// ===== 기술컨설팅 읽기/쓰기 =====
function readConsulting() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONSULTING_SHEET_NAME);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, CONSULTING_HEADERS.length).getValues();
  return data.map(row => consultingRowToObject(row));
}

function writeConsulting(dataArray) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONSULTING_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONSULTING_SHEET_NAME);
    sheet.getRange(1, 1, 1, CONSULTING_HEADERS.length).setValues([CONSULTING_HEADERS]);
    sheet.setFrozenRows(1);
    const headerRange = sheet.getRange(1, 1, 1, CONSULTING_HEADERS.length);
    headerRange.setBackground('#00695c');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(9);
  }
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, CONSULTING_HEADERS.length).clearContent();
  if (!dataArray || dataArray.length === 0) return;
  const rows = dataArray.map(obj => consultingObjectToRow(obj));
  sheet.getRange(2, 1, rows.length, CONSULTING_HEADERS.length).setValues(rows);
  // 금액 컬럼 서식 (낙찰금액, 실행금액)
  [6, 7].forEach(col => {
    sheet.getRange(2, col, rows.length, 1).setNumberFormat('#,##0');
  });
}

function consultingRowToObject(row) {
  return {
    id:             toStr(row[0]),
    year:           toStr(row[1]),
    type:           toStr(row[2]),
    site:           toStr(row[3]),
    project:        toStr(row[4]),
    bidAmount:      toNum(row[5]),
    execAmount:     toNum(row[6]),
    manager:        toStr(row[7]),
    company:        toStr(row[8]),
    companyManager: toStr(row[9]),
    warranty:       toStr(row[10]),
    fieldManager:   toStr(row[11]),
    photo:          toStr(row[12]),
    contractDate:   toStr(row[13]),
    startDate:      toStr(row[14]),
    endDate:        toStr(row[15]),
    note:           toStr(row[16]),
    updatedAt:      toNum(row[17]) || Date.now()
  };
}

function consultingObjectToRow(obj) {
  return [
    obj.id || '', obj.year || '', obj.type || '',
    obj.site || '', obj.project || '',
    toNum(obj.bidAmount), toNum(obj.execAmount),
    obj.manager || '', obj.company || '', obj.companyManager || '',
    obj.warranty || '', obj.fieldManager || '', obj.photo || '',
    obj.contractDate || '', obj.startDate || '', obj.endDate || '',
    obj.note || '', obj.updatedAt || Date.now()
  ];
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
  if (!metaSheet) { getMeta(); metaSheet = ss.getSheetByName('_meta'); }
  const currentVersion = metaSheet.getRange(1, 2).getValue() || 0;
  metaSheet.getRange(1, 2).setValue(currentVersion + 1);
  metaSheet.getRange(2, 2).setValue(Date.now());
  const idx = VALID_YEAR_TABS.indexOf(yearKey);
  if (idx >= 0) metaSheet.getRange(3 + idx, 2).setValue(Date.now());
}

// ===== 변환 함수 =====
function rowToObject(row) {
  return {
    no: row[0] || '', region: row[1] || '', site: row[2] || '', name: row[3] || '',
    type: row[4] || '', manager: row[5] || '', siteManager: row[6] || '',
    amount: toNum(row[7]), amountTax: toNum(row[8]),
    contractDate: toStr(row[9]), startDate: toStr(row[10]), endDate: toStr(row[11]),
    contract1: { amt: toNum(row[12]), bill: toStr(row[13]), paid: toStr(row[14]), isRed: toBool(row[15]), manualPaid: toBool(row[16]) },
    interim1:  { amt: toNum(row[17]), bill: toStr(row[18]), paid: toStr(row[19]), isRed: toBool(row[20]), manualPaid: toBool(row[21]) },
    interim2:  { amt: toNum(row[22]), bill: toStr(row[23]), paid: toStr(row[24]), isRed: toBool(row[25]), manualPaid: toBool(row[26]) },
    balance1:  { amt: toNum(row[27]), bill: toStr(row[28]), paid: toStr(row[29]), isRed: toBool(row[30]), manualPaid: toBool(row[31]) },
    updatedAt: toNum(row[32]) || Date.now()
  };
}

function objectToRow(obj) {
  const c1 = obj.contract1 || {}, i1 = obj.interim1 || {}, i2 = obj.interim2 || {}, b1 = obj.balance1 || {};
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
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function toNum(v) {
  if (v === '' || v === null || v === undefined) return 0;
  const n = Number(String(v).replace(/,/g, ''));
  return isNaN(n) ? 0 : n;
}
function toStr(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) {
    const y = v.getFullYear(), m = String(v.getMonth()+1).padStart(2,'0'), d = String(v.getDate()).padStart(2,'0');
    return y + '-' + m + '-' + d;
  }
  return String(v);
}
function toBool(v) {
  return v === true || v === 'TRUE' || v === 'true' || v === 1 || v === 'O';
}
function toLabel(v) { return v ? 'O' : ''; }
