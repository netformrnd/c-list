// ================================================================
// Triggers.gs - Google Sheets 직접 수정 시 자동 처리
// ================================================================

/**
 * 시트가 직접 수정될 때 updatedAt 타임스탬프 자동 갱신
 * 설치형 트리거로 등록 필요: setupOnEditTrigger() 실행
 */
function onSheetEdit(e) {
  try {
    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getName();

    // 유효한 연도 탭에서만 동작
    if (!VALID_YEAR_TABS.includes(sheetName)) return;

    const range = e.range;
    const row = range.getRow();

    // 헤더 행 무시
    if (row <= 1) return;

    // updatedAt 컬럼(AG, 33번째) 자동 갱신
    const updatedAtCol = HEADERS.length; // 33
    // updatedAt 컬럼 자체를 수정한 경우는 무시 (무한루프 방지)
    if (range.getColumn() === updatedAtCol && range.getNumColumns() === 1) return;

    // 수정된 행의 updatedAt 갱신
    const now = Date.now();
    const lastRow = Math.min(row + range.getNumRows() - 1, sheet.getLastRow());
    for (let r = row; r <= lastRow; r++) {
      sheet.getRange(r, updatedAtCol).setValue(now);
    }

    // 메타데이터 업데이트
    updateMeta(sheetName);
  } catch (err) {
    Logger.log('onSheetEdit error: ' + err.message);
  }
}

/**
 * 설치형 onEdit 트리거 등록 (최초 1회 실행)
 * Apps Script 에디터에서 이 함수를 수동 실행하세요.
 */
function setupOnEditTrigger() {
  // 기존 트리거 제거
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'onSheetEdit') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 새 트리거 등록
  ScriptApp.newTrigger('onSheetEdit')
    .forSpreadsheet(SPREADSHEET_ID)
    .onEdit()
    .create();

  Logger.log('onEdit 트리거 등록 완료');
}
