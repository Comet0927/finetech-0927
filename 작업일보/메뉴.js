const ss = SpreadsheetApp.getActiveSpreadsheet();

/**
 * 설치 트리거: 애드온을 설치하거나
 * 배포된 웹 애플리케이션을 설치할 때 한 번만 호출됩니다.
 */
function onInstall(e) {
  // onOpen에게 이벤트 넘겨서
  // 설치 후 바로 메뉴가 보이도록 설정
  onOpen(e);
}

/**
 * 스프레드시트 열 때마다 호출되는 onOpen
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('1. 작업 취합')
    .addItem('1. 작업 협의 시트 생성', 'showFormDialog')
    .addItem('2. 당일 작업 확정',       'showMoveExpectedForm')
    .addItem('3. 시트 내보내기',       'showExportForm')   // ← 이 줄 추가
    .addToUi();
}




// ① 설정: 내보낼 스프레드시트 ID
const DEST_SPREADSHEET_ID = '1hgaAw8LLMXla1gM8B3TnMO_VDx1jyZ3NXllvRN5W23U';

function showExportForm() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const allNames = ss.getSheets().map(s => s.getName());

  // 안심 리스트: "(안심)" 포함 시트 우선, 그 외 시트 뒤
  const sheetNames1 = [
    ...allNames.filter(n => n.includes('(안심)')),
    ...allNames.filter(n => !n.includes('(안심)'))
  ];

  // 작업 리스트: "(작업)" 포함 시트 우선, 그 외 시트 뒤
  const sheetNames2 = [
    ...allNames.filter(n => n.includes('(작업)')),
    ...allNames.filter(n => !n.includes('(작업)'))
  ];

  // 기본 선택값은 각 우선순위 첫 번째
  const default1 = sheetNames1[0];
  const default2 = sheetNames2[0];

  // 템플릿에 전달
  const tpl = HtmlService.createTemplateFromFile('ExportForm');
  tpl.sheetNames1   = sheetNames1;
  tpl.sheetNames2   = sheetNames2;
  tpl.default1      = default1;
  tpl.default2      = default2;
  tpl.destId        = DEST_SPREADSHEET_ID;

  const htmlOutput = tpl
    .evaluate()
    .setWidth(400)
    .setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '시트 내보내기');
}

/**
 * 클라이언트에서 전달된 두 시트를 같은 destId로 복사하고
 * 복사된 뒤에는 원본 시트를 삭제합니다.
 */
function exportTwoSheetsToSpreadsheet(sheetName1, sheetName2, destId) {
  const srcSs  = SpreadsheetApp.getActiveSpreadsheet();
  const destSs = SpreadsheetApp.openById(destId);

  [sheetName1, sheetName2].forEach(name => {
    // 1) 원본 시트 가져오기
    const sheet = srcSs.getSheetByName(name);
    if (!sheet) throw new Error(`원본에 "${name}" 시트가 없습니다.`);

    // 2) 대상 스프레드시트에 복사
    const copy = sheet.copyTo(destSs);
    copy.setName(name);               // 원본 이름 그대로 쓰기
    destSs.setActiveSheet(copy);
    destSs.moveActiveSheet(1);

    // 3) **원본 시트 삭제**
    srcSs.deleteSheet(sheet);
  });
}