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


  // 2. 작업 취합 메뉴
  ui.createMenu('1. 작업 취합')
    .addItem('1. 작업 협의 시트 생성', 'showFormDialog')
    .addItem('2. 당일 작업 확정', 'moveExpectedToToday')
    .addToUi();
}

// ✅ 2. 메뉴에서 호출되는 다이얼로그 열기
function showCombinedForm() {
  const html = HtmlService.createHtmlOutputFromFile('CombinedForm')
    .setWidth(450)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, "출력 확정 및 시트 복사");
}


